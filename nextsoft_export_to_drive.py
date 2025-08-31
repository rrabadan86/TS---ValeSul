# nextsoft_export_to_drive.py
# Fluxo:
# 1) Login (resiliente: retries, timeouts maiores, user-agent real)
# 2) Troca de loja (via secret APPNEXT_LOJA_DESTINO)
# 3) Vendas > Vendedor Analítico
# 4) Abre filtros (estabiliza grid)
# 5) Tenta exportar; se URL direta não for capturada (blob/POST), FALLBACK:
#    refaz a chamada de dados com dataInicial/dataFinal desejadas e gera o Excel localmente
# 6) Envia o arquivo ao Google Drive via rclone

import os
import sys
import subprocess
import traceback
import shutil
import json
import unicodedata
import re
from pathlib import Path
from datetime import datetime
from urllib.parse import urlparse, parse_qs, urlencode

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ================= LOG =================
def log(msg: str):
    now = datetime.now().strftime("%H:%M:%S")
    print(f"[{now}] {msg}", flush=True)


def step(title: str):
    log("=" * 70)
    log(title)
    log("=" * 70)


load_dotenv()

# ================ CONFIG ================
REDE = os.getenv("APPNEXT_REDE", "").strip()
USER = os.getenv("APPNEXT_USER", "").strip()
PASS = os.getenv("APPNEXT_PASS", "").strip()

DRIVE_REMOTE = os.getenv("DRIVE_REMOTE", "GDRIVE:")
DRIVE_FOLDER_ID = os.getenv("DRIVE_FOLDER_ID", "").strip()
DRIVE_FILE_NAME = os.getenv("DRIVE_FILE_NAME", "importacaoA.xlsx").strip()

DATA_INICIO = os.getenv("DATA_INICIO", "01/06/2025").strip()
DATA_FIM = os.getenv("DATA_FIM", "").strip()  # vazio => hoje 23:59:59

TEMPLATE_XLSX = os.getenv("TEMPLATE_XLSX", "importacaoA.xlsx")
TEMPLATE_HEADER_ROW = int(os.getenv("TEMPLATE_HEADER_ROW", "2"))

# Forçar lojaId diretamente no fallback (se quiser)
LOJA_ID_DESTINO = os.getenv("APPNEXT_LOJA_ID_DESTINO", "").strip()

RCLONE_PATH = os.getenv("RCLONE_PATH") or shutil.which("rclone") or r"C:\rclone\rclone.exe"
if not Path(RCLONE_PATH).exists() and shutil.which("rclone") is None:
    log("ERRO: rclone não encontrado. Ajuste RCLONE_PATH ou adicione ao PATH.")
    sys.exit(1)

if not (REDE and USER and PASS):
    log("ERRO: Preencha APPNEXT_REDE / APPNEXT_USER / APPNEXT_PASS")
    sys.exit(1)

TS = datetime.now().strftime("%Y%m%d_%H%M%S")
LOCAL_OUT = Path.cwd() / f"export_nextsoft_{TS}.xlsx"
LOGIN_URL = "https://www.appnext.com.br/#/login"

# -------- datas ----------
def to_dt(s: str, end=False):
    try:
        d = datetime.strptime(s, "%d/%m/%Y")
        return d.replace(hour=23, minute=59, second=59) if end else d
    except Exception:
        return None


if not DATA_FIM:
    DATA_FIM = datetime.now().strftime("%d/%m/%Y")

INI_DT = to_dt(DATA_INICIO) or datetime(2025, 6, 1)
FIM_DT = to_dt(DATA_FIM, end=True) or datetime.now().replace(hour=23, minute=59, second=59)

# ISO com Z (mesmo padrão visto no DevTools)
FMT_JSON_INI = INI_DT.strftime("%Y-%m-%dT%H:%M:%S.000Z")
FMT_JSON_FIM = FIM_DT.strftime("%Y-%m-%dT%H:%M:%S.000Z")
FMT_BR_INI = INI_DT.strftime("%d/%m/%Y, %H:%M:%S")
FMT_BR_FIM = FIM_DT.strftime("%d/%m/%Y, %H:%M:%S")

# -------- seletores ----------
SEL = {
    # login
    "rede": "input[placeholder='Rede']",
    "email": "input[placeholder='Email']",
    "senha": "input[placeholder='Senha']",
    "entrar": "button:has-text('Entrar'), button:has-text('Login')",
    # menu
    "menu_vendas": "nav >> text=Vendas, header >> text=Vendas, a[href*='vendas']:has-text('Vendas'), button:has-text('Vendas')",
    "item_vendedor_analitico": "a[href*='vendedor-analitico'], a:has-text('Vendedor Anal'), [role='menuitem']:has-text('Vendedor Anal'), li:has-text('Vendedor Anal')",
    # tela alvo
    "titulo_rel": "text=Listagem de Vendedor Analítico",
    # filtros / grid / export
    "btn_filtros": "[title*='Filtro'], [aria-label*='Filtro'], button:has(i.mdi-filter), button:has(i.fa-filter), button:has-text('Filtros')",
    "pane_filtros": "#filtrosForm",
    "btn_atualizar": "button:has-text('Atualizar Filtros')",
    "grid_row": "table tbody tr",
    "excel_btn": "#dataTableButtons button.buttons-excel, button.buttons-excel",
}

# =============== HELPERS ===============
def rclone_copy_latest(local_file: Path):
    cmd = [
        RCLONE_PATH,
        "copyto",
        str(local_file),
        f"{DRIVE_REMOTE}{DRIVE_FILE_NAME}",
        "--drive-root-folder-id",
        DRIVE_FOLDER_ID,
        "-v",
    ]
    log(f"rclone -> {' '.join(cmd)}")
    res = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if res.stdout.strip():
        print(res.stdout)
    if res.returncode != 0:
        print(res.stderr, file=sys.stderr)
        raise RuntimeError("Falha ao atualizar o arquivo no Drive.")


def _click_first(page, selectors):
    for sel in selectors:
        try:
            loc = page.locator(sel).first
            if loc.count():
                loc.scroll_into_view_if_needed(timeout=1500)
                loc.click(timeout=4000, force=True)
                return True
        except Exception:
            pass
    return False


def goto_vendedor_analitico_via_menu(page):
    # abre o menu Vendas
    for sel in [s.strip() for s in SEL["menu_vendas"].split(",")]:
        try:
            loc = page.locator(sel).first
            if loc.count():
                try:
                    loc.hover()
                except Exception:
                    pass
                loc.click(timeout=4000, force=True)
                break
        except Exception:
            continue
    # clica no item Vendedor Analítico
    for sel in [s.strip() for s in SEL["item_vendedor_analitico"].split(",")]:
        try:
            it = page.locator(sel).first
            if it.count():
                it.click(timeout=6000, force=True)
                break
        except Exception:
            continue
    # confirma rota/título
    try:
        page.wait_for_url("**/loja/vendas/**vendedor**", timeout=30000)
    except PWTimeout:
        page.wait_for_selector(SEL["titulo_rel"], timeout=20000)


def open_filters_pane(page):
    # se já estiver aberto, retorna
    try:
        page.wait_for_selector(f"{SEL['pane_filtros']}.show", timeout=1500)
        return
    except Exception:
        pass
    # clica no botão/ícone de filtros
    for sel in [s.strip() for s in SEL["btn_filtros"].split(",")]:
        try:
            loc = page.locator(sel).first
            if loc.count():
                loc.scroll_into_view_if_needed(timeout=1500)
                loc.click(timeout=2000, force=True)
                break
        except Exception:
            continue
    # espera abrir
    try:
        page.wait_for_selector(f"{SEL['pane_filtros']}.show", timeout=4000)
    except Exception:
        # força abrir (caso offcanvas)
        page.evaluate(
            """idSel => { const p=document.querySelector(idSel); if(p){p.classList.add('show'); p.style.display='block';} }""",
            SEL["pane_filtros"],
        )

# --------- CAPTURA DO FILTRO (debug/auxiliar) ----------
class Captured:
    filtro_url = None
    filtro_method = None
    filtro_ct = None
    filtro_body = None
    filtro_headers = None


def _clean_headers(h: dict) -> dict:
    bad = {"content-length", "host", "origin", "referer"}
    return {k: v for k, v in (h or {}).items() if k.lower() not in bad}


def hook_filter_capture(page):
    def on_request(req):
        url = (req.url or "").lower()
        if Captured.filtro_url is None and req.resource_type in ("xhr", "fetch"):
            # request principal do relatório (GET com query)
            is_report_data = ("vendedor" in url and "analit" in url) or "relatoriovendedoranalitico" in url
            if is_report_data:
                hdrs = _clean_headers(req.headers or {})
                body = ""
                try:
                    body = req.post_data or ""
                except Exception:
                    pass
                Captured.filtro_url = req.url
                Captured.filtro_method = (req.method or "GET").upper()
                Captured.filtro_body = body
                Captured.filtro_ct = hdrs.get("content-type", "")
                Captured.filtro_headers = hdrs
                log(f"[capture] filtro de DADOS {Captured.filtro_method} -> {req.url}")

    page.on("request", on_request)

# --------- CAPTURA DO EXPORT (se existir URL direta) ----------
class CapturedExport:
    url = None
    method = None
    headers = None


def hook_export_capture(page):
    def on_request(req):
        url = (req.url or "").lower()
        if req.resource_type in ("xhr", "fetch", "document"):
            if ("excel" in url or "export" in url or "xlsx" in url) and ("vendedor" in url or "analit" in url):
                CapturedExport.url = req.url
                CapturedExport.method = (req.method or "GET").upper()
                CapturedExport.headers = _clean_headers(req.headers or {})
                log(f"[capture] EXPORT {CapturedExport.method} -> {req.url}")

    page.on("request", on_request)

# --------- REPLAY DO EXPORT (se capturado) ---------
def replay_export_with_dates(page, destino_path: Path):
    if not CapturedExport.url:
        raise RuntimeError("URL de exportação não foi capturada. Clique 1x no botão Excel para eu aprender a URL.")
    parsed = urlparse(CapturedExport.url)
    qs = parse_qs(parsed.query)
    qs["dataInicial"] = [FMT_JSON_INI]
    qs["dataFinal"] = [FMT_JSON_FIM]
    target_url = f"{parsed.scheme}://{parsed.netloc}{parsed.path}?{urlencode(qs, doseq=True)}"
    log(f"[export-replay] GET -> {target_url}")
    with page.expect_download(timeout=180000) as dl:
        page.evaluate(
            """
            (url) => { const a=document.createElement('a'); a.href=url; a.target='_blank';
                       document.body.appendChild(a); a.click(); setTimeout(()=>a.remove(),500); }
            """,
            target_url,
        )
    download = dl.value
    download.save_as(destino_path.as_posix())
    log(f"Arquivo exportado (período aplicado): {destino_path.name}")

# --------- FALLBACK: baixar JSON e gerar Excel ----------
def fetch_report_json_with_dates(page):
    if not Captured.filtro_url:
        raise RuntimeError("Endpoint de dados não capturado ainda.")
    parsed = urlparse(Captured.filtro_url)
    qs = parse_qs(parsed.query)
    qs["dataInicial"] = [FMT_JSON_INI]
    qs["dataFinal"] = [FMT_JSON_FIM]
    if LOJA_ID_DESTINO:
        qs["lojaId"] = [LOJA_ID_DESTINO]
    target_url = f"{parsed.scheme}://{parsed.netloc}{parsed.path}?{urlencode(qs, doseq=True)}"
    headers = {**_clean_headers(Captured.filtro_headers or {})}
    log(f"[dados-replay] GET -> {target_url}")
    js_code = """
    async ([url, headers]) => {
        const res = await fetch(url, { method: "GET", headers });
        const text = await res.text();
        return { ok: res.ok, status: res.status, text };
    }
    """
    result = page.evaluate(js_code, [target_url, headers])
    if not result.get("ok"):
        raise RuntimeError(f"Falha ao obter dados: HTTP {result.get('status')}")
    try:
        data = json.loads(result["text"])
    except Exception as e:
        raise RuntimeError(f"Não consegui decodificar o JSON da API: {e}") from e
    return data

# ----------------- Normalização/Excel -----------------
MAPEAMENTO_COLUNAS = {
    "dataVenda": "Data",
    "numeroCupom": "Número Cupom",
    "qtdItens": "Qtd. Itens no Cupom",
    "quantidade": "Qtd. Itens no Cupom",
    "produto.nome": "Descrição",
    "produto": "Descrição",
    "item.descricao": "Descrição",
    "subGrupo.nome": "Sub-Grupo",
    "subgrupo.nome": "Sub-Grupo",
    "subGrupo": "Sub-Grupo",
    "subgrupo": "Sub-Grupo",
    "valorUnitario": "Valor Unitário",
    "precoUnitario": "Valor Unitário",
    "valorTotal": "Valor Total",
    "totalItem": "Valor Total",
    "valorLiquido": "Valor Total",
}

TEMPLATE_ALIASES = {
    "data": ["dataVenda", "data", "dataHora", "createdAt", "dt_venda"],
    "numero cupom": ["numeroCupom", "cupom.numero", "numero", "documento", "numeroDocumento"],
    "qtd itens no cupom": ["qtdItens", "quantidade", "qtde", "qtd", "itens", "itensQuantidade"],
    "descricao": ["descricao", "produto", "produto.descricao", "produtoNome", "item.descricao"],
    "sub-grupo": ["subGrupo", "subgrupo", "subGrupo.nome", "subgrupo.nome"],
    "valor unitario": ["valorUnitario", "precoUnitario", "valor.unitario", "preco"],
    "valor total": ["valorTotal", "valor.total", "total", "totalItem", "valorLiquido"],
    "vendedor": ["vendedor.nome", "vendedor", "vendedorNome", "colaborador", "usuario"],
    "grupo": ["grupo.nome", "grupo", "nomeGrupo"],
}

DATE_HEADER_HINTS = {"data", "data venda", "data de venda", "emissao", "data emissao", "data emissão"}


def _norm(s: str) -> str:
    s = (s or "").strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return " ".join("".join(ch if ch.isalnum() else " " for ch in s).split())


def _is_iso_datetime_str(x: str) -> bool:
    if not isinstance(x, str):
        return False
    return len(x) >= 19 and x[4] == "-" and x[7] == "-" and ("T" in x)


def _format_date_columns(df):
    import pandas as pd

    for col in df.columns:
        ncol = _norm(col)
        if (ncol in DATE_HEADER_HINTS) or any(h in ncol for h in DATE_HEADER_HINTS):
            try:
                df[col] = pd.to_datetime(df[col], errors="coerce", utc=False)
                if df[col].isna().all():
                    df[col] = df[col].apply(
                        lambda v: pd.to_datetime(v, errors="coerce") if _is_iso_datetime_str(str(v)) else pd.NaT
                    )
                df[col] = df[col].dt.strftime("%d/%m/%Y")
            except Exception:
                pass
    return df


def _load_template_headers(template_path: str) -> list[str]:
    import pandas as pd

    df_t = pd.read_excel(template_path, nrows=0, header=TEMPLATE_HEADER_ROW - 1, engine="openpyxl")
    return list(df_t.columns)


def _build_dataframe_with_template(df, template_cols: list[str]):
    import pandas as pd

    norm_template = {_norm(c): c for c in template_cols}

    idx_df_norm = {}
    for c in df.columns:
        idx_df_norm.setdefault(_norm(str(c)), []).append(c)

    # 1) Mapeamentos diretos
    mapping = {}
    for k_json, k_model in MAPEAMENTO_COLUNAS.items():
        if k_json in df.columns and k_model in template_cols:
            mapping[k_json] = k_model

    usados_df = set(mapping.keys())
    usados_model = set(mapping.values())

    # 2) Aliases por nome
    for col_model in template_cols:
        if col_model in usados_model:
            continue
        nmodel = _norm(col_model)
        for cand in TEMPLATE_ALIASES.get(nmodel, []):
            if cand in df.columns and cand not in usados_df:
                mapping[cand] = col_model
                usados_df.add(cand)
                usados_model.add(col_model)
                break
            for alt in idx_df_norm.get(_norm(cand), []):
                if alt not in usados_df:
                    mapping[alt] = col_model
                    usados_df.add(alt)
                    usados_model.add(col_model)
                    break
            if col_model in usados_model:
                break

    # 3) Heurística por similaridade simples
    for c in df.columns:
        if c in usados_df:
            continue
        nc = _norm(str(c))
        if nc in norm_template and norm_template[nc] not in usados_model:
            mapping[c] = norm_template[nc]
            usados_df.add(c)
            usados_model.add(norm_template[nc])
            continue
        for dst in [dst for nk, dst in norm_template.items() if nc and (nc in nk or nk in nc)]:
            if dst not in usados_model:
                mapping[c] = dst
                usados_df.add(c)
                usados_model.add(dst)
                break

    # 4) Construção do DF final
    out = pd.DataFrame()
    vazias = []
    for col_model in template_cols:
        origem = next((src for src, dst in mapping.items() if dst == col_model), None)
        if origem is None and col_model in df.columns:
            origem = col_model
        if origem is not None:
            out[col_model] = df[origem]
        else:
            out[col_model] = ""
            vazias.append(col_model)

    if vazias:
        log(f"[mapeamento] Colunas do MODELO sem origem (vazias): {', '.join(vazias)}")
    return out


def _apply_template_top_rows(out_path: Path):
    from openpyxl import load_workbook

    if TEMPLATE_HEADER_ROW <= 1:
        return

    wb_out = load_workbook(out_path.as_posix())
    ws_out = wb_out.active
    wb_tpl = load_workbook(TEMPLATE_XLSX)
    ws_tpl = wb_tpl.active

    # Inserir linhas superiores
    for _ in range(TEMPLATE_HEADER_ROW - 1):
        ws_out.insert_rows(1)

    # Copiar conteúdo das linhas de cabeçalho do template
    for r in range(1, TEMPLATE_HEADER_ROW):
        for c, cell in enumerate(ws_tpl[r], start=1):
            ws_out.cell(row=r, column=c, value=cell.value)

    # Mesclas que pertençam às linhas superiores
    for rng in ws_tpl.merged_cells.ranges:
        if rng.max_row <= (TEMPLATE_HEADER_ROW - 1):
            ws_out.merge_cells(
                start_row=rng.min_row,
                start_column=rng.min_col,
                end_row=rng.max_row,
                end_column=rng.max_col,
            )

    # Largura de colunas (se houver)
    try:
        for key, dim in ws_tpl.column_dimensions.items():
            if dim.width:
                ws_out.column_dimensions[key].width = dim.width
    except Exception:
        pass

    wb_out.save(out_path.as_posix())


def write_excel_from_json(data, out_path: Path):
    import pandas as pd

    if isinstance(data, list):
        rows = data
    elif isinstance(data, dict):
        rows = None
        for key in ("items", "data", "resultado", "rows", "content"):
            if key in data and isinstance(data[key], list):
                rows = data[key]
                break
        rows = rows or [data]
    else:
        rows = [{"raw": data}]

    df_raw = pd.json_normalize(rows)
    log("[debug] colunas do JSON normalizado: " + ", ".join(df_raw.columns.astype(str).tolist()))

    template_cols = _load_template_headers(TEMPLATE_XLSX)
    df_final = _build_dataframe_with_template(df_raw, template_cols)
    df_final = _format_date_columns(df_final)

    df_final.to_excel(out_path.as_posix(), index=False, engine="openpyxl")
    _apply_template_top_rows(out_path)
    log(f"Arquivo salvo com estrutura do modelo: {out_path.name}")

# -------------- TROCAR LOJA --------------
# -------------- TROCAR LOJA --------------
def trocar_loja(page, loja_nome=None):
    """
    Tenta trocar a loja pelo menu superior.
    Estratégia: procurar qualquer botão/área com 'TEA SHOP' ou nome da cidade e clicar.
    Depois selecionar o item do menu pelo texto (aceita ponto final).
    Retorna True/False indicando sucesso.
    """
    alvo = (
        loja_nome
        or os.getenv("APPNEXT_LOJA_DESTINO")
        or "GOIANIA - TEA SHOP FLAMBOYANT"
    )
    alvo_regex = re.compile(rf"^{re.escape(alvo)}\.?$", re.IGNORECASE)

    step(f"2) Trocando loja → {alvo}")
    try:
        opened = _click_first(
            page,
            [
                "button:has-text('TEA SHOP')",
                "button:has-text('VILA MADALENA')",
                "button:has-text('FLAMBOYANT')",
                "text=/S[ÂA]O PAULO - TEA SHOP|GOI[ÂA]NIA - TEA SHOP/i",
                "[data-bs-toggle='dropdown']",
                ".dropdown-toggle",
                "xpath=(//*[contains(translate(normalize-space(.),'áãéíóúâêô','aaeiouaeo'),'TEA SHOP')])[1]",
            ],
        )

        if not opened:
            # último recurso: clicar em algo com ícone de loja (emoji/mdi) + caret
            opened = _click_first(
                page,
                [
                    "xpath=(//i[contains(@class,'store') or contains(@class,'shop')]/ancestor::*[self::button or self::a])[1]"
                ],
            )

        # Seleciona a opção pelo texto
        try:
            page.get_by_role("menuitem", name=alvo_regex).first.click(timeout=6000, force=True)
        except Exception:
            clicked = _click_first(
                page,
                [
                    f"text=^{alvo}$",
                    f"text=^{re.escape(alvo)}",
                    f"xpath=//*[normalize-space()='{alvo}'] | //*[starts-with(normalize-space(), '{alvo}')]",
                ],
            )
            if not clicked:
                raise RuntimeError("Não encontrei a opção da loja no dropdown.")

        log(f"Loja selecionada: {alvo}")
        try:
            page.wait_for_load_state("networkidle", timeout=12000)
        except Exception:
            pass

        return True

    except Exception as e:
        log(f"Falha ao trocar loja: {e}")
        return False

# -------- Navegação resiliente: goto com retries --------
def goto_login_with_retries(page, url, tries=5):
    alt_urls = {url, url.replace("https://www.", "https://")}
    for i in range(1, tries + 1):
        for target in alt_urls:
            try:
                log(f"[goto] tentativa {i}/{tries} -> {target}")
                page.goto(target, wait_until="load", timeout=240_000)
                return
            except Exception as e:
                log(f"[goto] falhou em {target}: {e}")
        page.wait_for_timeout(i * 5000)  # backoff progressivo
    page.goto(url, wait_until="load", timeout=240_000)

# ================ MAIN =================
def main():
    with sync_playwright() as pw:
        context = pw.chromium.launch_persistent_context(
            user_data_dir=str(Path.cwd() / "pw_state"),
            headless=True,
            accept_downloads=True,
            viewport={"width": 1600, "height": 950},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36",
        )
        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_timeout(180_000)
        page.set_default_navigation_timeout(300_000)

        page.on("console", lambda m: log(f"[console] {m.type}: {m.text}"))
        page.on("pageerror", lambda e: log(f"[pageerror] {e}"))
        page.on("requestfailed", lambda r: log(f"[requestfailed] {r.url} -> {r.failure}"))

        hook_filter_capture(page)
        hook_export_capture(page)

        try:
            step("1) Login")
            goto_login_with_retries(page, LOGIN_URL)

            if "/#/login" in page.url:
                page.fill(SEL["rede"], REDE)
                page.fill(SEL["email"], USER)
                page.fill(SEL["senha"], PASS)
                page.click(SEL["entrar"])

            page.wait_for_url("**/#/loja/**", timeout=120_000)
            try:
                page.wait_for_load_state("networkidle", timeout=15_000)
            except Exception:
                pass

            # 2) Trocar loja via secret (habilite se quiser forçar)
            # trocar_loja(page, os.getenv("APPNEXT_LOJA_DESTINO"))

            # 3) Ir para Vendas > Vendedor Analítico
            step("3) Menu: Vendas → Vendedor Analítico")
            goto_vendedor_analitico_via_menu(page)
            page.wait_for_selector(SEL["titulo_rel"], timeout=40000)
            log("Tela confirmada.")

            # 4) Abrir filtros e estabilizar grid
            step("4) Abrindo filtros")
            open_filters_pane(page)
            try:
                page.locator(SEL["btn_atualizar"]).first.click(timeout=3000)
            except Exception:
                pass
            page.wait_for_selector(SEL["grid_row"], timeout=30000)

            # 5) Export
            step("5) Exportar Excel (tentar capturar URL)")
            captured_export_ok = False
            try:
                with page.expect_download(timeout=180000) as dl_tmp:
                    page.locator(SEL["excel_btn"]).first.click()
                tmp_download = dl_tmp.value
                tmp_path = Path.cwd() / f"_tmp_export_{TS}.xlsx"
                tmp_download.save_as(tmp_path.as_posix())
                log("Export padrão baixado (pode ser blob:/POST). Tentando reaproveitar URL capturada...")

                if CapturedExport.url:
                    step(f"5b) Reemitindo export com período {FMT_BR_INI} → {FMT_BR_FIM}")
                    replay_export_with_dates(page, LOCAL_OUT)
                    captured_export_ok = True
                else:
                    log("Não consegui capturar URL direta do Excel (provável blob:). Usarei fallback via API.")

                try:
                    tmp_path.unlink(missing_ok=True)
                except Exception:
                    pass
            except Exception as e:
                log(f"Aviso: falha ao acionar o botão Excel ou capturar URL: {e}. Vou usar fallback via API.")

            if not captured_export_ok:
                step(f"5c) Fallback: baixando dados da API {FMT_BR_INI} → {FMT_BR_FIM} e gerando Excel")
                data = fetch_report_json_with_dates(page)
                write_excel_from_json(data, LOCAL_OUT)

            # 6) Upload no Drive
            step("6) Enviando para Google Drive")
            rclone_copy_latest(LOCAL_OUT)
            log("Upload concluído.")

        except Exception:
            log("### ERRO DURANTE O FLUXO ###")
            print(traceback.format_exc())
        finally:
            try:
                context.close()
            except Exception:
                pass

    step("FINALIZADO")
    log(f"Local: {LOCAL_OUT.resolve()}")
    log(f"Drive: {DRIVE_FILE_NAME}  (pasta ID {DRIVE_FOLDER_ID})")


if __name__ == "__main__":
    main()
