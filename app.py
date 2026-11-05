import io
import re
import unicodedata
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Fluxo GN | Lista Quente 90d x Anomalias STC",
    page_icon="⚡",
    layout="wide",
)

NONE_OPTION = "(não usar)"
TIPOS = ["Poda", "Espaçador", "Ramal", "Outros"]


# ----------------------------------------------------------------------------
# Utilitários
# ----------------------------------------------------------------------------
def norm_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def norm_header(value):
    value = unicodedata.normalize("NFKD", str(value))
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", value).strip().lower()


def find_column(df, candidates):
    """Procura a coluna por nome exato e, depois, por nome que contenha o termo."""
    lookup = {norm_header(c): c for c in df.columns}
    for cand in candidates:
        key = norm_header(cand)
        if key in lookup:
            return lookup[key]
    for cand in candidates:
        key = norm_header(cand)
        for norm, original in lookup.items():
            if key in norm:
                return original
    return None


def clean_gn(value):
    if pd.isna(value):
        return ""
    match = re.search(r"GN\s*[-:]?\s*(\d+)", str(value).upper())
    return f"GN{match.group(1)}" if match else ""


def number_value(value):
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace("%", "").replace(",", ".")
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    return float(match.group()) if match else 0.0


def to_date(series):
    return pd.to_datetime(series, dayfirst=True, errors="coerce").dt.normalize()


def unique_join(series, separator=", "):
    values = []
    for value in series:
        value = norm_text(value)
        if value and value not in ("—", "-", "nan", "None") and value not in values:
            values.append(value)
    return separator.join(values)


def fmt_date(value):
    return "" if pd.isna(value) else value.strftime("%d/%m/%Y")


def for_screen(df):
    """Formata datas como dd/mm/aaaa para exibição na tela."""
    out = df.copy()
    for col in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[col]):
            out[col] = out[col].apply(fmt_date)
    return out


@st.cache_data(show_spinner=False)
def read_excel_bytes(content, expected_column):
    book = pd.ExcelFile(io.BytesIO(content), engine="openpyxl")
    for sheet in book.sheet_names:
        frame = pd.read_excel(book, sheet_name=sheet, dtype=object)
        if expected_column in {norm_header(c) for c in frame.columns}:
            return frame, sheet
    return pd.read_excel(book, sheet_name=book.sheet_names[0], dtype=object), book.sheet_names[0]


def classify_tipo(text):
    t = norm_header(text)
    if "poda" in t or "vegeta" in t or "arvore" in t:
        return "Poda"
    if "espa" in t:
        return "Espaçador"
    if "ramal" in t:
        return "Ramal"
    return "Outros"


def status_flags(status, exec_date):
    s = norm_header(status)
    executed = any(k in s for k in ("executad", "conclu", "finaliz", "realizad")) or pd.notna(exec_date)
    scheduled = (not executed) and any(k in s for k in ("programad", "agendad"))
    return executed, scheduled


# ----------------------------------------------------------------------------
# Regra de negócio
# ----------------------------------------------------------------------------
def build_flow(anom_raw, hot_raw, cols, today, window_days, stc_col, stc_value):
    """
    cols: dicionário com as colunas escolhidas (None quando não usada).
    Retorna (dicionário de abas, informações de apoio).
    """
    anom = anom_raw.copy()
    total_registros = len(anom)

    # 1) Somente informações do STC
    removidos_stc = 0
    if stc_col and stc_value.strip():
        alvo = norm_header(stc_value)
        mask = anom[stc_col].apply(lambda v: alvo in norm_header(v) if pd.notna(v) else False)
        removidos_stc = int((~mask).sum())
        anom = anom[mask].copy()

    # 2) GN de cada anomalia (Circuito; se vazio, Observação)
    def get_gn(row):
        gn = clean_gn(row.get(cols["a_circuito"]))
        if not gn and cols["a_obs"]:
            gn = clean_gn(row.get(cols["a_obs"]))
        return gn

    anom["_GN"] = anom.apply(get_gn, axis=1)
    sem_gn = int((anom["_GN"] == "").sum())
    anom = anom[anom["_GN"] != ""].copy()

    # 3) Campos derivados por anomalia
    blank = pd.Series([pd.NaT] * len(anom), index=anom.index)
    anom["_d_anom"] = to_date(anom[cols["a_data_anom"]]) if cols["a_data_anom"] else blank
    anom["_d_oe"] = to_date(anom[cols["a_data_oe"]]) if cols["a_data_oe"] else blank
    anom["_d_exec"] = to_date(anom[cols["a_data_exec"]]) if cols["a_data_exec"] else blank
    anom["_oe_num"] = anom[cols["a_oe"]].apply(norm_text) if cols["a_oe"] else ""
    anom["_tipo"] = anom[cols["a_tipo"]].apply(classify_tipo) if cols["a_tipo"] else "Outros"

    flags = anom.apply(
        lambda r: status_flags(r.get(cols["a_status"]) if cols["a_status"] else "", r["_d_exec"]),
        axis=1,
    )
    anom["_exec"] = [f[0] for f in flags]
    anom["_prog"] = [f[1] for f in flags]
    anom["_pend"] = ~(anom["_exec"] | anom["_prog"])
    anom["_has_oe"] = (anom["_oe_num"] != "") | anom["_d_oe"].notna()
    anom["_pend_sem_oe"] = (~anom["_exec"]) & (~anom["_has_oe"])
    anom["_d_exec_ef"] = anom["_d_exec"].where(anom["_exec"])
    for tipo in TIPOS:
        anom[f"_t_{tipo}"] = (anom["_tipo"] == tipo).astype(int)

    # 4) Consolidação por GN
    if len(anom):
        agg = anom.groupby("_GN").agg(
            Qtd_Anomalias=("_GN", "size"),
            Pendentes_Execucao=("_pend", "sum"),
            Programadas=("_prog", "sum"),
            Executadas=("_exec", "sum"),
            Sem_OE=("_pend_sem_oe", "sum"),
            Poda=("_t_Poda", "sum"),
            Espacador=("_t_Espaçador", "sum"),
            Ramal=("_t_Ramal", "sum"),
            Outros=("_t_Outros", "sum"),
            Data_Ultima_Anomalia=("_d_anom", "max"),
            Data_Ultima_OE=("_d_oe", "max"),
            Data_Ultima_Execucao=("_d_exec_ef", "max"),
            Qtd_OE=("_oe_num", lambda s: len({v for v in s if v})),
            OE_Numeros=("_oe_num", lambda s: unique_join(s, " | ")),
            Tem_OE=("_has_oe", "any"),
        ).reset_index().rename(columns={"_GN": "GN"})
        agg["Pendentes_Execucao"] = agg["Pendentes_Execucao"].astype(int)
        agg["Programadas"] = agg["Programadas"].astype(int)
        agg["Executadas"] = agg["Executadas"].astype(int)
        agg["Sem_OE"] = agg["Sem_OE"].astype(int)
    else:
        agg = pd.DataFrame(columns=[
            "GN", "Qtd_Anomalias", "Pendentes_Execucao", "Programadas", "Executadas", "Sem_OE",
            "Poda", "Espacador", "Ramal", "Outros", "Data_Ultima_Anomalia", "Data_Ultima_OE",
            "Data_Ultima_Execucao", "Qtd_OE", "OE_Numeros", "Tem_OE",
        ])

    # 5) Lista Quente (90 dias)
    hot = hot_raw.copy()
    hot["GN"] = hot[cols["h_gn"]].apply(clean_gn)
    hot = hot[hot["GN"] != ""].copy()
    hot_out = pd.DataFrame({"GN": hot["GN"]})
    hot_out["Prioridade"] = pd.to_numeric(hot[cols["h_prior"]], errors="coerce") if cols["h_prior"] else pd.NA
    hot_out["Reinc_90_dias"] = hot[cols["h_reinc"]].apply(number_value) if cols["h_reinc"] else 0.0
    hot_out["Data_Ultima_Ocorrencia"] = to_date(hot[cols["h_data_ocor"]]) if cols["h_data_ocor"] else pd.NaT
    hot_out["Conjunto"] = hot[cols["h_conjunto"]].apply(norm_text) if cols["h_conjunto"] else ""
    hot_out = hot_out.sort_values(["Reinc_90_dias", "Data_Ultima_Ocorrencia"], ascending=[False, False])
    hot_out = hot_out.drop_duplicates("GN", keep="first")
    hot_out["Na_Lista_Quente"] = "SIM"

    # 6) União das duas bases
    full = hot_out.merge(agg, on="GN", how="outer")
    full["Na_Lista_Quente"] = full["Na_Lista_Quente"].fillna("NÃO")
    full["Inspecionada_STC"] = full["Qtd_Anomalias"].notna().map({True: "SIM", False: "NÃO"})
    for c in ["Qtd_Anomalias", "Pendentes_Execucao", "Programadas", "Executadas", "Sem_OE",
              "Poda", "Espacador", "Ramal", "Outros", "Qtd_OE"]:
        full[c] = full[c].fillna(0).astype(int)
    full["Reinc_90_dias"] = full["Reinc_90_dias"].fillna(0)
    for c in ["Conjunto", "OE_Numeros"]:
        full[c] = full[c].fillna("")
    full["Tem_OE"] = full["Tem_OE"].fillna(False).astype(bool)

    # 7) Status da miscelânea
    def status_misc(r):
        if r["Inspecionada_STC"] == "NÃO":
            return "—"
        if r["Sem_OE"] > 0:
            return "PENDENTE DE ABERTURA"
        if r["Pendentes_Execucao"] + r["Programadas"] > 0:
            return "OE GERADA – EM EXECUÇÃO"
        return "TODAS EXECUTADAS"

    full["Status_Miscelanea"] = full.apply(status_misc, axis=1)

    # 8) Janela de 90 dias (última execução x nova OE)
    today = pd.Timestamp(today).normalize()
    exec_d = full["Data_Ultima_Execucao"]
    oe_d = full["Data_Ultima_OE"]
    full["Fim_Janela_90d"] = exec_d + pd.Timedelta(days=window_days)
    full["Dias_Desde_Execucao"] = (today - exec_d).dt.days
    full["Dias_Restantes_Janela"] = (full["Fim_Janela_90d"] - today).dt.days
    full["Dias_Execucao_ate_Nova_OE"] = (oe_d - exec_d).dt.days.where(oe_d > exec_d)
    full["Dias_Desde_Ultima_OE"] = (today - oe_d).dt.days
    ocor = full["Data_Ultima_Ocorrencia"]
    full["Ocorrencia_Apos_Execucao"] = (ocor > exec_d).map({True: "SIM", False: "NÃO"})
    full["Dias_Execucao_ate_Ocorrencia"] = (ocor - exec_d).dt.days.where(ocor > exec_d)

    def situacao(r):
        if r["Inspecionada_STC"] == "NÃO":
            return "SEM INSPEÇÃO STC"
        if pd.isna(r["Data_Ultima_Execucao"]):
            if r["Executadas"] > 0:
                return "EXECUTADA SEM DATA (informar)"
            return "AGUARDANDO EXECUÇÃO"
        nova_oe = r["Dias_Execucao_ate_Nova_OE"]
        if pd.notna(nova_oe) and nova_oe <= window_days:
            return f"REINCIDENTE – nova OE em até {window_days} dias"
        if today <= r["Fim_Janela_90d"]:
            return f"EM JANELA {window_days}d – NÃO GERAR NOVA OE"
        return "JANELA ENCERRADA"

    full["Situacao_90d"] = full.apply(situacao, axis=1)

    def acao(r):
        s = r["Situacao_90d"]
        if s == "SEM INSPEÇÃO STC":
            return "Delegar fiscal (Lista Quente)"
        if r["Status_Miscelanea"] == "PENDENTE DE ABERTURA":
            if s.startswith("EM JANELA") or s.startswith("REINCIDENTE"):
                return "Avaliar antes de abrir OE (dentro dos 90 dias da execução)"
            return "Gerar miscelânea / OE"
        if r["Status_Miscelanea"] == "OE GERADA – EM EXECUÇÃO":
            return "Acompanhar execução em campo"
        if s.startswith("EM JANELA"):
            return "Monitorar – não gerar nova OE até o fim da janela"
        if s.startswith("REINCIDENTE"):
            return "Analisar causa da reincidência"
        return "Sem ação – acompanhar Lista Quente"

    full["Acao_Recomendada"] = full.apply(acao, axis=1)

    ordered = [
        "GN", "Prioridade", "Conjunto", "Na_Lista_Quente", "Reinc_90_dias", "Data_Ultima_Ocorrencia",
        "Inspecionada_STC", "Qtd_Anomalias", "Poda", "Espacador", "Ramal", "Outros",
        "Pendentes_Execucao", "Programadas", "Executadas",
        "Status_Miscelanea", "Qtd_OE", "OE_Numeros", "Data_Ultima_Anomalia", "Data_Ultima_OE",
        "Data_Ultima_Execucao", "Fim_Janela_90d", "Dias_Desde_Execucao", "Dias_Restantes_Janela",
        "Dias_Execucao_ate_Nova_OE", "Dias_Desde_Ultima_OE", "Ocorrencia_Apos_Execucao",
        "Dias_Execucao_ate_Ocorrencia", "Situacao_90d", "Acao_Recomendada",
    ]
    full = full[ordered].sort_values(
        ["Reinc_90_dias", "Pendentes_Execucao"], ascending=[False, False]
    ).reset_index(drop=True)

    # 9) Abas
    pend_abertura = full[full["Status_Miscelanea"] == "PENDENTE DE ABERTURA"]
    em_execucao = full[full["Status_Miscelanea"] == "OE GERADA – EM EXECUÇÃO"]
    executadas = full[(full["Executadas"] > 0)].sort_values("Data_Ultima_Execucao", ascending=False)
    em_janela = full[full["Situacao_90d"].str.startswith("EM JANELA")].sort_values("Dias_Restantes_Janela")
    reincidentes = full[
        full["Situacao_90d"].str.startswith("REINCIDENTE") | (full["Ocorrencia_Apos_Execucao"] == "SIM")
    ]
    sem_inspecao = full[full["Situacao_90d"] == "SEM INSPEÇÃO STC"]

    detalhe_cols = [c for c in anom.columns if not c.startswith("_")]
    detalhe = anom[["_GN"] + detalhe_cols].rename(columns={"_GN": "GN"}).copy()
    detalhe.insert(1, "Tipo (classificado)", anom["_tipo"])
    detalhe.insert(2, "Situação do registro", anom.apply(
        lambda r: "EXECUTADA" if r["_exec"] else ("PROGRAMADA" if r["_prog"] else "PENDENTE"), axis=1))
    detalhe.insert(3, "OE gerada", anom["_has_oe"].map({True: "SIM", False: "NÃO"}))

    por_tipo = full[full["Inspecionada_STC"] == "SIM"][
        ["GN", "Qtd_Anomalias", "Poda", "Espacador", "Ramal", "Outros",
         "Pendentes_Execucao", "Programadas", "Executadas"]
    ]

    resumo = pd.DataFrame([
        ("Data de referência", today.strftime("%d/%m/%Y")),
        ("Janela de reincidência (dias)", window_days),
        ("Registros na base Anomalias", total_registros),
        ("Registros removidos pelo filtro STC", removidos_stc),
        ("Registros STC sem GN identificada", sem_gn),
        ("Registros STC considerados", len(anom)),
        ("GN na Lista Quente", int((full["Na_Lista_Quente"] == "SIM").sum())),
        ("GN inspecionadas pelo STC", int((full["Inspecionada_STC"] == "SIM").sum())),
        ("GN da Lista Quente sem inspeção STC", len(sem_inspecao)),
        ("GN pendentes de abertura de miscelânea", len(pend_abertura)),
        ("GN com OE gerada em execução", len(em_execucao)),
        ("GN com anomalia executada", len(executadas)),
        ("GN em janela de 90 dias (não gerar OE)", len(em_janela)),
        ("GN reincidentes", len(reincidentes)),
    ], columns=["Indicador", "Resultado"])

    criterios = pd.DataFrame({"Critério": [
        "Somente registros do STC entram no fluxo (filtro pela coluna escolhida na barra lateral).",
        "A GN da anomalia vem da coluna Circuito; se não houver GN, procura na coluna Observação.",
        "Tipo da anomalia: Poda, Espaçador, Ramal ou Outros (classificado pelo texto da coluna Tipo).",
        "Anomalia EXECUTADA: status contém executado/concluído/finalizado/realizado ou existe data de execução.",
        "Anomalia PROGRAMADA: status contém programado/agendado. As demais são PENDENTES de execução.",
        "Status da miscelânea: PENDENTE DE ABERTURA quando há anomalia não executada sem OE/data de OE.",
        f"Janela de {window_days} dias: conta a partir da data da última anomalia executada da GN.",
        f"REINCIDENTE: nova OE gerada até {window_days} dias depois da última execução.",
        f"EM JANELA: ainda dentro dos {window_days} dias da execução e sem nova OE – não gerar nova OE.",
        "JANELA ENCERRADA: passaram os dias da janela; nova OE deixa de contar como reincidência.",
        "Ocorrência após execução: data da última ocorrência (Lista Quente) posterior à última execução.",
        "Reinc. 90 dias e data da última ocorrência vêm da Lista Quente e servem para priorizar fiscais.",
    ]})

    sheets = {
        "Resumo": resumo,
        "Acompanhamento GN": full,
        "Gerar Miscelânea": pend_abertura,
        "OE em Execução": em_execucao,
        "Executadas": executadas,
        "Em Janela 90d": em_janela,
        "Reincidentes": reincidentes,
        "Sem Inspeção STC": sem_inspecao,
        "Anomalias por Tipo": por_tipo,
        "Detalhe Anomalias STC": detalhe,
        "Critérios": criterios,
    }
    return sheets


# ----------------------------------------------------------------------------
# Excel
# ----------------------------------------------------------------------------
SITUACAO_CORES = {
    "REINCIDENTE": ("F4CCCC", "990000"),
    "EM JANELA": ("FFF2CC", "7F6000"),
    "AGUARDANDO": ("FCE4D6", "C00000"),
    "EXECUTADA SEM": ("FCE4D6", "C00000"),
    "JANELA ENCERRADA": ("D9EAD3", "274E13"),
    "SEM INSPEÇÃO": ("EDEDED", "595959"),
}


def format_workbook(sheets):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name[:31], index=False)

        header_fill = PatternFill("solid", fgColor="17365D")
        header_font = Font(color="FFFFFF", bold=True)
        border = Border(bottom=Side(style="thin", color="D9E2F3"))

        for ws in writer.book.worksheets:
            ws.freeze_panes = "B2"
            ws.auto_filter.ref = ws.dimensions
            ws.sheet_view.showGridLines = False
            ws.row_dimensions[1].height = 34
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            for idx, cells in enumerate(ws.columns, start=1):
                lens = [len(str(c.value)) if c.value is not None else 0 for c in list(cells)[:250]]
                ws.column_dimensions[get_column_letter(idx)].width = min(max(max(lens, default=10) + 2, 12), 46)

            headers = {c.value: c.column for c in ws[1]}
            for header, col in headers.items():
                if header and ("Data" in str(header) or "Fim_Janela" in str(header)):
                    for row in range(2, ws.max_row + 1):
                        ws.cell(row, col).number_format = "DD/MM/YYYY"

            sit_col = headers.get("Situacao_90d")
            if sit_col:
                for row in range(2, ws.max_row + 1):
                    cell = ws.cell(row, sit_col)
                    for key, (bg, fg) in SITUACAO_CORES.items():
                        if str(cell.value).startswith(key):
                            cell.fill = PatternFill("solid", fgColor=bg)
                            cell.font = Font(color=fg, bold=True)
                            break

            misc_col = headers.get("Status_Miscelanea")
            if misc_col:
                for row in range(2, ws.max_row + 1):
                    cell = ws.cell(row, misc_col)
                    if cell.value == "PENDENTE DE ABERTURA":
                        cell.fill = PatternFill("solid", fgColor="FCE4D6")
                        cell.font = Font(color="C00000", bold=True)
    output.seek(0)
    return output


# ----------------------------------------------------------------------------
# Interface
# ----------------------------------------------------------------------------
st.title("⚡ Fluxo de GN – Lista Quente 90 dias x Anomalias STC")
st.caption(
    "Cruza a Lista Quente (top reincidentes) com as anomalias levantadas pelo fiscal do STC, "
    "acompanha miscelânea/OE, execução e a janela de 90 dias que define a reincidência."
)

with st.sidebar:
    st.header("Parâmetros")
    today_input = st.date_input("Data de referência", value=datetime.now().date(), format="DD/MM/YYYY")
    window_days = st.number_input("Janela de reincidência (dias)", min_value=1, max_value=365, value=90)
    st.divider()
    st.header("Como usar")
    st.markdown(
        "1. Anexe a **Lista Quente (90 dias)**.  \n"
        "2. Anexe a base **Anomalias**.  \n"
        "3. Confira o **mapeamento de colunas**.  \n"
        "4. Clique em **Gerar fluxo** e baixe o Excel."
    )

c1, c2 = st.columns(2)
with c1:
    hot_file = st.file_uploader("1. Lista Quente – 90 dias (.xlsx)", type=["xlsx"], key="hot")
with c2:
    anom_file = st.file_uploader("2. Base de Anomalias (.xlsx)", type=["xlsx"], key="anom")

if hot_file and anom_file:
    try:
        hot_raw, hot_sheet = read_excel_bytes(hot_file.getvalue(), "trafo")
        anom_raw, anom_sheet = read_excel_bytes(anom_file.getvalue(), "circuito")
    except Exception as exc:
        st.error(f"Não foi possível ler as planilhas: {exc}")
        st.stop()

    def pick(label, df, candidates, key, required=False):
        options = list(df.columns) if required else [NONE_OPTION] + list(df.columns)
        found = find_column(df, candidates)
        index = options.index(found) if found in options else 0
        choice = st.selectbox(label, options, index=index, key=key)
        return None if choice == NONE_OPTION else choice

    with st.expander("⚙️ Mapeamento de colunas (confira antes de gerar)", expanded=False):
        st.caption("O app tenta reconhecer as colunas sozinho. Ajuste aqui se alguma estiver errada.")
        ca, cb = st.columns(2)
        with ca:
            st.markdown(f"**Lista Quente** (aba: {hot_sheet})")
            cols = {
                "h_gn": pick("GN / Trafo", hot_raw, ["Trafo", "GN", "Transformador"], "h_gn", required=True),
                "h_prior": pick("Prioridade (ranking)", hot_raw, ["Prioridade", "Ranking"], "h_prior"),
                "h_reinc": pick("Reincidência 90 dias", hot_raw,
                                ["Reinc. 90 dias", "Reincidência 90 dias", "Reinc 90", "Reincid"], "h_reinc"),
                "h_data_ocor": pick("Data da última ocorrência", hot_raw,
                                    ["Última ocorrência", "Ultima ocorrencia", "Data última ocorrência", "Data ocorrência"],
                                    "h_data_ocor"),
                "h_conjunto": pick("Conjunto", hot_raw, ["Conjunto"], "h_conjunto"),
            }
        with cb:
            st.markdown(f"**Anomalias** (aba: {anom_sheet})")
            cols.update({
                "a_circuito": pick("Circuito (GN)", anom_raw, ["Circuito"], "a_circuito", required=True),
                "a_obs": pick("Observação (GN alternativa)", anom_raw, ["Observação", "Observacao"], "a_obs"),
                "a_tipo": pick("Tipo da anomalia", anom_raw, ["Tipo"], "a_tipo"),
                "a_status": pick("Status", anom_raw, ["Status"], "a_status"),
                "a_data_anom": pick("Data da anomalia / inspeção", anom_raw,
                                    ["Data anomalia", "Data inspeção", "Data inspecao", "Data registro", "Data"],
                                    "a_data_anom"),
                "a_oe": pick("Nº da OE / Miscelânea", anom_raw, ["OE", "Miscelânea", "Miscelanea"], "a_oe"),
                "a_data_oe": pick("Data da OE gerada", anom_raw,
                                  ["Data OE", "Data Miscelânea", "Data Miscelanea", "Data abertura", "Data geração"],
                                  "a_data_oe"),
                "a_data_exec": pick("Data da execução", anom_raw,
                                    ["Data execução", "Data execucao", "Data conclusão", "Data conclusao", "Executado em"],
                                    "a_data_exec"),
            })
        st.markdown("**Filtro STC**")
        fa, fb = st.columns(2)
        with fa:
            stc_col = pick("Coluna que identifica o STC", anom_raw,
                           ["Equipe", "Empresa", "Origem", "Contratada", "Executor", "Fonte", "Área", "Area", "STC"],
                           "stc_col")
        with fb:
            stc_value = st.text_input("Valor a manter (contém)", value="STC", key="stc_value")

    if not cols["a_data_exec"]:
        st.warning("Sem coluna de **data da execução** a janela de 90 dias não pode ser calculada.")
    if not cols["a_data_oe"] and not cols["a_oe"]:
        st.warning("Sem coluna de **OE/data da OE**, todas as GN não executadas ficam como PENDENTE DE ABERTURA.")
    if not stc_col:
        st.warning("Nenhuma coluna de STC escolhida: **todas** as linhas da base Anomalias serão usadas.")

    if st.button("Gerar fluxo", type="primary"):
        try:
            with st.spinner("Cruzando as bases e calculando a janela de 90 dias..."):
                sheets = build_flow(anom_raw, hot_raw, cols, today_input, int(window_days), stc_col, stc_value)
                st.session_state["flow_sheets"] = sheets
                st.session_state["flow_bytes"] = format_workbook(sheets).getvalue()
                st.session_state["flow_name"] = f"Fluxo_GN_Reincidencia_{datetime.now().strftime('%d-%m-%Y')}.xlsx"
            st.success("Fluxo gerado!")
        except Exception as exc:
            st.error(f"Não foi possível gerar o fluxo: {exc}")
            st.exception(exc)
elif hot_file or anom_file:
    st.warning("Falta anexar a outra planilha.")

if "flow_sheets" in st.session_state:
    sheets = st.session_state["flow_sheets"]
    res = sheets["Resumo"].set_index("Indicador")["Resultado"].to_dict()
    st.divider()
    st.subheader("Resultado")

    m = st.columns(5)
    cards = [
        ("Inspecionadas STC", res["GN inspecionadas pelo STC"]),
        ("Gerar miscelânea", res["GN pendentes de abertura de miscelânea"]),
        ("OE em execução", res["GN com OE gerada em execução"]),
        ("Em janela 90d", res["GN em janela de 90 dias (não gerar OE)"]),
        ("Reincidentes", res["GN reincidentes"]),
    ]
    for col, (label, value) in zip(m, cards):
        with col:
            with st.container(border=True):
                st.metric(label, f"{int(value):,}".replace(",", "."))

    st.download_button(
        "⬇️ Baixar Excel completo (todas as abas)",
        data=st.session_state["flow_bytes"],
        file_name=st.session_state["flow_name"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    tabs = st.tabs([
        "📋 Gerar miscelânea", "🛠️ OE em execução", "✅ Executadas",
        "⏳ Em janela 90d", "🔁 Reincidentes", "👷 Sem inspeção STC", "🗂️ Todas as abas",
    ])
    views = [
        ("Gerar Miscelânea", "GN inspecionadas pelo STC com anomalia ainda sem OE."),
        ("OE em Execução", "OE gerada; aguardando execução em campo."),
        ("Executadas", "GN com anomalia executada e data da última execução."),
        ("Em Janela 90d", "Executadas há menos de 90 dias: não podem gerar nova OE."),
        ("Reincidentes", "Nova OE (ou nova ocorrência) dentro da janela após a execução."),
        ("Sem Inspeção STC", "GN da Lista Quente ainda sem anomalia do STC – candidatas a delegar fiscal."),
    ]
    for tab, (name, caption) in zip(tabs[:6], views):
        with tab:
            st.caption(f"{caption}  Total: **{len(sheets[name])}**")
            st.dataframe(for_screen(sheets[name]), use_container_width=True, hide_index=True)
    with tabs[6]:
        aba = st.selectbox("Visualizar aba", list(sheets.keys()))
        st.dataframe(for_screen(sheets[aba]).head(500), use_container_width=True, hide_index=True)

st.divider()
st.caption("Aplicativo de apoio à gestão operacional. Confira as bases e valide os critérios antes de direcionar serviços.")
