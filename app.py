import io
import re
import unicodedata
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter


st.set_page_config(
    page_title="Cruzamento de GN | Lista Quente x Anomalias",
    page_icon="⚡",
    layout="wide",
)

st.title("⚡ Cruzamento de Lista Quente x Anomalias")
st.caption(
    "Carregue as duas bases em Excel (.xlsx). O aplicativo cruza as informações pela GN "
    "e gera um relatório gerencial com pendências de abertura de miscelânea, reincidência e DEC."
)


def norm_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def norm_header(value):
    value = unicodedata.normalize("NFKD", str(value))
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", value).strip().lower()


def find_column(df, candidates, required=True):
    lookup = {norm_header(col): col for col in df.columns}
    for candidate in candidates:
        key = norm_header(candidate)
        if key in lookup:
            return lookup[key]
    if required:
        raise ValueError(
            "Não encontrei a coluna '" + candidates[0] +
            "'. Colunas disponíveis: " + ", ".join(map(str, df.columns))
        )
    return None


def clean_gn(value):
    """Padroniza identificadores GN sem alterar a sequência numérica."""
    if pd.isna(value):
        return ""
    value = str(value).upper().strip()
    match = re.search(r"GN\s*[-:]?\s*(\d+)", value)
    return f"GN{match.group(1)}" if match else ""


def extract_gn(row, circuit_col, observation_col):
    # Prioridade: coluna Circuito; se não houver GN, procura na Observação.
    gn = clean_gn(row.get(circuit_col, ""))
    if gn:
        return gn
    if observation_col:
        text = norm_text(row.get(observation_col, ""))
        match = re.search(r"GN\s*[-:]?\s*(\d+)", text.upper())
        if match:
            return f"GN{match.group(1)}"
    return ""


def number_value(value):
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace("%", "").replace(",", ".")
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    return float(match.group()) if match else 0.0


def pct_text(value):
    if pd.isna(value) or str(value).strip() in ("", "—", "-"):
        return ""
    text = str(value).strip()
    if "%" in text:
        return text
    try:
        return f"{float(str(value).replace(',', '.')):.2f}%"
    except Exception:
        return text


def first_nonempty(series):
    for value in series:
        value = norm_text(value)
        if value and value not in ("—", "-", "nan", "None"):
            return value
    return ""


def unique_join(series, separator=", "):
    values = []
    for value in series:
        value = norm_text(value)
        if value and value not in ("—", "-", "nan", "None") and value not in values:
            values.append(value)
    return separator.join(values)


def read_excel(uploaded_file, expected_columns):
    try:
        book = pd.ExcelFile(uploaded_file, engine="openpyxl")
        for sheet in book.sheet_names:
            frame = pd.read_excel(book, sheet_name=sheet, dtype=object)
            normalized = {norm_header(c) for c in frame.columns}
            if all(norm_header(c) in normalized for c in expected_columns):
                return frame, sheet
        # Fallback: primeira aba, para mostrar erro útil abaixo.
        return pd.read_excel(book, sheet_name=book.sheet_names[0], dtype=object), book.sheet_names[0]
    except Exception as exc:
        raise ValueError(f"Não foi possível ler o arquivo Excel: {exc}") from exc


def build_report(anom_raw, hot_raw):
    # Localiza as colunas principais, tolerando variações de acentuação e espaços.
    a_circuit = find_column(anom_raw, ["Circuito"])
    a_group = find_column(anom_raw, ["Conjunto"], required=False)
    a_group_status = find_column(anom_raw, ["Status Conjunto"], required=False)
    a_dec = find_column(anom_raw, ["%DEC", "DEC"], required=False)
    a_type = find_column(anom_raw, ["Tipo"], required=False)
    a_status = find_column(anom_raw, ["Status"], required=False)
    a_observation = find_column(anom_raw, ["Observação", "Observacao"], required=False)
    a_irregularity = find_column(anom_raw, ["Obs. Irregularidade"], required=False)
    a_admin = find_column(anom_raw, ["Obs. Admin"], required=False)
    a_miscelanea = find_column(anom_raw, ["Miscelânea", "Miscelanea"], required=False)

    h_gn = find_column(hot_raw, ["Trafo", "GN", "Transformador"])
    h_priority = find_column(hot_raw, ["Prioridade"], required=False)
    h_rec = find_column(hot_raw, ["Reinc. 90 dias", "Reincidência 90 dias"], required=False)

    anom = anom_raw.copy()
    anom["_GN"] = anom.apply(
        lambda row: extract_gn(row, a_circuit, a_observation), axis=1
    )
    anom = anom[anom["_GN"] != ""].copy()

    # Agrega os registros da base de anomalias por GN.
    anomaly_rows = []
    for gn, group in anom.groupby("_GN", sort=False):
        statuses = group[a_status].astype(str).str.strip().str.lower() if a_status else pd.Series([], dtype=str)
        pending = int(statuses.isin(["pendente", "pendente de execução", "não iniciado", "nao iniciado"]).sum()) if a_status else 0
        completed = int(statuses.isin(["concluído", "concluido", "executado", "concluída", "concluida"]).sum()) if a_status else 0
        scheduled = int(statuses.isin(["programado", "programada", "agendado", "agendada"]).sum()) if a_status else 0

        group_names = group[a_group] if a_group else pd.Series([], dtype=object)
        dec_status = group[a_group_status] if a_group_status else pd.Series([], dtype=object)
        dec_values = group[a_dec] if a_dec else pd.Series([], dtype=object)

        # Conjunto(s) com DEC acima da meta / status de transgressão.
        transgression = []
        for idx in group.index:
            status = norm_text(group.at[idx, a_group_status]).lower() if a_group_status else ""
            dec = group.at[idx, a_dec] if a_dec else ""
            if "transgress" in status or number_value(dec) > 100:
                name = norm_text(group.at[idx, a_group]) if a_group else ""
                if name and name not in transgression:
                    transgression.append(name)

        all_groups = unique_join(group_names, " | ") if a_group else ""
        status_conjunto = "Transgressão" if transgression else first_nonempty(dec_status)
        dec_conjunto = ""
        if a_dec:
            for val in dec_values:
                if norm_text(val):
                    dec_conjunto = pct_text(val)
                    if number_value(val) > 100:
                        break

        anomaly_rows.append({
            "GN_match": gn,
            "Qtd_Anomalias": len(group),
            "Anomalias_Pendentes": pending,
            "Anomalias_Concluidas": completed,
            "Anomalias_Programadas": scheduled,
            "Tipos_Anomalia": unique_join(group[a_type]) if a_type else "",
            "Status_DEC_Conjunto": status_conjunto,
            "Conjunto": all_groups,
            "DEC_Conjunto": dec_conjunto,
            "DEC_Transgressao": "SIM" if transgression else "NÃO",
            "_Transgressao_Conjuntos": " | ".join(transgression),
            "Miscelanea_Origem": unique_join(group[a_miscelanea], " | ") if a_miscelanea else "",
            "Observacoes_Anomalia": unique_join(
                pd.concat([
                    group[a_observation] if a_observation else pd.Series([], dtype=object),
                    group[a_irregularity] if a_irregularity else pd.Series([], dtype=object),
                    group[a_admin] if a_admin else pd.Series([], dtype=object),
                ], ignore_index=True),
                " | "
            ),
        })
    anomaly_summary = pd.DataFrame(anomaly_rows)

    hot = hot_raw.copy()
    hot["_GN"] = hot[h_gn].apply(clean_gn)
    # Mantém todos os transformadores da Lista Quente, inclusive identificadores
    # que não começam por GN (AP, SW, TR etc.). Só os identificadores GN participam
    # do cruzamento com a base Anomalias.
    if h_priority:
        hot["_sort_priority"] = pd.to_numeric(hot[h_priority], errors="coerce")
        hot = hot.sort_values(["_sort_priority"], na_position="last")
    else:
        hot["_sort_priority"] = range(1, len(hot) + 1)

    # Remove apenas duplicidades reais: GN quando identificada; caso contrário,
    # usa o identificador original da coluna Trafo.
    hot["_JOIN_KEY"] = hot.apply(
        lambda row: row["_GN"] if row["_GN"] else "ID:" + norm_text(row[h_gn]).upper(),
        axis=1,
    )
    hot = hot.drop_duplicates("_JOIN_KEY", keep="first").copy()
    anomaly_map = anomaly_summary.set_index("GN_match").to_dict("index")

    merged = hot.copy()
    for col in anomaly_summary.columns:
        if col != "GN_match":
            merged[col] = merged["_GN"].map(
                {gn: row.get(col, "") for gn, row in anomaly_map.items()}
            )

    merged["Consta_na_Anomalias"] = merged["_GN"].map(
        lambda gn: "SIM" if gn in anomaly_map else "NÃO"
    )
    merged["Inspecao_Lista_Quente"] = merged.get(
        find_column(hot_raw, ["Origem Levantamento"], required=False), pd.Series("", index=merged.index)
    ).fillna("").astype(str).apply(lambda x: "SIM" if x.strip() and x.strip() not in ("—", "-") else "NÃO")
    merged["Status_Miscelanea"] = merged["Consta_na_Anomalias"].apply(
        lambda x: "PENDENTE DE ABERTURA" if x == "SIM" else "—"
    )

    # DEC fora da meta é um indicador do conjunto, não da GN individual.
    merged["Prioridade_Analise"] = merged.apply(
        lambda r: (
            "ALTA – DEC fora da meta + reincidência"
            if r.get("DEC_Transgressao") == "SIM"
            else "ALTA – reincidência + anomalia pendente"
            if r.get("Consta_na_Anomalias") == "SIM" and number_value(r.get("Anomalias_Pendentes", 0)) > 0
            else "ALTA – reincidência + anomalia registrada"
            if r.get("Consta_na_Anomalias") == "SIM"
            else "Acompanhar conforme ranking"
        ), axis=1
    )

    # Nomes originais da Lista Quente, para preservar os campos de origem.
    hot_cols = [c for c in hot_raw.columns]
    full = merged.copy()
    full["GN / Transformador"] = full["_GN"]

    # Cruzamento completo (colunas de origem + indicadores).
    complete_cols = hot_cols + [
        "Qtd_Anomalias", "Anomalias_Pendentes", "Anomalias_Concluidas",
        "Anomalias_Programadas", "Tipos_Anomalia", "Status_DEC_Conjunto",
        "Conjunto", "DEC_Conjunto", "Consta_na_Anomalias",
        "Inspecao_Lista_Quente", "DEC_Transgressao", "Prioridade_Analise",
        "Status_Miscelanea",
    ]
    complete = pd.DataFrame(index=full.index)
    for c in hot_cols:
        complete[c] = full[c]
    for c in complete_cols[len(hot_cols):]:
        complete[c] = full[c]

    # Pendências de miscelânea: toda GN presente na base Anomalias é considerada
    # inspecionada e pendente de abertura, independentemente do Status do registro.
    hot_index = full[full["_GN"] != ""].drop_duplicates("_GN").set_index("_GN", drop=False)
    pending_rows = []
    for _, ar in anomaly_summary.iterrows():
        gn = ar["GN_match"]
        is_hot = gn in hot_index.index
        hr = hot_index.loc[gn] if is_hot else None
        if isinstance(hr, pd.DataFrame):
            hr = hr.iloc[0]

        def hot_value(candidates):
            if not is_hot:
                return ""
            col = find_column(hot_raw, candidates, required=False)
            return hr[col] if col else ""

        pval = hot_value(["Prioridade"])
        pending_rows.append({
            "Prioridade Lista Quente": pval,
            "GN / Transformador": gn,
            "Status Miscelânea": "PENDENTE DE ABERTURA",
            "Inspeção": "SIM",
            "Na Lista Quente?": "SIM" if is_hot else "NÃO",
            "Reinc. 90 dias": hot_value(["Reinc. 90 dias", "Reincidência 90 dias"]),
            "Total Histórico": hot_value(["Total Histórico"]),
            "Clientes": hot_value(["Clientes"]),
            "Última Ocorrência": hot_value(["Última Ocorrência"]),
            "Causa Recente": hot_value(["Causa Recente"]),
            "Conjunto": ar["Conjunto"],
            "Status DEC Conjunto": ar["Status_DEC_Conjunto"],
            "% DEC Conjunto": ar["DEC_Conjunto"],
            "Qtd. Registros Anomalia": ar["Qtd_Anomalias"],
            "Tipos de Anomalia": ar["Tipos_Anomalia"],
            "Status dos Registros de Anomalia": (
                f"Concluído: {ar['Anomalias_Concluidas']}; Pendente: {ar['Anomalias_Pendentes']}; "
                f"Programado: {ar['Anomalias_Programadas']}"
            ).strip("; "),
            "Miscelânea (campo origem)": ar["Miscelanea_Origem"],
            "Observações de Anomalia": ar["Observacoes_Anomalia"],
        })
    pending = pd.DataFrame(pending_rows)
    pending["_priority_sort"] = pd.to_numeric(pending["Prioridade Lista Quente"], errors="coerce")
    pending["_hot_sort"] = pending["Na Lista Quente?"].map({"SIM": 0, "NÃO": 1})
    pending = pending.sort_values(
        ["_hot_sort", "_priority_sort", "GN / Transformador"],
        na_position="last"
    ).drop(columns=["_priority_sort", "_hot_sort"]).reset_index(drop=True)

    # Abas analíticas.
    top15 = complete.head(15).copy()
    top15["Consta_na_Anomalias"] = top15["Consta_na_Anomalias"]
    dec_priority = complete[complete["DEC_Transgressao"] == "SIM"].copy()
    both = complete[complete["Consta_na_Anomalias"] == "SIM"].copy()
    only_anomaly = anomaly_summary[
        ~anomaly_summary["GN_match"].isin(set(hot["_GN"]))
    ].copy()

    # Tabela por GN (para inspeção detalhada).
    anomaly_by_gn = anomaly_summary.drop(columns=["DEC_Transgressao", "_Transgressao_Conjuntos", "Miscelanea_Origem", "Observacoes_Anomalia"], errors="ignore")
    dec_groups = []
    if a_group:
        tmp = anom.copy()
        tmp["_DEC_NUM"] = tmp[a_dec].apply(number_value) if a_dec else 0
        for group_name, g in tmp.groupby(a_group, dropna=False):
            if not norm_text(group_name):
                continue
            status = first_nonempty(g[a_group_status]) if a_group_status else ""
            dec = first_nonempty(g[a_dec]) if a_dec else ""
            if ("transgress" in status.lower()) or (a_dec and g["_DEC_NUM"].max() > 100):
                gn_set = sorted(set(g["_GN"].dropna()) - {""})
                hot_set = sorted(set(gn_set) & set(hot["_GN"]))
                dec_groups.append({
                    "Conjunto": group_name,
                    "%DEC": pct_text(dec),
                    "Registros_Anomalia": len(g),
                    "GNs_identificadas": len(gn_set),
                    "GNs_na_Lista_Quente": len(hot_set),
                })
    dec_transgression = pd.DataFrame(
        dec_groups,
        columns=["Conjunto", "%DEC", "Registros_Anomalia", "GNs_identificadas", "GNs_na_Lista_Quente"]
    )

    total_hot = len(hot)
    total_anom_gn = len(anomaly_summary)
    total_both = len(both)
    total_anom_records = len(anom)
    # Indicadores de volume consideram todas as linhas da base original,
    # inclusive registros sem GN identificável; o cruzamento por GN considera
    # apenas as linhas em que foi possível extrair a chave.
    status_counts = anom_raw[a_status].astype(str).str.strip().str.lower().value_counts() if a_status else {}
    summary = pd.DataFrame([
        ["Transformadores na Lista Quente", total_hot],
        ["GN distintas identificadas na Anomalias", total_anom_gn],
        ["GN presentes nas duas listas", total_both],
        ["GN da Lista Quente sem correspondência na Anomalias", total_hot - total_both],
        ["GN da Lista Quente com anomalia pendente", int((both["Anomalias_Pendentes"].fillna(0).astype(float) > 0).sum()) if len(both) else 0],
        ["GN coincidentes com DEC em Transgressão", int((both["DEC_Transgressao"] == "SIM").sum()) if len(both) else 0],
        ["Registros de anomalia (linhas)", len(anom_raw)],
        ["Registros de anomalia pendentes", int(status_counts.get("pendente", 0))],
        ["Registros de anomalia concluídos", int(status_counts.get("concluído", 0) + status_counts.get("concluido", 0))],
        ["Registros de anomalia programados", int(status_counts.get("programado", 0) + status_counts.get("programada", 0))],
        ["GN pendentes de abertura de miscelânea (Lista Anomalias)", total_anom_gn],
        ["GN da Lista Anomalias também na Lista Quente", total_both],
    ], columns=["Indicador", "Resultado"])

    criteria = pd.DataFrame({"Observação": [
        f"Cruzamento gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}, utilizando a GN como chave.",
        "Na planilha Anomalias, a GN é identificada prioritariamente pela coluna Circuito; quando não encontrada, é extraída da coluna Observação.",
        "Todas as GN presentes na base Anomalias são consideradas já inspecionadas e PENDENTES DE ABERTURA DE MISCELÂNEA, conforme regra de gestão informada.",
        "O campo Status da base Anomalias descreve o status do registro/anomalia e NÃO altera o status da abertura de miscelânea.",
        "O DEC e o Status Conjunto são indicadores do conjunto, não do transformador individual. GN associada a conjunto em transgressão não significa DEC individual fora da meta.",
        "A quantidade de anomalias representa linhas/registros da base, não necessariamente defeitos distintos.",
        "A indicação de prioridade é uma triagem gerencial baseada nos dados disponíveis e não substitui validação técnica/operacional.",
    ]})

    # Ordena por ranking da Lista Quente (quando existe) e, depois, por reincidência.
    if "Prioridade" in complete.columns:
        complete["_p"] = pd.to_numeric(complete["Prioridade"], errors="coerce")
        complete = complete.sort_values("_p", na_position="last").drop(columns="_p")
        top15 = complete.head(15).copy()
        both = complete[complete["Consta_na_Anomalias"] == "SIM"].copy()
        dec_priority = complete[complete["DEC_Transgressao"] == "SIM"].copy()

    return {
        "Pend. Miscelânea": pending,
        "Resumo": summary,
        "Top 15": top15,
        "Prioridade DEC": dec_priority,
        "Cruzamento completo": complete,
        "GN nas duas listas": both,
        "Anomalias por GN": anomaly_by_gn,
        "Só Anomalias": only_anomaly,
        "DEC Transgressão": dec_transgression,
        "Critérios e ressalvas": criteria,
    }


def format_workbook(sheets):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=sheet_name[:31], index=False)

        wb = writer.book
        header_fill = PatternFill("solid", fgColor="17365D")
        header_font = Font(color="FFFFFF", bold=True)
        thin_gray = Side(style="thin", color="D9E2F3")

        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            ws.sheet_view.showGridLines = False
            ws.row_dimensions[1].height = 32
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = Border(bottom=thin_gray)

            for col_idx, cells in enumerate(ws.columns, start=1):
                values = [len(str(c.value)) if c.value is not None else 0 for c in list(cells)[:250]]
                width = min(max(max(values, default=10) + 2, 12), 42)
                ws.column_dimensions[get_column_letter(col_idx)].width = width

            # Realça pendências de abertura de miscelânea.
            headers = {cell.value: cell.column for cell in ws[1]}
            status_col = headers.get("Status Miscelânea")
            if status_col:
                for row in range(2, ws.max_row + 1):
                    cell = ws.cell(row, status_col)
                    cell.fill = PatternFill("solid", fgColor="FCE4D6")
                    cell.font = Font(color="C00000", bold=True)

            # Destaque de DEC em transgressão.
            for header in ("Status DEC Conjunto", "Status_DEC_Conjunto"):
                col = headers.get(header)
                if col:
                    for row in range(2, ws.max_row + 1):
                        cell = ws.cell(row, col)
                        if "transgress" in str(cell.value).lower():
                            cell.fill = PatternFill("solid", fgColor="F4CCCC")
                            cell.font = Font(color="990000", bold=True)

    output.seek(0)
    return output


with st.sidebar:
    st.header("Como usar")
    st.markdown(
        "1. Anexe a base **Anomalias**.  \n"
        "2. Anexe a **Lista Quente**.  \n"
        "3. Clique em **Gerar cruzamento**.  \n"
        "4. Baixe o Excel consolidado."
    )
    st.info(
        "Regra aplicada: toda GN encontrada na base Anomalias é tratada como "
        "inspecionada e pendente de abertura de miscelânea."
    )

col1, col2 = st.columns(2)
with col1:
    anom_file = st.file_uploader(
        "1. Base de Anomalias (.xlsx)",
        type=["xlsx"],
        key="anom_file",
        help="Planilha com os transformadores já inspecionados e pendentes de abertura de miscelânea.",
    )
with col2:
    hot_file = st.file_uploader(
        "2. Lista Quente (.xlsx)",
        type=["xlsx"],
        key="hot_file",
        help="Ranking de reincidência dos transformadores.",
    )

if anom_file and hot_file:
    st.success("As duas bases foram carregadas. Gere o cruzamento para montar o relatório.")
elif anom_file or hot_file:
    st.warning("Falta anexar a outra planilha para realizar o cruzamento.")

if st.button("Gerar cruzamento", type="primary", disabled=not (anom_file and hot_file)):
    try:
        with st.spinner("Lendo as bases e cruzando as GN..."):
            anom_df, anom_sheet = read_excel(anom_file, ["Circuito"])
            hot_df, hot_sheet = read_excel(hot_file, ["Trafo"])
            report = build_report(anom_df, hot_df)
            excel_file = format_workbook(report)
            st.session_state["report_bytes"] = excel_file.getvalue()
            st.session_state["report_sheets"] = report
            st.session_state["report_name"] = (
                f"Cruzamento_Lista_Quente_Anomalias_DEC_{datetime.now().strftime('%d-%m-%Y')}.xlsx"
            )
            st.session_state["source_sheets"] = (anom_sheet, hot_sheet)
        st.success("Cruzamento concluído!")
    except Exception as exc:
        st.error(f"Não foi possível gerar o cruzamento: {exc}")
        st.exception(exc)

if "report_bytes" in st.session_state:
    report = st.session_state["report_sheets"]
    summary = report["Resumo"].set_index("Indicador")["Resultado"].to_dict()
    st.divider()
    st.subheader("Resultado do cruzamento")

    metric_cols = st.columns(4)
    metrics = [
        ("GN na Lista Quente", summary.get("Transformadores na Lista Quente", 0)),
        ("GN na base Anomalias", summary.get("GN distintas identificadas na Anomalias", 0)),
        ("GN nas duas bases", summary.get("GN presentes nas duas listas", 0)),
        ("Pendentes de miscelânea", summary.get("GN pendentes de abertura de miscelânea (Lista Anomalias)", 0)),
    ]
    for col, (label, value) in zip(metric_cols, metrics):
        with col:
            with st.container(border=True):
                st.metric(label, f"{int(value):,}".replace(",", "."))

    st.download_button(
        label="⬇️ Baixar relatório Excel completo",
        data=st.session_state["report_bytes"],
        file_name=st.session_state["report_name"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    tab_pend, tab_top, tab_dec, tab_all = st.tabs(
        ["📋 Pendentes de miscelânea", "🔥 Top 15", "⚠️ Prioridade DEC", "🗂️ Todas as abas"]
    )
    with tab_pend:
        st.caption(
            "Todas as GN da base Anomalias aparecem aqui como INSPECIONADAS e "
            "PENDENTES DE ABERTURA, independentemente do status individual dos registros."
        )
        st.dataframe(report["Pend. Miscelânea"].head(50), use_container_width=True, hide_index=True)
    with tab_top:
        st.dataframe(report["Top 15"], use_container_width=True, hide_index=True)
    with tab_dec:
        st.caption("GN associadas a conjuntos em transgressão de DEC (indicador do conjunto, não do transformador).")
        st.dataframe(report["Prioridade DEC"].head(50), use_container_width=True, hide_index=True)
    with tab_all:
        st.write("Abas incluídas no Excel: " + ", ".join(report.keys()))
        aba = st.selectbox("Visualizar aba", list(report.keys()))
        st.dataframe(report[aba].head(100), use_container_width=True, hide_index=True)

st.divider()
st.caption("Aplicativo de apoio à gestão operacional. Confira as bases e valide os critérios técnicos antes de direcionar serviços.")
