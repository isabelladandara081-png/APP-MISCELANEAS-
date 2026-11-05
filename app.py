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
        col.metric(label, f"{int(value):,}".replace(",", "."))

    st.markdown("#### Pendentes de abertura de miscelânea")
    st.caption(
        "Todas as GN da base Anomalias aparecem nesta aba como INSPECIONADAS e "
        "PENDENTES DE ABERTURA, independentemente do status individual dos registros."
    )
    st.dataframe(report["Pend. Miscelânea"].head(30), use_container_width=True, hide_index=True)
    st.download_button(
        label="⬇️ Baixar relatório Excel completo",
        data=st.session_state["report_bytes"],
        file_name=st.session_state["report_name"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    with st.expander("Ver outras abas incluídas no Excel"):
        st.write(", ".join(report.keys()))
