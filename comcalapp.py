st.subheader(f"Admin Dashboard - CRD Summary ({CURRENT_YEAR})")
    st.caption("Click a CRD name to open the detailed commission view.")

    rows_html = []
    for crd in CRD_REPS:
        rep_display = match_rep_name(df_all, crd)
        rep_df = build_rep_df(df_all, rep_display)
        summary = compute_summary(rep_df, crd)


        href = f"?crd={quote_plus(crd)}"
        rows_html.append(
            "<tr>"
            f"<td><a class='crd-anchor' href='{href}'>{crd}</a></td>"
            f"<td>{fmt_money(summary['quota'])}</td>"
            f"<td>{fmt_money(summary['total_quota_credit'])}</td>"
            f"<td>{fmt_money(summary['eligible_comm'])}</td>"
            f"<td>{fmt_pct(summary['attainment'])}</td>"
            f"<td>{summary['payout_status']}</td>"
            "</tr>"
        )

    table_html = (
        "<table class='dashboard-table'>"
        "<thead><tr>"
        "<th>CRD Name</th><th>Annual Quota Goal</th><th>Total Quota Credit</th>"
        "<th>Eligible Comm YTD</th><th>YTD Attainment %</th><th>Payout Status</th>"
        "</tr></thead>"
        "<tbody>" + "".join(rows_html) + "</tbody></table>"
    )
    st.markdown(table_html, unsafe_allow_html=True)
