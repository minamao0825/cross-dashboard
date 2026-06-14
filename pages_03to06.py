
# —— 03-最低资本数据对比分析 ——
elif page == "03-最低资本数据对比分析":
    st.markdown(f"### ⚠️ 03 · 最低资本数据对比分析 · {q_label}")

    st.markdown("""
    <div class="metric-explain">
        <strong>指标含义：</strong><br>
        • <strong>最低资本</strong> 是监管要求的资本下限，由保险风险、市场风险、信用风险等构成。<br>
        • <strong>保险风险最低资本/认可负债</strong> 体现公司保险风险压力因子，比率越高说明负债对应的保险风险水平越高。<br>
        • <strong>市场风险</strong> 由利率风险、权益价格风险、房地产价格风险、汇率风险等构成。<br>
        • <strong>信用风险</strong> 由利差风险、交易对手违约风险等构成。
    </div>
    """, unsafe_allow_html=True)

    # 最低资本 箱线图
    if "最低资本" in df.columns:
        st.markdown("##### 📊 最低资本分布（箱线图）")
        fig_box = go.Figure()
        fig_box.add_trace(go.Box(
            y=df["最低资本"].dropna().tolist(),
            name="最低资本",
            boxmean="sd",
            marker_color="#185FA5",
            boxpoints="outliers"
        ))
        fig_box.update_layout(
            height=400,
            margin=dict(l=0,r=0,t=20,b=0),
            plot_bgcolor="white", paper_bgcolor="white",
            yaxis_title="最低资本（万元）"
        )
        st.plotly_chart(fig_box, use_container_width=True)

        # 最低资本 分布直方图
        st.markdown("##### 📊 最低资本分布直方图")
        fig_hist = px.histogram(
            df, x="最低资本", nbins=30,
            color_discrete_sequence=["#185FA5"],
            marginal="box"
        )
        fig_hist.update_layout(
            height=380,
            margin=dict(l=0,r=0,t=20,b=0),
            plot_bgcolor="white", paper_bgcolor="white",
            xaxis_title="最低资本（万元）", yaxis_title="公司数"
        )
        st.plotly_chart(fig_hist, use_container_width=True)

    st.divider()

    # 量化风险最低资本 构成饼图
    qc_col = "量化风险最低资本"
    if qc_col in df.columns:
        st.markdown("##### 📈 量化风险最低资本构成（行业合计）")

        risk_cols = [c for c in ["寿险业务保险风险最低资本合计","市场风险-最低资本合计","信用风险-最低资本合计"] if c in df.columns]
        if risk_cols:
            sums = {c.replace("最低资本合计","").replace("-",""): abs(df[c].sum()) for c in risk_cols}
            fig_pie = px.pie(
                names=list(sums.keys()),
                values=list(sums.values()),
                color_discrete_sequence=["#185FA5","#EF9F27","#1D9E75"]
            )
            fig_pie.update_layout(height=350, margin=dict(l=0,r=0,t=10,b=0))
            st.plotly_chart(fig_pie, use_container_width=True)

        # 量化风险最低资本 TOP15
        st.markdown("##### 📋 量化风险最低资本 TOP15（万元）")
        top_q = df.nlargest(15, qc_col)[["公司",qc_col]].copy()
        fig_bar = px.bar(
            top_q.sort_values(qc_col),
            x=qc_col, y="公司", orientation="h",
            color_discrete_sequence=["#185FA5"]
        )
        fig_bar.update_layout(
            height=420,
            margin=dict(l=0,r=0,t=10,b=0),
            plot_bgcolor="white", paper_bgcolor="white"
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    st.divider()

    # 保险风险分析
    ins_col = "寿险业务保险风险最低资本合计"
    if ins_col in df.columns:
        st.markdown("##### 📊 寿险业务保险风险最低资本分布")
        fig_ins = go.Figure()
        fig_ins.add_trace(go.Box(
            y=df[ins_col].dropna().tolist(),
            name="保险风险最低资本",
            boxmean="sd",
            marker_color="#378ADD",
            boxpoints="outliers"
        ))
        fig_ins.update_layout(
            height=380,
            margin=dict(l=0,r=0,t=20,b=0),
            plot_bgcolor="white", paper_bgcolor="white",
            yaxis_title="保险风险最低资本（万元）"
        )
        st.plotly_chart(fig_ins, use_container_width=True)

        # 保险风险压力因子：保险风险/认可负债
        if "认可负债" in df.columns:
            st.markdown("##### 📈 保险风险压力因子（保险风险/认可负债）")
            st.markdown("""
            <div class="metric-explain">
                <strong>指标含义：</strong>以该指标体现公司保险风险最低资本相对认可负债的压力因子值，比率越高说明负债对应的保险风险水平越高；银行系公司和中型公司的寿险保险风险压力因子明显低于其他类型公司的水平。
            </div>
            """, unsafe_allow_html=True)

            df_risk = df.dropna(subset=[ins_col, "认可负债"]).copy()
            df_risk["保险风险压力因子"] = df_risk[ins_col] / df_risk["认可负债"]

            fig_rf = go.Figure()
            fig_rf.add_trace(go.Box(
                y=df_risk["保险风险压力因子"].tolist(),
                name="保险风险压力因子",
                boxmean="sd",
                marker_color="#E24B4A",
                boxpoints="outliers"
            ))
            fig_rf.update_layout(
                height=380,
                margin=dict(l=0,r=0,t=20,b=0),
                plot_bgcolor="white", paper_bgcolor="white",
                yaxis_title="保险风险/认可负债"
            )
            st.plotly_chart(fig_rf, use_container_width=True)

            # 按公司分类对比
            st.markdown("##### 📊 各公司分类 · 保险风险压力因子")
            rf_stats = []
            for cat, cos in COMP_CATS.items():
                df_cat = df_risk[df_risk["公司"].isin(cos)]
                if not df_cat.empty:
                    vals = df_cat["保险风险压力因子"]
                    rf_stats.append({
                        "公司分类": cat,
                        "中位数": f"{vals.median():.2%}",
                        "上四分位": f"{vals.quantile(0.75):.2%}",
                        "下四分位": f"{vals.quantile(0.25):.2%}",
                        "公司数": len(vals)
                    })
            if rf_stats:
                st.dataframe(pd.DataFrame(rf_stats), use_container_width=True, hide_index=True)

    st.divider()

    # 市场风险分析
    mkt_col = "市场风险-最低资本合计"
    if mkt_col in df.columns:
        st.markdown("##### 📉 市场风险最低资本分布")
        st.caption("市场风险由利率风险、权益价格风险、房地产价格风险、境外固收/权益价格风险、汇率风险、风险分散效应构成。利率风险、权益价格风险是市场风险的主要组成部分。")

        fig_mkt = go.Figure()
        fig_mkt.add_trace(go.Box(
            y=df[mkt_col].dropna().tolist(),
            name="市场风险最低资本",
            boxmean="sd",
            marker_color="#EF9F27",
            boxpoints="outliers"
        ))
        fig_mkt.update_layout(
            height=380,
            margin=dict(l=0,r=0,t=20,b=0),
            plot_bgcolor="white", paper_bgcolor="white",
            yaxis_title="市场风险最低资本（万元）"
        )
        st.plotly_chart(fig_mkt, use_container_width=True)

        # 市场风险细分
        mkt_sub = [c for c in df.columns if "市场风险-" in c and "分散" not in c and "合计" not in c and df[c].notna().any()]
        if mkt_sub:
            st.markdown("##### 📊 市场风险细分（全行业合计，万元）")
            mkt_sums = {c.replace("市场风险-","").replace("最低资本",""): abs(df[c].sum()) for c in mkt_sub}
            fig_mkt_sub = px.bar(
                x=list(mkt_sums.keys()),
                y=list(mkt_sums.values()),
                color_discrete_sequence=["#EF9F27"]
            )
            fig_mkt_sub.update_layout(
                height=320,
                margin=dict(l=0,r=0,t=10,b=80),
                plot_bgcolor="white", paper_bgcolor="white",
                xaxis_tickangle=-30,
                yaxis_title="万元"
            )
            st.plotly_chart(fig_mkt_sub, use_container_width=True)

    st.divider()

    # 信用风险分析
    cr_col = "信用风险-最低资本合计"
    if cr_col in df.columns:
        st.markdown("##### 🏦 信用风险最低资本分布")
        st.caption("信用风险由利差风险、交易对手违约风险、风险分散效应组成。")

        fig_cr = go.Figure()
        fig_cr.add_trace(go.Box(
            y=df[cr_col].dropna().tolist(),
            name="信用风险最低资本",
            boxmean="sd",
            marker_color="#1D9E75",
            boxpoints="outliers"
        ))
        fig_cr.update_layout(
            height=380,
            margin=dict(l=0,r=0,t=20,b=0),
            plot_bgcolor="white", paper_bgcolor="white",
            yaxis_title="信用风险最低资本（万元）"
        )
        st.plotly_chart(fig_cr, use_container_width=True)

        # 信用风险细分
        cr_sub = [c for c in df.columns if "信用风险-" in c and "分散" not in c and "合计" not in c and df[c].notna().any()]
        if cr_sub:
            st.markdown("##### 📊 信用风险细分（全行业合计，万元）")
            cr_sums = {c.replace("信用风险-","").replace("最低资本",""): abs(df[c].sum()) for c in cr_sub}
            fig_cr_sub = px.bar(
                x=list(cr_sums.keys()),
                y=list(cr_sums.values()),
                color_discrete_sequence=["#1D9E75"]
            )
            fig_cr_sub.update_layout(
                height=320,
                margin=dict(l=0,r=0,t=10,b=80),
                plot_bgcolor="white", paper_bgcolor="white",
                xaxis_tickangle=-30,
                yaxis_title="万元"
            )
            st.plotly_chart(fig_cr_sub, use_container_width=True)


# —— 04-增资发债信息统计 ——
elif page == "04-增资发债信息统计":
    st.markdown(f"### 📋 04 · 增资发债信息统计 · {q_label}")

    st.markdown("""
    <div class="metric-explain">
        <strong>说明：</strong>本页面统计各公司在当季度的增资/发债情况，以及综合偿付能力充足率的变动情况。
    </div>
    """, unsafe_allow_html=True)

    # 从Excel中读取增资发债信息
    # 尝试从CROSS汇总表.xlsx的其他Sheet中读取
    try:
        xls = pd.ExcelFile(FILE_PATH)
        if "增资发债" in xls.sheet_names:
            df_cap = pd.read_excel(FILE_PATH, sheet_name="增资发债")
            st.markdown("##### 📋 增资发债明细表")
            st.dataframe(df_cap, use_container_width=True, hide_index=True)

            # 下载按钮
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                df_cap.to_excel(w, index=False, sheet_name="增资发债信息")
            buf.seek(0)
            st.download_button(
                "⬇ 下载增资发债信息（Excel）",
                data=buf,
                file_name=f"增资发债信息_{q_label}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("未在Excel中找到「增资发债」Sheet，请手动添加。")
    except Exception as e:
        st.warning(f"读取增资发债信息失败：{e}")

    st.divider()

    # 手动录入增资发债信息
    st.markdown("##### ✏️ 手动录入增资发债信息")
    st.caption("提示：如果Excel中没有增资发债数据，可以在此处手动录入。")

    with st.form("capital_form"):
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            inp_co  = st.text_input("公司名称")
        with col2:
            inp_type = st.selectbox("类型", ["增资", "发债"])
        with col3:
            inp_amt  = st.text_input("金额（亿元）")
        with col4:
            inp_eff  = st.text_input("综合充足率影响")

        submitted = st.form_submit_button("添加记录")
        if submitted:
            st.success(f"已添加：{inp_co} {inp_type} {inp_amt}亿元，综合充足率影响：{inp_eff}")
            st.info("提示：手动录入的数据仅在当前会话有效，如需永久保存，请更新Excel文件。")


# —— 05-公司级对标报告 ——
elif page == "05-公司级对标报告":
    st.markdown(f"### 📋 05 · 公司级对标报告 · {q_label}")

    if not target_co:
        st.info("请在左侧选择「目标公司」以生成对标报告。")
        st.stop()

    st.markdown(f"##### 🎯 目标公司：**{target_co}**")

    # 检查目标公司是否存在
    if target_co not in df["公司"].tolist():
        st.error(f"目标公司「{target_co}」不在数据中，请重新选择。")
        st.stop()

    trow = df[df["公司"] == target_co].iloc[0]

    # 同行公司
    same_cat = []
    for cat, cos in COMP_CATS.items():
        if target_co in cos:
            same_cat = [c for c in cos if c in df["公司"].tolist()]
            break

    # 生成报告按钮
    if st.button("📝 生成 Word 报告", use_container_width=True):
        with st.spinner("正在生成报告..."):
            try:
                doc = Document()

                # 标题
                doc.add_heading(f"CROSS 偿付能力对标报告", 0)
                doc.add_heading(f"{target_co} · {q_label}", 1)
                doc.add_paragraph(f"报告生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}")
                doc.add_paragraph()

                # 第一章：公司概况
                doc.add_heading("一、公司概况", 1)
                p1 = doc.add_paragraph()
                p1.add_run(f"公司名称：{target_co}\n")
                p1.add_run(f"公司分类：{trow.get('分类', '未知')}\n")
                if RATIO_COMP in df.columns:
                    v = trow.get(RATIO_COMP, None)
                    p1.add_run(f"综合偿付能力充足率：{pct(v) if pd.notna(v) else '—'}\n")
                if RATIO_CORE in df.columns:
                    v = trow.get(RATIO_CORE, None)
                    p1.add_run(f"核心偿付能力充足率：{pct(v) if pd.notna(v) else '—'}\n")
                doc.add_paragraph()

                # 第二章：核心指标
                doc.add_heading("二、核心偿付能力指标", 1)
                core_inds = [c for c in [RATIO_COMP, RATIO_CORE, "实际资本","最低资本","认可资产","认可负债"] if c in df.columns]
                tbl = doc.add_table(rows=1+len(core_inds), cols=3)
                tbl.style = "Light Grid Accent 1"
                hdr = tbl.rows[0].cells
                hdr[0].text = "指标"
                hdr[1].text = "目标公司"
                hdr[2].text = "行业中位数"

                for i, c in enumerate(core_inds):
                    row_cells = tbl.rows[i+1].cells
                    row_cells[0].text = c
                    v_t = trow.get(c, None)
                    row_cells[1].text = f"{v_t:,.0f}" if pd.notna(v_t) else "—"
                    v_m = df[c].median() if c in df.columns else None
                    row_cells[2].text = f"{v_m:,.0f}" if pd.notna(v_m) else "—"
                doc.add_paragraph()

                # 第三章：同行对比
                if same_cat:
                    doc.add_heading(f"三、同行对比（{len(same_cat)}家）", 1)
                    tbl2 = doc.add_table(rows=1+min(8, len(same_cat)), cols=3)
                    tbl2.style = "Light Grid Accent 1"
                    hdr2 = tbl2.rows[0].cells
                    hdr2[0].text = "排名"
                    hdr2[1].text = "公司"
                    hdr2[2].text = "综合偿付能力充足率"

                    same_df = df[df["公司"].isin(same_cat)].copy()
                    if RATIO_COMP in same_df.columns:
                        same_df = same_df.sort_values(RATIO_COMP, ascending=False)
                        for i, (_, row) in enumerate(same_df.head(8).iterrows()):
                            rcells = tbl2.rows[i+1].cells
                            rcells[0].text = str(i+1)
                            rcells[1].text = row["公司"]
                            v = row.get(RATIO_COMP, None)
                            rcells[2].text = pct(v) if pd.notna(v) else "—"
                    doc.add_paragraph()

                # 第四章：建议关注事项
                doc.add_heading("四、建议关注事项", 1)
                concerns = []
                if RATIO_COMP in df.columns:
                    v = trow.get(RATIO_COMP, None)
                    if pd.notna(v):
                        if v < THR_COMP_PASS:
                            concerns.append("⚠️ 综合偿付能力充足率低于监管下限100%，需立即关注。")
                        elif v < THR_COMP_WARN:
                            concerns.append("⚡ 综合偿付能力充足率处于预警区间（100%-150%），建议持续监测。")
                        else:
                            concerns.append("✅ 综合偿付能力充足率高于150%，资本充足水平良好。")
                if not concerns:
                    concerns.append("暂无特别关注事项。")
                for item in concerns:
                    doc.add_paragraph(item, style="List Bullet")
                doc.add_paragraph()

                # 页脚
                doc.add_paragraph()
                p5 = doc.add_paragraph()
                p5.add_run("（本报告由 CROSS 数智平台自动生成，仅供参考。）")

                # 保存
                buf = io.BytesIO()
                doc.save(buf)
                buf.seek(0)

                st.success("✅ 报告生成成功！")
                st.download_button(
                    "⬇ 下载 Word 报告",
                    data=buf,
                    file_name=f"CROSS报告_{target_co}_{q_label}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"报告生成失败：{e}")
                traceback.print_exc()

    st.divider()

    # 在线预览
    st.markdown("##### 📊 目标公司核心指标预览")
    core_inds = [c for c in [RATIO_COMP, RATIO_CORE, "实际资本","最低资本","认可资产","认可负债"] if c in df.columns]
    if core_inds:
        kcols = st.columns(min(6, len(core_inds)))
        for i, c in enumerate(core_inds[:6]):
            v = trow.get(c, None)
            disp = pct(v) if c in [RATIO_COMP,RATIO_CORE] else fmt_wan(v)
            kcols[i].metric(c, disp)

    # 同行对比图
    if same_cat:
        st.divider()
        st.markdown(f"##### 📊 同行对比：{target_co} vs 同行")

        cmp_cos = [target_co] + [c for c in same_cat if c != target_co][:7]
        df_cmp = df[df["公司"].isin(cmp_cos)].copy()

        if RATIO_COMP in df.columns:
            fig_cmp = go.Figure()
            for co in cmp_cos:
                row = df_cmp[df_cmp["公司"]==co]
                if row.empty: continue
                v = row[RATIO_COMP].values[0]
                fig_cmp.add_trace(go.Bar(
                    name=co, x=[co], y=[v*100],
                    marker_color="#185FA5" if co==target_co else "#ccc",
                    text=f"{v*100:.1f}%", textposition="outside"
                ))
            fig_cmp.add_hline(y=100, line_dash="dash", line_color="#E24B4A")
            fig_cmp.add_hline(y=150, line_dash="dot",  line_color="#EF9F27")
            fig_cmp.update_layout(height=360, margin=dict(l=0,r=0,t=10,b=80),
                                 plot_bgcolor="white", paper_bgcolor="white")
            st.plotly_chart(fig_cmp, use_container_width=True)


# —— 06-原始数据 ——
elif page == "06-原始数据":
    st.markdown(f"### 📋 06 · 原始数据浏览 · {q_label}")

    # 指标筛选
    if ind_choices == "ALL":
        show_cols = [c for c in df.columns if c not in ("分类",)]
    elif isinstance(ind_choices, list) and len(ind_choices) > 0:
        show_cols = ["公司"] + [c for c in ind_choices if c in df.columns]
    else:
        show_cols = ["公司", RATIO_COMP, RATIO_CORE, "实际资本","最低资本"]

    search = st.text_input("搜索公司名称", "")
    df_s = df[df["公司"].str.contains(search, na=False)] if search else df.copy()

    # 格式化显示
    df_disp = df_s[show_cols].copy() if all(c in df_s.columns for c in show_cols) else df_s.copy()
    for col in [RATIO_COMP, RATIO_CORE]:
        if col in df_disp.columns:
            df_disp[col] = df_disp[col].apply(lambda v: pct(v) if pd.notna(v) else "—")
    for col in df_disp.columns:
        if col not in ["公司",RATIO_COMP,RATIO_CORE] and pd.api.types.is_numeric_dtype(df_disp[col]):
            df_disp[col] = df_disp[col].apply(lambda v: f"{v:,.0f}" if pd.notna(v) and abs(float(v))>=10 else (f"{v:.4f}" if pd.notna(v) else "—"))

    st.markdown(f"共 **{len(df_s)}** 家公司 · **{len(df_disp.columns)-1}** 个指标")
    st.dataframe(df_disp, use_container_width=True, hide_index=True, height=600)

    # 下载
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df_s[show_cols].to_excel(w, index=False, sheet_name="CROSS数据")
    buf.seek(0)
    st.download_button(
        "⬇ 下载当前数据（Excel）",
        data=buf,
        file_name=f"CROSS_{q_label}_导出.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
