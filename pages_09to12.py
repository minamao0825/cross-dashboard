# —— 09-行业分类统计 ——
elif page == "09-行业分类统计":
    st.markdown(f"### 📊 09 · 行业分类统计分析 · {q_label}")
    st.caption("按公司类型统计偿付能力核心指标的分布情况")

    cat_stats = []
    for cat, cos in COMP_CATS.items():
        cos_in = [c for c in cos if c in df["公司"].tolist()]
        if not cos_in:
            continue
        df_cat = df[df["公司"].isin(cos_in)]
        row = {"公司类型": cat, "公司数量": len(df_cat)}
        if RATIO_COMP in df_cat.columns:
            row["综合充足率均值"] = df_cat[RATIO_COMP].mean()
            row["综合充足率中位数"] = df_cat[RATIO_COMP].median()
            row["综合充足率最小值"] = df_cat[RATIO_COMP].min()
            row["达标数量"] = (df_cat[RATIO_COMP] >= THR_COMP_PASS).sum()
        if "实际资本" in df_cat.columns:
            row["实际资本合计（亿元）"] = df_cat["实际资本"].sum() / 10000
        cat_stats.append(row)

    if not cat_stats:
        st.warning("暂无分类统计数据。")
    else:
        st.markdown("##### 各类型公司核心指标统计")
        df_stats = pd.DataFrame(cat_stats)
        df_stats = df_stats.round(4)
        # 格式化显示
        disp_cols = [c for c in df_stats.columns if c != "公司类型"]
        fmt_df = df_stats[["公司类型"] + list(disp_cols)].copy()
        if "综合充足率均值" in fmt_df.columns:
            for col in ["综合充足率均值","综合充足率中位数","综合充足率最小值"]:
                if col in fmt_df.columns:
                    fmt_df[col] = fmt_df[col].apply(lambda v: pct(v) if pd.notna(v) else "—")
        if "实际资本合计（亿元）" in fmt_df.columns:
            fmt_df["实际资本合计（亿元）"] = fmt_df["实际资本合计（亿元）"].apply(
                lambda v: f"{v:.2f}" if pd.notna(v) else "—")
        st.dataframe(fmt_df, use_container_width=True, hide_index=True)

        st.divider()
        # 综合充足率 by category 柱状图
        if "综合充足率中位数" in df_stats.columns:
            st.markdown("##### 各类型综合偿付能力充足率中位数")
            fig_cat = px.bar(
                df_stats.sort_values("综合充足率中位数", ascending=False),
                x="公司类型", y="综合充足率中位数",
                color_discrete_sequence=["#185FA5"],
                text=[pct(v) for v in df_stats.sort_values("综合充足率中位数", ascending=False)["综合充足率中位数"]]
            )
            fig_cat.add_hline(y=100, line_dash="dash", line_color="#b91c1c")
            fig_cat.add_hline(y=150, line_dash="dot",  line_color="#EF9F27")
            fig_cat.update_layout(height=380, margin=dict(l=0,r=0,t=10,b=80),
                                 plot_bgcolor="white", paper_bgcolor="white", yaxis_title="综合充足率(%)")
            st.plotly_chart(fig_cat, use_container_width=True)

        # 各类型公司数量
        st.markdown("##### 各类型公司数量分布")
        fig_n = px.pie(df_stats, names="公司类型", values="公司数量",
                        color_discrete_sequence=px.colors.qualitative.Set2)
        fig_n.update_layout(height=360, margin=dict(l=0,r=0,t=10,b=0))
        st.plotly_chart(fig_n, use_container_width=True)

        # 详细：每类公司充足率排名
        st.divider()
        st.markdown("##### 各类型公司综合充足率明细")
        for cat, cos in COMP_CATS.items():
            cos_in = [c for c in cos if c in df["公司"].tolist()]
            if not cos_in:
                continue
            df_cat = df[df["公司"].isin(cos_in)].copy()
            if RATIO_COMP not in df_cat.columns:
                continue
            df_cat = df_cat.sort_values(RATIO_COMP, ascending=False)[["公司", RATIO_COMP, COMP_PCT]]
            df_cat[COMP_PCT] = df_cat[COMP_PCT].apply(lambda v: f"{v:.1f}%")
            df_cat[RATIO_COMP] = df_cat[RATIO_COMP].apply(lambda v: pct(v))
            with st.expander(f"{cat}（{len(df_cat)}家）"):
                st.dataframe(df_cat, use_container_width=True, hide_index=True)

# —— 10-公司级对标报告 ——
elif page == "10-公司级对标报告":
    st.markdown(f"### 📋 10 · 公司级对标报告 · {q_label}")
    if not target_co:
        st.info("请在左侧选择「目标公司」以生成对标报告。")
        st.stop()

    st.caption(f"目标公司：**{target_co}**")

    col_btn1, col_btn2 = st.columns([1,2])
    with col_btn1:
        if st.button("📝 生成 Word 报告", use_container_width=True, type="primary"):
            with st.spinner("正在生成报告，请稍候..."):
                doc = generate_company_report(df, target_co, df_all, q_label)
                buf = io.BytesIO()
                doc.save(buf)
                buf.seek(0)
                st.session_state.report_buf = buf.read()
                st.session_state.report_co = target_co
                st.success("✅ 报告生成成功！")

    if "report_buf" in st.session_state and st.session_state.get("report_co") == target_co:
        buf = io.BytesIO(st.session_state.report_buf)
        st.download_button(
            "⬇ 下载 Word 报告",
            data=buf,
            file_name=f"CROSS_{target_co}_{q_label}_对标报告.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

    st.divider()

    # 在线预览
    st.markdown("##### 📋 报告预览（在线）")
    if target_co in df["公司"].tolist():
        trow = df[df["公司"]==target_co].iloc[0]
        cat = trow.get("分类", "未知")

        st.markdown(f"**公司名称：** {target_co}  \n**公司分类：** {cat}")

        # 核心指标卡片
        c1,c2,c3,c4 = st.columns(4)
        if RATIO_COMP in df.columns:
            v = trow[RATIO_COMP]
            label, color = risk_info(v)
            c1.metric("综合偿付能力充足率", pct(v), delta=label)
        if RATIO_CORE in df.columns:
            c2.metric("核心偿付能力充足率", pct(trow[RATIO_CORE]))
        if "实际资本" in df.columns:
            c3.metric("实际资本", fmt_wan(trow["实际资本"]))
        if "最低资本" in df.columns:
            c4.metric("最低资本", fmt_wan(trow["最低资本"]))

        # 同行对比
        st.divider()
        st.markdown("##### 同行对比")
        peers = []
        for ccat, cos in COMP_CATS.items():
            if target_co in cos:
                peers = [c for c in cos if c in df["公司"].tolist()]
                break
        if peers:
            cmp_cos = [target_co] + [c for c in peers if c != target_co][:6]
            df_cmp = df[df["公司"].isin(cmp_cos)].copy()
            if RATIO_COMP in df_cmp.columns:
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
                fig_cmp.add_hline(y=100, line_dash="dash", line_color="#b91c1c")
                fig_cmp.add_hline(y=150, line_dash="dot",  line_color="#EF9F27")
                fig_cmp.update_layout(height=360, margin=dict(l=0,r=0,t=10,b=80),
                                     plot_bgcolor="white", paper_bgcolor="white")
                st.plotly_chart(fig_cmp, use_container_width=True)

        # AI 分析按钮
        st.divider()
        st.markdown("##### 🤖 AI 辅助分析")
        if st.button("生成 AI 分析报告摘要", use_container_width=True):
            if not st.session_state.ai_api_key:
                st.warning("请先在左侧边栏配置 AI API Key。")
            else:
                prompt = f"""
                请作为资深保险精算分析师，对以下公司偿付能力数据进行分析：

                公司名称：{target_co}
                综合偿付能力充足率：{pct(trow.get(RATIO_COMP, None))}
                核心偿付能力充足率：{pct(trow.get(RATIO_CORE, None))}
                实际资本：{fmt_wan(trow.get('实际资本', None))}
                最低资本：{fmt_wan(trow.get('最低资本', None))}

                请从以下角度进行分析（500字以内）：
                1. 偿付能力充足率评价（充足/预警/不达标）
                2. 资本规模在行业中的位置
                3. 风险提示与关注事项
                4. 建议措施
                """
                with st.spinner("AI 正在分析..."):
                    ai_result = call_ai_api(st.session_state.ai_api_key, st.session_state.ai_model, prompt)
                    st.info(ai_result)

# —— 11-AI辅助分析 ——
elif page == "11-AI辅助分析":
    st.markdown(f"### 🤖 11 · AI 辅助分析 · {q_label}")
    st.caption("基于阿里云百炼（通义千问）大模型的智能分析")

    # API 配置状态
    if st.session_state.ai_api_key:
        st.success(f"✅ AI API Key 已配置（模型：{st.session_state.ai_model}）")
    else:
        st.warning("⚠️ 尚未配置 AI API Key，请在左侧边栏输入。")
        st.markdown("""
        **如何获取阿里云百炼 API Key？**
        1. 访问 [阿里云百炼控制台](https://bailian.console.aliyun.com/)
        2. 进入「API Key 管理」
        3. 创建新 API Key 并复制
        """)

    st.divider()

    # 分析类型选择
    analysis_type = st.selectbox(
        "选择分析类型",
        ["行业整体分析", "单公司深度分析", "同业对比分析", "风险预警分析"]
    )

    if analysis_type == "行业整体分析":
        st.markdown("##### 行业整体偿付能力分析")
        if st.button("开始 AI 分析", use_container_width=True, type="primary"):
            if not st.session_state.ai_api_key:
                st.error("请先配置 AI API Key。")
            else:
                n_total = len(df)
                n_pass = (df[RATIO_COMP] >= THR_COMP_WARN).sum() if RATIO_COMP in df.columns else 0
                n_warn = ((df[RATIO_COMP]>=THR_COMP_PASS)&(df[RATIO_COMP]<THR_COMP_WARN)).sum() if RATIO_COMP in df.columns else 0
                n_fail = (df[RATIO_COMP] < THR_COMP_PASS).sum() if RATIO_COMP in df.columns else 0
                med_comp = df[RATIO_COMP].median() if RATIO_COMP in df.columns else "N/A"
                prompt = f"""
                请作为资深保险精算分析师，对中国人身险行业{ q_label }偿付能力数据进行整体分析：

                行业数据概览：
                - 纳入公司数量：{n_total} 家
                - 综合充足率 ≥ 150%（充足）：{n_pass} 家（{n_pass/n_total*100:.1f}%)
                - 综合充足率 100-150%（预警）：{n_warn} 家（{n_warn/n_total*100:.1f}%)
                - 综合充足率 < 100%（不达标）：{n_fail} 家（{n_fail/n_total*100:.1f}%）
                - 行业综合充足率中位数：{pct(med_comp)}

                请从以下角度进行分析（800字以内）：
                1. 行业整体偿付能力充足率评价
                2. 风险集中度与系统性风险分析
                3. 与上一季度的变化趋势（如有）
                4. 监管建议与行业展望
                """
                with st.spinner("AI 正在分析行业数据..."):
                    ai_result = call_ai_api(st.session_state.ai_api_key, st.session_state.ai_model, prompt)
                    st.markdown("#### 🤖 AI 分析结果")
                    st.info(ai_result)

    elif analysis_type == "单公司深度分析":
        st.markdown("##### 单公司深度分析")
        if not target_co:
            st.info("请在左侧选择「目标公司」。")
        else:
            if st.button("开始 AI 深度分析", use_container_width=True, type="primary"):
                if not st.session_state.ai_api_key:
                    st.error("请先配置 AI API Key。")
                else:
                    trow = df[df["公司"]==target_co].iloc[0]
                    # 收集所有可用指标
                    info_lines = []
                    for col in df.columns:
                        if col in ["分类","公司"]: continue
                        v = trow.get(col, None)
                        if pd.notna(v):
                            if col in [RATIO_COMP, RATIO_CORE]:
                                info_lines.append(f"{col}：{pct(v)}")
                            elif isinstance(v, (int,float)):
                                info_lines.append(f"{col}：{v:,.2f}")
                    prompt = f"""
                    请作为资深保险精算分析师，对以下公司偿付能力数据进行深度分析：

                    公司名称：{target_co}
                    {chr(10).join(info_lines)}

                    请从以下角度进行深度分析（1000字以内）：
                    1. 偿付能力充足率评价与趋势判断
                    2. 资本结构与质量分析
                    3. 风险最低资本结构分析
                    4. 与同业相比的优势与短板
                    5. 风险提示与监管建议
                    """
                    with st.spinner("AI 正在进行深度分析..."):
                        ai_result = call_ai_api(st.session_state.ai_api_key, st.session_state.ai_model, prompt)
                        st.markdown("#### 🤖 AI 深度分析结果")
                        st.info(ai_result)

    elif analysis_type == "同业对比分析":
        st.markdown("##### 同业对比分析")
        if not target_co:
            st.info("请在左侧选择「目标公司」。")
        else:
            peers = []
            for ccat, cos in COMP_CATS.items():
                if target_co in cos:
                    peers = [c for c in cos if c in df["公司"].tolist()]
                    break
            if peers:
                st.caption(f"同行公司：{', '.join(peers[:6])}")
            if st.button("开始同业对比分析", use_container_width=True, type="primary"):
                if not st.session_state.ai_api_key:
                    st.error("请先配置 AI API Key。")
                else:
                    trow = df[df["公司"]==target_co].iloc[0]
                    peer_rows = []
                    for co in peers[:6]:
                        if co in df["公司"].tolist():
                            r = df[df["公司"]==co].iloc[0]
                            peer_rows.append(f"{co}：综合充足率 {pct(r.get(RATIO_COMP, None))}，实际资本 {fmt_wan(r.get('实际资本', None))}")
                    prompt = f"""
                    请作为资深保险精算分析师，对以下目标公司与同业进行对比分析：

                    目标公司：{target_co}
                    目标公司综合充足率：{pct(trow.get(RATIO_COMP, None))}
                    目标公司实际资本：{fmt_wan(trow.get('实际资本', None))}

                    同行业参考：
                    {chr(10).join(peer_rows)}

                    请从以下角度进行对比分析（800字以内）：
                    1. 目标公司与同业在偿付能力充足率上的对比
                    2. 资本规模的同业位置
                    3. 目标公司的相对优势与风险短板
                    4. 改进建议
                    """
                    with st.spinner("AI 正在进行同业对比分析..."):
                        ai_result = call_ai_api(st.session_state.ai_api_key, st.session_state.ai_model, prompt)
                        st.markdown("#### 🤖 AI 同业对比分析结果")
                        st.info(ai_result)

    elif analysis_type == "风险预警分析":
        st.markdown("##### 风险预警分析")
        if st.button("开始风险预警分析", use_container_width=True, type="primary"):
            if not st.session_state.ai_api_key:
                st.error("请先配置 AI API Key。")
            else:
                df_risk = df.copy()
                if RATIO_COMP in df_risk.columns:
                    df_risk["级别"] = df_risk[RATIO_COMP].apply(lambda v: 2 if v>=THR_COMP_WARN else (1 if v>=THR_COMP_PASS else 0))
                    n_p = (df_risk["级别"]==2).sum()
                    n_w = (df_risk["级别"]==1).sum()
                    n_f = (df_risk["级别"]==0).sum()
                    fail_list = df_risk[df_risk["级别"]==0]["公司"].tolist()
                else:
                    n_p, n_w, n_f = 0, 0, 0
                    fail_list = []
                prompt = f"""
                请作为资深保险精算分析师，对中国人身险行业{ q_label }偿付能力风险进行预警分析：

                风险统计数据：
                - 充足（≥150%）：{n_p} 家
                - 预警（100-150%）：{n_w} 家
                - 不达标（<100%）：{n_f} 家
                - 不达标公司名单：{', '.join(fail_list) if fail_list else '无'}

                请从以下角度进行风险预警分析（800字以内）：
                1. 行业整体风险等级评估
                2. 不达标/预警公司的主要风险成因分析
                3. 系统性风险识别
                4. 监管建议与风险缓释措施
                """
                with st.spinner("AI 正在进行风险预警分析..."):
                    ai_result = call_ai_api(st.session_state.ai_api_key, st.session_state.ai_model, prompt)
                    st.markdown("#### 🤖 AI 风险预警分析结果")
                    st.info(ai_result)

    st.divider()
    st.caption("⚠️ AI 生成内容仅供参考，不构成专业精算意见。")

# —— 12-原始数据 ——
elif page == "12-原始数据":
    st.markdown(f"### 📋 12 · 原始数据浏览 · {q_label}")

    if ind_choices == "ALL":
        show_cols = [c for c in df.columns if c not in ("分类",)]
    elif isinstance(ind_choices, list) and len(ind_choices) > 0:
        show_cols = ["公司"] + [c for c in ind_choices if c in df.columns]
    else:
        show_cols = ["公司", RATIO_COMP, RATIO_CORE, "实际资本","最低资本"]

    search = st.text_input("搜索公司名称", "")
    df_s = df[df["公司"].str.contains(search, na=False)] if search else df.copy()

    df_disp = df_s[show_cols].copy() if all(c in df_s.columns for c in show_cols) else df_s.copy()
    for col in [RATIO_COMP, RATIO_CORE]:
        if col in df_disp.columns:
            df_disp[col] = df_disp[col].apply(lambda v: pct(v) if pd.notna(v) else "—")
    for col in df_disp.columns:
        if col not in ["公司",RATIO_COMP,RATIO_CORE] and pd.api.types.is_numeric_dtype(df_disp[col]):
            df_disp[col] = df_disp[col].apply(lambda v: f"{v:,.0f}" if pd.notna(v) and abs(float(v))>=10 else (f"{v:.4f}" if pd.notna(v) else "—"))

    st.markdown(f"共 **{len(df_s)}** 家公司 · **{len(df_disp.columns)-1}** 个指标")
    st.dataframe(df_disp, use_container_width=True, hide_index=True, height=600)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df_s[show_cols].to_excel(w, index=False, sheet_name="CROSS数据")
    buf.seek(0)
    st.download_button("⬇ 下载当前数据（Excel）", data=buf,
                        file_name=f"CROSS_{q_label}_导出.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
