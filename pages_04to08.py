# —— 04-风险最低资本分析 ——
elif page == "04-风险最低资本分析":
    st.markdown(f"### ⚠️ 04 · 风险最低资本分析 · {q_label}")
    qc_col = "量化风险最低资本"
    if qc_col not in df.columns:
        st.warning("未找到「量化风险最低资本」字段。")
    else:
        risk_cols = [c for c in ["寿险业务保险风险最低资本合计","市场风险-最低资本合计","信用风险-最低资本合计"] if c in df.columns]
        if risk_cols:
            sums = {c: df[c].sum() for c in risk_cols}
            fig_pie = px.pie(names=list(sums.keys()), values=list(sums.values()),
                               color_discrete_sequence=["#185FA5","#EF9F27","#1D9E75"])
            fig_pie.update_layout(height=300, margin=dict(l=0,r=0,t=10,b=0))
            st.plotly_chart(fig_pie, use_container_width=True)
        st.markdown("##### 量化风险最低资本 TOP15（万元）")
        top_q = df.nlargest(15, qc_col)[["公司",qc_col]].copy()
        fig_q = px.bar(top_q.sort_values(qc_col), x=qc_col, y="公司", orientation="h",
                         color_discrete_sequence=["#185FA5"])
        fig_q.update_layout(height=420, margin=dict(l=0,r=0,t=10,b=0),
                             plot_bgcolor="white", paper_bgcolor="white")
        st.plotly_chart(fig_q, use_container_width=True)
        ins_cols = [c for c in df.columns if "保险风险" in c and "分散" not in c and "合计" not in c and df[c].notna().any()]
        if ins_cols:
            st.markdown("##### 寿险业务保险风险细分（全行业合计，万元）")
            ins_sums = {c: abs(df[c].sum()) for c in ins_cols}
            fig_ins = px.bar(x=list(ins_sums.keys()), y=list(ins_sums.values()),
                              color_discrete_sequence=["#378ADD"])
            fig_ins.update_layout(height=300, margin=dict(l=0,r=0,t=10,b=80),
                                   plot_bgcolor="white", paper_bgcolor="white", xaxis_tickangle=-30)
            st.plotly_chart(fig_ins, use_container_width=True)

# —— 05-市场风险分析 ——
elif page == "05-市场风险分析":
    st.markdown(f"### 📉 05 · 市场风险分析 · {q_label}")
    mkt_cols = [c for c in df.columns if "市场风险-" in c and "分散" not in c and "合计" not in c]
    mkt_total = "市场风险-最低资本合计"
    if not mkt_cols:
        st.warning("未找到市场风险细分字段。")
    else:
        st.markdown("##### 市场风险细分（全行业合计，万元）")
        mkt_sums = {c: abs(df[c].sum()) for c in mkt_cols if df[c].notna().any()}
        fig_mkt = px.bar(x=[c.replace("市场风险-","").replace("最低资本","") for c in mkt_sums.keys()],
                          y=list(mkt_sums.values()),
                          color_discrete_sequence=["#EF9F27"])
        fig_mkt.update_layout(height=320, margin=dict(l=0,r=0,t=10,b=80),
                               plot_bgcolor="white", paper_bgcolor="white", xaxis_tickangle=-30,
                               yaxis_title="万元")
        st.plotly_chart(fig_mkt, use_container_width=True)
        if mkt_total in df.columns and "最低资本" in df.columns:
            st.markdown("##### 市场风险占比（最低资本）")
            df_mkt_pct = df.dropna(subset=[mkt_total,"最低资本"]).copy()
            df_mkt_pct["市场风险占比"] = df_mkt_pct[mkt_total] / df_mkt_pct["最低资本"] * 100
            df_mkt_pct = df_mkt_pct.sort_values("市场风险占比", ascending=False).head(30)
            fig_mkt_p = px.bar(df_mkt_pct, x="市场风险占比", y="公司", orientation="h",
                                 color_discrete_sequence=["#EF9F27"])
            fig_mkt_p.update_layout(height=380, margin=dict(l=0,r=0,t=10,b=0),
                                     plot_bgcolor="white", paper_bgcolor="white")
            st.plotly_chart(fig_mkt_p, use_container_width=True)

# —— 06-信用风险分析 ——
elif page == "06-信用风险分析":
    st.markdown(f"### 🏦 06 · 信用风险分析 · {q_label}")
    cr_cols = [c for c in df.columns if "信用风险-" in c and "分散" not in c and "合计" not in c]
    cr_total = "信用风险-最低资本合计"
    if not cr_cols:
        st.warning("未找到信用风险细分字段。")
    else:
        st.markdown("##### 信用风险细分（全行业合计，万元）")
        cr_sums = {c: abs(df[c].sum()) for c in cr_cols if df[c].notna().any()}
        fig_cr = px.bar(x=[c.replace("信用风险-","").replace("最低资本","").replace("合计","") for c in cr_sums.keys()],
                         y=list(cr_sums.values()),
                         color_discrete_sequence=["#1D9E75"])
        fig_cr.update_layout(height=320, margin=dict(l=0,r=0,t=10,b=80),
                              plot_bgcolor="white", paper_bgcolor="white", xaxis_tickangle=-30,
                              yaxis_title="万元")
        st.plotly_chart(fig_cr, use_container_width=True)
        if cr_total in df.columns and "最低资本" in df.columns:
            st.markdown("##### 信用风险占比（最低资本）")
            df_cr_pct = df.dropna(subset=[cr_total,"最低资本"]).copy()
            df_cr_pct["信用风险占比"] = df_cr_pct[cr_total] / df_cr_pct["最低资本"] * 100
            df_cr_pct = df_cr_pct.sort_values("信用风险占比", ascending=False).head(30)
            fig_cr_p = px.bar(df_cr_pct, x="信用风险占比", y="公司", orientation="h",
                                 color_discrete_sequence=["#1D9E75"])
            fig_cr_p.update_layout(height=380, margin=dict(l=0,r=0,t=10,b=0),
                                     plot_bgcolor="white", paper_bgcolor="white")
            st.plotly_chart(fig_cr_p, use_container_width=True)

# —— 07-公司对比 ——
elif page == "07-公司对比":
    st.markdown(f"### 🔍 07 · 公司对比分析 · {q_label}")
    if not target_co:
        st.info("请在左侧选择「目标公司」以启用对比分析。")
        st.stop()
    same_cat = []
    for cat, cos in COMP_CATS.items():
        if target_co in cos:
            same_cat = [c for c in cos if c in df["公司"].tolist()]
            break
    cmp_default = list(dict.fromkeys([target_co] + same_cat[:6]))
    cmp_cos = st.multiselect("对比公司（最多10家）",
                               [c for c in df["公司"].tolist() if c != target_co],
                               default=[c for c in cmp_default if c != target_co][:5],
                               max_selections=9)
    cmp_cos = [target_co] + cmp_cos
    df_cmp = df[df["公司"].isin(cmp_cos)].copy()
    if RATIO_COMP in df.columns:
        st.markdown("##### 综合偿付能力充足率对比")
        fig_c = go.Figure()
        for co in cmp_cos:
            row = df_cmp[df_cmp["公司"]==co]
            if row.empty: continue
            v = row[RATIO_COMP].values[0]
            fig_c.add_trace(go.Bar(
                name=co, x=[co], y=[v*100],
                marker_color="#185FA5" if co==target_co else "#ccc",
                text=f"{v*100:.1f}%", textposition="outside"
            ))
        fig_c.add_hline(y=100, line_dash="dash", line_color="#b91c1c")
        fig_c.add_hline(y=150, line_dash="dot",  line_color="#EF9F27")
        fig_c.update_layout(height=360, margin=dict(l=0,r=0,t=10,b=80),
                             plot_bgcolor="white", paper_bgcolor="white")
        st.plotly_chart(fig_c, use_container_width=True)
    radar_ms = [COMP_PCT, CORE_PCT, "实际资本","最低资本"]
    radar_ms = [m for m in radar_ms if m in df_cmp.columns]
    if len(radar_ms) >= 3:
        st.markdown("##### 多维度雷达对比（归一化）")
        fig_rad = go.Figure()
        for co in cmp_cos:
            row = df_cmp[df_cmp["公司"]==co]
            if row.empty: continue
            vals = []
            for m in radar_ms:
                v = row[m].values[0]
                mx,mn = df[m].max(), df[m].min()
                vals.append((v-mn)/(mx-mn)*100 if mx!=mn else 50)
            fig_rad.add_trace(go.Scatterpolar(
                r=vals+[vals[0]],
                theta=[m.replace("(%)","").replace("综合偿付能力充足率","综合").replace("核心偿付能力充足率","核心") for m in radar_ms]+[radar_ms[0].replace("(%)","").replace("综合偿付能力充足率","综合").replace("核心偿付能力充足率","核心")],
                name=co, fill="toself", opacity=0.5,
                line=dict(width=3) if co==target_co else None
            ))
        fig_rad.update_layout(polar=dict(radialaxis=dict(visible=True,range=[0,100])),
                                height=420, margin=dict(l=0,r=0,t=20,b=0))
        st.plotly_chart(fig_rad, use_container_width=True)
    st.markdown("##### 对比数据表")
    tbl_cols = ["公司"] + [c for c in [
        RATIO_COMP,RATIO_CORE,"实际资本","最低资本","认可资产","认可负债"
    ] if c in df.columns]
    tbl = df_cmp[tbl_cols].copy()
    for c in [RATIO_COMP,RATIO_CORE]:
        if c in tbl.columns:
            tbl[c] = tbl[c].apply(lambda v: pct(v) if pd.notna(v) else "—")
    for c in ["实际资本","最低资本","认可资产","认可负债"]:
        if c in tbl.columns:
            tbl[c] = tbl[c].apply(lambda v: fmt_wan(v) if pd.notna(v) else "—")
    st.dataframe(tbl, use_container_width=True, hide_index=True)

# —— 08-风险预警 ——
elif page == "08-风险预警":
    st.markdown(f"### 🚨 08 · 风险预警看板 · {q_label}")
    if RATIO_COMP not in df.columns:
        st.warning("未找到「综合偿付能力充足率」字段。")
        st.stop()
    df["状态"]  = df[RATIO_COMP].apply(lambda v: risk_info(v)[0])
    df["级别"]  = df[RATIO_COMP].apply(lambda v: 2 if v>=THR_COMP_WARN else (1 if v>=THR_COMP_PASS else 0))
    n_p = (df["级别"]==2).sum()
    n_w = (df["级别"]==1).sum()
    n_f = (df["级别"]==0).sum()
    k1,k2,k3 = st.columns(3)
    k1.metric("充足（≥150%）", f"{n_p} 家")
    k2.metric("预警（100-150%）", f"{n_w} 家", delta=f"{n_w}家需关注" if n_w>0 else "无", delta_color="off")
    k3.metric("不达标（<100%）", f"{n_f} 家", delta=f"⚠{n_f}家" if n_f>0 else "无", delta_color="inverse")
    pie_data = {"充足":n_p,"预警":n_w,"不达标":n_f}
    pie_data = {k:v for k,v in pie_data.items() if v>0}
    if pie_data:
        fig_pie = px.pie(names=list(pie_data.keys()), values=list(pie_data.values()),
                           color_discrete_sequence=["#1D9E75","#EF9F27","#E24B4A"])
        fig_pie.update_layout(height=260, margin=dict(l=0,r=0,t=10,b=0))
        st.plotly_chart(fig_pie, use_container_width=True)
    st.divider()
    if n_f > 0:
        st.error(f"⚠️ {n_f} 家公司综合偿付能力充足率低于监管下限 100%")
        df_fail = df[df["级别"]==0][["公司","状态",COMP_PCT]+([CORE_PCT] if CORE_PCT in df.columns else [])].sort_values(COMP_PCT)
        for c in [COMP_PCT, CORE_PCT]:
            if c in df_fail.columns:
                df_fail[c] = df_fail[c].apply(lambda v: f"{v:.1f}%" if pd.notna(v) else "—")
        st.dataframe(df_fail, use_container_width=True, hide_index=True)
    if n_w > 0:
        st.warning(f"⚠️ {n_w} 家公司处于预警区间（100%-150%）")
        df_ww = df[df["级别"]==1][["公司","状态",COMP_PCT]+([CORE_PCT] if CORE_PCT in df.columns else [])].sort_values(COMP_PCT)
        for c in [COMP_PCT, CORE_PCT]:
            if c in df_ww.columns:
                df_ww[c] = df_ww[c].apply(lambda v: f"{v:.1f}%" if pd.notna(v) else "—")
        st.dataframe(df_ww, use_container_width=True, hide_index=True)
    st.divider()
    st.markdown("##### 全公司偿付能力充足率排名")
    rank_cols = ["公司","状态",COMP_PCT]+([CORE_PCT] if CORE_PCT in df.columns else [])
    df_rank = df.sort_values(COMP_PCT, ascending=False)[rank_cols].reset_index(drop=True)
    df_rank.index += 1
    for c in [COMP_PCT, CORE_PCT]:
        if c in df_rank.columns:
            df_rank[c] = df_rank[c].apply(lambda v: f"{v:.1f}%" if pd.notna(v) else "—")
    def color_row(row):
        styles = [""] * len(row)
        for i,c in enumerate(row.index):
            if c == COMP_PCT:
                try:
                    v = float(str(row[c]).replace("%",""))
                    if v < 100:   styles[i] = "background:#fce8e8;color:#b91c1c;font-weight:600"
                    elif v < 150: styles[i] = "background:#fff3cd;color:#856404"
                    else:         styles[i] = "background:#e6f4ea;color:#1e7e34"
                except: pass
        return styles
    st.dataframe(df_rank.style.apply(color_row, axis=1),
                 use_container_width=True, height=600)
