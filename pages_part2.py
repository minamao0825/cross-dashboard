# ============================================================
# 页面路由
# ============================================================
# —— 01-行业总览 ——
if page == "01-行业总览":
    st.markdown(f"### 📈 01 · 行业总览 · {q_label}")

    has_c = RATIO_COMP in df.columns
    has_k = RATIO_CORE in df.columns

    kcols = st.columns(6)
    kcols[0].metric("纳入公司数", f"{len(df)} 家")
    if has_c:
        med  = df[RATIO_COMP].median()
        n_p  = (df[RATIO_COMP] >= THR_COMP_WARN).sum()
        n_w  = ((df[RATIO_COMP]>=THR_COMP_PASS)&(df[RATIO_COMP]<THR_COMP_WARN)).sum()
        n_f  = (df[RATIO_COMP] < THR_COMP_PASS).sum()
        kcols[1].metric("综合充足率中位数", pct(med))
        kcols[2].metric("充足（≥150%）", f"{n_p} 家",  delta=f"{n_p/len(df)*100:.0f}%")
        kcols[3].metric("预警（100-150%）", f"{n_w} 家", delta=f"{n_w}家" if n_w>0 else "无", delta_color="off")
        kcols[4].metric("不达标（<100%）", f"{n_f} 家", delta=f"⚠{n_f}家" if n_f>0 else "无", delta_color="inverse")
    if has_k:
        med_k = df[RATIO_CORE].median()
        kcols[5].metric("核心充足率中位数", pct(med_k))

    st.divider()

    r1c1, r1c2 = st.columns(2)
    if has_c:
        with r1c1:
            st.markdown("##### 综合偿付能力充足率分布")
            fig_h = px.histogram(df, x=COMP_PCT, nbins=25, color_discrete_sequence=["#185FA5"])
            fig_h.add_vline(x=100, line_dash="dash", line_color="#b91c1c", annotation_text="100% 下限")
            fig_h.add_vline(x=150, line_dash="dot",  line_color="#EF9F27", annotation_text="150% 预警")
            fig_h.update_layout(height=300, margin=dict(l=0,r=0,t=20,b=0), plot_bgcolor="white", paper_bgcolor="white", yaxis_title="公司数")
            st.plotly_chart(fig_h, use_container_width=True)
        with r1c2:
            if has_k:
                st.markdown("##### 核心 vs 综合偿付能力充足率")
                fig_s = px.scatter(df, x=CORE_PCT, y=COMP_PCT, hover_name="公司", color=COMP_PCT, color_continuous_scale=["#E24B4A","#EF9F27","#1D9E75"])
                fig_s.add_hline(y=100, line_dash="dash", line_color="#b91c1c")
                fig_s.add_vline(x=50,  line_dash="dash", line_color="#b91c1c")
                fig_s.update_layout(height=300, margin=dict(l=0,r=0,t=20,b=0), plot_bgcolor="white", paper_bgcolor="white", coloraxis_showscale=False)
                st.plotly_chart(fig_s, use_container_width=True)

    if "实际资本" in df.columns:
        st.markdown("##### 实际资本 TOP15（亿元）")
        top = df.nlargest(min(15,len(df)), "实际资本")[["公司","实际资本"]].copy()
        top["实际资本_亿"] = top["实际资本"] / 10000
        fig_b = px.bar(top.sort_values("实际资本_亿"), x="实际资本_亿", y="公司", orientation="h", color_discrete_sequence=["#378ADD"])
        fig_b.update_layout(height=420, margin=dict(l=0,r=0,t=10,b=0), plot_bgcolor="white", paper_bgcolor="white")
        st.plotly_chart(fig_b, use_container_width=True)

    if has_c:
        st.markdown("##### 全公司综合偿付能力充足率排名（TOP 40）")
        rank_df = df.sort_values(COMP_PCT, ascending=False).head(40)
        colors  = [color_cells(v/100,"comp") for v in rank_df[COMP_PCT]]
        fig_r  = go.Figure(go.Bar(
            x=rank_df[COMP_PCT], y=rank_df["公司"],
            orientation="h", marker_color=colors,
            text=[f"{v:.1f}%" for v in rank_df[COMP_PCT]],
            textposition="outside"
        ))
        fig_r.add_vline(x=100, line_dash="dash", line_color="#b91c1c", annotation_text="100%")
        fig_r.add_vline(x=150, line_dash="dot",  line_color="#EF9F27", annotation_text="150%")
        fig_r.update_layout(height=max(400, len(rank_df)*22), margin=dict(l=0,r=60,t=10,b=0), plot_bgcolor="white", paper_bgcolor="white", yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig_r, use_container_width=True)

# —— 02-核心指标详情 ——
elif page == "02-核心指标详情":
    st.markdown(f"### 📊 02 · 核心指标详情 · {q_label}")
    core_inds = [c for c in [
        RATIO_COMP, RATIO_CORE,
        "综合偿付能力溢额","核心偿付能力溢额",
        "实际资本","最低资本","认可资产","认可负债"
    ] if c in df.columns]
    if not core_inds:
        st.warning("未找到核心指标字段。")
    else:
        sel_ind = st.selectbox("选择指标", core_inds, index=0)
        is_pct = sel_ind in [RATIO_COMP, RATIO_CORE]
        unit   = "%" if is_pct else "万元"
        df_s = df.dropna(subset=[sel_ind]).sort_values(sel_ind, ascending=False).head(40)
        col_series = COMP_PCT if sel_ind==RATIO_COMP else (CORE_PCT if sel_ind==RATIO_CORE else sel_ind)
        if col_series in df_s.columns or sel_ind in df_s.columns:
            yvals = df_s[col_series].tolist() if col_series in df_s.columns else df_s[sel_ind].tolist()
            colors = []
            for _, row in df_s.iterrows():
                vv = row.get(RATIO_COMP, 1)
                colors.append(color_cells(vv, "comp"))
            fig_b = go.Figure(go.Bar(
                x=yvals if not is_pct else [v for v in yvals],
                y=df_s["公司"], orientation="h",
                marker_color=colors,
                text=[f"{v:.1f}%" if is_pct else fmt_wan(v) for v in yvals],
                textposition="outside"
            ))
            fig_b.update_layout(height=max(400, len(df_s)*22),
                                 margin=dict(l=0,r=60,t=10,b=0),
                                 plot_bgcolor="white", paper_bgcolor="white",
                                 yaxis=dict(autorange="reversed"), xaxis_title=unit)
            st.plotly_chart(fig_b, use_container_width=True)
        if target_co and target_co in df["公司"].tolist():
            st.divider()
            st.markdown(f"##### 目标公司：{target_co}")
            trow = df[df["公司"]==target_co].iloc[0]
            k1,k2,k3,k4 = st.columns(4)
            for i,c in enumerate(core_inds[:4]):
                v = trow.get(c, None)
                disp = pct(v) if c in [RATIO_COMP,RATIO_CORE] else fmt_wan(v)
                if i==0: k1.metric(c, disp)
                elif i==1: k2.metric(c, disp)
                elif i==2: k3.metric(c, disp)
                else: k4.metric(c, disp)
        st.divider()
        st.markdown("##### 核心指标明细表")
        tbl_df = df[["公司"]+core_inds].copy()
        for c in [RATIO_COMP,RATIO_CORE]:
            if c in tbl_df.columns:
                tbl_df[c] = tbl_df[c].apply(lambda v: pct(v) if pd.notna(v) else "—")
        for c in core_inds:
            if c not in [RATIO_COMP,RATIO_CORE] and c in tbl_df.columns:
                tbl_df[c] = tbl_df[c].apply(lambda v: fmt_wan(v) if pd.notna(v) else "—")
        st.dataframe(tbl_df, use_container_width=True, hide_index=True, height=500)

# —— 03-资本效率分析 ——
elif page == "03-资本效率分析":
    st.markdown(f"### 💰 03 · 资本效率分析 · {q_label}")
    ratio_inds = [c for c in IND_CATS["03-资本效率比率"] if c in df.columns]
    if not ratio_inds:
        st.warning("未找到资本效率比率字段。")
    else:
        st.markdown("##### 资本效率比率行业分布（箱型图）")
        fig_box = go.Figure()
        colors_r = ["#185FA5","#378ADD","#1D9E75","#639922","#EF9F27","#E24B4A"]
        for i,c in enumerate(ratio_inds[:6]):
            vals = df[c].dropna()
            if vals.empty: continue
            fig_box.add_trace(go.Box(y=vals.tolist(), name=c, boxmean="sd",
                marker_color=colors_r[i%len(colors_r)]))
        fig_box.update_layout(height=380, margin=dict(l=0,r=0,t=10,b=80),
                                plot_bgcolor="white", paper_bgcolor="white", yaxis_title="比率值")
        st.plotly_chart(fig_box, use_container_width=True)
        if target_co and target_co in df["公司"].tolist():
            st.divider()
            st.markdown(f"##### 目标公司：{target_co} vs 行业中位数")
            trow = df[df["公司"]==target_co].iloc[0]
            cmp_rows = []
            for c in ratio_inds:
                v_t = trow.get(c, None)
                v_m = df[c].median() if c in df.columns else None
                cmp_rows.append({"指标": c, "目标公司": f"{v_t:.4f}" if pd.notna(v_t) else "—",
                                 "行业中位数": f"{v_m:.4f}" if pd.notna(v_m) else "—"})
            st.dataframe(pd.DataFrame(cmp_rows), use_container_width=True, hide_index=True)
    cap_cols = [c for c in ["核心一级资本","核心二级资本","附属一级资本","附属二级资本"] if c in df.columns]
    if cap_cols:
        st.divider()
        st.markdown("##### 资本分级结构（TOP20 实际资本）")
        top20 = df.nlargest(20, "实际资本")[["公司"]+cap_cols].copy()
        fig_stack = go.Figure()
        for c in cap_cols:
            fig_stack.add_trace(go.Bar(name=c, x=top20["公司"], y=top20[c].tolist()))
        fig_stack.update_layout(barmode="stack", height=400,
                                margin=dict(l=0,r=0,t=10,b=80),
                                plot_bgcolor="white", paper_bgcolor="white", xaxis_tickangle=-30)
        st.plotly_chart(fig_stack, use_container_width=True)
