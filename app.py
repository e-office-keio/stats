"""
APA Style Statistical Analysis Web Application using Streamlit.
"""

import io
import streamlit as st
import pandas as pd
import numpy as np

# カスタムモジュールの読み込み
import stats_engine
import apa_excel

st.set_page_config(
    page_title="APAスタイル 統計解析 Webアプリ",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# セッション状態の初期化
if "analysis_queue" not in st.session_state:
    st.session_state["analysis_queue"] = []

# カスタムCSSでデザインを学術風に整理
st.markdown("""
<style>
    .main-title {
        font-size: 2.2rem;
        font-weight: 700;
        color: #1e3a8a;
        margin-bottom: 0.2rem;
    }
    .sub-title {
        font-size: 1.05rem;
        color: #4b5563;
        margin-bottom: 1.5rem;
    }
    .card {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 1.2rem;
        margin-bottom: 1rem;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: #f1f5f9;
        border-radius: 6px 6px 0px 0px;
        gap: 1px;
        padding-top: 10px;
        padding-bottom: 10px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1e3a8a !important;
        color: white !important;
    }
</style>
""", unsafe_allow_html=True)


def load_csv(uploaded_file, encoding_choice):
    """エンコーディングに対応したCSV読み込み"""
    if encoding_choice == "自動判定 (Auto)":
        encodings_to_try = ["utf-8", "cp932", "shift_jis", "euc-jp", "utf-8-sig"]
        for enc in encodings_to_try:
            try:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, encoding=enc)
                return df, enc
            except Exception:
                continue
        raise ValueError("CSVファイルのエンコーディングを自動判定できませんでした。手動で選択してください。")
    else:
        uploaded_file.seek(0)
        df = pd.read_csv(uploaded_file, encoding=encoding_choice)
        return df, encoding_choice


def main():
    st.markdown('<div class="main-title">📊 APAスタイル 統計解析 Webアプリ</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">CSVデータをアップロードして、多様な統計分析（基本統計量・t検定・ANOVA・回帰分析・因子分析など）を行い、APA第7版形式のExcelレポート（表＋図）を出力します。</div>', unsafe_allow_html=True)
    
    # ------------------------------------------------------------------
    # サイドバー：ファイルアップロード & 分析レポート管理
    # ------------------------------------------------------------------
    with st.sidebar:
        st.header("📂 データアップロード")
        uploaded_file = st.file_uploader("CSVファイルを選択", type=["csv"])
        
        encoding_choice = st.selectbox(
            "文字コード (Encoding)",
            ["自動判定 (Auto)", "utf-8", "cp932 (Windows Shift-JIS)", "shift_jis", "utf-8-sig"]
        )
        
        st.divider()
        st.header("📋 出力レポート管理")
        queue_count = len(st.session_state["analysis_queue"])
        st.write(f"現在のレポート追加件数: **{queue_count}** 件")
        
        if queue_count > 0:
            if st.button("📥 全結果をAPAスタイルExcelでダウンロード", type="primary", use_container_width=True):
                excel_bytes = apa_excel.build_full_excel_report(st.session_state["analysis_queue"])
                st.download_button(
                    label="💾 Excelファイルを保存 (.xlsx)",
                    data=excel_bytes,
                    file_name="APA_Statistical_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            if st.button("🗑 レポートキューをクリア", use_container_width=True):
                st.session_state["analysis_queue"] = []
                st.rerun()

    if uploaded_file is None:
        st.info("👈 サイドバーから分析したいCSVファイルをアップロードしてください。")
        
        # サンプルデータの提供
        with st.expander("💡 テスト用のサンプルデータで試す"):
            st.write("ボタンを押すとダミーデータを生成してテストできます。")
            if st.button("サンプルデータを生成して読み込む"):
                np.random.seed(42)
                n = 120
                sample_data = pd.DataFrame({
                    "ID": range(1, n + 1),
                    "Group": np.random.choice(["Control", "Treatment_A", "Treatment_B"], size=n),
                    "Gender": np.random.choice(["Male", "Female"], size=n),
                    "Age": np.random.randint(20, 60, size=n),
                    "Pre_Test": np.random.normal(50, 10, size=n),
                    "Post_Test": np.random.normal(60, 12, size=n),
                    "Satisfaction": np.random.normal(3.8, 0.8, size=n),
                    "Q1_Motivation": np.random.randint(1, 6, size=n),
                    "Q2_Engagement": np.random.randint(1, 6, size=n),
                    "Q3_Performance": np.random.randint(1, 6, size=n)
                })
                # サンプルデータをセッションに保存
                csv_buffer = io.StringIO()
                sample_data.to_csv(csv_buffer, index=False)
                st.session_state["df"] = sample_data
                st.success("サンプルデータを読み込みました！下の分析タブで操作を試せます。")
                st.rerun()
                
        if "df" in st.session_state:
            df = st.session_state["df"]
        else:
            return
    else:
        try:
            df, enc_used = load_csv(uploaded_file, encoding_choice)
            st.session_state["df"] = df
            st.caption(f"読み込み成功: {uploaded_file.name} (文字コード: {enc_used}, 行数: {df.shape[0]}, 列数: {df.shape[1]})")
        except Exception as e:
            st.error(f"ファイルの読み込みエラー: {e}")
            return

    # データプレビュー
    with st.expander("🔍 データの先頭プレビュー & 変数情報"):
        col1, col2 = st.columns([3, 1])
        with col1:
            st.dataframe(df.head(10), use_container_width=True)
        with col2:
            st.write("**データ型一覧:**")
            st.dataframe(pd.DataFrame(df.dtypes, columns=["DataType"]), use_container_width=True)

    # 変数列の分離
    all_columns = df.columns.tolist()
    numeric_columns = df.select_dtypes(include=[np.number]).columns.tolist()
    categorical_columns = df.select_dtypes(exclude=[np.number]).columns.tolist()
    if not categorical_columns:
        categorical_columns = all_columns  # 数値型でもカテゴリ扱いできるようにフォールバック

    # ------------------------------------------------------------------
    # メイン分析タブ
    # ------------------------------------------------------------------
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
        "1. 単純集計",
        "2. 基本統計量",
        "3. クロス集計",
        "4. t検定/Welch",
        "5. 分散分析",
        "6. 相関分析",
        "7. 回帰分析",
        "8. 因子分析"
    ])

    # ------------------------------------------------------------------
    # TAB 1: 単純集計
    # ------------------------------------------------------------------
    with tab1:
        st.subheader("1. 単純集計 (Frequency Analysis)")
        target_var = st.selectbox("集計したいカテゴリ変数を選択:", all_columns, key="freq_var")
        
        if target_var:
            res_df, note, fig_bytes = stats_engine.analyze_frequency(df, target_var)
            
            c1, c2 = st.columns([2, 2])
            with c1:
                st.write("**集計表 (APA Style Table)**")
                st.dataframe(res_df, use_container_width=True)
                st.caption(f"Note. {note}")
            with c2:
                st.write("**図 (APA Style Figure)**")
                st.image(fig_bytes, use_container_width=True)
                
            if st.button("➕ この結果をExcelレポートに追加", key="btn_add_freq"):
                st.session_state["analysis_queue"].append({
                    "sheet_name": f"単純集計_{target_var}",
                    "title": f"Frequency Distribution for {target_var}",
                    "df": res_df,
                    "note": note,
                    "fig_bytes": fig_bytes
                })
                st.success(f"『単純集計_{target_var}』をレポートキューに追加しました！")

    # ------------------------------------------------------------------
    # TAB 2: 基本統計量
    # ------------------------------------------------------------------
    with tab2:
        st.subheader("2. 基本統計量 (Descriptive Statistics)")
        selected_num_vars = st.multiselect("分析する数値変数を選択:", numeric_columns, default=numeric_columns[:min(4, len(numeric_columns))], key="desc_vars")
        
        if selected_num_vars:
            res_df, note, fig_bytes = stats_engine.analyze_descriptives(df, selected_num_vars)
            
            c1, c2 = st.columns([2.5, 2])
            with c1:
                st.write("**基本統計量一覧**")
                st.dataframe(res_df, use_container_width=True)
                st.caption(f"Note. {note}")
            with c2:
                st.write("**分布プロット**")
                st.image(fig_bytes, use_container_width=True)
                
            if st.button("➕ この結果をExcelレポートに追加", key="btn_add_desc"):
                st.session_state["analysis_queue"].append({
                    "sheet_name": "基本統計量",
                    "title": "Descriptive Statistics for Selected Variables",
                    "df": res_df,
                    "note": note,
                    "fig_bytes": fig_bytes
                })
                st.success("『基本統計量』をレポートキューに追加しました！")

    # ------------------------------------------------------------------
    # TAB 3: クロス集計
    # ------------------------------------------------------------------
    with tab3:
        st.subheader("3. クロス集計 & カイ二乗検定 (Crosstab & Chi-Square)")
        c1, c2 = st.columns(2)
        with c1:
            row_var = st.selectbox("行変数 (Row):", all_columns, key="ct_row")
        with c2:
            col_var = st.selectbox("列変数 (Column):", [c for c in all_columns if c != row_var], key="ct_col")
            
        if row_var and col_var:
            ct_df, note, fig_bytes = stats_engine.analyze_crosstab(df, row_var, col_var)
            
            col_left, col_right = st.columns([2.5, 2])
            with col_left:
                st.write("**クロス度数表**")
                st.dataframe(ct_df, use_container_width=True)
                st.caption(f"Note. {note}")
            with col_right:
                st.write("**構成比積み上げグラフ**")
                st.image(fig_bytes, use_container_width=True)
                
            if st.button("➕ この結果をExcelレポートに追加", key="btn_add_ct"):
                st.session_state["analysis_queue"].append({
                    "sheet_name": f"クロス_{row_var}_vs_{col_var}",
                    "title": f"Crosstabulation between {row_var} and {col_var}",
                    "df": ct_df,
                    "note": note,
                    "fig_bytes": fig_bytes
                })
                st.success("『クロス集計』をレポートキューに追加しました！")

    # ------------------------------------------------------------------
    # TAB 4: t検定 / Welch検定
    # ------------------------------------------------------------------
    with tab4:
        st.subheader("4. 2群の平均値の差の検定 (t-Test / Welch Test)")
        c1, c2, c3 = st.columns(3)
        with c1:
            group_var = st.selectbox("グループ変数 (2カテゴリ):", all_columns, key="tt_group")
        with c2:
            num_var = st.selectbox("比較する数値変数:", numeric_columns, key="tt_num")
        with c3:
            equal_var_opt = st.selectbox("等分散性の仮定:", ["Welchのt検定 (推奨: 等分散非仮定)", "Studentのt検定 (等分散仮定)"])
            equal_var = True if "Student" in equal_var_opt else False
            
        if group_var and num_var:
            try:
                res_df, note, fig_bytes = stats_engine.analyze_ttest(df, group_var, num_var, equal_var=equal_var)
                
                col_left, col_right = st.columns([2.5, 2])
                with col_left:
                    st.write("**t検定分析結果**")
                    st.dataframe(res_df, use_container_width=True)
                    st.caption(f"Note. {note}")
                with col_right:
                    st.write("**平均値と比較グラフ (95% CI)**")
                    st.image(fig_bytes, use_container_width=True)
                    
                if st.button("➕ この結果をExcelレポートに追加", key="btn_add_tt"):
                    st.session_state["analysis_queue"].append({
                        "sheet_name": f"t検定_{num_var}",
                        "title": f"t-Test Comparison of {num_var} by {group_var}",
                        "df": res_df,
                        "note": note,
                        "fig_bytes": fig_bytes
                    })
                    st.success("『t検定結果』をレポートキューに追加しました！")
            except Exception as e:
                st.error(f"分析エラー: {e}")

    # ------------------------------------------------------------------
    # TAB 5: 分散分析 (ANOVA)
    # ------------------------------------------------------------------
    with tab5:
        st.subheader("5. 一元配置分散分析 & 多重比較 (One-Way ANOVA)")
        c1, c2 = st.columns(2)
        with c1:
            anova_group = st.selectbox("要因 (グループ変数):", all_columns, key="anova_grp")
        with c2:
            anova_num = st.selectbox("従属変数 (数値変数):", numeric_columns, key="anova_num")
            
        if anova_group and anova_num:
            desc_df, tukey_df, note, fig_bytes = stats_engine.analyze_anova(df, anova_group, anova_num)
            
            c_l, c_r = st.columns([2.5, 2])
            with c_l:
                st.write("**記述統計量 (各群の平均と標準偏差)**")
                st.dataframe(desc_df, use_container_width=True)
                st.caption(f"Note. {note}")
                
                st.write("**Tukey HSD 多重比較結果**")
                st.dataframe(tukey_df, use_container_width=True)
            with c_r:
                st.write("**各群の平均値比較プロット**")
                st.image(fig_bytes, use_container_width=True)
                
            if st.button("➕ この結果をExcelレポートに追加", key="btn_add_anova"):
                st.session_state["analysis_queue"].append({
                    "sheet_name": f"ANOVA_{anova_num}",
                    "title": f"One-Way ANOVA for {anova_num} across {anova_group}",
                    "df": desc_df,
                    "note": note,
                    "fig_bytes": fig_bytes
                })
                st.session_state["analysis_queue"].append({
                    "sheet_name": f"ANOVA_Tukey_{anova_num}",
                    "title": f"Tukey HSD Post-Hoc Comparisons for {anova_num}",
                    "df": tukey_df,
                    "note": "Tukey's HSD test at alpha = .05."
                })
                st.success("『ANOVA & 多重比較結果』をレポートキューに追加しました！")

    # ------------------------------------------------------------------
    # TAB 6: 相関分析
    # ------------------------------------------------------------------
    with tab6:
        st.subheader("6. 相関分析 (Correlation Analysis)")
        c1, c2 = st.columns([3, 1])
        with c1:
            corr_vars = st.multiselect("相関を算出する変数を選択 (2つ以上):", numeric_columns, default=numeric_columns[:min(5, len(numeric_columns))], key="corr_vars")
        with c2:
            corr_method = st.radio("相関係数の種類:", ["pearson", "spearman"])
            
        if len(corr_vars) >= 2:
            res_df, note, fig_bytes = stats_engine.analyze_correlation(df, corr_vars, method=corr_method)
            
            c_l, c_r = st.columns([2.5, 2])
            with c_l:
                st.write(f"**{corr_method.capitalize()} 相関係数行列**")
                st.dataframe(res_df, use_container_width=True)
                st.caption(f"Note. {note}")
            with c_r:
                st.write("**相関ヒートマップ**")
                st.image(fig_bytes, use_container_width=True)
                
            if st.button("➕ この結果をExcelレポートに追加", key="btn_add_corr"):
                st.session_state["analysis_queue"].append({
                    "sheet_name": f"相関分析_{corr_method}",
                    "title": f"{corr_method.capitalize()} Correlation Matrix",
                    "df": res_df,
                    "note": note,
                    "fig_bytes": fig_bytes
                })
                st.success("『相関係数行列』をレポートキューに追加しました！")

    # ------------------------------------------------------------------
    # TAB 7: 重回帰分析
    # ------------------------------------------------------------------
    with tab7:
        st.subheader("7. 重回帰分析 (Multiple Linear Regression)")
        c1, c2 = st.columns([1, 2])
        with c1:
            target_var = st.selectbox("目的変数 (Y):", numeric_columns, key="reg_target")
        with c2:
            avail_features = [c for c in numeric_columns if c != target_var]
            feature_vars = st.multiselect("説明変数 (X):", avail_features, default=avail_features[:min(3, len(avail_features))], key="reg_features")
            
        if target_var and feature_vars:
            res_df, note, fig_bytes = stats_engine.analyze_regression(df, target_var, feature_vars)
            
            c_l, c_r = st.columns([2.5, 2])
            with c_l:
                st.write("**回帰モデルモデル係数表 (APA Style)**")
                st.dataframe(res_df, use_container_width=True)
                st.caption(f"Note. {note}")
            with c_r:
                st.write("**実測値 vs 予測値**")
                st.image(fig_bytes, use_container_width=True)
                
            if st.button("➕ この結果をExcelレポートに追加", key="btn_add_reg"):
                st.session_state["analysis_queue"].append({
                    "sheet_name": f"回帰分析_{target_var}",
                    "title": f"Multiple Regression Predicting {target_var}",
                    "df": res_df,
                    "note": note,
                    "fig_bytes": fig_bytes
                })
                st.success("『回帰分析結果』をレポートキューに追加しました！")

    # ------------------------------------------------------------------
    # TAB 8: 因子分析
    # ------------------------------------------------------------------
    with tab8:
        st.subheader("8. 探索的因子分析 (Exploratory Factor Analysis)")
        c1, c2, c3 = st.columns(3)
        with c1:
            fa_vars = st.multiselect("因子分析に投入する観測変数:", numeric_columns, default=numeric_columns[:min(6, len(numeric_columns))], key="fa_vars")
        with c2:
            n_factors = st.number_input("抽出する因子数:", min_value=1, max_value=max(1, len(fa_vars)-1), value=min(2, max(1, len(fa_vars)-1)))
        with c3:
            rotation = st.selectbox("因子回転法:", ["promax", "varimax"])
            
        if len(fa_vars) >= 3:
            res_df, note, fig_bytes = stats_engine.analyze_factor_analysis(df, fa_vars, n_factors=int(n_factors), rotation=rotation)
            
            c_l, c_r = st.columns([2.5, 2])
            with c_l:
                st.write("**因子負荷量 & 寄与率**")
                st.dataframe(res_df, use_container_width=True)
                st.caption(f"Note. {note}")
            with c_r:
                st.write("**スクリープロット (Scree Plot)**")
                st.image(fig_bytes, use_container_width=True)
                
            if st.button("➕ この結果をExcelレポートに追加", key="btn_add_fa"):
                st.session_state["analysis_queue"].append({
                    "sheet_name": "因子分析",
                    "title": f"Exploratory Factor Analysis ({rotation.capitalize()} Rotation)",
                    "df": res_df,
                    "note": note,
                    "fig_bytes": fig_bytes
                })
                st.success("『因子分析結果』をレポートキューに追加しました！")


if __name__ == "__main__":
    main()
