"""
APA Style Statistical Analysis Web Application using Streamlit.
(Features: Instant Sidebar Refresh on Queue Add, Scrollable Plot Containers, Japanese Interface)
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
if "results" not in st.session_state:
    st.session_state["results"] = {}

# カスタムCSS
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
        margin-bottom: 1.2rem;
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
    st.markdown('<div class="sub-title">CSVをアップロード後、リコードや特定値の除外などのデータ加工を行い、変数と分析方法を選択して「分析を実行」ボタンでAPAスタイル（第7版）の表・図をExcel出力します。</div>', unsafe_allow_html=True)
    
    # ------------------------------------------------------------------
    # サイドバー：ファイルアップロード & 出力レポート管理
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
        st.info(f"追加済みレポート: **{queue_count}** 件")
        
        if queue_count > 0:
            if st.button("📥 全結果をAPAスタイルExcelで作成", type="primary", use_container_width=True):
                with st.spinner("⏳ APAスタイルのExcelレポートを作成中..."):
                    excel_bytes = apa_excel.build_full_excel_report(st.session_state["analysis_queue"])
                st.toast("🎉 Excelレポートの作成が完了しました！", icon="✅")
                st.download_button(
                    label="💾 Excelファイルを保存 (.xlsx)",
                    data=excel_bytes,
                    file_name="APA_統計解析結果レポート.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            if st.button("🗑 レポートリストをクリア", use_container_width=True):
                st.session_state["analysis_queue"] = []
                st.toast("レポートリストをクリアしました", icon="🧹")
                st.rerun()

        st.divider()
        with st.expander("🛠 日本語フォント診断 (Streamlit Cloud確認用)"):
            has_jap = getattr(stats_engine, "HAS_JAPANIZE", False)
            jap_err = getattr(stats_engine, "JAPANIZE_ERROR", "")
            if has_jap:
                st.success("✅ `japanize-matplotlib` 読み込み成功")
            else:
                st.error("❌ `japanize-matplotlib` 未適用")
                if jap_err:
                    st.code(f"エラー詳細: {jap_err}", language="text")
                
            import matplotlib.pyplot as plt
            current_font = plt.rcParams.get("font.family", ["不明"])
            if isinstance(current_font, list):
                current_font = current_font[0]
            st.caption(f"現在のMatplotlibフォント: `{current_font}`")

    # データフレームの初期ロード処理
    if uploaded_file is None:
        if "df" not in st.session_state:
            st.info("👈 サイドバーから分析したいCSVファイルをアップロードしてください。")
            with st.expander("💡 テスト用のサンプルデータで試す"):
                st.write("下のボタンを押すと動作確認用のサンプルデータを生成します。")
                if st.button("サンプルデータを生成して読み込む"):
                    with st.spinner("サンプルデータを生成中..."):
                        np.random.seed(42)
                        n = 120
                        sample_data = pd.DataFrame({
                            "被験者ID": range(1, n + 1),
                            "実験グループ": np.random.choice(["統制群", "介入A群", "介入B群"], size=n),
                            "性別": np.random.choice(["男性", "女性"], size=n),
                            "年齢": np.random.randint(20, 65, size=n),
                            "事前テスト得点": np.random.normal(50, 10, size=n),
                            "事後テスト得点": np.random.normal(60, 12, size=n),
                            "満足度": np.random.normal(3.8, 0.8, size=n),
                            "Q1_学習意欲": np.random.choice([1, 2, 3, 4, 5, -99], size=n, p=[0.1, 0.2, 0.4, 0.2, 0.08, 0.02]),
                            "Q2_集中度": np.random.choice([1, 2, 3, 4, 5, -99], size=n, p=[0.05, 0.15, 0.5, 0.25, 0.03, 0.02]),
                            "Q3_理解度": np.random.choice([1, 2, 3, 4, 5, -99], size=n, p=[0.08, 0.22, 0.45, 0.2, 0.03, 0.02])
                        })
                        st.session_state["raw_df"] = sample_data.copy()
                        st.session_state["df"] = sample_data.copy()
                    st.toast("サンプルデータを読み込みました！", icon="🚀")
                    st.rerun()
            return
        else:
            df = st.session_state["df"]
    else:
        if "last_uploaded" not in st.session_state or st.session_state["last_uploaded"] != uploaded_file.name:
            try:
                with st.spinner("データを読み込み中..."):
                    df_raw, enc_used = load_csv(uploaded_file, encoding_choice)
                    st.session_state["raw_df"] = df_raw.copy()
                    st.session_state["df"] = df_raw.copy()
                    st.session_state["last_uploaded"] = uploaded_file.name
                st.toast(f"✅ {uploaded_file.name} を読み込みました！", icon="📂")
            except Exception as e:
                st.error(f"ファイルの読み込みエラー: {e}")
                return
        df = st.session_state["df"]

    st.caption(f"現在の分析対象データ: {df.shape[0]}行 × {df.shape[1]}列")

    # ------------------------------------------------------------------
    # データ前処理・リコード・フィルタリング セクション
    # ------------------------------------------------------------------
    with st.expander("🛠️ データの加工・Recode（再符号化）・特定値の除外", expanded=False):
        t_proc1, t_proc2, t_proc3, t_proc4 = st.tabs([
            "① 特定値の除外",
            "② 値のリコード (置換)",
            "③ 数値変数のカテゴリ化 (ビン分割)",
            "④ 合成スコア (合計・平均)"
        ])
        
        # --- ① 特定値の除外 ---
        with t_proc1:
            st.markdown("**指定した変数から特定の無効値（例: -99, 99, '無回答' など）を除外します。**")
            col_ex1, col_ex2 = st.columns(2)
            with col_ex1:
                target_ex_var = st.selectbox("対象の変数を選択:", df.columns, key="ex_var")
            with col_ex2:
                unique_vals = df[target_ex_var].dropna().unique().tolist()
                vals_to_exclude = st.multiselect("除外したい値を選択:", unique_vals, key="ex_vals")
                
            if st.button("🚫 指定した値の行を除外してデータを更新", key="btn_apply_ex"):
                if vals_to_exclude:
                    new_df = stats_engine.filter_exclude_values(df, target_ex_var, vals_to_exclude)
                    st.session_state["df"] = new_df
                    st.toast(f"『{target_ex_var}』から {vals_to_exclude} を除外しました ({len(df)}行 → {len(new_df)}行)", icon="✂️")
                    st.rerun()

        # --- ② 値のリコード ---
        with t_proc2:
            st.markdown("**変数の特定の値を別の値（文字列・数値）に置き換えて新しい変数を作成します。**")
            rec_var = st.selectbox("リコード元の変数:", df.columns, key="rec_var")
            new_var_name = st.text_input("作成する新変数名:", value=f"{rec_var}_recoded", key="rec_new_name")
            
            unique_rec_vals = df[rec_var].dropna().unique().tolist()
            st.write("各値の置き換えルールを設定してください:")
            
            mapping_dict = {}
            col_a, col_b = st.columns(2)
            for i, val in enumerate(unique_rec_vals):
                with col_a:
                    st.write(f"旧値: `{val}`")
                with col_b:
                    new_val_str = st.text_input(f"`{val}` の新値:", value=str(val), key=f"rec_val_{i}")
                    try:
                        if "." in new_val_str:
                            val_conv = float(new_val_str)
                        else:
                            val_conv = int(new_val_str)
                    except ValueError:
                        val_conv = new_val_str
                    mapping_dict[val] = val_conv
                    
            if st.button("🔄 リコードを実行して新変数を作成", key="btn_apply_rec"):
                new_df, created_name = stats_engine.recode_values(df, rec_var, mapping_dict, new_var_name)
                st.session_state["df"] = new_df
                st.toast(f"新変数 『{created_name}』 を作成しました！", icon="✨")
                st.rerun()

        # --- ③ 数値変数のカテゴリ化 (ビン分割) ---
        with t_proc3:
            st.markdown("**連続数値変数（例: 年齢）を区切り値でカテゴリ変数（例: 年齢層）に変換します。**")
            num_cols_only = df.select_dtypes(include=[np.number]).columns.tolist()
            if num_cols_only:
                bin_var = st.selectbox("カテゴリ化する数値変数:", num_cols_only, key="bin_var")
                bin_new_name = st.text_input("作成する新変数名:", value=f"{bin_var}_層", key="bin_new_name")
                
                cuts_input = st.text_input("区切り値をカンマ区切りで入力 (例: 0, 30, 50, 100):", value="0, 30, 50, 100", key="bin_cuts")
                labels_input = st.text_input("ラベルをカンマ区切りで入力 (例: 若年, 中年, 高齢):", value="若年, 中年, 高齢", key="bin_labels")
                
                if st.button("📊 カテゴリ化（ビン分割）を実行", key="btn_apply_bin"):
                    try:
                        cuts = [float(x.strip()) for x in cuts_input.split(",")]
                        labels = [x.strip() for x in labels_input.split(",")]
                        new_df, created_name = stats_engine.create_binned_variable(df, bin_var, cuts, labels, bin_new_name)
                        st.session_state["df"] = new_df
                        st.toast(f"新変数 『{created_name}』 を作成しました！", icon="✨")
                        st.rerun()
                    except Exception as e:
                        st.error(f"エラー: {e}")

        # --- ④ 合成スコア ---
        with t_proc4:
            st.markdown("**複数の数値変数から「平均値」または「合計値」の新変数（尺度得点など）を作成します。**")
            if num_cols_only:
                source_vars = st.multiselect("合成する変数を選択 (複数):", num_cols_only, key="comp_vars")
                comp_method = st.radio("計算方法:", ["平均値 (Mean)", "合計値 (Sum)"], horizontal=True, key="comp_method")
                comp_new_name = st.text_input("作成する新変数名:", value="合成スコア", key="comp_new_name")
                
                if st.button("➕ 合成スコアを作成", key="btn_apply_comp"):
                    if len(source_vars) >= 2:
                        func_type = "mean" if "平均値" in comp_method else "sum"
                        new_df, created_name = stats_engine.create_composite_score(df, source_vars, func=func_type, new_var_name=comp_new_name)
                        st.session_state["df"] = new_df
                        st.toast(f"新変数 『{created_name}』 を作成しました！", icon="✨")
                        st.rerun()
                    else:
                        st.warning("合成には2つ以上の変数を選択してください。")

        st.divider()
        if st.button("↩️ データをアップロード直後の初期状態に戻す"):
            st.session_state["df"] = st.session_state["raw_df"].copy()
            st.toast("データを初期状態にリセットしました", icon="🔄")
            st.rerun()

    # データプレビュー
    with st.expander("🔍 現在のデータプレビュー & 変数一覧"):
        col1, col2 = st.columns([3, 1])
        with col1:
            st.dataframe(df.head(10), use_container_width=True)
        with col2:
            st.write("**データ型一覧:**")
            st.dataframe(pd.DataFrame(df.dtypes, columns=["データ型"]), use_container_width=True)

    # 変数列の分類
    all_columns = df.columns.tolist()
    numeric_columns = df.select_dtypes(include=[np.number]).columns.tolist()

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
        target_vars = st.multiselect("集計したいカテゴリ変数を選択 (複数選択可):", all_columns, default=[all_columns[0]] if all_columns else [], key="freq_vars")
        
        if st.button("🚀 単純集計を実行", key="run_freq", type="primary"):
            if target_vars:
                results_dict = {}
                with st.spinner(f"{len(target_vars)} 個の変数の単純集計を計算中..."):
                    for var in target_vars:
                        res_df, note, fig_bytes = stats_engine.analyze_frequency(df, var)
                        results_dict[var] = {
                            "res_df": res_df, "note": note, "fig_bytes": fig_bytes, "target_var": var
                        }
                st.session_state["results"]["freq"] = {
                    "results_dict": results_dict
                }
                st.toast(f"{len(results_dict)} 個の変数の単純集計が完了しました！", icon="✅")
            else:
                st.warning("1つ以上の変数を選択してください。")
            
        if "freq" in st.session_state["results"]:
            res_data = st.session_state["results"]["freq"]
            results_dict = res_data.get("results_dict", {})
            if not results_dict and "res_df" in res_data:
                results_dict = {res_data["target_var"]: res_data}
                
            if results_dict:
                c_btn1, _ = st.columns([2, 1])
                with c_btn1:
                    if st.button(f"➕ 選択した全変数 ({len(results_dict)}件) をExcelレポートに追加", key="btn_add_freq"):
                        for var, res in results_dict.items():
                            st.session_state["analysis_queue"].append({
                                "sheet_name": f"単純集計_{var}",
                                "title": f"{var} の度数分布表",
                                "df": res["res_df"],
                                "note": res["note"],
                                "fig_bytes": res["fig_bytes"]
                            })
                        st.toast(f"『単純集計 ({len(results_dict)}件)』をレポートリストに追加しました！", icon="📋")
                        st.rerun()

                active_var = st.selectbox("表示する変数の切替:", list(results_dict.keys()), key="select_freq_display")
                if active_var in results_dict:
                    res = results_dict[active_var]
                    c1, c2 = st.columns([2.2, 2])
                    with c1:
                        st.write(f"**集計表: {res['target_var']}**")
                        st.dataframe(res["res_df"], use_container_width=True)
                        st.caption(res["note"])
                    with c2:
                        st.write("**度数分布グラフ (APA Style)**")
                        with st.container(height=450):
                            st.image(res["fig_bytes"], use_column_width=True)

    # ------------------------------------------------------------------
    # TAB 2: 基本統計量
    # ------------------------------------------------------------------
    with tab2:
        st.subheader("2. 基本統計量 (Descriptive Statistics)")
        selected_num_vars = st.multiselect("分析する数値変数を選択:", numeric_columns, default=numeric_columns[:min(4, len(numeric_columns))], key="desc_vars")
        
        if st.button("🚀 基本統計量を計算", key="run_desc", type="primary"):
            if selected_num_vars:
                with st.spinner("基本統計量と分布プロットを作成中..."):
                    res_df, note, fig_bytes = stats_engine.analyze_descriptives(df, selected_num_vars)
                    st.session_state["results"]["desc"] = {
                        "res_df": res_df, "note": note, "fig_bytes": fig_bytes
                    }
                st.toast("基本統計量の計算が完了しました！", icon="✅")
            else:
                st.warning("1つ以上の数値変数を選択してください。")
                
        if "desc" in st.session_state["results"]:
            res = st.session_state["results"]["desc"]
            
            if st.button("➕ この結果をExcelレポートに追加", key="btn_add_desc"):
                st.session_state["analysis_queue"].append({
                    "sheet_name": "基本統計量",
                    "title": "選択変数の基本統計量一覧表",
                    "df": res["res_df"],
                    "note": res["note"],
                    "fig_bytes": res["fig_bytes"]
                })
                st.toast("『基本統計量』をレポートリストに追加しました！", icon="📋")
                st.rerun()

            c1, c2 = st.columns([2.5, 2])
            with c1:
                st.write("**基本統計量一覧**")
                st.dataframe(res["res_df"], use_container_width=True)
                st.caption(res["note"])
            with c2:
                st.write("**分布プロット (スクロール表示)**")
                with st.container(height=480):
                    st.image(res["fig_bytes"], use_column_width=True)

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
            
        if st.button("🚀 クロス集計を実行", key="run_ct", type="primary"):
            if row_var and col_var:
                with st.spinner("クロス集計とカイ二乗検定を計算中..."):
                    ct_df, note, fig_bytes = stats_engine.analyze_crosstab(df, row_var, col_var)
                    st.session_state["results"]["ct"] = {
                        "ct_df": ct_df, "note": note, "fig_bytes": fig_bytes, "row_var": row_var, "col_var": col_var
                    }
                st.toast("クロス集計が完了しました！", icon="✅")
                
        if "ct" in st.session_state["results"]:
            res = st.session_state["results"]["ct"]
            
            if st.button("➕ この結果をExcelレポートに追加", key="btn_add_ct"):
                st.session_state["analysis_queue"].append({
                    "sheet_name": f"クロス_{res['row_var']}_vs_{res['col_var']}",
                    "title": f"クロス集計表 ({res['row_var']} × {res['col_var']})",
                    "df": res["ct_df"],
                    "note": res["note"],
                    "fig_bytes": res["fig_bytes"]
                })
                st.toast("『クロス集計結果』をレポートリストに追加しました！", icon="📋")
                st.rerun()

            col_left, col_right = st.columns([2.5, 2])
            with col_left:
                st.write(f"**クロス度数表 ({res['row_var']} × {res['col_var']})**")
                st.dataframe(res["ct_df"], use_container_width=True)
                st.caption(res["note"])
            with col_right:
                st.write("**構成比グラフ**")
                with st.container(height=450):
                    st.image(res["fig_bytes"], use_column_width=True)

    # ------------------------------------------------------------------
    # TAB 4: t検定 / Welch検定
    # ------------------------------------------------------------------
    with tab4:
        st.subheader("4. 2群の平均値の差の検定 (t-Test / Welch Test)")
        c1, c2, c3 = st.columns([1.5, 2, 1.5])
        with c1:
            group_var = st.selectbox("グループ変数 (2カテゴリ):", all_columns, key="tt_group")
        with c2:
            num_vars = st.multiselect("比較する従属変数 (複数選択可):", numeric_columns, default=numeric_columns[:min(3, len(numeric_columns))], key="tt_nums")
        with c3:
            equal_var_opt = st.selectbox("等分散性の仮定:", ["Welchのt検定 (推奨: 等分散非仮定)", "Studentのt検定 (等分散仮定)"])
            equal_var = True if "Student" in equal_var_opt else False
            
        if st.button("🚀 t検定を実行", key="run_tt", type="primary"):
            if group_var and num_vars:
                results_dict = {}
                errors = []
                with st.spinner(f"{len(num_vars)} 個の従属変数のt検定を計算中..."):
                    for nv in num_vars:
                        try:
                            res_df, note, fig_bytes = stats_engine.analyze_ttest(df, group_var, nv, equal_var=equal_var)
                            results_dict[nv] = {
                                "res_df": res_df, "note": note, "fig_bytes": fig_bytes, "num_var": nv, "group_var": group_var
                            }
                        except Exception as e:
                            errors.append(f"{nv}: {e}")
                
                if results_dict:
                    st.session_state["results"]["tt"] = {
                        "results_dict": results_dict, "group_var": group_var
                    }
                    st.toast(f"{len(results_dict)} 個の従属変数のt検定が完了しました！", icon="✅")
                if errors:
                    for err in errors:
                        st.error(f"分析エラー ({err})")
            else:
                st.warning("グループ変数と1つ以上の従属変数を選択してください。")
                    
        if "tt" in st.session_state["results"]:
            res_data = st.session_state["results"]["tt"]
            results_dict = res_data.get("results_dict", {})
            if not results_dict and "res_df" in res_data:
                results_dict = {res_data["num_var"]: res_data}
                
            if results_dict:
                grp_v = res_data.get("group_var", "")
                c_btn1, _ = st.columns([2, 1])
                with c_btn1:
                    if st.button(f"➕ 選択した全従属変数 ({len(results_dict)}件) の結果をExcelレポートに追加", key="btn_add_tt"):
                        for nv, res in results_dict.items():
                            st.session_state["analysis_queue"].append({
                                "sheet_name": f"t検定_{nv}",
                                "title": f"2群の平均値の比較 (t検定: {nv} × {grp_v})",
                                "df": res["res_df"],
                                "note": res["note"],
                                "fig_bytes": res["fig_bytes"]
                            })
                        st.toast(f"『t検定結果 ({len(results_dict)}件)』をレポートリストに追加しました！", icon="📋")
                        st.rerun()

                active_nv = st.selectbox("表示する従属変数の切替:", list(results_dict.keys()), key="select_tt_display")
                if active_nv in results_dict:
                    res = results_dict[active_nv]
                    col_left, col_right = st.columns([2.5, 2])
                    with col_left:
                        st.write(f"**t検定結果 ({res['num_var']} × {res['group_var']})**")
                        st.dataframe(res["res_df"], use_container_width=True)
                        st.caption(res["note"])
                    with col_right:
                        st.write("**平均値と比較グラフ (95%信頼区間)**")
                        with st.container(height=450):
                            st.image(res["fig_bytes"], use_column_width=True)

    # ------------------------------------------------------------------
    # TAB 5: 分散分析 (ANOVA)
    # ------------------------------------------------------------------
    with tab5:
        st.subheader("5. 一元配置分散分析 & 多重比較 (One-Way ANOVA)")
        c1, c2 = st.columns([1.5, 2])
        with c1:
            anova_group = st.selectbox("要因 (グループ変数):", all_columns, key="anova_grp")
        with c2:
            anova_nums = st.multiselect("従属変数 (複数選択可):", numeric_columns, default=numeric_columns[:min(3, len(numeric_columns))], key="anova_nums")
            
        if st.button("🚀 分散分析(ANOVA)を実行", key="run_anova", type="primary"):
            if anova_group and anova_nums:
                results_dict = {}
                errors = []
                with st.spinner(f"{len(anova_nums)} 個の従属変数の分散分析を実行中..."):
                    for nv in anova_nums:
                        try:
                            desc_df, tukey_df, note, fig_bytes = stats_engine.analyze_anova(df, anova_group, nv)
                            results_dict[nv] = {
                                "desc_df": desc_df, "tukey_df": tukey_df, "note": note, "fig_bytes": fig_bytes,
                                "anova_num": nv, "anova_group": anova_group
                            }
                        except Exception as e:
                            errors.append(f"{nv}: {e}")
                
                if results_dict:
                    st.session_state["results"]["anova"] = {
                        "results_dict": results_dict, "anova_group": anova_group
                    }
                    st.toast(f"{len(results_dict)} 個の従属変数の分散分析が完了しました！", icon="✅")
                if errors:
                    for err in errors:
                        st.error(f"分析エラー ({err})")
            else:
                st.warning("要因変数と1つ以上の従属変数を選択してください。")
                
        if "anova" in st.session_state["results"]:
            res_data = st.session_state["results"]["anova"]
            results_dict = res_data.get("results_dict", {})
            if not results_dict and "desc_df" in res_data:
                results_dict = {res_data["anova_num"]: res_data}
                
            if results_dict:
                grp_v = res_data.get("anova_group", "")
                c_btn1, _ = st.columns([2, 1])
                with c_btn1:
                    if st.button(f"➕ 選択した全従属変数 ({len(results_dict)}件) の結果をExcelレポートに追加", key="btn_add_anova"):
                        for nv, res in results_dict.items():
                            st.session_state["analysis_queue"].append({
                                "sheet_name": f"分散分析_{nv}",
                                "title": f"一元配置分散分析表 ({nv} × {grp_v})",
                                "df": res["desc_df"],
                                "note": res["note"],
                                "fig_bytes": res["fig_bytes"]
                            })
                            st.session_state["analysis_queue"].append({
                                "sheet_name": f"多重比較_{nv}",
                                "title": f"Tukey HSD 多重比較結果 ({nv})",
                                "df": res["tukey_df"],
                                "note": "注. 有意水準 alpha = .05 におけるTukeyのHSD検定結果。"
                            })
                        st.toast(f"『分散分析 & 多重比較結果 ({len(results_dict)}件)』をレポートリストに追加しました！", icon="📋")
                        st.rerun()

                active_nv = st.selectbox("表示する従属変数の切替:", list(results_dict.keys()), key="select_anova_display")
                if active_nv in results_dict:
                    res = results_dict[active_nv]
                    c_l, c_r = st.columns([2.5, 2])
                    with c_l:
                        st.write("**記述統計量 (各群の平均値と標準偏差)**")
                        st.dataframe(res["desc_df"], use_container_width=True)
                        st.caption(res["note"])
                        
                        st.write("**Tukey HSD 多重比較結果**")
                        st.dataframe(res["tukey_df"], use_container_width=True)
                    with c_r:
                        st.write("**平均値比較プロット**")
                        with st.container(height=450):
                            st.image(res["fig_bytes"], use_column_width=True)
            with c_r:
                st.write("**平均値比較プロット**")
                with st.container(height=450):
                    st.image(res["fig_bytes"], use_column_width=True)

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
            
        if st.button("🚀 相関分析を実行", key="run_corr", type="primary"):
            if len(corr_vars) >= 2:
                with st.spinner("相関係数行列とヒートマップを生成中..."):
                    res_df, note, fig_bytes = stats_engine.analyze_correlation(df, corr_vars, method=corr_method)
                    st.session_state["results"]["corr"] = {
                        "res_df": res_df, "note": note, "fig_bytes": fig_bytes, "corr_method": corr_method
                    }
                st.toast("相関分析が完了しました！", icon="✅")
            else:
                st.warning("相関分析には2つ以上の変数を選択してください。")
                
        if "corr" in st.session_state["results"]:
            res = st.session_state["results"]["corr"]
            
            if st.button("➕ この結果をExcelレポートに追加", key="btn_add_corr"):
                method_jp = "ピアソン" if res["corr_method"] == "pearson" else "スピアマン"
                st.session_state["analysis_queue"].append({
                    "sheet_name": f"相関分析_{res['corr_method']}",
                    "title": f"{method_jp} 相関係数行列",
                    "df": res["res_df"],
                    "note": res["note"],
                    "fig_bytes": res["fig_bytes"]
                })
                st.toast("『相関係数行列』をレポートリストに追加しました！", icon="📋")
                st.rerun()

            c_l, c_r = st.columns([2.5, 2])
            method_jp = "ピアソン" if res["corr_method"] == "pearson" else "スピアマン"
            with c_l:
                st.write(f"**{method_jp} 相関係数行列**")
                st.dataframe(res["res_df"], use_container_width=True)
                st.caption(res["note"])
            with c_r:
                st.write("**相関ヒートマップ**")
                with st.container(height=450):
                    st.image(res["fig_bytes"], use_column_width=True)

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
            
        if st.button("🚀 回帰分析を実行", key="run_reg", type="primary"):
            if target_var and feature_vars:
                with st.spinner("重回帰モデルを推定中..."):
                    res_df, note, fig_bytes = stats_engine.analyze_regression(df, target_var, feature_vars)
                    st.session_state["results"]["reg"] = {
                        "res_df": res_df, "note": note, "fig_bytes": fig_bytes, "target_var": target_var
                    }
                st.toast("重回帰分析が完了しました！", icon="✅")
            else:
                st.warning("目的変数と1つ以上の説明変数を選択してください。")
                
        if "reg" in st.session_state["results"]:
            res = st.session_state["results"]["reg"]
            
            if st.button("➕ この結果をExcelレポートに追加", key="btn_add_reg"):
                st.session_state["analysis_queue"].append({
                    "sheet_name": f"回帰分析_{res['target_var']}",
                    "title": f"{res['target_var']} を目的変数とする重回帰分析表",
                    "df": res["res_df"],
                    "note": res["note"],
                    "fig_bytes": res["fig_bytes"]
                })
                st.toast("『重回帰分析結果』をレポートリストに追加しました！", icon="📋")
                st.rerun()

            c_l, c_r = st.columns([2.5, 2])
            with c_l:
                st.write(f"**重回帰モデル分析結果表 (目的変数: {res['target_var']})**")
                st.dataframe(res["res_df"], use_container_width=True)
                st.caption(res["note"])
            with c_r:
                st.write("**実測値 vs. 予測値プロット**")
                with st.container(height=450):
                    st.image(res["fig_bytes"], use_column_width=True)

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
            
        if st.button("🚀 因子分析を実行", key="run_fa", type="primary"):
            if len(fa_vars) >= 3:
                with st.spinner("因子分析とスクリープロットを実行中..."):
                    res_df, note, fig_bytes = stats_engine.analyze_factor_analysis(df, fa_vars, n_factors=int(n_factors), rotation=rotation)
                    st.session_state["results"]["fa"] = {
                        "res_df": res_df, "note": note, "fig_bytes": fig_bytes, "rotation": rotation
                    }
                st.toast("因子分析が完了しました！", icon="✅")
            else:
                st.warning("因子分析には3つ以上の観測変数を選択してください。")
                
        if "fa" in st.session_state["results"]:
            res = st.session_state["results"]["fa"]
            
            if st.button("➕ この結果をExcelレポートに追加", key="btn_add_fa"):
                rot_jp = "プロマックス回転" if res["rotation"] == "promax" else "バリマックス回転"
                st.session_state["analysis_queue"].append({
                    "sheet_name": "因子分析",
                    "title": f"探索的因子分析結果 ({rot_jp})",
                    "df": res["res_df"],
                    "note": res["note"],
                    "fig_bytes": res["fig_bytes"]
                })
                st.toast("『因子分析結果』をレポートリストに追加しました！", icon="📋")
                st.rerun()

            c_l, c_r = st.columns([2.5, 2])
            with c_l:
                rot_jp = "プロマックス回転" if res["rotation"] == "promax" else "バリマックス回転"
                st.write(f"**因子負荷量行列 & 寄与率 ({rot_jp})**")
                st.dataframe(res["res_df"], use_container_width=True)
                st.caption(res["note"])
            with c_r:
                st.write("**スクリープロット**")
                with st.container(height=450):
                    st.image(res["fig_bytes"], use_column_width=True)


if __name__ == "__main__":
    main()
