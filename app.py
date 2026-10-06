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


def load_data_file(uploaded_file, encoding_choice="自動判定 (Auto)", sheet_name=0):
    """CSVまたはExcelファイル (.xlsx, .xls) の読み込み"""
    file_name = uploaded_file.name.lower()
    
    if file_name.endswith(".xlsx") or file_name.endswith(".xls"):
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
        return df, f"Excel ({sheet_name})"
    else:  # CSV
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
    st.markdown('<div class="sub-title">CSV/Excelファイルをアップロード後、リコードや特定値の除外などのデータ加工を行い、変数と分析方法を選択して「分析を実行」ボタンでAPAスタイル（第7版）の表・図をExcel出力します。</div>', unsafe_allow_html=True)
    
    # ------------------------------------------------------------------
    # サイドバー：ファイルアップロード & 出力レポート管理
    # ------------------------------------------------------------------
    with st.sidebar:
        st.header("📂 データアップロード")
        uploaded_file = st.file_uploader("ファイルを選択 (CSV / Excel)", type=["csv", "xlsx", "xls"])
        
        encoding_choice = "自動判定 (Auto)"
        selected_sheet = 0
        if uploaded_file is not None:
            fname = uploaded_file.name.lower()
            if fname.endswith(".csv"):
                encoding_choice = st.selectbox(
                    "文字コード (CSV Encoding)",
                    ["自動判定 (Auto)", "utf-8", "cp932 (Windows Shift-JIS)", "shift_jis", "utf-8-sig"]
                )
            elif fname.endswith(".xlsx") or fname.endswith(".xls"):
                try:
                    excel_file = pd.ExcelFile(uploaded_file)
                    sheet_names = excel_file.sheet_names
                    if len(sheet_names) > 1:
                        selected_sheet = st.selectbox("読み込むシートを選択:", sheet_names)
                    else:
                        selected_sheet = sheet_names[0]
                except Exception:
                    selected_sheet = 0
        
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
            st.info("👈 サイドバーから分析したいCSVまたはExcelファイル（.xlsx / .xls）をアップロードしてください。")
            with st.expander("💡 演習用サンプルデータで試す (有意差・相関が出る実践データ)", expanded=True):
                st.markdown("""
                **統計分析の演習や動作確認に最適なサンプルデータ（$N=150$ 人の心理・教育実験データ）を生成できます。**
                - **t検定 / 分散分析**: 実験グループ（統制群・講義群・映像実験群）で事後テスト得点に明瞭な有意差あり
                - **対応のあるt検定**: 事前テスト vs 事後テストで有意な学習効果あり
                - **二元配置ANOVA**: 実験グループ × 指導経験（あり/なし）の主効果および交互作用あり
                - **相関 / 重回帰**: 事後テスト得点と学習満足度の間に中程度の正の相関 ($r \\approx 0.60$) あり
                - **ロジスティック回帰**: 事後テスト得点が合格判定（合格/不合格）に有意に寄与
                - **因子分析 / 信頼性分析**: 2因子構造（学習意欲・理解度）と高い内的一貫性 ($\alpha \\approx 0.85$)
                - **前処理演習**: 逆転項目（`Q7_退屈さ_逆転項目`）、無効値（`-99`）、欠損値（`NaN`）を適度に含有
                """)
                
                c_samp1, c_samp2 = st.columns(2)
                with c_samp1:
                    if st.button("🎲 毎回異なるランダムなサンプルデータを生成", type="primary", use_container_width=True):
                        with st.spinner("実践演習用サンプルデータを生成中..."):
                            sample_data = stats_engine.generate_sample_dataset(n=150, seed=None)
                            st.session_state["raw_df"] = sample_data.copy()
                            st.session_state["df"] = sample_data.copy()
                            st.session_state["last_uploaded"] = "sample_data_random"
                        st.toast("ランダムな演習用サンプルデータを読み込みました！", icon="🚀")
                        st.rerun()
                with c_samp2:
                    if st.button("🎯 再現可能な標準サンプルデータを生成 (Seed=42)", use_container_width=True):
                        with st.spinner("標準サンプルデータを生成中..."):
                            sample_data = stats_engine.generate_sample_dataset(n=150, seed=42)
                            st.session_state["raw_df"] = sample_data.copy()
                            st.session_state["df"] = sample_data.copy()
                            st.session_state["last_uploaded"] = "sample_data_fixed"
                        st.toast("標準サンプルデータを読み込みました！", icon="🎯")
                        st.rerun()
            return
        else:
            df = st.session_state["df"]
    else:
        file_identifier = f"{uploaded_file.name}_{selected_sheet}_{uploaded_file.size}"
        if "last_uploaded" not in st.session_state or st.session_state["last_uploaded"] != file_identifier:
            try:
                with st.spinner("データを読み込み中..."):
                    df_raw, load_info = load_data_file(uploaded_file, encoding_choice=encoding_choice, sheet_name=selected_sheet)
                    st.session_state["raw_df"] = df_raw.copy()
                    st.session_state["df"] = df_raw.copy()
                    st.session_state["last_uploaded"] = file_identifier
                st.toast(f"✅ {uploaded_file.name} ({load_info}) を読み込みました！", icon="📂")
            except Exception as e:
                st.error(f"ファイルの読み込みエラー: {e}")
                return
        df = st.session_state["df"]

    st.caption(f"現在の分析対象データ: {df.shape[0]}行 × {df.shape[1]}列")

    # ------------------------------------------------------------------
    # データ前処理・リコード・フィルタリング セクション
    # ------------------------------------------------------------------
    with st.expander("🛠️ データの加工・Recode（再符号化）・特定値の除外・欠損値処理", expanded=False):
        t_proc1, t_proc2, t_proc3, t_proc4, t_proc5, t_proc6, t_proc7, t_proc8, t_proc9 = st.tabs([
            "① 特定値の除外",
            "② 値のリコード (置換)",
            "③ 数値変数のカテゴリ化 (ビン分割)",
            "④ 合成スコア (合計・平均)",
            "⑤ 逆転項目の反転",
            "⑥ 標準化 / 正規化",
            "⑦ 欠損値処理",
            "⑧ 複数条件フィルタ",
            "⑨ ダミー変数化"
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

        # --- ⑤ 逆転項目の反転 ---
        with t_proc5:
            st.markdown("**アンケートの逆転項目（例: 1~5件法で 1↔5, 2↔4 に反転）を一括処理します。**")
            if num_cols_only:
                c_r1, c_r2, c_r3 = st.columns(3)
                with c_r1:
                    rev_vars = st.multiselect("反転する変数を選択:", num_cols_only, key="rev_vars")
                with c_r2:
                    min_val = st.number_input("尺度の最小値 (例: 1):", value=1.0, key="rev_min")
                with c_r3:
                    max_val = st.number_input("尺度の最大値 (例: 5や7):", value=5.0, key="rev_max")
                
                rev_suffix = st.text_input("新変数の末尾プレフィックス (例: _rev):", value="_rev", key="rev_suffix")
                st.caption(f"計算式: 新値 = ({min_val} + {max_val}) - 元の値 = {min_val + max_val} - 元の値")
                
                if st.button("🔄 逆転項目の反転を実行", key="btn_apply_rev"):
                    if rev_vars:
                        new_df, created_names = stats_engine.reverse_code_values(df, rev_vars, min_val, max_val, rev_suffix)
                        st.session_state["df"] = new_df
                        st.toast(f"{len(created_names)} 個の反転変数を作成しました！ ({', '.join(created_names)})", icon="✨")
                        st.rerun()
                    else:
                        st.warning("1つ以上の変数を選択してください。")

        # --- ⑥ 標準化 / 正規化 ---
        with t_proc6:
            st.markdown("**数値変数を「標準化（Zスコア: 平均0, 分散1）」または「正規化（0〜1スケーリング）」します。**")
            if num_cols_only:
                c_s1, c_s2 = st.columns(2)
                with c_s1:
                    std_vars = st.multiselect("変換する数値変数を選択:", num_cols_only, key="std_vars")
                with c_s2:
                    std_method = st.radio("変換方法:", ["標準化 (Zスコア化)", "正規化 (0-1 Min-Max)"], key="std_method")
                
                method_type = "standardize" if "標準化" in std_method else "normalize"
                default_suffix = "_z" if method_type == "standardize" else "_norm"
                std_suffix = st.text_input("新変数の末尾 (suffix):", value=default_suffix, key="std_suffix")
                
                if st.button("📐 標準化 / 正規化を実行", key="btn_apply_std"):
                    if std_vars:
                        new_df, created_names = stats_engine.standardize_normalize_variables(df, std_vars, method=method_type, suffix=std_suffix)
                        st.session_state["df"] = new_df
                        st.toast(f"{len(created_names)} 個の変換変数を作成しました！ ({', '.join(created_names)})", icon="✨")
                        st.rerun()
                    else:
                        st.warning("1つ以上の変数を選択してください。")

        # --- ⑦ 欠損値処理 ---
        with t_proc7:
            st.markdown("**データ内の欠損値（NaN/空白）を削除または補完します。**")
            c_m1, c_m2 = st.columns(2)
            with c_m1:
                na_strategy = st.selectbox(
                    "処理方法を選択:",
                    [
                        "リストワイズ削除 (欠損のある行を削除)",
                        "平均値で補完 (数値変数のみ)",
                        "中央値で補完 (数値変数のみ)",
                        "最頻値で補完 (カテゴリ/数値)",
                        "指定の固定値で補完"
                    ],
                    key="na_strategy"
                )
            with c_m2:
                na_target_vars = st.multiselect("対象とする変数 (未指定の場合は全変数):", df.columns, key="na_vars")
                
            fill_val = None
            if "指定の固定値" in na_strategy:
                fill_val_str = st.text_input("補完する固定値を入力:", value="0", key="na_fill_val")
                try:
                    fill_val = float(fill_val_str)
                except ValueError:
                    fill_val = fill_val_str
                    
            if st.button("🧹 欠損値処理を実行", key="btn_apply_na"):
                strat_map = {
                    "リストワイズ削除 (欠損のある行を削除)": "listwise",
                    "平均値で補完 (数値変数のみ)": "mean",
                    "中央値で補完 (数値変数のみ)": "median",
                    "最頻値で補完 (カテゴリ/数値)": "mode",
                    "指定の固定値で補完": "constant"
                }
                strat_code = strat_map[na_strategy]
                target_v = na_target_vars if na_target_vars else None
                new_df = stats_engine.handle_missing_values(df, strategy=strat_code, target_vars=target_v, fill_val=fill_val)
                st.session_state["df"] = new_df
                st.toast(f"欠損値処理を適用しました ({len(df)}行 → {len(new_df)}行)", icon="🧼")
                st.rerun()

        # --- ⑧ 複数条件フィルタ ---
        with t_proc8:
            st.markdown("**複数条件（AND / OR）を組み合わせて分析対象データを抽出します。**")
            filt_logic = st.radio("条件の結合方法:", ["すべての条件を満たす (AND)", "いずれかの条件を満たす (OR)"], horizontal=True, key="filt_logic")
            logic_code = "AND" if "AND" in filt_logic else "OR"
            
            if "num_conditions" not in st.session_state:
                st.session_state["num_conditions"] = 1
                
            c_f_btn1, c_f_btn2, _ = st.columns([1, 1, 3])
            with c_f_btn1:
                if st.button("➕ 条件を追加", key="btn_add_cond"):
                    st.session_state["num_conditions"] += 1
                    st.rerun()
            with c_f_btn2:
                if st.button("➖ 条件を減らす", key="btn_sub_cond") and st.session_state["num_conditions"] > 1:
                    st.session_state["num_conditions"] -= 1
                    st.rerun()
                    
            conditions = []
            for c_idx in range(st.session_state["num_conditions"]):
                cf1, cf2, cf3 = st.columns([2, 1.5, 2.5])
                with cf1:
                    col_name = st.selectbox(f"条件{c_idx+1} 変数:", df.columns, key=f"f_col_{c_idx}")
                with cf2:
                    op = st.selectbox(f"条件{c_idx+1} 演算子:", ["==", "!=", ">", ">=", "<", "<=", "contains", "in"], key=f"f_op_{c_idx}")
                with cf3:
                    val_str = st.text_input(f"条件{c_idx+1} 比較値 (カンマ区切り可):", value="", key=f"f_val_{c_idx}")
                if val_str:
                    conditions.append((col_name, op, val_str))
                    
            if st.button("🔍 フィルタを実行してデータを抽出", key="btn_apply_filter"):
                if conditions:
                    new_df = stats_engine.filter_advanced(df, conditions, logic=logic_code)
                    st.session_state["df"] = new_df
                    st.toast(f"フィルタを適用しました ({len(df)}行 → {len(new_df)}行)", icon="🎯")
                    st.rerun()
                else:
                    st.warning("比較値を入力してください。")

        # --- ⑨ ダミー変数化 ---
        with t_proc9:
            st.markdown("**カテゴリ変数（名義尺度）を 0 と 1 のダミー変数（One-Hot Encoding）に変換します。**")
            cat_vars = [c for c in df.columns if c not in num_cols_only or df[c].nunique() <= 10]
            if cat_vars:
                c_d1, c_d2 = st.columns(2)
                with c_d1:
                    dummy_var = st.selectbox("ダミー変数化するカテゴリ変数:", cat_vars, key="dummy_var")
                    dummy_prefix = st.text_input("プレフィックス (変数名の接頭辞):", value=dummy_var, key="dummy_prefix")
                with c_d2:
                    drop_first = st.checkbox("最初のカテゴリを変数から除外 (参照カテゴリ・多重共線性対策)", value=False, key="dummy_drop_first")
                    st.caption("※ 回帰分析等で多重共線性（マルチコ）を防ぐ場合はチェックを推奨します。")
                    
                if st.button("🏷️ ダミー変数を作成", key="btn_apply_dummy"):
                    new_df, created_cols = stats_engine.create_dummy_variables(df, dummy_var, drop_first=drop_first, prefix=dummy_prefix)
                    st.session_state["df"] = new_df
                    st.toast(f"{len(created_cols)} 個のダミー変数を作成しました！ ({', '.join(created_cols)})", icon="✨")
                    st.rerun()

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
    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10, tab11, tab12 = st.tabs([
        "1. 単純集計",
        "2. 基本統計量",
        "3. クロス集計",
        "4. 独立t検定",
        "5. 対応ありt検定",
        "6. 一元配置ANOVA",
        "7. 二元配置ANOVA",
        "8. 相関分析",
        "9. 重回帰分析",
        "10. ロジスティック回帰",
        "11. 因子分析",
        "12. 信頼性分析 (α)"
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
                    res_df, hist_dict, note, fig_bytes = stats_engine.analyze_descriptives(df, selected_num_vars)
                    st.session_state["results"]["desc"] = {
                        "res_df": res_df, "hist_dict": hist_dict, "note": note, "fig_bytes": fig_bytes
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
                    "note": res["note"]
                })
                for var, hist_df in res.get("hist_dict", {}).items():
                    st.session_state["analysis_queue"].append({
                        "sheet_name": f"分布_{var}",
                        "title": f"{var} の度数分布 (ヒストグラム)",
                        "df": hist_df,
                        "note": f"注. {var} の階級別度数分布。"
                    })
                st.toast("『基本統計量 & ヒストグラム分布』をレポートリストに追加しました！", icon="📋")
                st.rerun()

            c1, c2 = st.columns([2.5, 2])
            with c1:
                st.write("**基本統計量一覧**")
                st.dataframe(res["res_df"], use_container_width=True)
                st.caption(res["note"])
            with c2:
                st.write("**分布プロット (ヒストグラム & 確率密度)**")
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
                    ct_formatted_df, pct_df, note, fig_bytes = stats_engine.analyze_crosstab(df, row_var, col_var)
                    st.session_state["results"]["ct"] = {
                        "ct_formatted_df": ct_formatted_df, "pct_df": pct_df, "note": note, "fig_bytes": fig_bytes,
                        "row_var": row_var, "col_var": col_var
                    }
                st.toast("クロス集計が完了しました！", icon="✅")
                
        if "ct" in st.session_state["results"]:
            res = st.session_state["results"]["ct"]
            
            if st.button("➕ この結果をExcelレポートに追加", key="btn_add_ct"):
                st.session_state["analysis_queue"].append({
                    "sheet_name": f"クロス_{res['row_var']}_vs_{res['col_var']}",
                    "title": f"クロス集計表 ({res['row_var']} × {res['col_var']})",
                    "df": res["ct_formatted_df"],
                    "note": res["note"]
                })
                st.session_state["analysis_queue"].append({
                    "sheet_name": f"構成比_{res['row_var']}_vs_{res['col_var']}",
                    "title": f"構成比 (%) 表 ({res['row_var']} × {res['col_var']})",
                    "df": res["pct_df"],
                    "note": "注. 数値は列方向の構成比 (%) を表します。"
                })
                st.toast("『クロス集計表 & 構成比グラフ』をレポートリストに追加しました！", icon="📋")
                st.rerun()

            col_left, col_right = st.columns([2.5, 2])
            with col_left:
                st.write(f"**クロス度数表 [度数 (列%)] ({res['row_var']} × {res['col_var']})**")
                st.dataframe(res["ct_formatted_df"], use_container_width=True)
                st.caption(res["note"])
            with col_right:
                st.write("**構成比グラフ (100% 積み上げ)**")
                with st.container(height=450):
                    st.image(res["fig_bytes"], use_column_width=True)

    # ------------------------------------------------------------------
    # TAB 4: 独立2群のt検定
    # ------------------------------------------------------------------
    with tab4:
        st.subheader("4. 独立2群の平均値の差の検定 (Independent Samples t-Test / Welch)")
        c1, c2, c3 = st.columns([1.5, 2, 1.5])
        with c1:
            group_var = st.selectbox("グループ変数 (2カテゴリ):", all_columns, key="tt_group")
        with c2:
            num_vars = st.multiselect("比較する従属変数 (複数選択可):", numeric_columns, default=numeric_columns[:min(3, len(numeric_columns))], key="tt_nums")
        with c3:
            equal_var_opt = st.selectbox("等分散性の仮定:", ["Welchのt検定 (推奨: 等分散非仮定)", "Studentのt検定 (等分散仮定)"], key="tt_eq_opt")
            equal_var = True if "Student" in equal_var_opt else False
            
        if st.button("🚀 独立t検定を実行", key="run_tt", type="primary"):
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
                combined_df = pd.concat([res["res_df"] for res in results_dict.values()])
                combined_notes = " | ".join([f"{nv}: {res['note']}" for nv, res in results_dict.items()]) if len(results_dict) > 1 else list(results_dict.values())[0]["note"]
                
                c_btn1, _ = st.columns([2, 1])
                with c_btn1:
                    if st.button(f"➕ 選択した全従属変数 ({len(results_dict)}件) の結果をExcelレポートに追加", key="btn_add_tt"):
                        st.session_state["analysis_queue"].append({
                            "sheet_name": f"t検定_{grp_v}",
                            "title": f"2群の平均値の比較 (t検定: {grp_v})",
                            "df": combined_df,
                            "note": combined_notes,
                            "fig_bytes": list(results_dict.values())[0]["fig_bytes"]
                        })
                        st.toast(f"『t検定一括結果 ({len(results_dict)}件)』をレポートリストに追加しました！", icon="📋")
                        st.rerun()

                st.markdown(f"**📊 t検定 要約一覧表 (グループ変数: {grp_v})**")
                st.dataframe(combined_df, use_container_width=True)

                st.divider()
                active_nv = st.selectbox("個別グラフ表示の従属変数を選択:", list(results_dict.keys()), key="select_tt_display")
                if active_nv in results_dict:
                    res = results_dict[active_nv]
                    col_left, col_right = st.columns([2.5, 2])
                    with col_left:
                        st.write(f"**t検定詳細 ({res['num_var']} × {res['group_var']})**")
                        st.dataframe(res["res_df"], use_container_width=True)
                        st.caption(res["note"])
                    with col_right:
                        st.write("**平均値比較グラフ (95%信頼区間)**")
                        with st.container(height=450):
                            st.image(res["fig_bytes"], use_column_width=True)

    # ------------------------------------------------------------------
    # TAB 5: 対応のあるt検定 (Paired t-test)
    # ------------------------------------------------------------------
    with tab5:
        st.subheader("5. 対応のある2群の平均値の差の検定 (Paired Samples t-Test)")
        st.caption("同一の被験者における前後比較（例: 事前テスト vs 事後テスト）や対応する2条件間の平均値の差を検定します。")
        c1, c2 = st.columns(2)
        with c1:
            pair_var1 = st.selectbox("条件1 (例: 事前):", numeric_columns, key="pt_v1")
        with c2:
            pair_var2 = st.selectbox("条件2 (例: 事後):", [c for c in numeric_columns if c != pair_var1], key="pt_v2")
            
        if st.button("🚀 対応のあるt検定を実行", key="run_pt", type="primary"):
            if pair_var1 and pair_var2:
                try:
                    with st.spinner("対応のあるt検定を計算中..."):
                        pt_df, note, fig_bytes = stats_engine.analyze_paired_ttest(df, pair_var1, pair_var2)
                        st.session_state["results"]["pt"] = {
                            "pt_df": pt_df, "note": note, "fig_bytes": fig_bytes,
                            "pair_var1": pair_var1, "pair_var2": pair_var2
                        }
                    st.toast("対応のあるt検定が完了しました！", icon="✅")
                except Exception as e:
                    st.error(f"分析エラー: {e}")
            else:
                st.warning("2つの数値変数を選択してください。")
                
        if "pt" in st.session_state["results"]:
            res = st.session_state["results"]["pt"]
            
            if st.button("➕ この結果をExcelレポートに追加", key="btn_add_pt"):
                st.session_state["analysis_queue"].append({
                    "sheet_name": f"対応ありt_{res['pair_var1']}_vs_{res['pair_var2']}",
                    "title": f"対応のあるt検定結果表 ({res['pair_var1']} vs. {res['pair_var2']})",
                    "df": res["pt_df"],
                    "note": res["note"],
                    "fig_bytes": res["fig_bytes"]
                })
                st.toast("『対応のあるt検定結果』をレポートリストに追加しました！", icon="📋")
                st.rerun()

            c_l, c_r = st.columns([2.5, 2])
            with c_l:
                st.write(f"**対応のあるt検定 結果表 ({res['pair_var1']} vs. {res['pair_var2']})**")
                st.dataframe(res["pt_df"], use_container_width=True)
                st.caption(res["note"])
            with c_r:
                st.write("**平均値の変化プロット (95%信頼区間)**")
                with st.container(height=450):
                    st.image(res["fig_bytes"], use_column_width=True)

    # ------------------------------------------------------------------
    # TAB 6: 一元配置分散分析 (ANOVA)
    # ------------------------------------------------------------------
    with tab6:
        st.subheader("6. 一元配置分散分析 & 多重比較 (One-Way ANOVA)")
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
                            anova_df, tukey_df, note, fig_bytes = stats_engine.analyze_anova(df, anova_group, nv)
                            results_dict[nv] = {
                                "anova_df": anova_df, "tukey_df": tukey_df, "note": note, "fig_bytes": fig_bytes,
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
            if not results_dict and "anova_df" in res_data:
                results_dict = {res_data["anova_num"]: res_data}
                
            if results_dict:
                grp_v = res_data.get("anova_group", "")
                combined_anova_df = pd.concat([res["anova_df"] for res in results_dict.values()])
                combined_notes = " | ".join([f"{nv}: {res['note']}" for nv, res in results_dict.items()]) if len(results_dict) > 1 else list(results_dict.values())[0]["note"]
                
                tukey_list = []
                for nv, res in results_dict.items():
                    tdf = res["tukey_df"].reset_index()
                    tdf.insert(0, "従属変数", nv)
                    tukey_list.append(tdf)
                combined_tukey_df = pd.concat(tukey_list).set_index(["従属変数", "グループ1", "グループ2"]) if tukey_list else None
                
                c_btn1, _ = st.columns([2, 1])
                with c_btn1:
                    if st.button(f"➕ 選択した全従属変数 ({len(results_dict)}件) の結果をExcelレポートに追加", key="btn_add_anova"):
                        st.session_state["analysis_queue"].append({
                            "sheet_name": f"分散分析_{grp_v}",
                            "title": f"一元配置分散分析表 ({grp_v})",
                            "df": combined_anova_df,
                            "note": combined_notes,
                            "fig_bytes": list(results_dict.values())[0]["fig_bytes"]
                        })
                        if combined_tukey_df is not None:
                            st.session_state["analysis_queue"].append({
                                "sheet_name": f"多重比較_{grp_v}",
                                "title": f"Tukey HSD 多重比較結果 ({grp_v})",
                                "df": combined_tukey_df,
                                "note": "注. 有意水準 alpha = .05 におけるTukeyのHSD検定結果。"
                            })
                        st.toast(f"『分散分析 & 多重比較一括結果 ({len(results_dict)}件)』をレポートリストに追加しました！", icon="📋")
                        st.rerun()

                st.markdown(f"**📊 分散分析 要約一覧表 (要因: {grp_v})**")
                st.dataframe(combined_anova_df, use_container_width=True)

                st.divider()
                active_nv = st.selectbox("個別結果表示の従属変数を選択:", list(results_dict.keys()), key="select_anova_display")
                if active_nv in results_dict:
                    res = results_dict[active_nv]
                    c_l, c_r = st.columns([2.5, 2])
                    with c_l:
                        st.write(f"**分散分析詳細 ({res['anova_num']} × {res['anova_group']})**")
                        st.dataframe(res["anova_df"], use_container_width=True)
                        st.caption(res["note"])
                        
                        st.write("**Tukey HSD 多重比較結果**")
                        st.dataframe(res["tukey_df"], use_container_width=True)
                    with c_r:
                        st.write("**平均値比較プロット**")
                        with st.container(height=450):
                            st.image(res["fig_bytes"], use_column_width=True)

    # ------------------------------------------------------------------
    # TAB 7: 二元配置分散分析 (Two-way ANOVA)
    # ------------------------------------------------------------------
    with tab7:
        st.subheader("7. 二元配置分散分析 (Two-Way ANOVA)")
        st.caption("2つの要因（独立変数）による主効果および交互作用効果を検定します。")
        c1, c2, c3 = st.columns(3)
        with c1:
            two_f1 = st.selectbox("要因1 (因子A):", all_columns, key="two_f1")
        with c2:
            two_f2 = st.selectbox("要因2 (因子B):", [c for c in all_columns if c != two_f1], key="two_f2")
        with c3:
            two_dep = st.selectbox("従属変数 (目的変数):", numeric_columns, key="two_dep")
            
        if st.button("🚀 二元配置分散分析を実行", key="run_two_anova", type="primary"):
            if two_f1 and two_f2 and two_dep:
                try:
                    with st.spinner("二元配置分散分析と交互作用プロットを生成中..."):
                        anova_df, cell_desc, note, fig_bytes = stats_engine.analyze_two_way_anova(df, two_f1, two_f2, two_dep)
                        st.session_state["results"]["two_anova"] = {
                            "anova_df": anova_df, "cell_desc": cell_desc, "note": note, "fig_bytes": fig_bytes,
                            "two_f1": two_f1, "two_f2": two_f2, "two_dep": two_dep
                        }
                    st.toast("二元配置分散分析が完了しました！", icon="✅")
                except Exception as e:
                    st.error(f"分析エラー: {e}")
            else:
                st.warning("要因1、要因2、および従属変数を選択してください。")
                
        if "two_anova" in st.session_state["results"]:
            res = st.session_state["results"]["two_anova"]
            
            if st.button("➕ この結果をExcelレポートに追加", key="btn_add_two_anova"):
                st.session_state["analysis_queue"].append({
                    "sheet_name": f"二元配置ANOVA_{res['two_dep']}",
                    "title": f"二元配置分散分析表 (従属変数: {res['two_dep']})",
                    "df": res["anova_df"],
                    "note": res["note"],
                    "fig_bytes": res["fig_bytes"],
                    "extra_df": res["cell_desc"],
                    "extra_title": f"セル別平均値 (M (SD): {res['two_f1']} × {res['two_f2']})"
                })
                st.toast("『二元配置分散分析結果』をレポートリストに追加しました！", icon="📋")
                st.rerun()

            c_l, c_r = st.columns([2.5, 2])
            with c_l:
                st.write(f"**二元配置分散分析表 (従属変数: {res['two_dep']})**")
                st.dataframe(res["anova_df"], use_container_width=True)
                st.caption(res["note"])
                
                st.write(f"**セル別記述統計量 [M (SD)] ({res['two_f1']} × {res['two_f2']})**")
                st.dataframe(res["cell_desc"], use_container_width=True)
            with c_r:
                st.write("**交互作用プロット (Interaction Plot)**")
                with st.container(height=450):
                    st.image(res["fig_bytes"], use_column_width=True)

    # ------------------------------------------------------------------
    # TAB 8: 相関分析
    # ------------------------------------------------------------------
    with tab8:
        st.subheader("8. 相関分析 (Correlation Analysis)")
        c1, c2 = st.columns([2.5, 1])
        with c1:
            corr_vars = st.multiselect("相関を計算する数値変数 (複数選択):", numeric_columns, default=numeric_columns[:min(4, len(numeric_columns))], key="corr_vars")
        with c2:
            corr_method = st.radio("相関係数の種類:", ["pearson", "spearman"], key="corr_method_opt")
            
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
    # TAB 9: 重回帰分析
    # ------------------------------------------------------------------
    with tab9:
        st.subheader("9. 重回帰分析 (Multiple Linear Regression)")
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
    # TAB 10: ロジスティック回帰分析
    # ------------------------------------------------------------------
    with tab10:
        st.subheader("10. 二項ロジスティック回帰分析 (Binary Logistic Regression)")
        st.caption("二値カテゴリ変数（例: 購買あり/なし、合格/不合格、0/1）の生起確率を予測し、オッズ比を算出します。")
        c1, c2 = st.columns([1, 2])
        with c1:
            logit_target = st.selectbox("目的変数 (二値カテゴリ):", all_columns, key="logit_target")
        with c2:
            logit_avail_features = [c for c in numeric_columns if c != logit_target]
            logit_features = st.multiselect("説明変数 (連続またはダミー変数):", logit_avail_features, default=logit_avail_features[:min(3, len(logit_avail_features))], key="logit_features")
            
        if st.button("🚀 ロジスティック回帰を実行", key="run_logit", type="primary"):
            if logit_target and logit_features:
                try:
                    with st.spinner("ロジスティック回帰モデルを推定中..."):
                        logit_df, note, fig_bytes = stats_engine.analyze_logistic_regression(df, logit_target, logit_features)
                        st.session_state["results"]["logit"] = {
                            "logit_df": logit_df, "note": note, "fig_bytes": fig_bytes,
                            "logit_target": logit_target
                        }
                    st.toast("ロジスティック回帰分析が完了しました！", icon="✅")
                except Exception as e:
                    st.error(f"分析エラー: {e}")
            else:
                st.warning("目的変数と1つ以上の説明変数を選択してください。")
                
        if "logit" in st.session_state["results"]:
            res = st.session_state["results"]["logit"]
            
            if st.button("➕ この結果をExcelレポートに追加", key="btn_add_logit"):
                st.session_state["analysis_queue"].append({
                    "sheet_name": f"ロジスティック_{res['logit_target']}",
                    "title": f"ロジスティック回帰分析結果表 (目的変数: {res['logit_target']})",
                    "df": res["logit_df"],
                    "note": res["note"],
                    "fig_bytes": res["fig_bytes"]
                })
                st.toast("『ロジスティック回帰結果』をレポートリストに追加しました！", icon="📋")
                st.rerun()

            c_l, c_r = st.columns([2.5, 2])
            with c_l:
                st.write(f"**ロジスティック回帰パラメータ表 (オッズ比・95%CI)**")
                st.dataframe(res["logit_df"], use_container_width=True)
                st.caption(res["note"])
            with c_r:
                st.write("**オッズ比 フォレストプロット (95%信頼区間)**")
                with st.container(height=450):
                    st.image(res["fig_bytes"], use_column_width=True)

    # ------------------------------------------------------------------
    # TAB 11: 因子分析
    # ------------------------------------------------------------------
    with tab11:
        st.subheader("11. 探索的因子分析 (Exploratory Factor Analysis)")
        c1, c2, c3 = st.columns(3)
        with c1:
            fa_vars = st.multiselect("因子分析に投入する観測変数:", numeric_columns, default=numeric_columns[:min(6, len(numeric_columns))], key="fa_vars")
        with c2:
            n_factors = st.number_input("抽出する因子数:", min_value=1, max_value=max(1, len(fa_vars)-1), value=min(2, max(1, len(fa_vars)-1)))
        with c3:
            rotation = st.selectbox("因子回転法:", ["promax", "varimax"], key="fa_rot")
            
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

    # ------------------------------------------------------------------
    # TAB 12: 尺度信頼性分析
    # ------------------------------------------------------------------
    with tab12:
        st.subheader("12. 尺度信頼性分析 (Scale Reliability / Cronbach's Alpha)")
        st.caption("アンケート尺度などの内的一貫性をクロンバックのα係数、項目-全体相関、項目削除時αで評価します。")
        rel_vars = st.multiselect("信頼性を分析する尺度項目 (数値変数):", numeric_columns, default=numeric_columns[:min(5, len(numeric_columns))], key="rel_vars")
        
        if st.button("🚀 信頼性分析を実行", key="run_rel", type="primary"):
            if len(rel_vars) >= 2:
                try:
                    with st.spinner("クロンバックのα係数と項目統計量を計算中..."):
                        rel_df, note, fig_bytes, alpha_val = stats_engine.analyze_reliability(df, rel_vars)
                        st.session_state["results"]["rel"] = {
                            "rel_df": rel_df, "note": note, "fig_bytes": fig_bytes, "alpha_val": alpha_val, "rel_vars": rel_vars
                        }
                    st.toast(f"信頼性分析が完了しました！ (全体 α = {alpha_val:.3f})", icon="✅")
                except Exception as e:
                    st.error(f"分析エラー: {e}")
            else:
                st.warning("信頼性分析には2つ以上の変数を選択してください。")
                
        if "rel" in st.session_state["results"]:
            res = st.session_state["results"]["rel"]
            
            if st.button("➕ この結果をExcelレポートに追加", key="btn_add_rel"):
                st.session_state["analysis_queue"].append({
                    "sheet_name": "信頼性分析",
                    "title": f"尺度項目の信頼性分析表 (全体 α = {res['alpha_val']:.3f})",
                    "df": res["rel_df"],
                    "note": res["note"],
                    "fig_bytes": res["fig_bytes"]
                })
                st.toast("『尺度信頼性分析結果』をレポートリストに追加しました！", icon="📋")
                st.rerun()

            c_l, c_r = st.columns([2.5, 2])
            with c_l:
                st.write(f"**項目統計量および項目削除時クロンバックのα係数 (全体 α = {res['alpha_val']:.3f})**")
                st.dataframe(res["rel_df"], use_container_width=True)
                st.caption(res["note"])
            with c_r:
                st.write("**項目削除時αプロット**")
                with st.container(height=450):
                    st.image(res["fig_bytes"], use_column_width=True)


if __name__ == "__main__":
    main()
