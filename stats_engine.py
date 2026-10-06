"""
Statistical Engine & Data Preprocessing for Streamlit Data Analysis App.
Calculates statistical analyses and generates Japanese APA-styled figures.
"""

import io
import numpy as np
import pandas as pd
import scipy.stats as stats
import statsmodels.api as sm
from statsmodels.stats.multicomp import pairwise_tukeyhsd
from statsmodels.stats.outliers_influence import variance_inflation_factor

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import seaborn as sns

# japanize_matplotlib を最優先適用 (完全文字化け防止)
HAS_JAPANIZE = False
JAPANIZE_ERROR = ""
try:
    import japanize_matplotlib
    HAS_JAPANIZE = True
except Exception as e:
    JAPANIZE_ERROR = f"{type(e).__name__}: {str(e)}"

def setup_japanese_font():
    """環境に合わせた日本語フォントの設定"""
    if HAS_JAPANIZE:
        try:
            japanize_matplotlib.japanize()
        except Exception:
            pass
        plt.rcParams['font.family'] = 'IPAexGothic'
        plt.rcParams['font.sans-serif'] = ['IPAexGothic', 'DejaVu Sans', 'sans-serif']
        plt.rcParams['axes.unicode_minus'] = False
        return

    # japanize_matplotlib がない場合のOS別フォールバック (実際にシステムに存在するフォントのみ設定)
    system_fonts = [f.name for f in fm.fontManager.ttflist]
    candidates = [
        'IPAexGothic', 'IPAGothic', 'TakaoPGothic',
        'Hiragino Sans', 'Hiragino Kaku Gothic ProN', 'Yu Gothic', 'Meiryo',
        'DejaVu Sans'
    ]
    matched_font = 'sans-serif'
    for font in candidates:
        if font in system_fonts:
            matched_font = font
            break
            
    plt.rcParams['font.family'] = matched_font
    plt.rcParams['font.sans-serif'] = [matched_font, 'DejaVu Sans', 'sans-serif']
    plt.rcParams['axes.unicode_minus'] = False

setup_japanese_font()

def set_apa_plot_style():
    """APA形式のグラフスタイルを設定"""
    setup_japanese_font()
    font_name = 'IPAexGothic' if HAS_JAPANIZE else plt.rcParams.get('font.family', ['sans-serif'])
    if isinstance(font_name, list):
        font_name = font_name[0]
        
    plt.rcParams.update({
        'font.family': font_name,
        'font.sans-serif': [font_name, 'DejaVu Sans', 'sans-serif'],
        'font.size': 11,
        'axes.labelsize': 11,
        'axes.titlesize': 12,
        'xtick.labelsize': 10,
        'ytick.labelsize': 10,
        'legend.fontsize': 10,
        'figure.titlesize': 13,
        'axes.spines.top': False,
        'axes.spines.right': False,
        'axes.edgecolor': '#333333',
        'axes.linewidth': 1.0,
        'grid.color': '#e0e0e0',
        'grid.linestyle': '--',
        'grid.alpha': 0.5
    })
    setup_japanese_font()

def fig_to_bytes(fig):
    """Matplotlib Figureオブジェクトをbytesへ変換"""
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=300, bbox_inches='tight')
    plt.close(fig)
    return buf.getvalue()



# ----------------------------------------------------------------------
# データ前処理・リコード・除外機能
# ----------------------------------------------------------------------
def filter_exclude_values(df, var_name, exclude_values):
    """特定の値を指定して該当行を除外"""
    new_df = df.copy()
    new_df = new_df[~new_df[var_name].isin(exclude_values)]
    return new_df

def recode_values(df, target_var, mapping_dict, new_var_name=None):
    """既存変数の値をマッピング辞書に従って置換し新変数を作成"""
    if not new_var_name:
        new_var_name = f"{target_var}_recoded"
    new_df = df.copy()
    new_df[new_var_name] = new_df[target_var].replace(mapping_dict)
    return new_df, new_var_name

def create_binned_variable(df, target_var, bins, labels, new_var_name=None):
    """数値変数を任意のビンで区切ってカテゴリ化"""
    if not new_var_name:
        new_var_name = f"{target_var}_binned"
    new_df = df.copy()
    new_df[new_var_name] = pd.cut(new_df[target_var], bins=bins, labels=labels, include_lowest=True)
    return new_df, new_var_name

def create_composite_score(df, source_vars, func="mean", new_var_name="Composite_Score"):
    """複数変数の合成スコア（平均値または合計値）を作成"""
    new_df = df.copy()
    if func == "mean":
        new_df[new_var_name] = new_df[source_vars].mean(axis=1)
    else:
        new_df[new_var_name] = new_df[source_vars].sum(axis=1)
    return new_df, new_var_name


# ----------------------------------------------------------------------
# 1. 単純集計 (Frequencies)
# ----------------------------------------------------------------------
def analyze_frequency(df, var_name):
    """単一カテゴリ変数の単純集計"""
    set_apa_plot_style()
    series = df[var_name].dropna()
    counts = series.value_counts().sort_index()
    percentages = (counts / len(series)) * 100
    cum_counts = counts.cumsum()
    cum_percentages = percentages.cumsum()
    
    res_df = pd.DataFrame({
        "カテゴリ": counts.index,
        "度数 (N)": counts.values,
        "割合 (%)": percentages.values,
        "累積度数": cum_counts.values,
        "累積割合 (%)": cum_percentages.values
    }).set_index("カテゴリ")
    
    fig, ax = plt.subplots(figsize=(6, 4))
    bars = ax.bar(res_df.index.astype(str), res_df["度数 (N)"], color="#2b5c8f", edgecolor="black", width=0.5)
    ax.set_ylabel("度数 (N)")
    ax.set_xlabel(str(var_name))
    ax.set_title(f"{var_name} の度数分布")
    
    for bar in bars:
        yval = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2, yval + (max(counts)*0.01), f"{int(yval)}", ha='center', va='bottom', fontsize=10)
        
    fig_bytes = fig_to_bytes(fig)
    note = f"全サンプル数 N = {len(series)}。欠損値は除外されています。"
    
    return res_df, note, fig_bytes


# ----------------------------------------------------------------------
# 2. 基本統計量 (Descriptive Statistics)
# ----------------------------------------------------------------------
def analyze_descriptives(df, num_vars):
    """選択した数値変数の基本統計量およびヒストグラム度数分布"""
    set_apa_plot_style()
    records = []
    hist_dict = {}
    
    for var in num_vars:
        s = df[var].dropna()
        if len(s) == 0:
            continue
        records.append({
            "変数名": var,
            "サンプルサイズ (N)": len(s),
            "平均値 (M)": s.mean(),
            "標準偏差 (SD)": s.std(),
            "中央値 (Mdn)": s.median(),
            "四分位範囲 (IQR)": s.quantile(0.75) - s.quantile(0.25),
            "最小値": s.min(),
            "最大値": s.max(),
            "歪度": s.skew(),
            "尖度": s.kurtosis()
        })
        
        # ヒストグラムの度数分布テーブル計算
        counts, bin_edges = np.histogram(s, bins="auto")
        if len(counts) > 12:
            counts, bin_edges = np.histogram(s, bins=10)
            
        hist_records = []
        for i in range(len(counts)):
            label = f"{bin_edges[i]:.2f} ~ {bin_edges[i+1]:.2f}"
            hist_records.append({"階級 (区間)": label, "度数 (N)": int(counts[i])})
        hist_dict[var] = pd.DataFrame(hist_records).set_index("階級 (区間)")
        
    res_df = pd.DataFrame(records).set_index("変数名")
    
    # Web画面用: ヒストグラム & 確率密度曲線 (KDE)
    fig, axes = plt.subplots(len(num_vars), 1, figsize=(6, 3.2 * len(num_vars)))
    if len(num_vars) == 1:
        axes = [axes]
        
    for i, var in enumerate(num_vars):
        ax = axes[i]
        s = df[var].dropna()
        sns.histplot(s, kde=True, ax=ax, color="#34495e", edgecolor="white", linewidth=0.5)
        ax.set_title(f"{var} のデータ分布 (ヒストグラム & 確率密度)")
        ax.set_xlabel(str(var))
        ax.set_ylabel("度数 / 密度")
        
    plt.tight_layout()
    fig_bytes = fig_to_bytes(fig)
    
    note = "M = 平均値; SD = 標準偏差; Mdn = 中央値; IQR = 四分位範囲。"
    return res_df, hist_dict, note, fig_bytes


# ----------------------------------------------------------------------
# 3. クロス集計 & カイ二乗検定 (Crosstab & Chi-Square)
# ----------------------------------------------------------------------
def analyze_crosstab(df, row_var, col_var):
    """2変数のクロス集計 (度数 (列%)) と カイ二乗検定 & 構成比グラフ"""
    set_apa_plot_style()
    clean_df = df[[row_var, col_var]].dropna()
    ct_raw = pd.crosstab(clean_df[row_var], clean_df[col_var])
    
    col_sums = ct_raw.sum(axis=0)
    row_sums = ct_raw.sum(axis=1)
    grand_total = ct_raw.values.sum()
    
    # 度数 (列%) フォーマットのテーブル構築
    formatted_matrix = []
    for r in ct_raw.index:
        row_cells = []
        for c in ct_raw.columns:
            cnt = ct_raw.loc[r, c]
            col_tot = col_sums[c]
            pct = (cnt / col_tot * 100) if col_tot > 0 else 0.0
            row_cells.append(f"{cnt} ({pct:.1f}%)")
        row_tot = row_sums[r]
        row_pct = (row_tot / grand_total * 100) if grand_total > 0 else 0.0
        row_cells.append(f"{row_tot} ({row_pct:.1f}%)")
        formatted_matrix.append(row_cells)
        
    # 合計行の追加
    col_tot_cells = []
    for c in ct_raw.columns:
        col_tot_cells.append(f"{col_sums[c]} (100.0%)")
    col_tot_cells.append(f"{grand_total} (100.0%)")
    formatted_matrix.append(col_tot_cells)
    
    row_labels = list(ct_raw.index) + ["合計"]
    col_labels = list(ct_raw.columns) + ["合計"]
    
    ct_formatted_df = pd.DataFrame(formatted_matrix, index=row_labels, columns=col_labels)
    ct_formatted_df.index.name = f"{row_var} \\ {col_var}"
    
    # 数値のみの構成比 (%) テーブル（列方向の%）
    pct_df = (ct_raw.div(col_sums, axis=1) * 100).round(1)
    pct_df.index.name = f"{row_var} \\ {col_var}"
    
    # カイ二乗検定
    chi2, p, dof, ex = stats.chi2_contingency(ct_raw)
    min_dim = min(ct_raw.shape) - 1
    cramers_v = np.sqrt(chi2 / (grand_total * min_dim)) if min_dim > 0 else 0.0
    
    # Web画面用 構成比グラフ (100% 積み上げ棒グラフ)
    fig, ax = plt.subplots(figsize=(6.5, 4))
    pct_df_T = pct_df.T  # 列カテゴリごとに積み上げ
    pct_df_T.plot(kind='bar', stacked=True, ax=ax, colormap='Blues', edgecolor='black', width=0.5)
    ax.set_ylabel("構成比 (%)")
    ax.set_xlabel(str(col_var))
    ax.set_title(f"構成比グラフ ({row_var} × {col_var})")
    ax.legend(title=str(row_var), bbox_to_anchor=(1.05, 1), loc='upper left')
    plt.tight_layout()
    fig_bytes = fig_to_bytes(fig)
    
    p_str = "< .001" if p < 0.001 else f"= {p:.3f}"
    note = f"注. 表内の数値は 度数 (列方向の%) を表します。カイ二乗検定: χ²({dof}) = {chi2:.2f}, p {p_str}, クラメールのV = {cramers_v:.2f}。"
    
    return ct_formatted_df, pct_df, note, fig_bytes


# ----------------------------------------------------------------------
# 4. t検定 / Welch検定 (t-Test)
# ----------------------------------------------------------------------
def analyze_ttest(df, group_var, num_var, equal_var=False):
    """2群の平均値の差の検定 (Student or Welch) - 横並びグループM (SD)表示"""
    set_apa_plot_style()
    groups = df[group_var].dropna().unique()
    if len(groups) != 2:
        raise ValueError("グループ変数のカテゴリ数はちょうど2つである必要があります。")
        
    g1_val, g2_val = groups[0], groups[1]
    s1 = df[df[group_var] == g1_val][num_var].dropna()
    s2 = df[df[group_var] == g2_val][num_var].dropna()
    
    n1, m1, sd1 = len(s1), s1.mean(), s1.std()
    n2, m2, sd2 = len(s2), s2.mean(), s2.std()
    
    res = stats.ttest_ind(s1, s2, equal_var=equal_var)
    t_val = float(res.statistic)
    p_val = float(res.pvalue)
    
    if equal_var:
        df_val = float(n1 + n2 - 2)
        s_pooled = np.sqrt(((n1 - 1) * sd1**2 + (n2 - 1) * sd2**2) / df_val)
        cohens_d = float((m1 - m2) / s_pooled) if s_pooled != 0 else 0.0
        test_type = "Studentのt検定 (等分散仮定)"
    else:
        v1, v2 = sd1**2 / n1, sd2**2 / n2
        df_val = float((v1 + v2)**2 / ((v1**2 / (n1 - 1)) + (v2**2 / (n2 - 1))))
        s_pooled = np.sqrt((sd1**2 + sd2**2) / 2)
        cohens_d = float((m1 - m2) / s_pooled) if s_pooled != 0 else 0.0
        test_type = "Welchのt検定 (等分散非仮定)"
        
    tt_df = pd.DataFrame([{
        "従属変数": str(num_var),
        f"{g1_val} M (SD)": f"{m1:.2f} ({sd1:.2f})",
        f"{g2_val} M (SD)": f"{m2:.2f} ({sd2:.2f})",
        "t値": t_val,
        "自由度 (df)": df_val,
        "p値": p_val,
        "効果量 (d)": cohens_d
    }]).set_index("従属変数")
    
    fig, ax = plt.subplots(figsize=(5, 4))
    means = [m1, m2]
    errors = [1.96 * sd1 / np.sqrt(n1), 1.96 * sd2 / np.sqrt(n2)]
    labels = [str(g1_val), str(g2_val)]
    
    bars = ax.bar(labels, means, yerr=errors, capsize=5, color=['#4c72b0', '#55a868'], edgecolor='black', width=0.4)
    ax.set_ylabel(str(num_var))
    ax.set_xlabel(str(group_var))
    ax.set_title(f"{group_var} による {num_var} の平均値比較 (95%CI)")
    plt.tight_layout()
    fig_bytes = fig_to_bytes(fig)
    
    p_str = "< .001" if p_val < 0.001 else f"= {p_val:.3f}"
    note = f"注. {test_type}: t({df_val:.2f}) = {t_val:.2f}, p {p_str}, Cohen's d = {cohens_d:.2f}。サンプルサイズ: {g1_val} (N={n1}), {g2_val} (N={n2})。"
    
    return tt_df, note, fig_bytes


# ----------------------------------------------------------------------
# 5. 一元配置分散分析 (One-way ANOVA)
# ----------------------------------------------------------------------
def analyze_anova(df, group_var, num_var):
    """一元配置分散分析 & Tukey HSD多重比較 - 横並びグループ M (SD) 表示"""
    set_apa_plot_style()
    clean_df = df[[group_var, num_var]].dropna()
    group_names = clean_df[group_var].dropna().unique()
    groups = [group[num_var].values for name, group in clean_df.groupby(group_var, sort=False)]
    
    f_val, p_val = stats.f_oneway(*groups)
    
    grand_mean = clean_df[num_var].mean()
    ss_total = np.sum((clean_df[num_var] - grand_mean)**2)
    ss_between = np.sum([len(g) * (np.mean(g) - grand_mean)**2 for g in groups])
    eta_sq = float(ss_between / ss_total) if ss_total != 0 else 0.0
    
    df_between = len(groups) - 1
    df_within = len(clean_df) - len(groups)
    
    anova_row = {"従属変数": str(num_var)}
    for g_name in group_names:
        g_series = clean_df[clean_df[group_var] == g_name][num_var]
        m_g = g_series.mean()
        sd_g = g_series.std()
        anova_row[f"{g_name} M (SD)"] = f"{m_g:.2f} ({sd_g:.2f})"
        
    anova_row["F値"] = float(f_val)
    anova_row["自由度 (df)"] = f"{df_between}, {df_within}"
    anova_row["p値"] = float(p_val)
    anova_row["効果量 (η²)"] = eta_sq
    
    anova_df = pd.DataFrame([anova_row]).set_index("従属変数")
    
    tukey = pairwise_tukeyhsd(endog=clean_df[num_var], groups=clean_df[group_var], alpha=0.05)
    tukey_df = pd.DataFrame(data=tukey._results_table.data[1:], columns=["グループ1", "グループ2", "平均値の差", "p値", "下限(95%CI)", "上限(95%CI)", "有意差"])
    tukey_df = tukey_df.set_index(["グループ1", "グループ2"])
    
    fig, ax = plt.subplots(figsize=(6, 4))
    sns.barplot(data=clean_df, x=group_var, y=num_var, ax=ax, capsize=0.1, palette="Blues_d", edgecolor="black")
    ax.set_title(f"一元配置分散分析: {group_var} × {num_var}")
    ax.set_xlabel(str(group_var))
    ax.set_ylabel(str(num_var))
    plt.tight_layout()
    fig_bytes = fig_to_bytes(fig)
    
    p_str = "< .001" if p_val < 0.001 else f"= {p_val:.3f}"
    note = f"注. 分散分析: F({df_between}, {df_within}) = {f_val:.2f}, p {p_str}, 効果量 η² = {eta_sq:.2f}。"
    
    return anova_df, tukey_df, note, fig_bytes


# ----------------------------------------------------------------------
# 6. 相関分析 (Correlation Analysis)
# ----------------------------------------------------------------------
def analyze_correlation(df, num_vars, method="pearson"):
    """相関係数行列と有意確率の算出"""
    set_apa_plot_style()
    clean_df = df[num_vars].dropna()
    
    n_vars = len(num_vars)
    corr_matrix = np.zeros((n_vars, n_vars))
    p_matrix = np.zeros((n_vars, n_vars))
    
    for i in range(n_vars):
        for j in range(n_vars):
            if i == j:
                corr_matrix[i, j] = 1.0
                p_matrix[i, j] = 0.0
            else:
                if method == "pearson":
                    r, p = stats.pearsonr(clean_df.iloc[:, i], clean_df.iloc[:, j])
                else:
                    r, p = stats.spearmanr(clean_df.iloc[:, i], clean_df.iloc[:, j])
                corr_matrix[i, j] = r
                p_matrix[i, j] = p
                
    display_df = pd.DataFrame(index=num_vars, columns=num_vars)
    for i in range(n_vars):
        for j in range(n_vars):
            r = corr_matrix[i, j]
            p = p_matrix[i, j]
            if i == j:
                display_df.iloc[i, j] = "-"
            else:
                stars = ""
                if p < 0.001:
                    stars = "***"
                elif p < 0.01:
                    stars = "**"
                elif p < 0.05:
                    stars = "*"
                display_df.iloc[i, j] = f"{r:.2f}{stars}"
                
    method_jp = "ピアソン" if method == "pearson" else "スピアマン"
    fig, ax = plt.subplots(figsize=(6, 5))
    sns.heatmap(corr_matrix, annot=True, fmt=".2f", cmap="coolwarm", vmin=-1, vmax=1,
                xticklabels=num_vars, yticklabels=num_vars, ax=ax, cbar_kws={'label': f'{method_jp} 相関係数 r'})
    ax.set_title(f"{method_jp} 相関係数ヒートマップ")
    plt.tight_layout()
    fig_bytes = fig_to_bytes(fig)
    
    note = f"全サンプル数 N = {len(clean_df)}。* p < .05, ** p < .01, *** p < .001。"
    return display_df, note, fig_bytes


# ----------------------------------------------------------------------
# 7. 重回帰分析 (Multiple Linear Regression)
# ----------------------------------------------------------------------
def analyze_regression(df, target_var, feature_vars):
    """重回帰分析"""
    set_apa_plot_style()
    clean_df = df[[target_var] + feature_vars].dropna()
    
    X = clean_df[feature_vars]
    y = clean_df[target_var]
    
    X_const = sm.add_constant(X)
    model = sm.OLS(y, X_const).fit()
    
    y_std = (y - y.mean()) / y.std()
    X_std = (X - X.mean()) / X.std()
    model_std = sm.OLS(y_std, X_std).fit()
    betas = model_std.params
    
    vifs = [variance_inflation_factor(X_const.values, i) for i in range(1, X_const.shape[1])]
    
    reg_table = []
    reg_table.append({
        "要因 / 変数名": "切片 (Intercept)",
        "非標準化係数 (B)": model.params["const"],
        "標準誤差 (SE)": model.bse["const"],
        "標準化係数 (β)": "-",
        "t値": model.tvalues["const"],
        "p値": model.pvalues["const"],
        "VIF": "-"
    })
    
    for i, var in enumerate(feature_vars):
        reg_table.append({
            "要因 / 変数名": var,
            "非標準化係数 (B)": model.params[var],
            "標準誤差 (SE)": model.bse[var],
            "標準化係数 (β)": betas[var],
            "t値": model.tvalues[var],
            "p値": model.pvalues[var],
            "VIF": vifs[i]
        })
        
    res_df = pd.DataFrame(reg_table).set_index("要因 / 変数名")
    
    fig, ax = plt.subplots(figsize=(5.5, 4.5))
    y_pred = model.predict(X_const)
    ax.scatter(y, y_pred, color="#2b5c8f", alpha=0.7, edgecolors="none")
    ax.plot([y.min(), y.max()], [y.min(), y.max()], 'r--', lw=1.5, label="理想線")
    ax.set_xlabel(f"実測値 ({target_var})")
    ax.set_ylabel(f"予測値 ({target_var})")
    ax.set_title(f"実測値 vs. 予測値の散布図 ({target_var})")
    ax.legend()
    plt.tight_layout()
    fig_bytes = fig_to_bytes(fig)
    
    r2 = model.rsquared
    adj_r2 = model.rsquared_adj
    f_stat = model.fvalue
    f_p = model.f_pvalue
    p_str = "< .001" if f_p < 0.001 else f"= {f_p:.3f}"
    
    note = f"目的変数: {target_var}。決定係数 R² = {r2:.3f}, 自由度調整済み R² = {adj_r2:.3f}, F({len(feature_vars)}, {len(clean_df)-len(feature_vars)-1}) = {f_stat:.2f}, p {p_str}。"
    return res_df, note, fig_bytes


# ----------------------------------------------------------------------
# 8. 因子分析 (Exploratory Factor Analysis)
# ----------------------------------------------------------------------
def analyze_factor_analysis(df, num_vars, n_factors=2, rotation="promax"):
    """探索的因子分析"""
    set_apa_plot_style()
    clean_df = df[num_vars].dropna()
    
    factor_cols = [f"第{i+1}因子" for i in range(n_factors)]
    
    try:
        from factor_analyzer import FactorAnalyzer
        fa = FactorAnalyzer(n_factors=n_factors, rotation=rotation, method='principal')
        fa.fit(clean_df)
        
        loadings = pd.DataFrame(
            fa.loadings_,
            index=num_vars,
            columns=factor_cols
        )
        
        ev, v = fa.get_eigenvalues()
        var_explained = fa.get_factor_variance()
        variance_df = pd.DataFrame(
            var_explained,
            index=["因子負荷量二乗和", "分散説明率 (寄与率)", "累積分散説明率 (累積寄与率)"],
            columns=factor_cols
        )
        
    except ImportError:
        from sklearn.decomposition import PCA
        pca = PCA(n_components=n_factors)
        pca.fit((clean_df - clean_df.mean()) / clean_df.std())
        loadings = pd.DataFrame(
            pca.components_.T,
            index=num_vars,
            columns=factor_cols
        )
        ev = pca.explained_variance_
        variance_df = pd.DataFrame([
            pca.explained_variance_,
            pca.explained_variance_ratio_,
            np.cumsum(pca.explained_variance_ratio_)
        ], index=["固有値", "分散説明率 (寄与率)", "累積分散説明率 (累積寄与率)"], columns=factor_cols)

    fig, ax = plt.subplots(figsize=(6, 4))
    ax.plot(range(1, len(ev) + 1), ev, marker='o', color='#2b5c8f', linewidth=2)
    ax.axhline(1.0, color='red', linestyle='--', label='カイザー基準 (固有値 = 1.0)')
    ax.set_title("スクリープロット (固有値の推移)")
    ax.set_xlabel("因子番号")
    ax.set_ylabel("固有値 (Eigenvalue)")
    ax.legend()
    plt.tight_layout()
    fig_bytes = fig_to_bytes(fig)
    
    rot_jp = "プロマックス回転" if rotation == "promax" else "バリマックス回転"
    note = f"因子回転法: {rot_jp}。総サンプル数 N = {len(clean_df)}。"
    
    combined_df = pd.concat([loadings, variance_df], axis=0)
    return combined_df, note, fig_bytes
