"""
Statistical Engine for Streamlit Data Analysis App.
Calculates statistical analyses and generates APA-styled figures.
"""

import io
import numpy as np
import pandas as pd
import scipy.stats as stats
import statsmodels.api as sm
import statsmodels.formula.api as smf
from statsmodels.stats.multicomp import pairwise_tukeyhsd
from statsmodels.stats.outliers_influence import variance_inflation_factor

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns

# 日本語フォントの設定
plt.rcParams['font.sans-serif'] = ['Hiragino Sans', 'Yu Gothic', 'Meiryo', 'IPAexGothic', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False

def set_apa_plot_style():
    """APA形式のグラフスタイルを設定"""
    plt.rcParams.update({
        'font.size': 11,
        'axes.labelsize': 12,
        'axes.titlesize': 13,
        'xtick.labelsize': 10,
        'ytick.labelsize': 10,
        'legend.fontsize': 10,
        'figure.titlesize': 14,
        'axes.spines.top': False,
        'axes.spines.right': False,
        'axes.edgecolor': '#333333',
        'axes.linewidth': 1.0,
        'grid.color': '#e0e0e0',
        'grid.linestyle': '--',
        'grid.alpha': 0.5
    })

def fig_to_bytes(fig):
    """Matplotlib FigureオブジェクトをBytesIOへ変換"""
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=300, bbox_inches='tight')
    buf.seek(0)
    plt.close(fig)
    return buf


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
        "Category": counts.index,
        "Frequency (N)": counts.values,
        "Percent (%)": percentages.values,
        "Cumulative N": cum_counts.values,
        "Cumulative %": cum_percentages.values
    }).set_index("Category")
    
    # グラフ作成
    fig, ax = plt.subplots(figsize=(6, 4))
    bars = ax.bar(res_df.index.astype(str), res_df["Frequency (N)"], color="#2b5c8f", edgecolor="black", width=0.5)
    ax.set_ylabel("Frequency (N)")
    ax.set_xlabel(var_name)
    ax.set_title(f"Frequency Distribution: {var_name}")
    
    # バーの上に数値表示
    for bar in bars:
        yval = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2, yval + (max(counts)*0.01), f"{int(yval)}", ha='center', va='bottom', fontsize=10)
        
    fig_bytes = fig_to_bytes(fig)
    note = f"Total N = {len(series)}. Missing values excluded."
    
    return res_df, note, fig_bytes


# ----------------------------------------------------------------------
# 2. 基本統計量 (Descriptive Statistics)
# ----------------------------------------------------------------------
def analyze_descriptives(df, num_vars):
    """選択した数値変数の基本統計量"""
    set_apa_plot_style()
    records = []
    
    for var in num_vars:
        s = df[var].dropna()
        if len(s) == 0:
            continue
        n = len(s)
        mean = s.mean()
        std = s.std()
        median = s.median()
        q25 = s.quantile(0.25)
        q75 = s.quantile(0.75)
        iqr = q75 - q25
        min_v = s.min()
        max_v = s.max()
        skew = s.skew()
        kurt = s.kurtosis()
        
        records.append({
            "Variable": var,
            "N": n,
            "M": mean,
            "SD": std,
            "Mdn": median,
            "IQR": iqr,
            "Min": min_v,
            "Max": max_v,
            "Skewness": skew,
            "Kurtosis": kurt
        })
        
    res_df = pd.DataFrame(records).set_index("Variable")
    
    # ヒストグラム＋KDEグラフ（最初の3変数程度）
    fig, axes = plt.subplots(len(num_vars), 1, figsize=(6, 3.2 * len(num_vars)))
    if len(num_vars) == 1:
        axes = [axes]
        
    for i, var in enumerate(num_vars):
        ax = axes[i]
        s = df[var].dropna()
        sns.histplot(s, kde=True, ax=ax, color="#34495e", edgecolor="white", linewidth=0.5)
        ax.set_title(f"Distribution of {var}")
        ax.set_xlabel(var)
        ax.set_ylabel("Density / Count")
        
    plt.tight_layout()
    fig_bytes = fig_to_bytes(fig)
    
    note = "M = Mean; SD = Standard Deviation; Mdn = Median; IQR = Interquartile Range."
    return res_df, note, fig_bytes


# ----------------------------------------------------------------------
# 3. クロス集計 & カイ二乗検定 (Crosstab & Chi-Square)
# ----------------------------------------------------------------------
def analyze_crosstab(df, row_var, col_var):
    """2変数のクロス集計とカイ二乗検定"""
    set_apa_plot_style()
    ct = pd.crosstab(df[row_var], df[col_var], margins=True, margins_name="Total")
    
    # カイ二乗検定 (Total行・列を除く)
    ct_clean = pd.crosstab(df[row_var], df[col_var])
    chi2, p, dof, ex = stats.chi2_contingency(ct_clean)
    n_total = ct_clean.sum().sum()
    min_dim = min(ct_clean.shape) - 1
    cramers_v = np.sqrt(chi2 / (n_total * min_dim)) if min_dim > 0 else 0
    
    # グラフ作成（積み上げ棒グラフ）
    ct_prop = pd.crosstab(df[row_var], df[col_var], normalize='index') * 100
    fig, ax = plt.subplots(figsize=(6.5, 4))
    ct_prop.plot(kind='bar', stacked=True, ax=ax, colormap='Blues', edgecolor='black', width=0.5)
    ax.set_ylabel("Percentage (%)")
    ax.set_xlabel(row_var)
    ax.set_title(f"Crosstab: {row_var} vs {col_var}")
    ax.legend(title=col_var, bbox_to_anchor=(1.05, 1), loc='upper left')
    plt.tight_layout()
    fig_bytes = fig_to_bytes(fig)
    
    p_str = "< .001" if p < 0.001 else f"= {p:.3f}"
    note = f"χ²({dof}) = {chi2:.2f}, p {p_str}, Cramer's V = {cramers_v:.2f}."
    
    return ct, note, fig_bytes


# ----------------------------------------------------------------------
# 4. t検定 / Welch検定 (t-Test)
# ----------------------------------------------------------------------
def analyze_ttest(df, group_var, num_var, equal_var=False):
    """2群の平均値の差の検定 (Student or Welch)"""
    set_apa_plot_style()
    groups = df[group_var].dropna().unique()
    if len(groups) != 2:
        raise ValueError("グループ変数のカテゴリ数はちょうど2つである必要があります。")
        
    g1_val, g2_val = groups[0], groups[1]
    s1 = df[df[group_var] == g1_val][num_var].dropna()
    s2 = df[df[group_var] == g2_val][num_var].dropna()
    
    n1, m1, sd1 = len(s1), s1.mean(), s1.std()
    n2, m2, sd2 = len(s2), s2.mean(), s2.std()
    
    # 検定
    res = stats.ttest_ind(s1, s2, equal_var=equal_var)
    t_val = res.statistic
    p_val = res.pvalue
    
    if equal_var:
        df_val = n1 + n2 - 2
        # Pooled SD for Cohen's d
        s_pooled = np.sqrt(((n1 - 1) * sd1**2 + (n2 - 1) * sd2**2) / df_val)
        cohens_d = (m1 - m2) / s_pooled if s_pooled != 0 else 0
        test_type = "Student's t-test"
    else:
        # Welch-Satterthwaite df
        v1, v2 = sd1**2 / n1, sd2**2 / n2
        df_val = (v1 + v2)**2 / ((v1**2 / (n1 - 1)) + (v2**2 / (n2 - 1)))
        s_pooled = np.sqrt((sd1**2 + sd2**2) / 2)
        cohens_d = (m1 - m2) / s_pooled if s_pooled != 0 else 0
        test_type = "Welch's t-test"
        
    summary_df = pd.DataFrame([
        {"Group": f"{group_var} = {g1_val}", "N": n1, "M": m1, "SD": sd1},
        {"Group": f"{group_var} = {g2_val}", "N": n2, "M": m2, "SD": sd2},
    ]).set_index("Group")
    
    test_df = pd.DataFrame([{
        "Test Type": test_type,
        "t": t_val,
        "df": df_val,
        "p": p_val,
        "Cohen's d": cohens_d
    }]).set_index("Test Type")
    
    # グラフ（95% CIつきエラーバー棒グラフ）
    fig, ax = plt.subplots(figsize=(5, 4))
    means = [m1, m2]
    errors = [1.96 * sd1 / np.sqrt(n1), 1.96 * sd2 / np.sqrt(n2)]
    labels = [str(g1_val), str(g2_val)]
    
    bars = ax.bar(labels, means, yerr=errors, capsize=5, color=['#4c72b0', '#55a868'], edgecolor='black', width=0.4)
    ax.set_ylabel(num_var)
    ax.set_xlabel(group_var)
    ax.set_title(f"Comparison of {num_var} by {group_var}")
    plt.tight_layout()
    fig_bytes = fig_to_bytes(fig)
    
    p_str = "< .001" if p_val < 0.001 else f"= {p_val:.3f}"
    note = f"{test_type}: t({df_val:.2f}) = {t_val:.2f}, p {p_str}, Cohen's d = {cohens_d:.2f}. Error bars indicate 95% CI."
    
    # 総合結果表を作成
    combined_df = pd.concat([summary_df, test_df], axis=0)
    return combined_df, note, fig_bytes


# ----------------------------------------------------------------------
# 5. 一元配置分散分析 (One-way ANOVA)
# ----------------------------------------------------------------------
def analyze_anova(df, group_var, num_var):
    """一元配置分散分析 & Tukey HSD多重比較"""
    set_apa_plot_style()
    clean_df = df[[group_var, num_var]].dropna()
    groups = [group[num_var].values for name, group in clean_df.groupby(group_var)]
    group_names = list(clean_df[group_var].unique())
    
    # ANOVA計算
    f_val, p_val = stats.f_oneway(*groups)
    
    # 効果量 Eta-squared (η²)
    grand_mean = clean_df[num_var].mean()
    ss_total = np.sum((clean_df[num_var] - grand_mean)**2)
    ss_between = np.sum([len(g) * (np.mean(g) - grand_mean)**2 for g in groups])
    eta_sq = ss_between / ss_total if ss_total != 0 else 0
    
    # 各群のDescriptives
    desc = clean_df.groupby(group_var)[num_var].agg(['count', 'mean', 'std']).reset_index()
    desc.columns = ["Group", "N", "M", "SD"]
    desc_df = desc.set_index("Group")
    
    # Tukey HSD
    tukey = pairwise_tukeyhsd(endog=clean_df[num_var], groups=clean_df[group_var], alpha=0.05)
    tukey_df = pd.DataFrame(data=tukey._results_table.data[1:], columns=tukey._results_table.data[0])
    tukey_df = tukey_df.set_index(["group1", "group2"])
    
    # グラフ
    fig, ax = plt.subplots(figsize=(6, 4))
    sns.barplot(data=clean_df, x=group_var, y=num_var, ax=ax, capsize=0.1, palette="Blues_d", edgecolor="black")
    ax.set_title(f"One-way ANOVA: {num_var} across {group_var}")
    plt.tight_layout()
    fig_bytes = fig_to_bytes(fig)
    
    df_between = len(groups) - 1
    df_within = len(clean_df) - len(groups)
    p_str = "< .001" if p_val < 0.001 else f"= {p_val:.3f}"
    note = f"ANOVA: F({df_between}, {df_within}) = {f_val:.2f}, p {p_str}, η² = {eta_sq:.2f}."
    
    return desc_df, tukey_df, note, fig_bytes


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
                
    # 表示用フォーマット (r値に*を付与)
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
                
    # ヒートマップ
    fig, ax = plt.subplots(figsize=(6, 5))
    sns.heatmap(corr_matrix, annot=True, fmt=".2f", cmap="coolwarm", vmin=-1, vmax=1,
                xticklabels=num_vars, yticklabels=num_vars, ax=ax, cbar_kws={'label': f'{method.capitalize()} r'})
    ax.set_title(f"{method.capitalize()} Correlation Matrix")
    plt.tight_layout()
    fig_bytes = fig_to_bytes(fig)
    
    note = f"N = {len(clean_df)}. * p < .05, ** p < .01, *** p < .001."
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
    
    # 標準化回帰係数 (Beta)
    y_std = (y - y.mean()) / y.std()
    X_std = (X - X.mean()) / X.std()
    model_std = sm.OLS(y_std, X_std).fit()
    betas = model_std.params
    
    # VIF
    vifs = [variance_inflation_factor(X_const.values, i) for i in range(1, X_const.shape[1])]
    
    reg_table = []
    # Intercept
    reg_table.append({
        "Variable": "Intercept",
        "B": model.params["const"],
        "SE": model.bse["const"],
        "β": "-",
        "t": model.tvalues["const"],
        "p": model.pvalues["const"],
        "VIF": "-"
    })
    
    for i, var in enumerate(feature_vars):
        reg_table.append({
            "Variable": var,
            "B": model.params[var],
            "SE": model.bse[var],
            "β": betas[var],
            "t": model.tvalues[var],
            "p": model.pvalues[var],
            "VIF": vifs[i]
        })
        
    res_df = pd.DataFrame(reg_table).set_index("Variable")
    
    # 観測値 vs 予測値プロット
    fig, ax = plt.subplots(figsize=(5.5, 4.5))
    y_pred = model.predict(X_const)
    ax.scatter(y, y_pred, color="#2b5c8f", alpha=0.7, edgecolors="none")
    ax.plot([y.min(), y.max()], [y.min(), y.max()], 'r--', lw=1.5, label="Ideal")
    ax.set_xlabel(f"Actual ({target_var})")
    ax.set_ylabel(f"Predicted ({target_var})")
    ax.set_title("Observed vs. Predicted Values")
    ax.legend()
    plt.tight_layout()
    fig_bytes = fig_to_bytes(fig)
    
    r2 = model.rsquared
    adj_r2 = model.rsquared_adj
    f_stat = model.fvalue
    f_p = model.f_pvalue
    p_str = "< .001" if f_p < 0.001 else f"= {f_p:.3f}"
    
    note = f"Dependent variable: {target_var}. R² = {r2:.3f}, Adjusted R² = {adj_r2:.3f}, F({len(feature_vars)}, {len(clean_df)-len(feature_vars)-1}) = {f_stat:.2f}, p {p_str}."
    return res_df, note, fig_bytes


# ----------------------------------------------------------------------
# 8. 因子分析 (Exploratory Factor Analysis)
# ----------------------------------------------------------------------
def analyze_factor_analysis(df, num_vars, n_factors=2, rotation="promax"):
    """探索的因子分析 (Factor Analyzer ライブラリまたは SciPy / PCAフォールバック)"""
    set_apa_plot_style()
    clean_df = df[num_vars].dropna()
    
    try:
        from factor_analyzer import FactorAnalyzer
        fa = FactorAnalyzer(n_factors=n_factors, rotation=rotation, method='principal')
        fa.fit(clean_df)
        
        loadings = pd.DataFrame(
            fa.loadings_,
            index=num_vars,
            columns=[f"Factor {i+1}" for i in range(n_factors)]
        )
        
        ev, v = fa.get_eigenvalues()
        var_explained = fa.get_factor_variance()
        variance_df = pd.DataFrame(
            var_explained,
            index=["SS Loadings", "Proportion Var", "Cumulative Var"],
            columns=[f"Factor {i+1}" for i in range(n_factors)]
        )
        
    except ImportError:
        # factor_analyzer がない場合は SVD / PCA による簡易代替
        from sklearn.decomposition import PCA
        pca = PCA(n_components=n_factors)
        transformed = pca.fit_transform((clean_df - clean_df.mean()) / clean_df.std())
        loadings = pd.DataFrame(
            pca.components_.T,
            index=num_vars,
            columns=[f"Factor {i+1}" for i in range(n_factors)]
        )
        ev = pca.explained_variance_
        variance_df = pd.DataFrame([
            pca.explained_variance_,
            pca.explained_variance_ratio_,
            np.cumsum(pca.explained_variance_ratio_)
        ], index=["Eigenvalue", "Proportion Var", "Cumulative Var"], columns=[f"Factor {i+1}" for i in range(n_factors)])

    # スクリープロット
    fig, ax = plt.subplots(figsize=(6, 4))
    ax.plot(range(1, len(ev) + 1), ev, marker='o', color='#2b5c8f', linewidth=2)
    ax.axhline(1.0, color='red', linestyle='--', label='Kaiser Criterion (Eigenvalue = 1)')
    ax.set_title("Scree Plot")
    ax.set_xlabel("Factor Number")
    ax.set_ylabel("Eigenvalue")
    ax.legend()
    plt.tight_layout()
    fig_bytes = fig_to_bytes(fig)
    
    note = f"Rotation method: {rotation.capitalize()}. Total N = {len(clean_df)}."
    
    # 因子負荷量と寄与率を統合したDataFrame
    combined_df = pd.concat([loadings, variance_df], axis=0)
    return combined_df, note, fig_bytes
