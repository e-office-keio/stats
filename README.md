# 📊 APAスタイル 統計解析 Webアプリケーション (stats)

CSVファイルをアップロードしてデータ前処理（Recodeや特定値の除外など）を行い、目的の統計分析（基本統計量・検定・回帰・因子分析など）を選択して「分析を実行」ボタンを押すと、**APAスタイル（第7版）の表と図を統合した日本語Excelファイル (.xlsx)** を作成・ダウンロードできるWebアプリです。

---

## 📁 フォルダ構成

本アプリケーションは `stats` フォルダー内に格納されています。

- `stats/app.py`: メイン画面・Streamlit UI・前処理ツール・オンデマンド分析実行
- `stats/stats_engine.py`: 統計計算・データ編集（Recode/Filter/Binning/合成スコア）・APAグラフ生成エンジン
- `stats/apa_excel.py`: APA7thスタイル（日本語対応・罫線・注釈・画像埋め込み）Excelレポート作成機能
- `stats/requirements.txt`: 依存ライブラリ定義ファイル
- `stats/README.md`: 説明ドキュメント

---

## 🚀 アプリの起動手順

### 1. ディレクトリの移動
ターミナルで `stats` サブフォルダーに移動します。

```bash
cd stats
```

### 2. 依存ライブラリのインストール
```bash
pip install -r requirements.txt
```

### 3. アプリの実行
```bash
streamlit run app.py
```

ブラウザが立ち上がり、 `http://localhost:8501` でアプリをご利用いただけます。
