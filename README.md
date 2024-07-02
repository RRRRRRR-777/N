# 使用技術一覧
## 言語
Python 
## ライブラリ
pandas / numpy / openpyxl / selenium / request / beautifulsoup

# TL;DR
このコードを実行すると`xlsxファイル`が出力され、
そのファイルを見ることで株式投資のスクリーニング時間を大幅に短縮できます。
極端な話し、Ticker列が青色の銘柄のテクニカル分析が優れていればその銘柄を購入することができます。
#### ※ テクニカル分析とは
[過去の値動きをチャートで表して、そこからトレンドやパターンなどを把握し、今後の株価、為替動向を予想するものです。](https://info.monex.co.jp/technical-analysis/column/001.html#:~:text=%E3%83%86%E3%82%AF%E3%83%8B%E3%82%AB%E3%83%AB%E5%88%86%E6%9E%90%E3%81%A8%E3%81%AF%E3%81%99%E3%82%99%E3%81%AF%E3%82%99%E3%82%8A%E3%80%81%E9%81%8E%E5%8E%BB%E3%81%AE%E5%80%A4%E5%8B%95%E3%81%8D%E3%82%92%E3%83%81%E3%83%A3%E3%83%BC%E3%83%88%E3%81%A6%E3%82%99%E8%A1%A8%E3%81%97%E3%81%A6%E3%80%81%E3%81%9D%E3%81%93%E3%81%8B%E3%82%89%E3%83%88%E3%83%AC%E3%83%B3%E3%83%88%E3%82%99%E3%82%84%E3%83%8F%E3%82%9A%E3%82%BF%E3%83%BC%E3%83%B3%E3%81%AA%E3%81%A8%E3%82%99%E3%82%92%E6%8A%8A%E6%8F%A1%E3%81%97%E3%80%81%E4%BB%8A%E5%BE%8C%E3%81%AE%E6%A0%AA%E4%BE%A1%E3%80%81%E7%82%BA%E6%9B%BF%E5%8B%95%E5%90%91%E3%82%92%E4%BA%88%E6%83%B3%E3%81%99%E3%82%8B%E3%82%82%E3%81%AE%E3%81%A6%E3%82%99%E3%81%99%E3%80%82)

# 投資戦略

# ER図

# 実行方法

# 環境変数
|役割|変数|値|
|---|---|---|
|ログの出力設定|OUTPUT_LOGGER_LEVEL|20|
|finvizのサイトURL|FINVIZ_URL||
|NASDAQヒストリカルデータのURL|YFINANCE_NASDAQ_URL||
|各銘柄ヒストリカルデータのURL|YFINANCE_STOCKS_URL||
|アナリスト情報のURL|ANALYSTS_URL||
|過去決算情報のURL|FINANCE_URL||
|機関投資家情報のURL|INSTITUTE_URL||
|機関投資家情報のURLにアクセスする情報|INSTITUTE_USERNAME||
|機関投資家情報のURLにアクセスする情報|INSTITUTE_PASSWORD||
