# 技術トレンド収集システム

AI・エンジニアリング全般のトレンドを毎週自動収集してGitHub Pagesで公開します。

## 収集ソース
- GitHub Trending（週間・全言語 + Python/TypeScript/Rust/Go）
- Qiita API（LLM/ML/インフラ/セキュリティ/フロント/組み込み）
- Zenn RSS（全体 + カテゴリ別）
- npm 週間DL数（AI/インフラ/ツール系）

## 自動実行
毎週月曜 10:00 JST に GitHub Actions が自動実行されます。

## ローカル実行
```bash
pip install -r requirements.txt
python trend_collector/collect.py
```

結果は `trend_output/report_YYYYMMDD.html` と `.xlsx` に出力されます。
