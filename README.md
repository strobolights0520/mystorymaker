# MY STORY MAKER

学生が質問に答えながら、自分らしさが伝わる自己紹介PowerPointを作成できるWebアプリです。

## 主な機能

- 「MY FEATURE」（3・4要素）の作成
- 「モチベーショングラフ」（5〜7エピソード）の作成
- モチベーションが上がる時・下がる時の全体まとめ
- シンプル・クール・ファッションの3テンプレート
- JPEG・PNG・WebP写真のアップロード、拡大、位置調整
- ROOKIES関連ロゴの選択
- 16:9・1ページのリアルタイムプレビュー
- 編集可能なPowerPoint（`.pptx`）出力
- LocalStorageへの自動保存と再開
- 文字数・未入力・画像形式・容量のバリデーション
- PowerPointダウンロード成功時のGoogleスプレッドシート記録

## 開発

```bash
npm install
npm run dev
```

本番ビルド:

```bash
npm run build
```

## 入力記録の環境変数

VercelのEnvironment Variablesに以下を設定してください。

```text
SHEETS_WEB_APP_URL=https://script.google.com/macros/s/.../exec
SHEETS_INGEST_SECRET=Apps Script側と同じ秘密文字列
```

入力記録では氏名と写真ファイルを送信しません。写真については、各項目に写真が設定されているかどうかだけを記録します。同じ内容を続けてダウンロードした場合は重複記録しません。

## ロゴについて

ROOKIES、CAREER ROOKIES GP 2024・2025・2026、ROOKIES GUILD、ROOKIES WORLD SERIESは、提供済みの正式画像を使用しています。
