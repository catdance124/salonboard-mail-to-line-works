# Gmail → LINE WORKS 転送スクリプト

サロンボード（`yoyaku_system@salonboard.com`）からの予約・キャンセル通知メールを、LINE WORKS のトークルームに自動転送する Google Apps Script です。

---

## 前提条件

- Google アカウント（Gmail + Google Apps Script）
- LINE WORKS の以下の設定が完了していること
  - Developer Console でアプリ作成済み（Client ID / Client Secret 取得済み）
  - Service Account 作成済み・秘密鍵（PEM）発行済み
  - Bot 作成済み・トークルームへの招待済み

---

## セットアップ

### 1. GAS プロジェクトにファイルを貼り付ける

[Google Apps Script](https://script.google.com) で新規プロジェクトを作成し、デフォルトの `コード.gs` の中身を `main.js` の内容で上書きします。

### 2. スクリプトプロパティを登録する

「プロジェクトの設定」→「スクリプトプロパティ」に以下を登録します。

| キー | 値 |
|------|----|
| `CLIENT_ID` | LINE WORKS アプリの Client ID |
| `CLIENT_SECRET` | LINE WORKS アプリの Client Secret |
| `SERVICE_ACCOUNT` | `xxx@xxx.service.worksmobile.com` 形式 |
| `PRIVATE_KEY` | 秘密鍵の PEM 全文（注意: 下記参照） |
| `BOT_ID` | Bot の数字 ID |
| `CHANNEL_ID` | 送信先トークルームの ID |

> **PRIVATE_KEY の貼り付けについて**
> GAS の UI で貼り付けると改行がスペースになる場合がありますが、スクリプト側で自動的に PEM 形式へ再構築するため、そのまま貼り付けて問題ありません。

### 3. トリガーを設定する

「トリガー」→「トリガーを追加」から以下のように設定します。

- 実行する関数: `forwardGmailToLineWorks`
- イベントのソース: 時間主導型
- 種類: 分ベースのタイマー（例: 5 分おき）

---

## テスト

実際のメールや API を使わず、フォーマットだけ確認したい場合：

```
testFormat() を実行 → ログにメッセージ内容が出力される
```

LINE WORKS への送信まで含めてテストしたい場合：

```
testSend() を実行 → テストデータを実際にトークルームへ送信する
```

---

## 転送対象メール

| 件名に含まれる文字列 | 動作 |
|---------------------|------|
| `予約連絡` | 予約通知としてフォーマット・送信 |
| `キャンセル連絡` | キャンセル通知としてフォーマット・送信 |
| それ以外 | スキップ（ログに記録） |

転送済みメールには `LW転送済み` ラベルが付き、二重送信を防止します。
