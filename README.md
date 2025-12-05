# GAS Recruitment Mailer

スプレッドシートの候補者リストを読み込み、PDF資料を添付した「面談お礼 & 体験予約案内メール」を一括送信する Google Apps Script (GAS) ツールです。

## 🚀 Features

* **Bulk Sending**: リストを上から順に処理し、未送信の候補者にのみメールを送信します。
* **Attachment Handling**: 指定されたGoogle DriveのPDFファイルを自動添付します。
* **Status Tracking**: 送信成功/失敗のステータスをシートに自動記録します。
* **Robust Error Handling**: 特定の行でエラー（PDFが見つからない、アドレス不備など）が発生しても、他の候補者への送信は止まりません。

## ⚙️ Setup

### 1. Script Properties
セキュリティ確保のため、メールアドレスや企業名はコードに直書きせず、GASの「スクリプトプロパティ」で管理します。
以下のプロパティを設定してください。

| Property Key | Description | Example |
| :--- | :--- | :--- |
| `MAIL_ALIAS` | Gmailの送信元メールアドレス（エイリアス） | `recruit@example.com` |
| `RECRUIT_MAIL_ADDR` | 署名欄に記載する問い合わせ用アドレス | `recruit@example.com` |
| `COMPANY_NAME` | (Optional) 会社名。未設定時はデフォルト値を使用 | `株式会社Example` |

### 2. Spreadsheet Structure
アクティブなシートが以下のカラム構成であることを前提としています。
※ 1行目はヘッダーとして無視されます。

| Column (A=0) | Field Name | Description |
| :--- | :--- | :--- |
| A | 候補者名 | 送信先のお名前 |
| B | Email | 送信先メールアドレス |
| C | 予約URL | スタジオ予約フォームのURL |
| D | 採用担当者 | 署名および本文で使用する担当者名 |
| E | 店舗名 | 配属予定/体験予定のスタジオ名 |
| F | PDFファイルID | Google Drive上の添付ファイルID |
| G | 送信ステータス | 送信後に「送信済」またはエラー内容が記載されます |

## 💻 Usage

1.  対象のスプレッドシートを開きます。
2.  スクリプトエディタ、または登録したカスタムメニューから `sendThankYouAndBookingEmails` 関数を実行します。
3.  処理完了後、ポップアップで実行結果（成功数・エラー数）が表示されます。

## 🛡 Security & Privacy

* **OAuth Scopes**: 本スクリプトはメール送信(`GmailApp`)およびドライブ参照(`DriveApp`)の権限を使用します。
* **Environment Variables**: 機密情報はすべてスクリプトプロパティ経由で注入されるため、コード自体を公開しても個人情報や社内アドレスは流出しません。

## 📝 License

This project is licensed under the MIT License.
