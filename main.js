/**
 * Recruitment Mailer Automation (Production Ready)
 */

// ==========================================
// 1. Configuration & Constants
// ==========================================

const APP_CONFIG = {
  MENU_NAME: "▼ 採用アクション",
  // スプレッドシート上のヘッダー名定義（ここを実際のシートと合わせる）
  HEADER_MAP: {
    NAME: "名前",
    EMAIL: "メールアドレス",
    BOOKING_URL: "予約URL",
    RECRUITER: "担当者",
    STUDIO_NAME: "スタジオ名",
    PDF_FILE_ID: "添付ファイルID",
    STATUS: "送信ステータス"
  }
};

/**
 * スプレッドシートを開いた時に実行（メニュー追加）
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(APP_CONFIG.MENU_NAME)
    .addItem('📧 メールを一括送信する', 'sendThankYouAndBookingEmails')
    .addSeparator()
    .addItem('⚙️ 初期セットアップ', 'setupEnvironment')
    .addToUi();
}

// ==========================================
// 2. Setup Logic
// ==========================================

/**
 * 初回のみ実行する設定用関数
 * スクリプトプロパティを対話形式で保存する
 */
function setupEnvironment() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const questions = [
    { key: 'COMPANY_NAME', label: '会社名（例: 〇〇株式会社）' },
    { key: 'MAIL_ALIAS', label: '送信元のエイリアス（Gmailで設定済みのもの）' },
    { key: 'RECRUIT_MAIL_ADDR', label: '返信先・署名用メールアドレス' }
  ];

  for (const q of questions) {
    const response = ui.prompt(`${q.label} を入力してください`, ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() === ui.Button.OK) {
      props.setProperty(q.key, response.getResponseText());
    } else {
      ui.alert('セットアップを中断しました。');
      return;
    }
  }
  ui.alert('セットアップが完了しました！これでメール送信が可能です。');
}

// ==========================================
// 3. Core Logic (Refactored)
// ==========================================

function sendThankYouAndBookingEmails() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties().getProperties();
  
  // バリデーション
  if (!props.MAIL_ALIAS || !props.RECRUIT_MAIL_ADDR) {
    ui.alert('設定が不足しています。「初期セットアップ」を実行してください。');
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // ヘッダーから列番号を特定（動的取得）
  const col = {};
  for (const [key, text] of Object.entries(APP_CONFIG.HEADER_MAP)) {
    const index = headers.indexOf(text);
    if (index === -1) {
      ui.alert(`シートに「${text}」という列が見つかりません。`);
      return;
    }
    col[key] = index;
  }

  const service = new MailService(props);
  const result = { success: 0, skipped: 0, error: 0 };

  // データ処理ループ
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[col.STATUS] === '送信済') {
      result.skipped++;
      continue;
    }
    if (!row[col.EMAIL] || !row[col.NAME]) continue; // 必須項目空きはスキップ

    try {
      service.sendRecruitmentMail(row, col);
      sheet.getRange(i + 1, col.STATUS + 1).setValue('送信済');
      result.success++;
    } catch (e) {
      console.error(e);
      sheet.getRange(i + 1, col.STATUS + 1).setValue(`エラー: ${e.message}`);
      result.error++;
    }
  }

  ui.alert(`処理完了\n成功: ${result.success}件\nエラー: ${result.error}件`);
}

// ==========================================
// 4. Classes (Logic remains mostly same)
// ==========================================

class MailService {
  constructor(props) {
    this.props = props;
    this.template = new MailTemplate(props);
  }

  sendRecruitmentMail(row, colMap) {
    const candidate = {
      name:       row[colMap.NAME],
      email:      row[colMap.EMAIL],
      bookingUrl: row[colMap.BOOKING_URL],
      recruiter:  row[colMap.RECRUITER],
      studioName: row[colMap.STUDIO_NAME],
      pdfId:      row[colMap.PDF_FILE_ID]
    };

    const subject = `【${this.props.COMPANY_NAME}】カジュアル面談のお礼 & ${candidate.studioName}の体験予約について`;
    const body = this.template.buildBody(candidate);
    const attachment = DriveApp.getFileById(candidate.pdfId).getBlob();

    GmailApp.sendEmail(candidate.email, subject, body, {
      from: this.props.MAIL_ALIAS,
      name: `${this.props.COMPANY_NAME} 採用担当`,
      attachments: [attachment]
    });
  }
}

class MailTemplate {
  constructor(props) {
    this.props = props;
  }
  buildBody(data) {
    // 既存のテンプレートロジック
    return `
${data.name} 様

お世話になっております。
${this.props.COMPANY_NAME} 採用担当の${data.recruiter}です。
... (以下略) ...
`.trim();
  }
}
