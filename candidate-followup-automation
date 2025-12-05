/**
 * Recruitment Mailer Automation
 * カジュアル面談お礼 & 体験予約案内メール送信スクリプト
 *
 * アクティブシートのリストをもとに、PDFを添付した案内メールを一括送信する。
 * 送信完了後、ステータスを更新する。
 */

// ==========================================
// 1. Configuration
// ==========================================

const CONFIG = Object.freeze({
  // スクリプトプロパティから設定値をロード
  ENV: {
    SENDER_NAME:      '〇〇株式会社 採用担当', // 送信者表示名
    SENDER_ALIAS:     PropertiesService.getScriptProperties().getProperty('MAIL_ALIAS'),
    RECRUIT_ADDR:     PropertiesService.getScriptProperties().getProperty('RECRUIT_MAIL_ADDR'),
    COMPANY_NAME:     PropertiesService.getScriptProperties().getProperty('COMPANY_NAME') || '〇〇株式会社'
  },
  // カラム定義 (0-based index)
  COLS: {
    NAME: 0,
    EMAIL: 1,
    BOOKING_URL: 2,
    RECRUITER: 3,
    STUDIO_NAME: 4,
    PDF_FILE_ID: 5,
    STATUS: 6
  },
  STATUS: {
    SENT: '送信済',
    ERROR: 'エラー'
  }
});

// ==========================================
// 2. Main Entry Point
// ==========================================

function sendThankYouAndBookingEmails() {
  const ui = SpreadsheetApp.getUi();
  const result = { success: 0, skipped: 0, error: 0 };

  try {
    // 1. 設定検証
    validateEnvironment();

    // 2. データ取得
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();

    // 3. 送信処理実行
    const service = new MailService();
    
    // ヘッダー(1行目)をスキップしてループ
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = row[CONFIG.COLS.STATUS];

      // 送信済みはスキップ
      if (status === CONFIG.STATUS.SENT) {
        result.skipped++;
        continue;
      }

      // 必須データのバリデーション
      if (!isValidRow(row)) {
        console.warn(`Row ${i + 1}: Missing required fields.`);
        continue;
      }

      try {
        // メール送信処理
        service.sendRecruitmentMail(row);
        
        // ステータス更新 (成功時のみ即時書き込み)
        sheet.getRange(i + 1, CONFIG.COLS.STATUS + 1).setValue(CONFIG.STATUS.SENT);
        result.success++;
        
      } catch (e) {
        console.error(`Row ${i + 1} Error: ${e.message}`);
        sheet.getRange(i + 1, CONFIG.COLS.STATUS + 1).setValue(`${CONFIG.STATUS.ERROR}: ${e.message}`);
        result.error++;
      }
    }

    // 4. 完了通知
    const msg = `処理完了\n成功: ${result.success}件\nスキップ: ${result.skipped}件\nエラー: ${result.error}件`;
    console.log(msg);
    ui.alert(msg);

  } catch (error) {
    console.error('Fatal Error:', error);
    ui.alert(`重大なエラーが発生しました:\n${error.message}`);
  }
}

// ==========================================
// 3. Service Layer
// ==========================================

class MailService {
  constructor() {
    this.template = new MailTemplate();
  }

  /**
   * 個別のメール送信処理を行う
   * @param {Array} row 行データ
   */
  sendRecruitmentMail(row) {
    const candidate = {
      name:       row[CONFIG.COLS.NAME],
      email:      row[CONFIG.COLS.EMAIL],
      bookingUrl: row[CONFIG.COLS.BOOKING_URL],
      recruiter:  row[CONFIG.COLS.RECRUITER],
      studioName: row[CONFIG.COLS.STUDIO_NAME],
      pdfId:      row[CONFIG.COLS.PDF_FILE_ID]
    };

    // 本文生成
    const subject = `【${CONFIG.ENV.COMPANY_NAME}】カジュアル面談のお礼 & ${candidate.studioName}の体験予約について`;
    const body = this.template.buildBody(candidate);

    // 添付ファイル取得
    const attachment = this._getAttachment(candidate.pdfId);

    // Gmail送信オプション
    const options = {
      from: CONFIG.ENV.SENDER_ALIAS,
      name: CONFIG.ENV.SENDER_NAME,
      attachments: [attachment]
    };

    GmailApp.sendEmail(candidate.email, subject, body, options);
  }

  _getAttachment(fileId) {
    try {
      return DriveApp.getFileById(fileId).getBlob();
    } catch (e) {
      throw new Error(`添付ファイル取得失敗 (ID: ${fileId})`);
    }
  }
}

class MailTemplate {
  buildBody(data) {
    return `
${data.name} 様

お世話になっております。
${CONFIG.ENV.COMPANY_NAME} 採用担当の${data.recruiter}です。

本日はカジュアル面談にご参加いただき、誠にありがとうございました。
お話しできたこと、大変嬉しく思っております。

弊社のカルチャーや働き方についてご興味を持っていただけましたでしょうか。
もし何かご不明点がございましたら、お気軽にご連絡ください。

カジュアル面談の中でもお伝えしたとおり、${data.studioName}のスタジオ体験を通じて、
より具体的にイメージを持っていただければと考えております。

以下のリンクからスタジオをご予約いただけます。
ぜひ実際の雰囲気を体験し、${CONFIG.ENV.COMPANY_NAME}の魅力を感じていただければ幸いです。

${data.studioName} スタジオ予約はこちら
${data.bookingUrl}

※体験レッスン終了後、店舗案内や面接のお時間をいただく関係により、
　可能な範囲で「平日」でのご予約にご協力いただけますと幸いです。
　土日祝をご希望の場合は、お手数ですがメールにてご相談ください。

また、スタジオ予約の方法を添付のPDFにまとめておりますので、あわせてご確認ください。

ご不明な点がございましたら、お気軽にご連絡ください。

ご予約が完了しましたら、ご一報いただけますと幸いです。
スタジオでお会いできることを楽しみにしております。

どうぞよろしくお願いいたします。

――――――――――――――――――
${CONFIG.ENV.COMPANY_NAME}
採用担当 ${data.recruiter}
${CONFIG.ENV.RECRUIT_ADDR}
――――――――――――――――――
`.trim();
  }
}

// ==========================================
// 4. Utilities
// ==========================================

/**
 * 必須環境変数のチェック
 */
function validateEnvironment() {
  const missing = [];
  if (!CONFIG.ENV.SENDER_ALIAS) missing.push('MAIL_ALIAS');
  if (!CONFIG.ENV.RECRUIT_ADDR) missing.push('RECRUIT_MAIL_ADDR');
  
  if (missing.length > 0) {
    throw new Error(`スクリプトプロパティが設定されていません: ${missing.join(', ')}`);
  }
}

/**
 * 必須カラムの値が存在するかチェック
 */
function isValidRow(row) {
  // 名前、Email、URL、担当者、店舗名、PDF ID が全て存在するか
  const requiredCols = [
    CONFIG.COLS.NAME,
    CONFIG.COLS.EMAIL,
    CONFIG.COLS.BOOKING_URL,
    CONFIG.COLS.RECRUITER,
    CONFIG.COLS.STUDIO_NAME,
    CONFIG.COLS.PDF_FILE_ID
  ];
  return requiredCols.every(idx => row[idx] && row[idx].toString().trim() !== '');
}
