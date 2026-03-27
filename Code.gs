function onFormSubmit(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();

  // 回答データ取得
  const values = sheet.getRange(row, 1, 1, 5).getValues()[0];
  const timestamp = values[0];
  const name = values[1];
  const email = values[2];
  const category = values[3];
  const inquiry = values[4];

  // ▼ここが変更ポイント
  const configSheet = e.source.getSheetByName("設定");
  const configData = configSheet.getDataRange().getValues();

  let assignee = "info@example.com"; // デフォルト

  for (let i = 1; i < configData.length; i++) {
    const configCategory = configData[i][0];
    const configEmail = configData[i][1];

    if (configCategory === category) {
      assignee = configEmail;
      break;
    }
  }

  const judgedCategory = category || "未分類";

  const subject = `【新規問い合わせ】${judgedCategory} / ${name}`;

  const body = `
新しい問い合わせがありました。

【氏名】${name}
【メール】${email}
【種別】${judgedCategory}

【内容】
${inquiry}
`;

  let status = "送信成功";
  let sentAt = new Date();

  try {
    GmailApp.sendEmail(assignee, subject, body);
  } catch (error) {
    status = `送信失敗: ${error.message}`;
    sentAt = "";
  }

  sheet.getRange(row, 6).setValue(judgedCategory);
  sheet.getRange(row, 7).setValue(assignee);
  sheet.getRange(row, 8).setValue(sentAt);
  sheet.getRange(row, 9).setValue(status);
}
