function classifyInquiryWithAzure(inquiryText) {
  const props = PropertiesService.getScriptProperties();
  const endpoint = props.getProperty("AZURE_OPENAI_ENDPOINT");
  const apiKey = props.getProperty("AZURE_OPENAI_API_KEY");
  const deployment = props.getProperty("AZURE_OPENAI_DEPLOYMENT");

  if (!endpoint || !apiKey || !deployment) {
    throw new Error("Azure OpenAI のスクリプトプロパティが未設定です。");
  }

  const url =
    endpoint.replace(/\/$/, "") +
    "/openai/deployments/" +
    deployment +
    "/chat/completions?api-version=2024-10-21";

  const payload = {
    messages: [
      {
        role: "system",
        content:
          "あなたは問い合わせ分類AIです。必ず次の4つの語のうち1つだけを返してください。説明や理由は不要です。返答候補: 資料請求 / 操作方法 / 不具合報告 / その他"
      },
      {
        role: "user",
        content: inquiryText
      }
    ],
    temperature: 0,
    max_tokens: 20
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "api-key": apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const statusCode = response.getResponseCode();
  const bodyText = response.getContentText();

  Logger.log("statusCode: " + statusCode);
  Logger.log("bodyText: " + bodyText);

  if (statusCode !== 200) {
    throw new Error("Azure OpenAI API error: " + statusCode + " / " + bodyText);
  }

  const json = JSON.parse(bodyText);
  const result = json.choices?.[0]?.message?.content?.trim() || "";
  Logger.log("AI raw result: [" + result + "]");

  if (result.includes("資料請求")) return "資料請求";
  if (result.includes("操作方法")) return "操作方法";
  if (result.includes("不具合")) return "不具合報告";
  return "その他";
}
// ここまでAI処理

// テスト用
function testAzureClassification() {
  const testText = "画面が真っ白で保存ができません";
  const result = classifyInquiryWithAzure(testText);
  Logger.log("分類結果: " + result);
}

// 呼び出し
function onFormSubmit(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();

  const values = sheet.getRange(row, 1, 1, 5).getValues()[0];
  const name = values[1];
  const email = values[2];
  const category = values[3] ? values[3].trim() : "";
  const inquiry = values[4] ? values[4].trim() : "";

  Logger.log("category: [" + category + "]");
  Logger.log("inquiry: [" + inquiry + "]");

  let judgedCategory = "その他";

  if (inquiry) {
    try {
      judgedCategory = classifyInquiryWithAzure(inquiry);
    } catch (error) {
      Logger.log("AI分類失敗: " + error.message);
      judgedCategory = category || "その他";
    }
  } else if (category) {
    judgedCategory = category;
  } else {
    judgedCategory = "その他";
  }

  const configSheet = e.source.getSheetByName("設定");
  const configData = configSheet.getDataRange().getValues();

  let assignee = "info@example.com";

  for (let i = 1; i < configData.length; i++) {
    const configCategory = configData[i][0];
    const configEmail = configData[i][1];

    if (configCategory === judgedCategory && configEmail) {
      assignee = configEmail;
      break;
    }
  }

  const subject = `【新規問い合わせ】${judgedCategory} / ${name}`;
  const body = `
新しい問い合わせがありました。

【氏名】${name}
【メール】${email}
【申告種別】${category || "未選択"}
【AI判定種別】${judgedCategory}

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
  sheet.getRange(row, 10).setValue("未対応");
  sheet.getRange(row, 11).setValue("");
}

//チェック
function checkAzureProps() {
  const props = PropertiesService.getScriptProperties();
  Logger.log("ENDPOINT=" + props.getProperty("AZURE_OPENAI_ENDPOINT"));
  Logger.log("DEPLOYMENT=" + props.getProperty("AZURE_OPENAI_DEPLOYMENT"));
}

// 項目名設定
function setupInquirySheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("フォームの回答 1");

  const headers = [
    "Timestamp",
    "名前",
    "メールアドレス",
    "フォーム選択カテゴリ",
    "問い合わせ内容",
    "AI判定カテゴリ",
    "担当者メール",
    "通知日時",
    "通知結果",
    "対応ステータス",
    "対応メモ"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

// 選択肢追加
function setupStatusDropdown() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("フォームの回答 1");

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["未対応", "対応中", "完了", "保留"], true)
    .setAllowInvalid(false)
    .build();

  sheet.getRange("J2:J").setDataValidation(rule);
}

// 書式設定
function setupStatusColors() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("フォームの回答 1");

  const range = sheet.getRange("J2:J");

  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("未対応")
      .setBackground("#d9d9d9")
      .setRanges([range])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("対応中")
      .setBackground("#fff2cc")
      .setRanges([range])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("完了")
      .setBackground("#d9ead3")
      .setRanges([range])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("保留")
      .setBackground("#f4cccc")
      .setRanges([range])
      .build()
  ];

  sheet.setConditionalFormatRules(rules);
}
