/**
 * ==============================================================
 * SERENE Horizon 耳機預購系統 - Google Apps Script (GAS) 後端處理程式
 * ==============================================================
 * 
 * 功能說明：
 * 1. 接收前端網頁送出的顧客預購訂單資料 (doPost)。
 * 2. 自動將訂單寫入 Google 試算表（首次執行會自動建立格式化欄位標題）。
 * 3. 自動寄送「預購成功確認信」給顧客的 Email。
 * 4. 支援 CORS 跨網域請求與除錯測試 (doGet)。
 * 
 * 部署指引：
 * 1. 在 Google 試算表上方選單點選「擴充功能」->「Apps Script」。
 * 2. 將此檔案全部程式碼貼上覆蓋原始內容。
 * 3. 點選右上角「部署」->「新增部署作業」。
 * 4. 種類選擇「網路應用程式 (Web App)」。
 * 5. 執行身分選擇「我 (自己的 Google 帳號)」。
 * 6. 誰可以存取選擇「所有人 (Anyone)」（重要：務必選所有人，前端才能送出資料）。
 * 7. 點選「部署」並授權存取，複製產生的「網路應用程式網址 (Web App URL)」。
 * 8. 將該網址貼入 index.html 的 GAS_ENDPOINT_URL 變數中。
 */

// 試算表欄位定義
const HEADERS = [
  "訂單編號",
  "下單時間",
  "預購方案",
  "訂購數量",
  "單價",
  "應付總額",
  "收件人姓名",
  "聯絡電話",
  "電子信箱",
  "收件地址",
  "付款方式",
  "發票類型",
  "訂單備註",
  "付款狀態",
  "出貨狀態"
];

// ==============================================================
// 【LINE Messaging API 推播設定】
// 請填入您在 LINE Developers 申請的 Channel Access Token 與您的 User ID (或群組 Group ID)
// ==============================================================
const LINE_CHANNEL_ACCESS_TOKEN = "qVRDzJN8nj1Dc+Y5MI5dPIZyFBouLIllUB0lF0lhdALIoWv4LLiLHDM0Dzm5liU4yBs39i3Jdi1j+YLRpkpZD5ppPTdIa3CFlIxXkhwqCZSvB5X+I2wpiOxdHb8pP5YxgHCww4UKcZNy4wdvohHcFgdB04t89/1O/w1cDnyilFU=";
const LINE_USER_ID = "U0670003881102246a2ccab750e1bd01c"; // 您的 LINE User ID 或 群組 ID (以 U 或 C 開頭)


/**
 * 處理 POST 請求 (接收訂單並寫入試算表 + 寄信)
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  // 最多等待 10 秒取得鎖定，避免高併發重複寫入問題
  lock.tryLock(10000);

  try {
    // 1. 判斷是否為綠界科技付款結果回傳 (ReturnURL)
    if (e && e.parameter && e.parameter.RtnCode) {
      return handleEcpayCallback(e.parameter);
    }

    // 2. 否則為前端顧客預購表單送出
    let data;
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (jsonErr) {
        data = e.parameter;
      }
    } else if (e.parameter) {
      data = e.parameter;
    } else {
      throw new Error("No data received");
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // 若試算表為空，自動建立欄位標頭
    if (sheet.getLastRow() === 0) {
      setupHeaders(sheet);
    }

    // 格式化當前時間 (台灣時區 GMT+8)
    const formattedTime = data.timestamp || Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");

    // 準備寫入試算表的一筆資料
    const row = [
      data.orderId || "SRN-" + Date.now(),
      formattedTime,
      data.plan || "尊榮旗艦 · 高階降噪版",
      data.quantity || 1,
      data.unitPrice || 5999,
      data.totalAmount || 5999,
      data.customerName || "",
      data.customerPhone || "",
      data.customerEmail || "",
      data.customerAddress || "",
      data.paymentMethod || "信用卡線上刷卡",
      data.invoiceType || "個人雲端發票",
      data.notes || "無",
      "待付款 (信用卡轉跳中)",
      "待處理"
    ];

    // 新增至試算表最末列
    sheet.appendRow(row);

    // 寄送顧客確認信
    if (data.customerEmail && data.customerEmail.includes("@")) {
      sendOrderConfirmationEmail(data);
    }

    // 發送 LINE 即時通知給管理員
    if (LINE_CHANNEL_ACCESS_TOKEN && !LINE_CHANNEL_ACCESS_TOKEN.includes("YOUR_LINE_")) {
      sendLineNotification(data);
    }

    // 回傳成功訊息
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "預購訂單已成功儲存並發送確認信",
      orderId: data.orderId
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log("doPost Error: " + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}

/**
 * 處理 GET 請求 (可用於測試服務是否運作正常)
 */
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "ok",
    service: "SERENE Horizon Pre-order Backend API",
    time: Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss")
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * 自動建立並美化標題列
 */
function setupHeaders(sheet) {
  sheet.appendRow(HEADERS);
  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setBackground("#2C221E"); // 摩卡深色底
  headerRange.setFontColor("#FFFFFF"); // 白色文字
  headerRange.setFontWeight("bold");
  headerRange.setHorizontalAlignment("center");
  sheet.setFrozenRows(1); // 凍結第一列
}

/**
 * 寄送美觀的 HTML 預購成功確認信
 */
function sendOrderConfirmationEmail(data) {
  const subject = `【SERENE Horizon】預購成功確認通知（訂單編號：${data.orderId}）`;
  
  const formattedTotal = Number(data.totalAmount).toLocaleString();
  const formattedUnitPrice = Number(data.unitPrice).toLocaleString();

  const htmlBody = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, "Noto Sans TC", sans-serif; background-color: #f7f7f9; margin: 0; padding: 20px; color: #333333; }
        .container { max-width: 600px; margin: 0 auto; background: #ffffff; border-radius: 16px; overflow: hidden; box-shadow: 0 4px 20px rgba(0,0,0,0.06); border: 1px solid #eaeaea; }
        .header { background: linear-gradient(135deg, #1C1A17 0%, #2C221E 100%); color: #ffffff; padding: 36px 30px; text-align: center; }
        .header h1 { margin: 0; font-size: 24px; letter-spacing: 2px; color: #E8D5A0; }
        .header p { margin: 8px 0 0 0; font-size: 13px; color: #d0c8be; letter-spacing: 1px; }
        .content { padding: 32px 30px; }
        .greeting { font-size: 16px; font-weight: bold; margin-bottom: 16px; color: #222; }
        .intro { font-size: 14px; line-height: 1.7; color: #555; margin-bottom: 24px; }
        .card { background: #faf8f5; border: 1px solid #eee5dc; border-radius: 12px; padding: 20px; margin-bottom: 24px; }
        .card-title { font-size: 14px; font-weight: bold; color: #8C6D53; margin-bottom: 12px; text-transform: uppercase; letter-spacing: 1px; }
        .item-row { display: flex; justify-content: space-between; font-size: 13px; padding: 6px 0; border-bottom: 1px dashed #e5ddd3; }
        .item-row:last-child { border-bottom: none; }
        .item-label { color: #777; }
        .item-value { font-weight: 600; color: #222; text-align: right; }
        .total-row { display: flex; justify-content: space-between; font-size: 16px; font-weight: bold; padding-top: 12px; margin-top: 8px; border-top: 2px solid #D4AF37; color: #222; }
        .total-price { color: #B8860B; }
        .notice-box { background: #fffdf7; border-left: 4px solid #D4AF37; padding: 14px 16px; border-radius: 4px; font-size: 13px; line-height: 1.6; color: #66522c; margin-bottom: 24px; }
        .footer { background: #f2f0ec; padding: 20px 30px; text-align: center; font-size: 12px; color: #888; }
        .footer a { color: #8C6D53; text-decoration: none; }
      </style>
    </head>
    <body>
      <div class="container">
        <!-- Header -->
        <div class="header">
          <h1>SERENE Horizon</h1>
          <p>空間自適應降噪無線耳機 · 官方預購確認</p>
        </div>

        <!-- Content -->
        <div class="content">
          <div class="greeting">${data.customerName} 您好：</div>
          <div class="intro">
            感謝您預購 <strong>SERENE Horizon</strong> 空間自適應降噪無線耳機！<br>
            我們已成功收到您的預購資訊，以下為您的詳細訂購清單，請妥善留存此郵件以便後續查詢。
          </div>

          <!-- Order Summary Card -->
          <div class="card">
            <div class="card-title">預購訂單摘要</div>
            <table width="100%" cellpadding="6" cellspacing="0" style="font-size: 13px;">
              <tr>
                <td style="color: #777;">訂單編號：</td>
                <td style="font-weight: bold; text-align: right; font-family: monospace;">${data.orderId}</td>
              </tr>
              <tr>
                <td style="color: #777;">下單時間：</td>
                <td style="text-align: right;">${data.timestamp}</td>
              </tr>
              <tr>
                <td style="color: #777;">預購方案：</td>
                <td style="font-weight: bold; text-align: right; color: #3A2A20;">${data.plan}</td>
              </tr>
              <tr>
                <td style="color: #777;">預購數量：</td>
                <td style="text-align: right;">${data.quantity} 台</td>
              </tr>
              <tr>
                <td style="color: #777;">單價：</td>
                <td style="text-align: right;">NT$ ${formattedUnitPrice}</td>
              </tr>
              <tr>
                <td style="color: #777;">運費：</td>
                <td style="text-align: right; color: #2e7d32; font-weight: bold;">全台免運（NT$ 0）</td>
              </tr>
              <tr style="border-top: 2px solid #D4AF37;">
                <td style="font-size: 15px; font-weight: bold; padding-top: 10px;">應付總額：</td>
                <td style="font-size: 16px; font-weight: bold; text-align: right; color: #B8860B; padding-top: 10px;">NT$ ${formattedTotal}</td>
              </tr>
            </table>
          </div>

          <!-- Shipping Info Card -->
          <div class="card">
            <div class="card-title">配送與開立資訊</div>
            <table width="100%" cellpadding="6" cellspacing="0" style="font-size: 13px;">
              <tr>
                <td style="color: #777; width: 30%;">收件人：</td>
                <td style="font-weight: 600;">${data.customerName}</td>
              </tr>
              <tr>
                <td style="color: #777;">聯絡電話：</td>
                <td>${data.customerPhone}</td>
              </tr>
              <tr>
                <td style="color: #777;">收件地址：</td>
                <td>${data.customerAddress}</td>
              </tr>
              <tr>
                <td style="color: #777;">付款方式：</td>
                <td>${data.paymentMethod}</td>
              </tr>
              <tr>
                <td style="color: #777;">發票類型：</td>
                <td>${data.invoiceType}</td>
              </tr>
              ${data.notes && data.notes !== "無" ? `
              <tr>
                <td style="color: #777;">備註事項：</td>
                <td>${data.notes}</td>
              </tr>` : ""}
            </table>
          </div>

          <!-- Notice -->
          <div class="notice-box">
            <strong>★ 預購權益與後續流程叮嚀：</strong><br>
            • <strong>信用卡付款</strong>：若您尚未完成線上刷卡授權，請透過官方結帳頁面提供的專屬連結前往銀行安全通道完成刷卡。<br>
            • <strong>出貨通知</strong>：商品出貨前我們將寄發簡訊與 Email 附帶黑貓宅急便貨態追蹤碼。<br>
            • <strong>安心保障</strong>：享有 1 年原廠硬體保固與收到商品後 7 天試聽鑑賞期無條件退換保障。
          </div>

          <div style="font-size: 13px; color: #666; line-height: 1.6;">
            如有任何訂單疑問或欲變更收件資訊，歡迎隨時回覆此信件或聯繫客服中心。<br>
            祝您聆聽愉快，生活美滿！
          </div>
        </div>

        <!-- Footer -->
        <div class="footer">
          SERENE Horizon Acoustic Taiwan · 官方旗艦直營<br>
          客服信箱：support@serene-audio.tw ｜ 服務時間：週一至週五 09:30 - 18:30
        </div>
      </div>
    </body>
    </html>
  `;

  // 寄出郵件
  MailApp.sendEmail({
    to: data.customerEmail,
    subject: subject,
    htmlBody: htmlBody
  });
}

/**
 * ==============================================================
 * LINE 即時訂單通知模組 (LINE Messaging API)
 * ==============================================================
 */

/**
 * 發送 LINE 即時通知給管理員
 * @param {Object} data 訂單資料
 */
function sendLineNotification(data) {
  if (!LINE_CHANNEL_ACCESS_TOKEN || !LINE_USER_ID) return;

  const url = "https://api.line.me/v2/bot/message/push";
  
  const messageText = 
`🔔【SERENE Horizon 收到新預購訂單】

📋 訂單編號：${data.orderId || "未提供"}
🕒 下單時間：${data.timestamp || "即時"}
🎧 預購方案：${data.plan}
📦 訂購數量：${data.quantity || 1} 台
💰 應付總額：NT$ ${Number(data.totalAmount || 0).toLocaleString()}

👤 顧客姓名：${data.customerName}
📞 聯絡電話：${data.customerPhone}
✉️ 電子信箱：${data.customerEmail}
📍 配送地址：${data.customerAddress}
🧾 發票類型：${data.invoiceType || "個人雲端發票"}
📝 訂單備註：${data.notes || "無"}

💳 付款狀態：待付款（準備前往綠界信用卡刷卡）
🚚 出貨狀態：待處理`;

  const payload = {
    to: LINE_USER_ID,
    messages: [
      {
        type: "text",
        text: messageText
      }
    ]
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: "post",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + LINE_CHANNEL_ACCESS_TOKEN
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    Logger.log("LINE 推播回應碼：" + response.getResponseCode());
  } catch (err) {
    Logger.log("LINE 推播失敗：" + err.toString());
  }
}

/**
 * 測試 LINE 推播功能是否正常運作
 * （可在 Google Apps Script 介面上方下拉選單選取「testLinePush」並點擊「執行」進行測試）
 */
function testLinePush() {
  const testData = {
    orderId: "SRN-TEST-" + Math.floor(1000 + Math.random() * 9000),
    timestamp: Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss"),
    plan: "尊榮旗艦 · 高階降噪版 (Pro ANC)",
    quantity: 1,
    totalAmount: 5999,
    customerName: "測試管理員",
    customerPhone: "0912-345-678",
    customerEmail: "admin@test.com",
    customerAddress: "台北市信義區信義路五段7號",
    invoiceType: "手機條碼載具",
    notes: "這是一則測試通知"
  };
  sendLineNotification(testData);
}

/**
 * ==============================================================
 * 綠界科技 (ECPay) 金流伺服器回傳處理模組 (ReturnURL)
 * ==============================================================
 */

/**
 * 處理綠界金流伺服器背景回傳 (ReturnURL)
 * @param {Object} params 綠界 POST 回傳的表單參數
 */
function handleEcpayCallback(params) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const tradeNo = String(params.MerchantTradeNo || "").trim();
  const rtnCode = String(params.RtnCode || "").trim(); // '1' 代表刷卡交易成功
  const rtnMsg = params.RtnMsg || "交易完成";
  const tradeAmt = params.TradeAmt || "0";
  const paymentDate = params.PaymentDate || Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");
  const paymentType = params.PaymentType || "Credit_CreditCard";

  Logger.log("收到綠界付款回傳通知: TradeNo=" + tradeNo + ", RtnCode=" + rtnCode + ", RtnMsg=" + rtnMsg);

  if (rtnCode === "1") {
    // 遍歷試算表尋找該筆訂單編號 (位於 A 欄，第 1 欄)
    const dataRange = sheet.getDataRange().getValues();
    let updated = false;

    for (let i = 1; i < dataRange.length; i++) {
      if (String(dataRange[i][0]).trim() === tradeNo) {
        const rowNumber = i + 1;
        // 欄位 N (第 14 欄) 為「付款狀態」
        sheet.getRange(rowNumber, 14).setValue("已付款 (綠界授權成功)");
        // 欄位 O (第 15 欄) 為「出貨狀態」
        sheet.getRange(rowNumber, 15).setValue("備貨中");
        updated = true;
        Logger.log("試算表第 " + rowNumber + " 列訂單狀態已更新為【已付款】");
        break;
      }
    }

    // 發送 LINE 刷卡成功即時推播
    sendLinePaymentSuccessNotification({
      orderId: tradeNo,
      amount: tradeAmt,
      paymentDate: paymentDate,
      paymentType: paymentType,
      rtnMsg: rtnMsg
    });
  } else {
    Logger.log("綠界交易非成功代碼: " + rtnCode + " / " + rtnMsg);
  }

  // 綠界規範：伺服器端必須回傳純文字 1|OK，否則綠界會視為失敗並持續重試發送
  return ContentService.createTextOutput("1|OK");
}

/**
 * 發送綠界刷卡付款成功 LINE 通知
 * @param {Object} payData 付款資訊
 */
function sendLinePaymentSuccessNotification(payData) {
  if (!LINE_CHANNEL_ACCESS_TOKEN || !LINE_USER_ID) return;

  const url = "https://api.line.me/v2/bot/message/push";
  const messageText = 
`💳【SERENE Horizon 綠界刷卡成功！】

🎉 賀！顧客已順利完成信用卡付款！

📋 訂單編號：${payData.orderId}
💰 刷卡金額：NT$ ${Number(payData.amount || 0).toLocaleString()}
🕒 付款時間：${payData.paymentDate}
💳 支付方式：綠界信用卡線上刷卡（測試環境）
✅ 交易訊息：${payData.rtnMsg}

📊 Google 試算表狀態已自動更新：
【付款狀態】➔ 已付款 (綠界授權成功)
【出貨狀態】➔ 備貨中`;

  const payload = {
    to: LINE_USER_ID,
    messages: [
      {
        type: "text",
        text: messageText
      }
    ]
  };

  try {
    const res = UrlFetchApp.fetch(url, {
      method: "post",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + LINE_CHANNEL_ACCESS_TOKEN
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    Logger.log("LINE 付款成功推播回應碼：" + res.getResponseCode());
  } catch (err) {
    Logger.log("LINE 付款推播失敗：" + err.toString());
  }
}

/**
 * 測試綠界刷卡成功推播
 * （可在 Google Apps Script 介面上方下拉選單選取「testEcpayPaymentSuccess」並點擊「執行」進行測試）
 */
function testEcpayPaymentSuccess() {
  sendLinePaymentSuccessNotification({
    orderId: "SRN" + Date.now().toString().slice(-9) + "9999",
    amount: "5999",
    paymentDate: Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss"),
    paymentType: "Credit_CreditCard",
    rtnMsg: "Succeeded"
  });
}
