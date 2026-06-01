/**
 * Этот скрипт разворачивается в Google Apps Script (GAS)
 * Инструкции:
 * 1. Создайте Google Таблицу
 * 2. Перейдите: Расширения -> Apps Script
 * 3. Вставьте этот код
 * 4. Нажмите СУПЕРВАЖНО: Начать развертывание -> Новое развертывание -> Веб-приложение (доступ: Все)
 */

const SHEET_PARTICIPANTS = "Участники";
const SHEET_WINNERS = "Победители";

// Данные администратора рекомендуется хранить в Script Properties (PropertiesService)
// Для первоначальной установки можно задать их здесь или в настройках скрипта.
// PropertiesService.getScriptProperties().setProperty("ADMIN_LOGIN", "admin");
// PropertiesService.getScriptProperties().setProperty("ADMIN_PASSWORD", "password123");

// Функция для первоначальной генерации листов, если их нет
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEET_PARTICIPANTS)) {
    let s = ss.insertSheet(SHEET_PARTICIPANTS);
    s.appendRow(["Номер чека", "Имя", "Телефон", "Дата", "Статус"]);
  }
  if (!ss.getSheetByName(SHEET_WINNERS)) {
    let s = ss.insertSheet(SHEET_WINNERS);
    s.appendRow(["Номер чека", "Имя", "Телефон", "Дата розыгрыша"]);
  }
}

// Принимает GET-запросы
function doGet(e) {
  const action = e.parameter.action;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (action === "winners") {
    const sheet = ss.getSheetByName(SHEET_WINNERS);
    if (!sheet)
      return ContentService.createTextOutput(
        JSON.stringify({ success: true, winners: [] }),
      ).setMimeType(ContentService.MimeType.JSON);

    const data = sheet.getDataRange().getValues();
    const winners = [];

    for (let i = 1; i < data.length; i++) {
      winners.push({
        receipt: data[i][0],
        name: data[i][1],
        phone: data[i][2],
        date: data[i][3],
      });
    }

    return ContentService.createTextOutput(
      JSON.stringify({ success: true, winners: winners }),
    ).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "participants") {
    const token = e.parameter.token;
    const cache = CacheService.getScriptCache();
    if (!token || cache.get("auth_" + token) !== "valid") {
      return ContentService.createTextOutput(
        JSON.stringify({ success: false, message: "Необходима авторизация" }),
      ).setMimeType(ContentService.MimeType.JSON);
    }

    const sheet = ss.getSheetByName(SHEET_PARTICIPANTS);
    if (!sheet)
      return ContentService.createTextOutput(
        JSON.stringify({ success: true, participants: [] }),
      ).setMimeType(ContentService.MimeType.JSON);

    const data = sheet.getDataRange().getValues();
    const participants = [];

    // Номер чека (0), Имя (1), Телефон (2), Дата (3), Статус (4)
    for (let i = 1; i < data.length; i++) {
      participants.push({
        receipt: data[i][0],
        name: data[i][1],
        phone: data[i][2],
        date: data[i][3],
        won: data[i][4] === "Победитель",
      });
    }

    return ContentService.createTextOutput(
      JSON.stringify({ success: true, participants: participants }),
    ).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput("Google Script is Works!");
}

// Принимает POST-запросы
function doPost(e) {
  try {
    let data;
    try {
      data = JSON.parse(e.postData.contents);
    } catch (err) {
      data = JSON.parse(e.parameter.data || "{}");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    setup();

    // Проверка логина и пароля администратора
    if (data.action === "login") {
      const p = PropertiesService.getScriptProperties();
      const adminLogin = p.getProperty("ADMIN_LOGIN") || "admin";
      const adminPass = p.getProperty("ADMIN_PASSWORD") || "password123";

      if (data.login === adminLogin && data.password === adminPass) {
        const token = Utilities.getUuid();
        const cache = CacheService.getScriptCache();
        // Сохраняем токен в кэш на 6 часов (максимум для CacheService - 21600 секунд)
        cache.put("auth_" + token, "valid", 21600);
        return ContentService.createTextOutput(
          JSON.stringify({
            success: true,
            message: "Авторизация успешна",
            token: token,
          }),
        ).setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(
          JSON.stringify({
            success: false,
            message: "Неверный логин или пароль",
          }),
        ).setMimeType(ContentService.MimeType.JSON);
      }
    }

    if (data.action === "register") {
      if (!data.receipt || !data.name || !data.phone) {
        return ContentService.createTextOutput(
          JSON.stringify({
            success: false,
            message: "Не заполнены обязательные поля (receipt, name, phone)",
          }),
        ).setMimeType(ContentService.MimeType.JSON);
      }

      const receipt = String(data.receipt).trim();

      // Ровно 12 цифр
      if (!/^\d{12}$/.test(receipt)) {
        return ContentService.createTextOutput(
          JSON.stringify({
            success: false,
            message: "Неверный номер чека",
          }),
        ).setMimeType(ContentService.MimeType.JSON);
      }

      // Первые 6 цифр должны быть 000081
      if (!receipt.startsWith("000081")) {
        return ContentService.createTextOutput(
          JSON.stringify({
            success: false,
            message: "Неверный номер чека",
          }),
        ).setMimeType(ContentService.MimeType.JSON);
      }

      const phone = String(data.phone).trim();
      if (!/^\+?[0-9]{8,15}$/.test(phone)) {
        return ContentService.createTextOutput(
          JSON.stringify({
            success: false,
            message: "Некорректный номер телефона",
          }),
        ).setMimeType(ContentService.MimeType.JSON);
      }

      const sheet = ss.getSheetByName(SHEET_PARTICIPANTS);

      const values = sheet.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]).trim() === String(data.receipt).trim()) {
          return ContentService.createTextOutput(
            JSON.stringify({
              success: false,
              message: "Такой номер чека уже зарегистрирован",
            }),
          ).setMimeType(ContentService.MimeType.JSON);
        }
      }

      const todayDate = Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        "yyyy-MM-dd HH:mm:ss",
      );
      sheet.appendRow([
        "'" + data.receipt,
        data.name,
        "'" + data.phone,
        todayDate,
        "",
      ]);

      return ContentService.createTextOutput(
        JSON.stringify({ success: true, message: "Успешно зарегистрировано" }),
      ).setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === "drawWinner") {
      const token = data.token;
      const cache = CacheService.getScriptCache();
      if (!token || cache.get("auth_" + token) !== "valid") {
        return ContentService.createTextOutput(
          JSON.stringify({ success: false, message: "Необходима авторизация" }),
        ).setMimeType(ContentService.MimeType.JSON);
      }

      const sheet = ss.getSheetByName(SHEET_PARTICIPANTS);
      const winSheet = ss.getSheetByName(SHEET_WINNERS);
      const values = sheet.getDataRange().getValues();

      let eligible = [];
      for (let i = 1; i < values.length; i++) {
        if (values[i][4] !== "Победитель") {
          eligible.push({ rowIndex: i + 1, data: values[i] });
        }
      }

      if (eligible.length === 0) {
        return ContentService.createTextOutput(
          JSON.stringify({
            success: false,
            message: "Нет доступных участников для розыгрыша",
          }),
        ).setMimeType(ContentService.MimeType.JSON);
      }

      const winnerObj = eligible[Math.floor(Math.random() * eligible.length)];
      const winnerData = winnerObj.data;

      // Обновляем статус
      sheet.getRange(winnerObj.rowIndex, 5).setValue("Победитель");

      const drawDate = Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        "yyyy-MM-dd HH:mm:ss",
      );
      winSheet.appendRow([
        "'" + winnerData[0],
        winnerData[1],
        "'" + winnerData[2],
        drawDate,
      ]);

      return ContentService.createTextOutput(
        JSON.stringify({
          success: true,
          message: "Победитель выбран",
          winner: {
            receipt: winnerData[0],
            name: winnerData[1],
            phone: winnerData[2],
          },
        }),
      ).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(
      JSON.stringify({
        success: false,
        message: "Неизвестное действие (action)",
      }),
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ success: false, message: err.toString() }),
    ).setMimeType(ContentService.MimeType.JSON);
  }
}
