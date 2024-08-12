function checkIndexing() {
  var sheetName = "Main"; // Назва аркуша, що використовується
  var rangeNotation = "A2:A10"; // Діапазон комірок із URL-адресами
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    Logger.log("Аркуш із назвою '" + sheetName + "' не знайдено.");
    return;
  }
  
  var range = sheet.getRange(rangeNotation);
  var values = range.getValues();
  var startRow = range.getRow(); // Отримуємо початковий рядок діапазону
  
  for (var i = 0; i < values.length; i++) {
    var url = values[i][0];
    Logger.log("Перевірка URL: " + url);
    if (url !== "") {
      var attempts = 0;
      var maxAttempts = 3;
      var indexed = false;
      var captchaDetected = false;
      
      while (attempts < maxAttempts) {
        attempts++;
        try {
          var searchUrl = "https://www.google.com/search?q=site:" + encodeURIComponent(url);
          var response = UrlFetchApp.fetch(searchUrl, {muteHttpExceptions: true});
          var html = response.getContentText();
          
          // Перевірка наявності фрази "No results found for site:" і "did not match any documents"
          var noResults = html.indexOf("No results found for") !== -1 || html.indexOf("did not match any documents") !== -1;
          captchaDetected = html.indexOf("Our systems have detected unusual traffic from your computer network") !== -1;
          
          if (!captchaDetected) {
            indexed = !noResults;
            Logger.log("Проіндексовано: " + indexed);
            break;
          } else {
            Logger.log("Виявлено капчу, спроба повторити");
          }
        } catch (e) {
          Logger.log("Помилка для " + url + ": " + e.message);
        }
        // Затримка перед повторною спробою
        Utilities.sleep(5000);
      }
      
      var statusCell = sheet.getRange(startRow + i, 2); // Запис у колонку B з урахуванням стартового рядка
      var linkCell = sheet.getRange(startRow + i, 3); // Запис у колонку C з урахуванням стартового рядка
      
      if (captchaDetected) {
        statusCell.setValue("Виявлено капчу").setFontColor("orange");
      } else if (indexed) {
        statusCell.setValue("Проіндексовано").setFontColor("green");
      } else {
        statusCell.setValue("Не проіндексовано").setFontColor("red");
      }
      
      // Запис посилання на пошук у третю колонку
      linkCell.setFormula('HYPERLINK("' + searchUrl + '"; "Перевірити")');
      
      // Додавання затримки між запитами
      Utilities.sleep(5000);
    }
  }
}
