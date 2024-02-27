function getDatesAndEmails() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Munkalap1"); //"Munkalap1" helyett munkalap neve
  const startColumn = 1; // Kezdő dátum oszlop
  const endColumn = 2; // Végdátum oszlop
  const emailColumn = 4; // Email oszlop
  const url = "https://jsonplaceholder.typicode.com/comments";

  // Az utolsó sor beolvasása
  const lastRow = sheet.getLastRow();

  // i=1 ha nincs címsor, i=2 ha van
  for (let i = 1; i <= lastRow; i++) {
    // Kezdő és végdátum beolvasása
    const startDate = sheet.getRange(i, startColumn).getValue();
    const endDate = sheet.getRange(i, endColumn).getValue();

    // Dátumok közt eltelt nap
    const daysDiff = Math.floor((endDate - startDate) / (1000 * 60 * 60 * 24));

    // Random email cím lekérdezése az URL-ből
    const emailResponse = UrlFetchApp.fetch(url);
    const emailData = JSON.parse(emailResponse.getContentText());
    const randomEmail = emailData[Math.floor(Math.random() * emailData.length)].email;

    // Sorbaírás
    sheet.getRange(i, emailColumn -1).setValue(daysDiff);
    sheet.getRange(i, emailColumn).setValue(randomEmail);
  }
}