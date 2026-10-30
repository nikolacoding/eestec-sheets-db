const SHEET_FORM = "Unos";
const SHEET_DATABASE = "Lista";

// novi meni u traci iznad
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Upravljaj bazom članova')
    .addItem('Dodaj člana ', 'addMember')
    .addItem('Obriši člana (odabrani red)', 'deleteSelectedRow')
    .addSeparator()
    .addItem('Sortiraj po fakultetu', 'sortByFaculty')
    .addItem('Sortiraj po datumu', 'sortByDate')
    .addItem('Sortiraj po statusu', 'sortByStatus')
    .addToUi();
}

function addMember() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName(SHEET_FORM);
  const dbSheet = ss.getSheetByName(SHEET_DATABASE);
  
  if (!formSheet || !dbSheet) {
    SpreadsheetApp.getUi().alert("Greška: Skripta traži ali ne može pronaći dva specifično imenovana sheet-a: 'Forma' i 'Baza'.");
    return;
  }
  
  const name = formSheet.getRange("C3").getValue();
  const faculty = formSheet.getRange("C4").getValue();
  const indexNum = formSheet.getRange("C5").getValue();
  const phone = formSheet.getRange("C6").getValue();
  const email = formSheet.getRange("C7").getValue();
  let joinDate = formSheet.getRange("C8").getValue();
  const status = formSheet.getRange("C9").getValue();
  
  if (!name || !faculty || !phone || !email) {
    SpreadsheetApp.getUi().alert("Greška: Minimalni podaci da se kreira zapis su ime, fakultet+smjer, broj telefona i e-mail.");
    return;
  }
  
  dbSheet.appendRow([
    name,
    faculty,
    indexNum,
    phone,
    email,
    joinDate,
    status
  ]);
  
  const lastRow = dbSheet.getLastRow();
  dbSheet.getRange(lastRow, 6).setNumberFormat("dd.mm.yyyy.");
  
  clearForm();
  
  SpreadsheetApp.getUi().alert(name + " je uspješno registrovan.");
}

function clearForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getSheetByName(SHEET_FORM);
  formSheet.getRange("C3:C9").clearContent();
}

// brisanje trenutno selektovanog reda, mada ovo se može i ručno uraditi za neki pojedinačan zapis
// fazon može se pejstovati prazan red preko i onda sortirati po nekom kriterijumu
function deleteSelectedRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  
  if (activeSheet.getName() !== SHEET_DATABASE) {
    SpreadsheetApp.getUi().alert("Greška: Prebaci na listu članova da izvršiš ovu operaciju.");
    return;
  }
  
  const activeRow = activeSheet.getActiveCell().getRow();
  
  // mjera da se ne počisti zaglavlje tabele
  const targetMember = activeSheet.getRange(activeRow, 1).getValue();
  if (activeRow > 1) {
    const ui = SpreadsheetApp.getUi();

    if (targetMember != ""){
      const response = ui.alert(
        "Potvrdi brisanje",
        `Da li sigurno hoćeš da ukloniš člana '${targetMember}' uz sve njegove informacije? Ova operacija je neopoziva.`,
        ui.ButtonSet.YES_NO
      );
      
      if (response === ui.Button.YES) {
        activeSheet.deleteRow(activeRow);
        ui.alert(`Član '${targetMember}' je uspješno obrisan.`);
      }
    } else {
      // zapravo sa praznim imenom ali mi je mrsko pokrivati ostale nevalidnosti, nek to ostane na dusi onog ko unese pogresno i ne primjeti
      SpreadsheetApp.getUi().alert("Greška: Ciljaš red bez validnog zapisa.");
    }

  } else {
    SpreadsheetApp.getUi().alert("Greška: Ne možeš obrisati header, uozbilji se.");
  }
}

function sortDatabase(columnIndex, ascending) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName(SHEET_DATABASE);
  
  if (!dbSheet) return;
  
  const lastRow = dbSheet.getLastRow();
  const lastCol = dbSheet.getLastColumn();
  
  if (lastRow > 1) {
    const dataRange = dbSheet.getRange(2, 1, lastRow - 1, lastCol);
    dataRange.sort({ column: columnIndex, ascending: ascending });
    // SpreadsheetApp.getUi().alert("debug: sortirano");
  }
}

// asc
function sortByFaculty() {
  sortDatabase(2, true);
}

// desc
function sortByDate() {
  sortDatabase(6, false);
}

// asc
function sortByStatus() {
  sortDatabase(7, true);
}