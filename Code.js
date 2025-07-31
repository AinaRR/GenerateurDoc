function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("📄 Génération de documents")
    .addItem("🔁 Un document par client", "generationDocParClient")
    .addItem("📄 Tous les clients sur une page (PDF)", "generationDocSurPage")
    .addItem("🧪 Générer client test (numéro de la ligne)", "testClient")
    .addToUi();

    creerModeleDocument();
}

/**
 * Crée un modèle de document Google Docs avec des balises de remplacement.
 * @param {string} titre - Le titre du document à créer.
 * @return {DocumentApp.Document} - Le document Google Docs créé.
 **/
function creerModeleDocument(titre) {
  var doc = DocumentApp.create(titre || "Google Form Test");
  var body = doc.getBody();

  // Insertion de champs personnalisés avec des balises
  body.appendParagraph("Client ID: {{Id}}").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph("Nom : {{Nom_client}}");
  body.appendParagraph("Adresse : {{Address_client}}");
  body.appendParagraph("Téléphone : {{Telephone}}");
  body.appendParagraph("Email : {{Email}}");
  body.appendParagraph("Date d’échéance : {{Date_échéance}}");

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // ou getSheetByName("NomFeuille")
  var headers = ["Id", "Nom_client", "Address_client", "Telephone", "Email", "Date_échéance", "URL Document"];

  // Écriture des en-têtes
  var range = sheet.getRange(1, 1, 1, headers.length);
  range.setValues([headers]);

  // Application du style : Gras, Noir, Georgia, Centré
  range.setFontWeight("bold");
  range.setFontColor("black");
  range.setFontSize(12);
  range.setFontFamily("Georgia");
  range.setHorizontalAlignment("center");
  range.setBackground("#edbb99");

  // Fixer les largeurs des colonnes A à G manuellement (en pixels)
  var columnWidths = [60, 120, 180, 100, 180, 140, 600];  // Exemple de tailles

  for (var col = 1; col <= columnWidths.length; col++) {
    sheet.setColumnWidth(col, columnWidths[col - 1]);
  }

  return doc;
}

/**
 * Remplit un modèle de document pour un client spécifique (à partir d'une ligne de la feuille).
 * @param {number} rowIndex - L’index de la ligne (>= 2) correspondant au client.
 **/
function remplirDocumentDepuisSheet(rowIndex) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Données");

  if (!sheet) return;

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (!headers || headers.length === 0) return;
  if (rowIndex < 2 || rowIndex > sheet.getLastRow()) return;

  var values = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  var data = {};

  for (var i = 0; i < headers.length; i++) {
    var cle = headers[i];
    var valeur = values[i];

    // Ajouter un 0 devant les numéros de téléphone numériques
    if (cle === "Telephone" && typeof valeur === "number") {
      valeur = "0" + valeur.toString();
    }

    // Formater les dates
    if (valeur instanceof Date) {
      valeur = Utilities.formatDate(valeur, Session.getScriptTimeZone(), "dd/MM/yyyy");
    }

    data[cle] = valeur;
  }

  var nomDoc = "Client_" + data["Id"] + "_" + data["Nom_client"];
  var doc = creerModeleDocument(nomDoc);
  var body = doc.getBody();

  // Remplacement des balises par les vraies données
  for (var key in data) {
    body.replaceText("{{" + key + "}}", data[key]);
  }

  doc.saveAndClose();

  SpreadsheetApp.getUi().alert("Document généré :\n" + doc.getUrl());
}


/**
 * Génère un document personnalisé pour chaque ligne (client) dans la feuille "Données".
 * Enregistre les URL des documents générés dans la feuille sur chaque ligne correspondante.
 **/
function generationDocParClient() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Confirmation", "Voulez-vous générer les documents pour tous les clients ?", ui.ButtonSet.YES_NO);
  if (response == ui.Button.NO) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Données");
  if (!sheet) return;

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // 🧹 Réinitialisation de la colonne G (URL Document), ligne 2 à la dernière ligne
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 7, lastRow - 1).clearContent(); // colonne G = 7
  }

  var dataRange = sheet.getRange(2, 1, lastRow - 1, headers.length);
  var toutesLesLignes = dataRange.getValues();

  var urlColIndex = headers.indexOf("URL Document");
  if (urlColIndex === -1) {
    sheet.getRange(1, headers.length + 1).setValue("URL Document");
    urlColIndex = headers.length;
  }

  var fichier = DriveApp.getFileById(ss.getId());
  var dossier;
  var parents = fichier.getParents();
  if (parents.hasNext()) {
    dossier = parents.next();
  } else {
    dossier = DriveApp.createFolder("Documents générés");
  }

  for (var i = 0; i < toutesLesLignes.length; i++) {
    var values = toutesLesLignes[i];
    var data = {};

    for (var j = 0; j < headers.length; j++) {
      var cle = headers[j];
      var valeur = values[j];

      if (cle === "Telephone" && typeof valeur === "number") {
        valeur = "0" + valeur.toString();
      }

      if (valeur instanceof Date) {
        valeur = Utilities.formatDate(valeur, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }

      data[cle] = valeur;
    }

    var nomDoc = "Client_" + data["Id"] + "_" + data["Nom_client"];
    var doc = creerModeleDocument(nomDoc);
    var body = doc.getBody();

    for (var key in data) {
      body.replaceText("{{" + key + "}}", data[key]);
    }

    doc.saveAndClose();

    var fichierDoc = DriveApp.getFileById(doc.getId());
    dossier.addFile(fichierDoc);
    DriveApp.getRootFolder().removeFile(fichierDoc);

    var url = doc.getUrl();
    sheet.getRange(i + 2, urlColIndex + 1).setValue(url);
  }

  SpreadsheetApp.getUi().alert("Tous les documents ont été générés avec leurs liens !");
}

/**
 * Génère un seul document contenant tous les clients et l’exporte en PDF.
 **/
function generationDocSurPage() {
    // Confirmation de l'utilisateur 
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Confirmation", "Voulez-vous générer les documents pour tous les clients sur une page ?", ui.ButtonSet.YES_NO);
  if (response == ui.Button.NO) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Données");
  if (!sheet) return;

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

  var docTitle = "Tous les clients - " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var doc = DocumentApp.create(docTitle);
  var body = doc.getBody();

  // Remplir le document avec les données de tous les clients
  data.forEach(function(row) {
    var entry = {};
    headers.forEach(function(cle, i) {
      var val = row[i];
      if (cle === "Telephone" && typeof val === "number") val = "0" + val.toString();
      if (val instanceof Date) val = Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
      entry[cle] = val;
    });

    body.appendParagraph("Client ID: " + entry["Id"]).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph("Nom : " + entry["Nom_client"]);
    body.appendParagraph("Adresse : " + entry["Address_client"]);
    body.appendParagraph("Téléphone : " + entry["Telephone"]);
    body.appendParagraph("Email : " + entry["Email"]);
    body.appendParagraph("Date d’échéance : " + entry["Date_échéance"]);
    body.appendParagraph("");
    body.appendHorizontalRule(); // Séparateur entre clients
  });

  doc.saveAndClose();

  // Convertir le document en PDF
  var pdfFile = DriveApp.getFileById(doc.getId()).getAs("application/pdf");

  var file = DriveApp.getFileById(ss.getId());
  var parents = file.getParents();
  var folder;

  if (parents.hasNext()) {
    folder = parents.next();
  } else {
    folder = DriveApp.createFolder("Documents générés");
  }

  var pdf = folder.createFile(pdfFile).setName(docTitle + ".pdf");

  SpreadsheetApp.getUi().alert("Document PDF généré :\n" + pdf.getUrl());
}

/**
 * Fonction de test pour générer un document pour le client de la 2eme ligne.
 **/
function testClient() {
  remplirDocumentDepuisSheet(2);//Modifiable selon préférence
}
