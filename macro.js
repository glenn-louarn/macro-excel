
function sheetsNamesMagasin() {
    var out = new Array()
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for (var i = 0; i < sheets.length; i++) {
        if (sheets[i].getName() != "Produit")
            out.push(sheets[i].getName())
    }
    return out
}

/*
  return nombre de produit
*/
function nbProduit(ss) {
    var searchVide = ss.getRange(depart).getValue();
    var position = ss.getRange(depart).getA1Notation();
    while (searchVide != word) {
        position = ss.getRange(position).offset(0, 1).getA1Notation();
        searchVide = ss.getRange(position).getValue();
    }
    return position
}
/*
  return position du word
*/
function findWordInLineV2(depart, word, spreadsheet) {
    var searchVide = spreadsheet.getRange(depart).getValue();
    var position = spreadsheet.getRange(depart).getA1Notation();
    while (searchVide != word) {
        position = spreadsheet.getRange(position).offset(0, 1).getA1Notation();
        searchVide = spreadsheet.getRange(position).getValue();
    }
    return position
}
function ajouterUnProduitUneFeuille(nameFeuille, ss) {
    ss.setActiveSheet(ss.getSheetByName(nameFeuille), true);
    ss.getRange('A3').activate();
    ss.getActiveSheet().insertRowsAfter(ss.getActiveRange().getLastRow(), 1);
    ss.getActiveRange().offset(ss.getActiveRange().getNumRows(), 0, 1, ss.getActiveRange().getNumColumns()).activate();
    ss.getRange('A4').activate();
    ss.getRange('3:3').copyTo(ss.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
}

//document a partir d'un template https://stackoverflow.com/questions/58611018/how-to-create-a-new-document-from-a-template-with-placeholders
function Ajouterunproduit() {
    var ss = SpreadsheetApp.getActive();
    var name = Browser.inputBox('Nom du produit', Browser.Buttons.OK_CANCEL);
    var prix = Browser.inputBox('Prix du produit', Browser.Buttons.OK_CANCEL);
    ss.setActiveSheet(ss.getSheetByName('Produit'), true);
    ss.getRange('A3').activate();
    ss.getActiveSheet().insertRowsAfter(ss.getActiveRange().getRow(), 1);
    ss.getRange('A4').activate();
    ss.getCurrentCell().setValue(name);
    ss.getRange('B4').activate();
    ss.getCurrentCell().setValue(prix);
    var namesMagasin = sheetsNamesMagasin();
    for (var i = 0; i < namesMagasin.length; i++) {
        ajouterUnProduitUneFeuille(namesMagasin[i], ss)
    }
}

function AjouterVentes() {
    AjouterAction('Ventes', 'L1')
};
function Ajoutercommandes() {
    AjouterAction('Commandes', 'E1')
};

function AjouterAction(nom, depart) {
    var ss = SpreadsheetApp.getActive();
    //recule de 2 column et monte D'un ligne 
    ss.getRange(findWordInLineV2(ss.getRange(findWordInLineV2(depart, nom, ss)).offset(1, 0).getA1Notation(), '', ss)).offset(1, -2).activate();

    //ajoute 2 column à la fin du tableau du tableau d'action 
    var sheet = ss.getActiveSheet();
    sheet.getRange(1, ss.getCurrentCell().getColumn(), sheet.getMaxRows(), 2).activate();
    ss.getActiveSheet().insertColumnsAfter(ss.getActiveRange().getLastColumn(), 2);
    ss.getActiveRange().offset(0, ss.getActiveRange().getNumColumns(), ss.getActiveRange().getNumRows(), 2).activate();
    ss.getCurrentCell().activate();
    //Copie et colle la dernieres action
    ss.getCurrentCell().offset(0, -2, 31, 2).copyTo(ss.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    //reset les valeurs de la nouvelle action
    ss.getCurrentCell().setValue(new Date());
    ss.getActiveRange().offset(2, 1).activate();
    ss.getCurrentCell().setValue(0);
    var destinationRange = ss.getActiveRange().offset(0, 0, 27);
    ss.getActiveRange().autoFill(destinationRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    var newActionCellule = ss.getCurrentCell().getA1Notation();
    //add ref contage coussin 
    ss.getRange(findWordInLineV2(depart, nom, ss)).offset(2, 3).activate();
    ss.getCurrentCell().setFormula(ss.getCurrentCell().getFormula() + '+' + newActionCellule);
    var destinationRange = ss.getActiveRange().offset(0, 0, 27);
    ss.getActiveRange().autoFill(destinationRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
};
