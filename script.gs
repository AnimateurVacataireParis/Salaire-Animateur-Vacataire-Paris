/** @OnlyCurrentDoc */


/** Pour lancer le script sur mobile */
function onEdit(e) {
    var feuilleApp = e.source.getActiveSheet();
  
    if(feuilleApp.getSheetId()!=0) return;
  
    if(feuilleApp.getRange('D23').getValue() == "Ajouter"){
      InsererUnService();
      feuilleApp.getRange('D23').setValue('Sur mobile CLIQUEZ ICI');
    }
    
    if(feuilleApp.getRange('J27').getValue() == "Ajouter"){
      InsererUnePaie();
      feuilleApp.getRange('J27').setValue('Sur mobile CLIQUEZ ICI');
    }

  }
  /** Pour lancer le script sur mobile */
  
  
  /** Réinitialise la date à l'ouverture du document */
  function onOpen(e) {
    var app = getSheetById(0); //Séléctionne la page du formulaire par id
    app.getRange('D12').setFormula('=TODAY()');
    Popup();
  }
  /** Réinitialise la date à l'ouverture du document */
  
  
  /** Récupère une feuille par ID */
  function getSheetById(gid){
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  
    return sheets.filter(function(sheet){
      return sheet.getSheetId() === gid;
    })[0];
  }
  /** Récupère une feuille par ID */
  
  
  /** Insere un service */
  function InsererUnService() {
    SpreadsheetApp.getActiveSpreadsheet().toast('Ajout en cours...');
    var app = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var date = app.getRange('D12').getDisplayValue().replace(/\//g, "").toString();
    var service = app.getRange('D8').getValue();
    var ligneService = LigneService(service);
    date = date.substring(2, 4)+"/"+date.substring(6, 8); //Récupère la date (mois et année) cellule D12
    var salaire = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(date); //Séléctionne la feuille du mois en cours
    if (salaire){//Ajoute une ligne si la feuille existe
      // salaire.insertRowsAfter(1,1);  //Ajoute une ligne complète
      salaire.getRange('A2:H2').insertCells(SpreadsheetApp.Dimension.ROWS); //Ajoute un interval de cellules
    }
    else if (!salaire) {    //Crée la feuille du moi en cours si elle n'existe pas
      SpreadsheetApp.getActiveSpreadsheet().insertSheet(date);
      var salaire = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(date);
      MiseEnForme(date); //Met en forme la feuille
      TropPercu(date);
      AjouterSalaire(date);
      AjouterHeuresNonPayees(date);
    }
    app.getRange('D12').copyTo( salaire.getRange('A2'), {contentsOnly:true} );
    app.getRange('D16').copyTo( salaire.getRange('B2'), {contentsOnly:true} );
    app.getRange('D8').copyTo( salaire.getRange('C2'), {contentsOnly:true} );
    salaire.getRange('D2').setValue(Horaires(ligneService));
    salaire.getRange('E2').setValue(Code(ligneService));
  
    var tauxH = TauxH(ligneService);
    var nombreH = NombreH(ligneService);
    var salaireBrut = tauxH * nombreH;
  
    salaire.getRange('F2').setValue(tauxH.toString().replace(".",",")); 
    salaire.getRange('G2').setValue(nombreH.toString().replace(".",","));
    salaire.getRange('H2').setValue(salaireBrut.toString().replace(".",","));  
  
    var ecole = app.getRange('D16').getValue();
    RechercheEcole(ecole);
    app.activate();
    SpreadsheetApp.getActiveSpreadsheet().toast('Ajouté avec succès !');
  
  };
  /** Insere un service */
  
  /** Insere une fiche de paie */
  function InsererUnePaie() {
    SpreadsheetApp.getActiveSpreadsheet().toast('Ajout en cours...');
    var app = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var date = app.getRange('J16').getDisplayValue().replace(/\//g, "").toString();
    
    date = date.substring(2, 4)+"/"+date.substring(0, 2); //Récupère la date (mois et année) cellule J16
    var salaire = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(date); //Séléctionne la feuille du mois en cours

    if (salaire){//Ajoute une ligne si la feuille existe
      // salaire.insertRowsAfter(1,1);  //Ajoute une ligne complète
      salaire.getRange('J2:S2').insertCells(SpreadsheetApp.Dimension.ROWS); //Ajoute un interval de cellules
    }
    else if (!salaire) {    //Crée la feuille du mois en cours si elle n'existe pas
      SpreadsheetApp.getActiveSpreadsheet().insertSheet(date);
      var salaire = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(date);
      MiseEnForme(date); //Met en forme la feuille
      TropPercu(date);
      AjouterSalaire(date);
      AjouterHeuresNonPayees(date)
    }
    app.getRange('J8').copyTo( salaire.getRange('J2'), {contentsOnly:true} );
    app.getRange('J16').copyTo( salaire.getRange('K2'), {contentsOnly:true} );

    var indexColonne = 12;
    var bonCode = "";
    while(bonCode == ""){
      if( salaire.getRange(1, indexColonne).getValue() == app.getRange('J12').getValue() ){
        bonCode = app.getRange('J20').getDisplayValues().toString().replace(".",",");
        salaire.getRange(2, indexColonne).setValue(bonCode);
      }
      indexColonne++;
    }
  
    app.activate();
    SpreadsheetApp.getActiveSpreadsheet().toast('Ajouté avec succès !');
  
  };
  /** Insere une fiche de paie */
  
  /** Met une nouvelle feuille */
  function MiseEnForme(feuilleSalaire) {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(feuilleSalaire);
    spreadsheet.getRange('J1').setValue('Date fiche de paie');
    spreadsheet.getRange('K1').setValue('Date d\'origine');
    spreadsheet.getRange('L1').setValue('V85');
    spreadsheet.getRange('M1').setValue('VAH');
    spreadsheet.getRange('N1').setValue('VAC');
    spreadsheet.getRange('O1').setValue('V87');
    spreadsheet.getRange('P1').setValue('V83');
    spreadsheet.getRange('Q1').setValue('V90');
    spreadsheet.getRange('R1').setValue('V67');
    spreadsheet.getRange('S1').setValue('VRW');
    spreadsheet.getRange('K3').setValue('Total');
    spreadsheet.getRange('K3:S3')
    .setFontWeight('bold')
    .setFontColor('BACKGROUND')
    .setBackground('TEXT');

    spreadsheet.getRangeList(['L:L', 'S:S']).setNumberFormat('#,##0.00');
    spreadsheet.getRange('J1:S2').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
    spreadsheet.getRange('L3').setFormula('=IF(SUM(L$1:L2)>0;SUM(L$1:L2);"")');
    spreadsheet.getRange('M3').setFormula('=IF(SUM(M$1:M2)>0;SUM(M$1:M2);"")');
    spreadsheet.getRange('N3').setFormula('=IF(SUM(N$1:N2)>0;SUM(N$1:N2);"")');
    spreadsheet.getRange('O3').setFormula('=IF(SUM(O$1:O2)>0;SUM(O$1:O2);"")');
    spreadsheet.getRange('P3').setFormula('=IF(SUM(P$1:P2)>0;SUM(P$1:P2);"")');
    spreadsheet.getRange('Q3').setFormula('=IF(SUM(Q$1:Q2)>0;SUM(Q$1:Q2);"")');
    spreadsheet.getRange('R3').setFormula('=IF(SUM(R$1:R2)>0;SUM(R$1:R2);"")');
    spreadsheet.getRange('S3').setFormula('=IF(SUM(S$1:S2)>0;SUM(S$1:S2);"")');
    spreadsheet.getRange('A1').setValue('Date');
    spreadsheet.getRange('B1').setValue('École');
    spreadsheet.getRange('C1').setValue('Service');
    spreadsheet.getRange('D1').setValue('Horaires');
    spreadsheet.getRange('E1').setValue('Code');
    spreadsheet.getRange('F1').setValue('Taux horaire (brut)');
    spreadsheet.getRange('G1').setValue('Nombre d\'heure');
    spreadsheet.getRange('H1').setValue('Salaire (brut)');
    spreadsheet.getRange('F3').setValue('Total');
    spreadsheet.getRange('F3:H3')
    .setFontWeight('bold')
    .setFontColor('BACKGROUND')
    .setBackground('TEXT');
    spreadsheet.getRangeList(['E:E', 'H:H']).setNumberFormat('#,##0.00\\ [$€-1]');
    spreadsheet.getRange('G:G').setNumberFormat('General');
    spreadsheet.getRange('A:A').setNumberFormat('dd/MM/yyyy');
    spreadsheet.getRange('A1:H2').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
    spreadsheet.getRange('A1:H2').createFilter();
    spreadsheet.getRange('G3').setFormula('=SUBTOTAL(109;G$1:G2)');
    spreadsheet.getRange('H3').setFormula('=SUBTOTAL(109;H$1:H2)');
    spreadsheet
    .setColumnWidth(4, 233)
    .setColumnWidth(6, 145)
    .setColumnWidth(7, 129)
    .setColumnWidth(8, 109);
  };
  /** Met une nouvelle feuille */
  
  
  /** Définie les horaires */
  function Horaires(ligneService) {
    var spreadsheet = getSheetById(0);
    var code = spreadsheet.getRange(ligneService, 15).getValue();
    return code
  };
  /** Définie les horaires */
  
  
  /** Définie le code du service */
  function Code(ligneService) {
    var spreadsheet = getSheetById(0);
    var code = spreadsheet.getRange(ligneService, 17).getValue();
    return code
  };
  /** Définie le code du service */
  
  
  /** Définie le taux horaire */
  function TauxH(ligneService) {
    var spreadsheet = getSheetById(0);
    var TauxH = spreadsheet.getRange(ligneService, 19).getValue();
    return TauxH;
  };
  /** Définie le taux horaire */
  
  
  /** Définie le nombre d'heure */
  function NombreH(ligneService) {
    var spreadsheet = getSheetById(0);
    var NombreH = spreadsheet.getRange(ligneService, 21).getValue();
    return NombreH;
  };
  /** Définie le nombre d'heure */
  
  
  /** Recherche la ligne du service sélectionné */
  function LigneService(service) {
    var spreadsheet = getSheetById(0);
    var indexDebut = spreadsheet.getRange('N6').getRowIndex();
    var indexFin = spreadsheet.getRange('N6').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRowIndex();
    var index = null;
  
    while ( index == null || indexDebut<indexFin ) {
      
      if(service == spreadsheet.getRange(indexDebut, 14).getValue()){
        index = indexDebut;
        return index;
      }
      indexDebut++;
      
    }
  };
  /** Recherche la ligne du service sélectionné */
  
  
  /** Ajouter une école */
  function AjouterEcole(ecole) {
    var spreadsheet = getSheetById(0);
    var ligneTitre = 1;
    var cellule = spreadsheet.getRange(ligneTitre, 24).getValue();
    var ligneEcole;
  
    while ( "Mes écoles" != cellule ) {
      
      cellule = spreadsheet.getRange(ligneTitre, 24).getValue();
      ligneTitre++;
      
    }
  
    
    ligneEcole = ligneTitre+4;
    if( spreadsheet.getRange(ligneEcole, 24).getValue() ){
      spreadsheet.getRange(ligneEcole, 24, 2, 9).insertCells(SpreadsheetApp.Dimension.ROWS).merge().setValue(ecole);
    }
    spreadsheet.getRange(ligneEcole, 24).setValue(ecole);
  };
  /** Ajouter une école */
  
  
  /** Ajouter un trop perçu */
  function TropPercu(date) {
    var spreadsheet = getSheetById(0);
    var ligneTitre = 1;
    var cellule = spreadsheet.getRange(ligneTitre, 24).getValue();
    var ligneEcole;
  
    while ( "Heures trop perçues" != cellule ) {
      
      cellule = spreadsheet.getRange(ligneTitre, 24).getValue();
      ligneTitre++;
      
    }
  
    ligneTropPercue = ligneTitre+4;
    if( spreadsheet.getRange(ligneTropPercue, 24, 1, 9).getValue() ){
      spreadsheet.getRange(ligneTropPercue, 24, 1, 9).insertCells(SpreadsheetApp.Dimension.ROWS);
    }

    /** Ajoute formule trop perçu & date */

    spreadsheet.getRange(ligneTropPercue, 24).setValue(date);
    spreadsheet.getRange(ligneTropPercue, 25).setFormula('=IF(\'' + date + '\'!L$3-SUMIF(\'' + date + '\'!$E$1:$E;Y$5;\'' + date + '\'!$G$1:$G)>0;\'' + date + '\'!L$3-SUMIF(\'' + date + '\'!$E$1:$E;Y$5;\'' + date + '\'!$G$1:$G);"")');
    spreadsheet.getRange(ligneTropPercue, 25).autoFill(spreadsheet.getRange(ligneTropPercue, 25, 1, 8), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
    
    /** Ajoute formule trop perçu & date */
    
  };
  /** Ajouter un trop perçu */
  
  
  /** Ajouter un salaire */
  function AjouterSalaire(date) {
    var spreadsheet = getSheetById(0);
    var ligneTitre = 1;
    var cellule = spreadsheet.getRange(ligneTitre, 24).getValue();
    var ligneEcole;
  
    while ( "Estimation de salaire (brut)" != cellule ) {
      
      cellule = spreadsheet.getRange(ligneTitre, 24).getValue();
      ligneTitre++;
      
    }
  
    
    ligneSalaire = ligneTitre+4;
    if( spreadsheet.getRange(ligneSalaire, 24, 1, 9).getValue() ){
      spreadsheet.getRange(ligneSalaire, 24, 1, 9).insertCells(SpreadsheetApp.Dimension.ROWS);
      spreadsheet.getRange(ligneSalaire, 25, 1, 8).merge();
    }

    spreadsheet.getRange(ligneSalaire, 24).setValue(date);
    spreadsheet.getRange(ligneSalaire, 25).setFormula('=sum(\'' + date + '\'!H1:H)-\''+ date + '\'!H3');
  };
  /** Ajouter un salaire */
  
  
  /** Ajouter heures non payées */
  function AjouterHeuresNonPayees(date) {
    var spreadsheet = getSheetById(0);
    var ligneTitre = 1;
    var cellule = spreadsheet.getRange(ligneTitre, 14).getValue();
    var ligneEcole;
  
    while ( "Heures non payées" != cellule ) {
      
      cellule = spreadsheet.getRange(ligneTitre, 14).getValue();
      ligneTitre++;
      
    }
  
    
    ligneHeuresNonPayees = ligneTitre+4;
    if( spreadsheet.getRange(ligneHeuresNonPayees, 14, 1, 9).getValue() ){
      spreadsheet.getRange(ligneHeuresNonPayees, 14, 1, 9).insertCells(SpreadsheetApp.Dimension.ROWS);
    }

    /** Ajoute formule heures non payées & date */
    
    spreadsheet.getRange(ligneHeuresNonPayees, 14).setValue(date);
    spreadsheet.getRange(ligneHeuresNonPayees, 15).setFormula('=IF(SUMIF(\'' + date + '\'!$E$1:$E;Y$5;\'' + date + '\'!$G$1:$G)-\'' + date + '\'!L$3>0;SUMIF(\'' + date + '\'!$E$1:$E;Y$5;\'' + date + '\'!$G$1:$G)-\'' + date + '\'!L$3;"")');
    spreadsheet.getRange(ligneHeuresNonPayees, 15).autoFill(spreadsheet.getRange(ligneHeuresNonPayees, 15, 1, 8), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  /** Ajoute formule heures non payées & date */

  };
    /** Ajouter heures non payées */
  
  
  /** Recherche une école */
  function RechercheEcole(ecole) {
    var spreadsheet = getSheetById(0);
    var range = spreadsheet.getRange("X1:X");
    if(ecole){
      var ecoles = range.createTextFinder(ecole).findAll();
    }
    if(ecoles==""){
      AjouterEcole(ecole);
    }
  };
  /** Recherche une école */


  /** Pop-up bienvenue */
  function Popup() {
    var spreadsheet = getSheetById(0);

    if( spreadsheet.getRange('D31').getValue() == "" ){

      SpreadsheetApp.getUi().alert("Bienvenu ! \r\n\r\n Ici, vous pourrez visualiser en un coup d'oeil une estimation de votre salaire brut, vos heures manquantes et vos heures en trop. Controlez dès maintenant vos fiches de paies en toute simplicité.\r\n\r\n Si vous souhaitez lire le guide d'utilisation ou apporter votre contribution au projet le lien est à la cellule D31.\r\n\r\nBonne utilisation !");

      var lien = SpreadsheetApp.newRichTextValue()
    .setText("Voir les instructions")
    .setLinkUrl("https://github.com/AnimateurVacataireParis/Salaire-Animateur-Vacataire-Paris")
    .build();

      spreadsheet.getRange('B27:F29')
        .merge()
        .setBorder(true, true, true, true, null, null, '#e2e2e4', SpreadsheetApp.BorderStyle.SOLID)
        .setBackground('#7380ff');
      spreadsheet.getRange('B30:F36')
        .setBorder(true, true, true, true, null, null, '#e2e2e4', SpreadsheetApp.BorderStyle.SOLID)
        .setBackground('BACKGROUND');
      spreadsheet.getRange('D31:D32')
        .mergeVertically()
        .setBorder(true, true, true, true, null, null, '#7380ff', SpreadsheetApp.BorderStyle.SOLID);
      spreadsheet.getRange('D34')
        .setFontColor('#000')
        .setFontSize(11)
        .setFontWeight('bold')
        .setHorizontalAlignment('left')
        .setValue('Contact');
      spreadsheet.getRange('D35')
        .setFontColor('#000')
        .setFontSize(11)
        .setFontWeight('normal')
        .setHorizontalAlignment('left')
        .setValue('salaire.animateur.vacataire.paris@gmail.com');
      spreadsheet.getRange('B27:F29')
        .setFontColor('#fff')
        .setFontSize(14)
        .setFontWeight('bold')
        .setValue('Comment ça marche ?');
      spreadsheet.getRange('D31:D32')
        .setFontColor('#4285f4')
        .setFontSize(11)
        .setFontWeight('normal')
        .setRichTextValue(lien);
    }

  };
  /** Pop-up bienvenue */
