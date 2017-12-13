//Santi 12/12/2017
// Mètodes per crear un full de càlcul en una pàgina per alumne a partir dels
// comentaris dels alumnes posats en format de taula al full de càlcul on executem l'script
/*
Per exemple, aplicant l'script sobre les següents 2 columnes d'un full de càlcul:

ALUMNE	                    COMENTARIS
Spencer, John           	  "Assistència: 7 faltes d'assistència injustificades i 5 retrassos.
                            Pràctiques: 100% fetes. Nota mitja: 7.5
                            Projecte: OK"
Bici , Sebastia             "Assistència: 5 faltes d'assistència injustificades.
                            Pràctiques: 75% fetes.  Nota mitja: 6.75
                            Projecte: inacabat"
Hairss, David	              "Assistència: 1 falta d'assistència injustificada.
                            Pràctiques: 75% fetes. Nota mitja: 7
                            Projecte: OK"

Ens crearà un altre full de càlcul en 3 pàgines que contindran el nom i comentaris de cadascun d'ells per separat. A l'script
l'indicarem la casella on està situat el text ALUMNE.
*/
/**
 * Si executem l'script onOpen() des de l'editor d'scripts se'ns crearà l'opció de menú al full de càlcul per executar l'script
 Eventhandler for spreadsheet opening - add a menu. Crea l'opció de menú de la fulla
 */

function onOpen() {
  SpreadsheetApp
    .getUi()
    .createMenu('Crear comentaris')
    .addItem('Crear comentaris', 'demanarDades')    //Opció de menú que executa la funció demanarDades
    .addToUi();

} // onOpen()

//Demana el nom del fitxer i la cel·la inicial de dades per poder fer el llistat. Si tot va bé crida al mètode que genera el full de càlcul en comentaris
function demanarDades(nomFitxer,celaInicial){

  // Display a dialog box with a title, message, inputs fields, and "Accept" and "Cancel" buttons. The
  // user can also close the dialog by clicking the close button in its title bar.
  var ui = SpreadsheetApp.getUi();
  var resposta = ui.prompt('Dades per generar el document', 'Nom del document', ui.ButtonSet.OK_CANCEL);
  // Process the user's response.
  if (resposta.getSelectedButton() == ui.Button.OK) {
    nomFitxer = resposta.getResponseText();
    resposta = ui.prompt('Dades per generar el document', 'Cel·la inicial', ui.ButtonSet.OK_CANCEL);
    celaInicial = resposta.getResponseText();
 
    // Process the user's response.
    if (resposta.getSelectedButton() == ui.Button.OK) {
      crearComentaris(nomFitxer,celaInicial);
    }
  }
}

//Funció que crea el full de càlcul a partir dels comentaris que troba a la cel·la indicada
function crearComentaris(nomFitxer,celaInicial){
  // Creem la fulla en lo nom indicat
  var fulla= SpreadsheetApp.create(nomFitxer);
  var pagina=fulla.getActiveSheet();
  // Selecciono la cel·la indicada
  var range=fulla.getRange(celaInicial);
  // Obtenim valor necessaris per crear el full
  var activeSheet = SpreadsheetApp.getActiveSheet();
  // Fila de la cel·la indicada 
  var activeRowIndex = range.getRow();
  // Columna de la cel·la indicada 
  var activeColumnIndex = range.getColumn();  
  // Número de columnes totals a imprimir (solen ser 2) 
  var numberOfColumns = activeSheet.getLastColumn()-activeColumnIndex+1;
  // Número de files totals a imprimir, que serà el número de pàgines del nou full de càlcul 
  var numberOfRows = activeSheet.getLastRow()-activeRowIndex;
  // Dades a copiar guardades a una matriu
  var activeRow = activeSheet.getRange(activeRowIndex+1, activeColumnIndex, numberOfRows, numberOfColumns).getValues();
  // Textos de la capçalera de les columnes, que copiarem a cada pàgina
  var headerRow = activeSheet.getRange(activeRowIndex, activeColumnIndex, 1, numberOfColumns).getValues();
  
  // Recorrem les dades de la matriu
  for(var i in activeRow){
    // Posem el nom de l'alumne com a nom de la pàgina
    pagina.setName(activeRow[i][0]);
    
    // Recorrem les diferents columnes    
    for(var col=1;col<=numberOfColumns;col++){
      // Posem els títols de les columnes
      pagina.setActiveRange(pagina.getRange(1,col));
      pagina.getActiveCell().setValue(headerRow[0][col-1]);
    
      // Ara les dades dels alumnes
      pagina.setActiveRange(pagina.getRange(2,col));
      pagina.getActiveCell().setValue(activeRow[i][col-1]);
    
      //Fem que la columna ajuste l'amplada al seu contingut
      pagina.autoResizeColumn(col);    
    }
  
    //Insertem una nova pàgina
    pagina=fulla.insertSheet();
    
  }
  //La última pàgina insertada la borro ja que està en blanc
  fulla.deleteSheet(pagina);
  
  //Mostro missatge de que s'ha creat el nou full de càlcul
  SpreadsheetApp.getUi().alert('Nou full de càlcul '+nomFitxer+' creat a l\'arrel del teu Google Drive');
};


//Santi 2017
/**
 * Counts the occurrencies of the "text" within the "values" interval.
 * The interval can also contain in its cells the text with a one digit number before it.
 * If so we add this one digit number to the counting variable.
 * For example:
 * calculateSum(A1:A5, "FJ") would return 12 in this context:
 *    1   2   3   4   5
 *A   FJ     2FJ     9FJ
 */
function countText(values,text) {
 var sum=0;
 for (var i=0; i<values[0].length; i++) {
   if(values[0][i]==text) sum++;
   else 
     if(String(values[0][i]).substring(1)==text) sum+=Number(String(values[0][i]).substring(0,1));
 }
 return sum;
};
