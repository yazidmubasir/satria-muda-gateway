/** SATRIA MUDA WRITE GATEWAY — minimal functional gateway. */
const GATEWAY_PROP_TOKEN='GATEWAY_TOKEN';
function doGet(){return ContentService.createTextOutput(JSON.stringify({ok:true,service:'SATRIA MUDA GATEWAY',version:'1.0.0'})).setMimeType(ContentService.MimeType.JSON);}
function doPost(e){try{const body=JSON.parse(e.postData.contents||'{}');const expected=PropertiesService.getScriptProperties().getProperty(GATEWAY_PROP_TOKEN)||'SATRIA-GATEWAY-2026';if(String(body.token||'')!==expected)throw new Error('Token Gateway tidak valid.');const action=String(body.action||'');let data;switch(action){case'SPREADSHEET_READ':data=readSheet_(body);break;case'SPREADSHEET_APPEND':data=appendRow_(body);break;case'SPREADSHEET_DELETE_ROW':data=deleteRow_(body);break;case'GATEWAY_PING':data={pong:true};break;default:throw new Error('Action tidak dikenal: '+action)}return json_({ok:true,data});}catch(err){return json_({ok:false,message:err.message||String(err)})}}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function ss_(id){if(!id)throw new Error('spreadsheetId wajib.');return SpreadsheetApp.openById(id)}
function sh_(b){const sh=ss_(b.spreadsheetId).getSheetByName(String(b.sheet||''));if(!sh)throw new Error('Sheet tidak ditemukan: '+b.sheet);return sh}
function readSheet_(b){const sh=sh_(b);return {headers:sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0],rows:sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues():[]}}
function appendRow_(b){const sh=sh_(b);sh.appendRow(b.row||[]);return {rowNumber:sh.getLastRow()}}
function deleteRow_(b){const sh=sh_(b);const n=Number(b.rowNumber);if(n<2||n>sh.getLastRow())throw new Error('Nomor baris tidak valid.');sh.deleteRow(n);return {deletedRow:n}}
function setGatewayToken(token){if(!token)throw new Error('Token kosong.');PropertiesService.getScriptProperties().setProperty(GATEWAY_PROP_TOKEN,String(token));return{ok:true}}
