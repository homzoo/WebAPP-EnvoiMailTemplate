// Code pour envoyer des mails en automatique en fonction d'une date définie dans le spredsheet
// Le contenu du mail est défini dans un fichier docs

function sendBdayWishes(){
    var ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1fpXpudBDaHa0rR4-1Dezj_-SvW6_Ub1FmZsCvLOpAHs/edit#gid=1548574907");
    var sheet = ss.getSheetByName("Réponses au formulaire 1");// Nom de la feuille
    var templateId = '1goMIW8QhNPn2WSeYxVLA2Qm_m9kWcSXkVZP2nPv0Ozo';// Clé du fichier doc Template
    
     var cDate = new Date(); //Date du jour
     for(var i =2 ;i<=sheet.getLastRow(); i++){
  
        var bDate = sheet.getRange(i,4).getValue(); // Date du SpreadSheet 
  
        if(cDate.getDate()==bDate.getDate()){
            if(cDate.getMonth()==bDate.getMonth()){
                var name = sheet.getRange(i,2).getValue();
                var toMail= sheet.getRange(i,3).getValue();
                sendMail(sheet,templateId,name,toMail);
                sheet.getRange(i,5).setValue("Bday wishes sent"); 
            }
        }
     }
}



function sendMail(sheet,templateId,name,toMail){
 
  var docId = DriveApp.getFileById(templateId).makeCopy('temp').getId();
  var doc = DocumentApp.openById(docId);// Création du fichier temporaire
  var body = doc.getBody();
  
    body.replaceText('#name#',name);// maj du fichier temporaire
    doc.saveAndClose();// Enregistrer les modification
      
var url = "https://docs.google.com/feeds/download/documents/export/Export?id="+docId+"&exportFormat=html"; // Exporter le template en HTML via son ID
    var param = {
    method      : "get",
    headers     : {"Authorization": "Bearer " +     ScriptApp.getOAuthToken()}
    };
    
    var htmlBody =     UrlFetchApp.fetch(url,param).getContentText();
    
    var trashed = DriveApp.getFileById(docId).setTrashed(true);// Supprimer le fichier temporaire
    
    MailApp.sendEmail(toMail,'Happy BirthDay '+name,' ' ,{htmlBody : htmlBody});// Envoi du mail

     
    }
  
