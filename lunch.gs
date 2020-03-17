//Brute force search to remove duplicates
function removeDuplicates(){
  var ss = SpreadsheetApp.getActive();
  var sheet1 = ss.getSheetByName('responses');
  var sheet2 = ss.getSheetByName('avmelding');
   
  while (sheet2.getLastRow() >= 2){
    for (var j = 2; j <= sheet1.getLastRow(); j++){
      var d1 = sheet2.getRange(sheet2.getLastRow(),2).getValue().replace(/\s+/g, '').toUpperCase();
      var d2 = sheet1.getRange(j,2).getValue().replace(/\s+/g, '').toUpperCase();
      var d3 = sheet2.getRange(sheet2.getLastRow(),3).getValue().replace(/\s+/g, '').toUpperCase();
      var d4 = sheet1.getRange(j,3).getValue().replace(/\s+/g, '').toUpperCase();
      
      if ((d1 == d2) || (d3 == d4)){
        
        if (j != sheet1.getLastRow()){          
          sheet1.getRange(j, 2).setValue(sheet1.getRange(sheet1.getLastRow(),2).getValue());
          sheet1.getRange(j, 3).setValue(sheet1.getRange(sheet1.getLastRow(),3).getValue());
          sheet1.getRange(j, 4).setValue(sheet1.getRange(sheet1.getLastRow(),4).getValue());
        }
        
        sheet1.deleteRow(sheet1.getLastRow());
        j--;
       }
    }
    sheet2.deleteRow(sheet2.getLastRow());
  }
}


function sendEmail() {
  
  removeDuplicates();
  
  //Get spreadsheet access
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('responses');
  
  const rows = sheet.getLastRow() -1;
  const cols = sheet.getLastColumn() ;
  var range = sheet.getRange(2,1,rows,cols);
  
  //Randomize row-wise
  range.randomize();
   
  //How large you want the groups
  const groupSize = 4;
 
  //How many extra people there are
  const extras = rows % groupSize; 
  
  //How many groups there are
  const groups = (rows-extras)/groupSize;

  //Arrays containing the extra people
  var extras_name = [];
  var extras_age = [];
  var extras_email = [];
  var extras_study = [];
  var extras_msg = [];
 
  for (var j = extras-1; j >= 0; j--){
    var email_working = 1;
     try{
      MailApp.sendEmail(sheet.getRange(sheet.getLastRow()-j,3).getValue().replace(/\s+/g, '').toLowerCase(),'test', 'test-message');
    }
    catch(err){
      email_working = 0;
    }
    if (email_working){
      extras_name.push(sheet.getRange(sheet.getLastRow()-j,2).getValue());
      extras_email.push(sheet.getRange(sheet.getLastRow()-j,3).getValue().replace(/\s+/g, '').toLowerCase());
      extras_msg.push(sheet.getRange(sheet.getLastRow()-j,4).getValue());
    }
  }
  
  var iter = 0;
  var end = rows - groupSize + 2;
  
  for (var i = 2; i <= end; i = i + groupSize){
    var greeting = 'Hei '
    var about = 'Dette er hva dere har skrevet om dere selv:\n\n';
    var email ='';
    for (var t = 0; t< 4; t++){
      var email_working = 1;
      try{
        MailApp.sendEmail(sheet.getRange(i+t,3).getValue().replace(/\s+/g, '').toLowerCase(),'test', 'test-message');
      }
      catch(err){
        email_working = 0;
      }
      if (email_working){
        greeting += sheet.getRange(i+t,2).getValue();
        email += sheet.getRange(i+t,3).getValue().replace(/\s+/g, '').toLowerCase();
        if (t < 3){
          greeting += ', ';
          email += ',';
        }
        about += sheet.getRange(i+t,2).getValue() + ':\n' + sheet.getRange(i+t,4).getValue() + '\n ---------\n';
      }
    }
      
    if (extras > 0){
      if (groups == 1){
        for (var j = 0; j < extras; j++){
          greeting += ', ' + extras_name[j];
          about += extras_name[j] + ': \n' + extras_msg[j]  + '\n ---------\n';
          email += ','+extras_email[j];
        }        
      }
      else if(groups == 2){
        if (iter == 0) {
          greeting += ', ' + extras_name[iter];
          about += extras_name[iter] + ': \n' + extras_msg[iter]  + '\n ---------\n';
          email += ','+extras_email[iter];
        }
        else if (iter == 1){
          for (var j = 1; j < extras; j++){
            greeting += ', ' + extras_name[j];
            about += extras_name[j] + ': \n' + extras_msg[j]  + '\n ---------\n';
            email += ','+extras_email[j];
          }
        }
      }
      else if (groups > 2 && iter < extras){
        greeting += ', ' + extras_name[iter];
        about += extras_name[iter] + ': \n' + extras_msg[iter]  + '\n ---------\n';
        email += ','+extras_email[iter];
      }
    }
    
    greeting += '!\n';
    
    //What the message should contain
    const feedback = 'Send meg en melding om dere får denne!\n';
    const subject = 'ELSK 4.1!';
    const intro = 'Så kult at dere ville ha en ElektroniskLunsjSamtaleKamerat! \n';
    const contact = 'Ta kontakt med kameratene dine ved å trykke "Svar Alle" på denne eposten. \n';
    
    //Constructs the message
    const message = [feedback,greeting, intro, contact, about].join('\n');
    Logger.log(message);
    MailApp.sendEmail(email, subject, message);
      
  } 
}
    
//Triggers daily between 8 and 9 am and does the job                 
function triggerDaily() {
  ScriptApp.newTrigger('sendEmail')
      .timeBased()
      .everyDays(1)
      .atHour(8)
      .create();
}
