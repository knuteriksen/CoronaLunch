//Brute force search to remove duplicates
function removeDuplicates(){
  var ss = SpreadsheetApp.getActive();
  var sheet1 = ss.getSheetByName('responses');
  var sheet2 = ss.getSheetByName('avmelding');
  const rows1 = sheet1.getLastRow() -1;
  const rows2 = sheet2.getLastRow() -1;
  const cols1 = sheet1.getLastColumn() ;
  const cols2 = sheet2.getLastColumn() ;
  var range1 = sheet1.getRange(2,1,rows1,cols1);
  var range2 = sheet2.getRange(2,1,rows2,cols2);
  var data1 = range1.getValues();
  var data2 = range2.getValues();
   
  for (var i = 0; i < sheet2.getLastRow()-1; i++){
    for (var j = 0; j < sheet1.getLastRow()-1; j++){
      var d1 = sheet2.getRange(i+2,2).getValue().replace(/\s+/g, '').toUpperCase();
      var d2 = sheet1.getRange(j+2,2).getValue().replace(/\s+/g, '').toUpperCase();
      var d3 = sheet2.getRange(i+2,3).getValue().replace(/\s+/g, '').toUpperCase();
      var d4 = sheet1.getRange(j+2,3).getValue().replace(/\s+/g, '').toUpperCase();
      
      if ((d1 == d2) || (d3 == d4)){
          
          var cell1 = sheet1.getRange(j+2, 2);
          var cell2 = sheet1.getRange(j+2, 3);
          var cell3 = sheet1.getRange(j+2, 4);
          
          var value1 = data1[sheet1.getLastRow()-2][1];
          var value2 = data1[sheet1.getLastRow()-2][2];
          var value3 = data1[sheet1.getLastRow()-2][3];
          
          cell1.setValue(value1);          
          cell2.setValue(value2);
          cell3.setValue(value3);
       
          var delRow = sheet1.getLastRow();
          sheet1.deleteRow(sheet1.getLastRow());
          j--;
//          sheet2.deleteRow(i)
       }
    }
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
  
  //Get data in spreadsheet
  var data = range.getValues();
   
  //How large you want the groups
  const groupSize = 4;
 
  //How many extra people there are (3/2/1 or 0)
  const extras = rows % groupSize; 
  
  //How many groups there are
  const groups = (rows-extras)/groupSize;
  Logger.log('Groups');
  Logger.log(groups);

  //Arrays containing the extra people
  var extras_name = [];
  var extras_age = [];
  var extras_email = [];
  var extras_study = [];
  var extras_msg = [];
  
  for (var j = extras; j > 0; j--){
    extras_name.push(data[rows-j][1]);
    extras_email.push(data[rows-j][2]);
    extras_msg.push(data[rows-j][3]);
  }
  var iter = 0;
  var end = rows -groupSize;
 
  for (var i = 0; i <= end; i = i + groupSize){
   
    const name1 = data[i][1];
    const name2 = data[i+1][1];
    const name3 = data[i+2][1];
    const name4 = data[i+3][1];
        
    const email1 = data[i][2];
    const email2 = data[i+1][2];
    const email3 = data[i+2][2];
    const email4 = data[i+3][2];
    
    const msg1 = data[i][3];
    const msg2 = data[i+1][3];
    const msg3 = data[i+2][3];
    const msg4 = data[i+3][3];
   
    var greeting = 'Hei ' + name1 + ', ' + name2 + ', ' + name3 + ', ' + name4;
    var about ='Dette er hva dere har skrevet om dere selv:\n\n' + 
        name1 + ': \n' + msg1  + '\n ---------\n' + 
        name2 + ': \n' + msg2  + '\n ---------\n' +
        name3 + ': \n' + msg3  + '\n ---------\n' + 
        name4 + ': \n' + msg4  + '\n ---------\n' ;
    var email = email1 + ',' + email2 + ',' + email3 + ',' + email4;
    
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
    const subject = 'ELSK 3.0!';
    const intro = 'Så kult at dere ville ha en ElektroniskLunsjSamtaleKamerat! \n';
    const contact = 'Ta kontakt med kameraten dine ved å trykke "Svar Alle" på denne eposten. \n';
  
    
    //Constructs the message
    const message = [greeting, intro, contact, about].join('\n');
    Logger.log('Message');
    Logger.log(message);
    Logger.log(email);
        
    //MailApp.sendEmail(email, subject, message);
    
    iter ++;
    }
}
    
//Triggers daily at 9 am and does the job                 
function triggerDaily() {
  ScriptApp.newTrigger('sendEmail')
      .timeBased()
      .everyDays(1)
      .atHour(14)
      .create();
}
