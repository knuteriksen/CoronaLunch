function sendEmail() {
  
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
    extras_name.push(data[rows-j][0]);
    extras_age.push(data[rows-j][1]);
    extras_email.push(data[rows-j][2]);
    extras_msg.push(data[rows-j][3]);
  }
  var iter = 0;
  var end = rows -groupSize;
 
  for (var i = 0; i <= end; i = i + groupSize){
   
    const name1 = data[i][0];
    const name2 = data[i+1][0];
    const name3 = data[i+2][0];
    const name4 = data[i+3][0];
    
    const age1 = data[i][1];
    const age2 = data[i+1][1];
    const age3 = data[i+2][1];
    const age4 = data[i+3][1];
    
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
    
    if (extras > 0){
      if (groups == 1){
        for (var j = 0; j < extras; j++){
          greeting += ', ' + extras_name[j];
          about += extras_name[j] + ': \n' + extras_msg[j]  + '\n ---------\n';
        }        
      }
      else if(groups == 2){
        if (iter == 0) {
          greeting += ', ' + extras_name[iter];
          about += extras_name[iter] + ': \n' + extras_msg[iter]  + '\n ---------\n';
        }
        else if (iter == 1){
          for (var j = 1; j < extras; j++){
            greeting += ', ' + extras_name[j];
            about += extras_name[j] + ': \n' + extras_msg[j]  + '\n ---------\n';
          }
        }
      }
      else if (groups > 2 && iter < extras){
        greeting += ', ' + extras_name[iter];
        about += extras_name[iter] + ': \n' + extras_msg[iter]  + '\n ---------\n';
      }
    }
    
    greeting += '!\n';
    
    //What the message should contain
    const subject = 'ELSK 3.0!';
    const intro = 'Så kult at dere ville ha en ElektroniskLunsjSamtaleKamerat! \n';
    const contact = 'Ta kontakt med kameraten dine ved å trykke "Svar Alle" på denne eposten. \n';
    const email = 'knutvagneseriksen@gmail.com,knut@organisasjonskollegiet.no';
    
    //Constructs the message
    const message = [greeting, intro, contact, about].join('\n');
    Logger.log('Message');
    Logger.log(message);
        
    MailApp.sendEmail(email, subject, message);
    
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
