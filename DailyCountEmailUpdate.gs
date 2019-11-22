/**
 * Sends emails with data from the current spreadsheet.
 */
function sendEmails() {
  
  var filteredColumIndex = 5;
  var subject = 'TEST'
  var emailAddress =['test@host.com']
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; // Data - Start index
  var numRecord = sheet.getLastRow() - 1; // All the rows
  var filteredColum = sheet.getRange(1, filteredColumIndex).getValue()
  var dataRange = sheet.getRange(startRow, filteredColumIndex, numRecord);
  var html = ["<h2><ul>Registeration Status based on Location</ul></h2>"];
  
  //---------- TABLE -------------
  html.push(["<table border='1' style='border-collapse:collapse'><thead><tr><td>" + filteredColum + "</td><td>Count</td></tr></thead>"]);
  html.push(["<tbody>"]);
  
  var counter = {};
  
  // Filtering Data  
  var data = dataRange.getValues();
  for (var i = 0 ; i < data.length; i++) {
    var value = data[i][0];
    counter[value] = 1 + (counter[value] || 0);
  }
    
  keys = Object.keys(counter);
  for (var i = 0 ; i < data.length; i++){
    html.push(["<tr><td>" + keys[i]  + "</td><td>" + counter[keys[i]] + "</td></tr>"]);
  }
  
  html.push(["<tr><td>TOTAL</td><td>" + numRecord + "</td></tr>"]);
  
  // Close the table
  html.push(["</tbody>"]);
  html.push(["</table>"]);
  
  //---------- FOOTER -------------
  html.push("<hr>");
  html.push("<p>This is an auto-generated email for regualar updates. For unsubscribing, kindly drop a mail to author@host.com </p>");

  html = html.join('');
  // Send Mail
  MailApp.sendEmail({
    to: emailAddress.join(","),
    subject: subject,
    htmlBody: html
  });
}

function createTrigger() {
  
  // Trigger every 1 minute
  ScriptApp.newTrigger('sendEmails')
      .timeBased()
      .everyMinutes(1)
      .create();
}
