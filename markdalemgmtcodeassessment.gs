// The three columns I have chosen to convert are:
// Column L: Given the best- and worst-case returns of the four investment choices below, which would you prefer?
// Column M: In addition to whatever you own, you have been given $1,000. You are now asked to choose between:
// Column N: In addition to whatever you own, you have been given $2,000. You are now asked to choose between:

function autoFillIPSTemplateGoogleDoc(e) {
  // declare variables from Google Sheet
  let investorName = e.values[1];
  let timeStamp = e.values[0];
  let emailID = e.values[2];

  // declare three columns to be converted to floats from Google Sheet
  let bestWorstReturn = e.values[11];
  let given1000 = e.values[12];
  let given2000 = e.values[13];

  // convert values from column 3 of Google Sheet to string
  const goals = e.values[4].toString();

  // declare goal variables
  let goal1 = ""
  let goal2 = ""
  let goal3 = ""

  //create an array and parse values from CSV format, store them in an array
  goalsArr = goals.split(',')
  if (goalsArr.length >= 1)
    goal1 = goalsArr[0]
  if (goalsArr.length >= 2)
    goal2 = goalsArr[1]
  if (goalsArr.length >= 3)
    goal3 = goalsArr[2]

  // (1) converts bestWorstReturn to array with two floats, best and worst selections by user

  //remove commas to prevent error when converting from string to float
  let bestWorst = bestWorstReturn.replaceAll(",","")

  //regex statement to find all numbers
  let gainLossValues = gainLoss.match( /\d+/g);

  //assuming first found number represents gain, and second number represents loss
  gainLossValues[0] = +gainLossValues[0]
  gainLossValues[1] = -(+gainLossValues[1])

  // (2 & 3) converts given1000 and given2000 from string to float

  //gain or loss float return values
  let gainLoss1000 = 0
  let gainLoss2000 = 0
  
  //remove commas to prevent error when converting from string to float
  let get1000 = given1000.replace(",", "");
  let get2000 = given2000.replace(",", "");
  get1000 = +given1000.split('$').pop();
  get2000 = +given2000.split('$').pop();

  if (given1000.includes("gain")) {
    gainLoss1000 = -get1000
  }
  else {
    gainLoss1000 = get1000
  }

  if (given2000.includes("gain")) {
    gainLoss2000 = -get2000
  }
  else {
    gainLoss2000 = get2000
  }

// By this point the three variables for the converted google sheet columns are: gainLossValues, gainLoss1000, gainLoss2000

//grab the template file ID to modify
  const file = DriveApp.getFileById(templateID);

//grab the Google Drive folder ID to place the modied file into
  var folder = DriveApp.getFolderById(folderID)

//create a copy of the template file to modify, save using the naming conventions below
  var copy = file.makeCopy(investorName + ' Investment Policy', folder);

//modify the Google Drive file
  var doc = DocumentApp.openById(copy.getId());

  var body = doc.getBody();

  body.replaceText('%InvestorName%', investorName);
  body.replaceText('%Date%', timeStamp);

  body.replaceText('%Goal1%', goal1.trim())
  body.replaceText('%Goal2%', goal2.trim())
  body.replaceText('%Goal3%', goal3.trim())

  doc.saveAndClose();

//find the file that was just modified, convert to PDF, attach to e-mail, send e-mail
  var attach = DriveApp.getFileById(copy.getId());
  var pdfattach = attach.getAs(MimeType.PDF);
  MailApp.sendEmail(emailID, subject, emailBody, { attachments: [pdfattach] });
}