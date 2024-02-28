function dateDifference() {

  var sheet = SpreadsheetApp.getActiveSheet();

  x_start = 2;   // Row coordinate for first start date
  y_start = 1;   // Column coordinate for first start date
  x_end = 2;     // Row coordinate for first end date
  y_end = 2;     // Column coordinate for first end date
  x_diff = 2;    // Row coordinate for first difference
  y_diff = 3;    // Column coordinate for first difference

  var startDate = sheet.getRange(x_start,y_start).getValue();
  var endDate = sheet.getRange(x_end,y_end).getValue();

  // Calculate number of rows (dates) until blank row
  var j = x_start;

  while(sheet.getRange(j,y_start).isBlank() == false){
    j = j+1;
  }

  // Subtract x_start (initial value) 
  var number_of_rows = j-x_start;

  // Calculate the difference for each interval
  for(var i = 0; i<number_of_rows; i++){

    var startDate = sheet.getRange(x_start + i,y_start).getValue();
    var endDate = sheet.getRange(x_end + i,y_end).getValue();

    var date = new Date(endDate - startDate);
    var diff = Math.floor((date) / (1000*60*60*24));  // Convert to integer (day)

    sheet.getRange(x_diff + i, y_diff).setValue(Math.floor(diff));

  }

  // Fetch the data from the API
  const apiUrl = 'https://jsonplaceholder.typicode.com/comments';
  const response = UrlFetchApp.fetch(apiUrl);
  const jsonResponse = response.getContentText();
  const data = JSON.parse(jsonResponse);
  const emailArray = [];

  // Populate an array with the email addresses
  data.forEach(item => {
    const email = item.email;
    emailArray.push(email);
  });

  // Write a random email address next to each difference
  for(var k = 0; k < number_of_rows; k++){
    var randomIndex = Math.floor(Math.random() * emailArray.length);
    sheet.getRange(x_diff + k, y_diff + 1).setValue(emailArray[randomIndex]);
    emailArray.splice(randomIndex, 1);  // Remove used email addresses
  }
}
