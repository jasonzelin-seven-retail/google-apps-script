function myFunction() {
  const webhookUrl = "https://chat.googleapis.com/v1/spaces/AAQAPf3YqnE/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=nTG-sd4RgOLouq3-mzi4dkos6ldbm4k58WAkMPkQL9E"
  var message = {
    "text": "List of non-uniform headers\n- BGR: Trial Date\n- BGR: Trial Schedule"
  };

  // Options for the HTTP POST request
  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(message) 
  };

  // Send the message using the UrlFetch service
  UrlFetchApp.fetch(webhookUrl, options);
}