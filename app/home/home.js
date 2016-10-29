oldselection = "";
(function(){
  'use strict';

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function(reason){
    jQuery(document).ready(function(){
      app.initialize();

      // jQuery('#get-data-from-selection').click(getDataFromSelection);
      setInterval(getDataFromSelection, 2000);
    });
  };

  // Reads data from current document selection and displays a notification
  function getDataFromSelection() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
      function(result){
        // We don't want to make API calls if the value selected is not changing.
        if (oldselection != result.value) {
          oldselection = result.value;
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log('The selected text is:', '"' + result.value + '"');
            getKeywords(result.value);
          } else {
            console.log('Error:', result.error.message);
          }
        }
      }
    );
  }


  // REST API calls go here.
  function getKeywords(sentence){
    var xhrObj = new XMLHttpRequest();
    var url = "https://westus.api.cognitive.microsoft.com/text/analytics/v2.0/keyPhrases?";

    var data = '{ "documents": [ { "language":"en", "id":"1", "text":"' + sentence + '" } ] }';

    xhrObj.onreadystatechange = function() {
      if (xhrObj.readyState == 4) {
        var responsedata = JSON.parse(xhrObj.response);
        responsedata = responsedata.documents[0].keyPhrases;
        console.log(responsedata);
      }
    }

    xhrObj.open("POST", url, true);
    xhrObj.setRequestHeader("Content-Type","application/json");
    xhrObj.setRequestHeader("Ocp-Apim-Subscription-Key","7514572919c74af9a97d235f9a9355d4");
    xhrObj.send(data);

  }
})();
