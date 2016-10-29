oldselection = "";
state = "empty";

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
          var resultvalue = result.value.trim();

          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log('The selected text is:', '"' + resultvalue + '"');
            if (resultvalue.indexOf(" ")>=0) {
              state = "sentence";
              getKeywords(resultvalue);
            } else if (resultvalue=="") {
              state = "empty";
              cleanScreen();
            } else {
              state = "word";
              getSynonyms(resultvalue);
            }
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
        var searchstring = "";
        for (var i=0;responsedata[i];++i) {
          searchstring = searchstring+" "+responsedata[i];
        }

        bingsearch(searchstring.trim());
        bingimgsearch(searchstring.trim());
      }
    }

    xhrObj.open("POST", url, true);
    xhrObj.setRequestHeader("Content-Type","application/json");
    xhrObj.setRequestHeader("Ocp-Apim-Subscription-Key","7514572919c74af9a97d235f9a9355d4");
    xhrObj.send(data);

  }

  function getSynonyms(word){

    // var XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;
    var xhrObj = new XMLHttpRequest();
    var url = "http://thesaurus.altervista.org/thesaurus/v1?word=" + word + "&language=en_US&output=json&key=X6cKt6oU9RJjmYsBQqRw";

    xhrObj.onreadystatechange = function() {
        // console.log('readyState '+xhrObj.readyState);
        if (xhrObj.readyState == 4) {
          console.log(xhrObj);
          var synonymdata = JSON.parse(xhrObj.responseText);
          synonymdata = synonymdata.response;
          var isnoun = false;

          // Get first three synonyms.
          for(var i=0;i<3;++i) {
            if (synonymdata && synonymdata[i]) {
              if (synonymdata[i].list.category.indexOf("(noun)")>=0) {
                isnoun = true;
              }
              var syn = synonymdata[i].list.category + " " + synonymdata[i].list.synonyms;
            } else {
              var syn = "";
            }

            if (i==0)
              document.getElementById("first-h-text1").innerHTML = syn;
            else if (i==1)
              document.getElementById("first-h-text2").innerHTML = syn;
            else
              document.getElementById("first-h-text3").innerHTML = syn;
          }
        }

        if (synonymdata[0]) {
          document.getElementById("first-h-subh1").innerHTML = "Synonyms";
          document.getElementById("first-heading").innerHTML = "Information";
        }

        if (isnoun) {
          bingsearch(word);
          bingimgsearch(word);
        }
    }

    xhrObj.open("GET", url, true);
    xhrObj.send(null);
  }

  function bingsearch(phrase){
  	var xhrObj = new XMLHttpRequest();
  	var url = "https://api.cognitive.microsoft.com/bing/v5.0/search?q="+phrase+"&count=10&offset=0&mkt=en-us&safesearch=Moderate";

  	xhrObj.onreadystatechange = function() {
  		// console.log('readyState '+xhrObj.readyState);
  	    if (xhrObj.readyState == 4) {
          var searchresponse = JSON.parse(xhrObj.responseText);
          var viewmore = searchresponse.webPages.webSearchUrl;
          searchresponse = searchresponse.webPages.value[0];

          document.getElementById("second-h-text1").innerHTML = searchresponse.name;
          document.getElementById("second-h-text1").href = searchresponse.url;
          document.getElementById("second-h-text2").innerHTML = searchresponse.snippet;
          document.getElementById("second-h-text3").href = viewmore;
          document.getElementById("second-h-text3").innerHTML = "View more on Bing! ...";
          document.getElementById("second-heading").innerHTML = "Bing!";
          document.getElementById("second-h-subh1").innerHTML = "Web";
  	      console.log(searchresponse);
  	    }
  	}

  	xhrObj.open("GET", url, true);
  	xhrObj.setRequestHeader("Ocp-Apim-Subscription-Key","3afb767cdfd3453aaff41d92fb571965");
  	xhrObj.send(null);
  }

  function bingimgsearch(phrase){
  	var xhrObj = new XMLHttpRequest();
  	var url = "https://api.cognitive.microsoft.com/bing/v5.0/images/search?q="+phrase+"&count=10&offset=0&mkt=en-us&safeSearch=Moderate";

  	xhrObj.onreadystatechange = function() {
  		// console.log('readyState '+xhrObj.readyState);
  	    if (xhrObj.readyState == 4) {
          var imgresponse = JSON.parse(xhrObj.responseText);
          var webSearchUrl = imgresponse.webSearchUrl;

          document.getElementById('second-h-subh2').innerHTML = "Images";
          document.getElementById('img-href-1').href = imgresponse.value[0].hostPageUrl;
          document.getElementById('img-1').src = imgresponse.value[0].contentUrl;
          document.getElementById('img-href-2').href = imgresponse.value[1].hostPageUrl;
          document.getElementById('img-2').src = imgresponse.value[1].contentUrl;
          document.getElementById('img-href-3').href = imgresponse.value[2].hostPageUrl;
          document.getElementById('img-3').src = imgresponse.value[2].contentUrl;
          document.getElementById('img-href-4').href = imgresponse.value[3].hostPageUrl;
          document.getElementById('img-4').src = imgresponse.value[3].contentUrl;
          document.getElementById('second-h-text4').href = webSearchUrl;
          document.getElementById('second-h-text4').innerHTML = "View more on Bing! ...";
  	      console.log(xhrObj.responseText);
  	    }
  	}

  	xhrObj.open("GET", url, true);
  	xhrObj.setRequestHeader("Ocp-Apim-Subscription-Key","3afb767cdfd3453aaff41d92fb571965");
  	xhrObj.send(null);
  }

  function cleanScreen() {
    document.getElementById("first-heading").innerHTML = "";
    document.getElementById("first-h-text1").innerHTML = "";
    document.getElementById("first-h-text2").innerHTML = "";
    document.getElementById("first-h-text3").innerHTML = "";
    document.getElementById("second-h-text1").innerHTML = "";
    document.getElementById("second-h-text2").innerHTML = "";
    document.getElementById("second-h-text3").innerHTML = "";
    document.getElementById("first-h-subh1").innerHTML = "";
    document.getElementById("second-heading").innerHTML = "";
    document.getElementById("second-h-subh1").innerHTML = "";
    document.getElementById("second-h-text4").innerHTML = "";

    //images
    document.getElementById('second-h-subh2').innerHTML = "";
    document.getElementById('img-1').src = "";
    document.getElementById('img-2').src = "";
    document.getElementById('img-3').src = "";
    document.getElementById('img-4').src = "";
  }
})();
