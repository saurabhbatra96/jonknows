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
              if ((resultvalue.substring(resultvalue.indexOf(".")+1)).indexOf(".")>=0)
                parasenti(resultvalue);
              else
                sentencesenti(resultvalue);
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

  function parasenti(paragraph) {

  	var sentences = paragraph.split(".");
  	var sentis = new Array();
    var i;

  	for(i=0;i<sentences.length;i++) {
  		if(sentences[i].length == 0) {
  			sentences.splice(i,1);
  		}
  	}

  	for(i = 0; i<sentences.length; i++) { // sentiment analysis of each sentence of paragraph
  		var xhrObj = new XMLHttpRequest();
  		var url = "https://westus.api.cognitive.microsoft.com/text/analytics/v2.0/sentiment";
  		var data = '{ "documents": [ { "language":"en", "id":"1", "text":"' + sentences[i] + '" } ] }';
  		xhrObj.open("POST", url, false);
  		xhrObj.setRequestHeader("Content-Type","application/json");
  		xhrObj.setRequestHeader("Ocp-Apim-Subscription-Key","7514572919c74af9a97d235f9a9355d4");
  		xhrObj.send(data);

  		var JSONobj = JSON.parse(xhrObj.responseText);
  		sentis.push(JSONobj.documents[0].score);
  	}

  	// console.log(sentis);

  	var min = sentis[0]; var min_pos = 0;
  	for(i=0;i<sentis.length;i++) {
  		if(sentis[i] < min) {
  			min = sentis[i];
  			min_pos = i;
  		}
  	}

    sentences[min_pos] = sentences[min_pos].trim();
    document.getElementById('third-heading').innerHTML = "Sentiment Analysis";
    document.getElementById('parasenti-quote').innerHTML = "\"" + sentences[min_pos] + ".\"";
    document.getElementById('parasenti-darkside').innerHTML = "Looks like the dark side is strong in this line. Select it for better suggestions."
  }

  function sentencesenti(sentence) {
  	var words = sentence.split(" ");
  	var i,j;
  	var sentis = new Array();

  	for(i = 0; i<words.length; i++) { // sentiment analysis of each word of sentence
  		var xhrObj = new XMLHttpRequest();
  		var url = "https://westus.api.cognitive.microsoft.com/text/analytics/v2.0/sentiment";
  		var data = '{ "documents": [ { "language":"en", "id":"1", "text":"' + words[i] + '" } ] }';
  		xhrObj.open("POST", url, false);
  		xhrObj.setRequestHeader("Content-Type","application/json");
  		xhrObj.setRequestHeader("Ocp-Apim-Subscription-Key","7514572919c74af9a97d235f9a9355d4");
  		xhrObj.send(data);

  		var JSONobj = JSON.parse(xhrObj.responseText);
  		sentis.push(JSONobj.documents[0].score);
  	}

  	// console.log(sentis);

  	// getting synonyms of lowest sentiment word

  	var min = 1.0, min_pos = 0;
  	for(i=0; i<sentis.length; i++) {
  		if(sentis[i]<min) {
  			min = sentis[i];
  			min_pos = i;
  		}
  	}

  	var xhrObj = new XMLHttpRequest();
  	var url = "http://thesaurus.altervista.org/thesaurus/v1?word=" + words[min_pos] + "&language=en_US&output=json&key=X6cKt6oU9RJjmYsBQqRw";
  	xhrObj.open("GET", url, false);
  	xhrObj.send(null);

  	var syns = JSON.parse(xhrObj.responseText).response;

  	var synonyms = [];

  	for(i=0; i<syns.length; i++) {
  		var s = (syns[i].list.synonyms).split("|");
  		// console.log(s);
  		for(j=0;j<s.length;j++) {
  			synonyms.push(s[j]);
  		}
  	}

  	for(i=0;i<synonyms.length;i++) {
  		if(synonyms[i].split(" ").length > 1) {
  			synonyms.splice(i,1);
  		}
  	}

  	// console.log(synonyms);

  	// replacing the word with all synonyms got and calling the joint dist API

  	var lang_queries = [];
  	for(i=0;i<synonyms.length;i++) {
  		var str = "";
  		for(j=0;j<words.length;j++) {
  			if(j == min_pos)
  				str = str+synonyms[i]+" ";
  			else
  				str = str+words[j]+" ";
  		}
  		lang_queries.push(str);
  	}

  	xhrObj = new XMLHttpRequest();
  	var url = "https://api.projectoxford.ai/text/weblm/v1.0/calculateJointProbability?model=body";
  	var data = '{ "queries": [';
  	for(i=0;i<lang_queries.length;i++) {
  		data = data + '"' + lang_queries[i]+ '", ';
  	}
  	data = data + '] }';
  	xhrObj.open("POST", url, false);
  	xhrObj.setRequestHeader("Content-Type","application/json");
  	xhrObj.setRequestHeader("Ocp-Apim-Subscription-Key","96d1620188ce4d4c8d60074af57689c9");
  	xhrObj.send(data);

    // console.log(xhrObj.responseText);
    var xmlDoc = jQuery.parseXML(xhrObj.responseText);
    console.log(xmlDoc);
    var results = xmlDoc.getElementsByTagName("JointProbabilityResult");
    var suggestions = new Array();
    var noofsugg = 2;

    for (j=0;j<noofsugg;++j) {
      min = results[0].childNodes[0].childNodes[0].nodeValue;
      min_pos = 0;
      for (i=0;i<results.length;++i) {
        if (results[i].childNodes[0].childNodes[0].nodeValue < min) {
          min = results[i].childNodes[0].childNodes[0].nodeValue;
          min_pos = i;
        }
      }
      suggestions.push(results[min_pos].childNodes[1].childNodes[0].nodeValue);
      results[min_pos].childNodes[0].childNodes[0].nodeValue = 9999;
    }

    console.log(suggestions);
    document.getElementById('third-heading').innerHTML = "Sentiment Analysis";
    document.getElementById('sentsenti-1').innerHTML = "\"" + suggestions[0] + "\"";
    document.getElementById('sentsenti-2').innerHTML = "\"" + suggestions[1] + "\"";
    document.getElementById('sentsenti-sugg').innerHTML = "You might want to change a couple of words in there.";
    document.getElementById('sentor').innerHTML = "OR";
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

    //parasenti
    document.getElementById('parasenti-quote').innerHTML = "";
    document.getElementById('parasenti-darkside').innerHTML = "";
    document.getElementById('third-heading').innerHTML = "";

    //sentsenti
    document.getElementById('third-heading').innerHTML = "";
    document.getElementById('sentsenti-1').innerHTML = "";
    document.getElementById('sentsenti-2').innerHTML = "";
    document.getElementById('sentsenti-sugg').innerHTML = "";
    document.getElementById('sentor').innerHTML = "";
  }

})();
