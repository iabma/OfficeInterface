(function () {
    var content = "",
        id = 0;

    Office.onReady(function () {
        $(document).ready(function () {
            if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                // Do something that is only available via the new APIs
                $('#scan').click(readDocumentContent);
                $('#supportedVersion').html('This code is using Word 2016 or later.');
                
            }
            else
                $('#supportedVersion').html('This code requires Word 2016 or later.');
        });
    });

    function readDocumentContent() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed)
                content = asyncResult.error.message;
            else
                content = asyncResult.value;
            scan(content)
        });
    }

    function scan(con) {
        Word.run(function (context) {
            let sugg = document.getElementById("suggestions");
            sugg.innerHTML = "";

            let doc = context.document;
            let words = con.split(" ");

            doc.body.clear();

            for (var i = 0; i < words.length; i++) {
                if (words[i] == "demo") {
                    id++;
                    let _id = id;
                    addSuggestion(words, i, _id, "demonstration");
                    //addAndBindControl(doc, _id.toString(), "demonstration");
                    //sugg.innerHTML += Office.context.document.bindings;
                    /* var myOOXMLRequest = new XMLHttpRequest();
                    var myXML;
                    myOOXMLRequest.open('GET', "insertion.xml", false);
                    myOOXMLRequest.send();
                    if (myOOXMLRequest.status === 200) {
                        myXML = myOOXMLRequest.responseText.replace("!#PTH#!", "demonstration");
                    } */
                    doc.body.insertText("demonstration", "End");//insertOoxml(myXML, "End");
                } else
                    doc.body.insertText(words[i], "End");
                if (i < words.length - 1)
                    doc.body.insertText(" ", "End");
            }

            /*if (sugg.innerHTML == "")
                sugg.innerHTML = "<span>No suggestions.</span>";
            else {
                let lastChild = sugg.children[sugg.children.length - 1];
                lastChild.style.marginBottom = "0px";
            }*/

            return context.sync();
        })
            .catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
    }

    function addSuggestion(arr, i, _id, sugg) {
        let prev = arr[i];
        var previewing = false;
        var suggestion = document.createElement("div");
        suggestion.className = "suggestion";
        suggesiont.id = _id;
        suggestion.innerHTML = "<span>" + prev + " -> " + sugg + "</span>";
        let suggestions = $("#suggestions");
        suggestions.append(suggestion);
        suggestion.onclick = function () {
            replace(arr, i, sugg);
            readDocumentContent();
        }
        suggestion.onmouseenter = function () {
            replace(arr, i, sugg);
            previewing = true;
        }
        suggestion.onmouseleave = function () {
            if (previewing)
                replace(arr, i, prev)
            previewing = false;
        }
    }

    function addAndBindControl(doc, id, initial) {
        /*
        var myOOXMLRequest = new XMLHttpRequest();
        var myXML;
        myOOXMLRequest.open('GET', 'insertion.xml', false);
        myOOXMLRequest.send();
        if (myOOXMLRequest.status === 200) {
            myXML = myOOXMLRequest.responseText.replace("!#PTH#!", content);
        }
        doc.body.insertOoxml(myXML, "End");*/
        Office.context.document.bindings.addFromNamedItemAsync("suggestion", "text", { id: id }, function (result) {
            if (result.status == "failed") {
                if (result.error.message == "The named item does not exist.")
                    var myOOXMLRequest = new XMLHttpRequest();
                var myXML;
                myOOXMLRequest.open('GET', 'content_control.xml', false);
                myOOXMLRequest.send();
                if (myOOXMLRequest.status === 200)
                    myXML = myOOXMLRequest.responseText.replace("!#PTH#!", initial);
                doc.body.insertOoxml(myXML, "End");
                //Office.context.document.setSelectedDataAsync(myXML, { coercionType: 'ooxml' }, function (result) {
                Office.context.document.bindings.addFromNamedItemAsync("suggestion", "text", { id: id });
                //});
                populateBinding("insertion.xml", id, "demonstrat");
            }
        });
    }

    function populateBinding(filename, id, content) {
        var myOOXMLRequest = new XMLHttpRequest();
        var myXML;
        myOOXMLRequest.open('GET', filename, false);
        myOOXMLRequest.send();
        if (myOOXMLRequest.status === 200) {
            myXML = myOOXMLRequest.responseText.replace("!#PTH#!", content);
        }
        Office.select("bindings#" + id).setDataAsync(myXML, { coercionType: 'ooxml' });
    }

/* function writeContent() {
    var myOOXMLRequest = new XMLHttpRequest();
    var myXML;
    myOOXMLRequest.open('GET', 'insertion.xml', false);
    myOOXMLRequest.send();
    if (myOOXMLRequest.status === 200) {
        myXML = myOOXMLRequest.responseText;
    }
    Office.context.document.setSelectedDataAsync(myXML, { coercionType: 'ooxml' });
} */

    function replace(arr, i, toReplace) {
        arr[i] = toReplace;//"!r" + id++;
        let reconstructed = arr.join(" ");
        Office.context.document.setSelectedDataAsync(reconstructed.substring(0, reconstructed.length - 1));
    }
})();