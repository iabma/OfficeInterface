(function () {
    var content = "";
    console.log("yo")

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
        $("#suggestions .suggestion").remove();
        let words = con.split(" ");
        for (var i = 0; i < words.length; i++) {
            if (words[i] == "demo")
                addSuggestion(words, i, "demonstration");
        }
    }

    function addSuggestion(arr, i, sugg) {
        let prev = arr[i];
        var previewing = false;
        var suggestion = document.createElement("div");
        suggestion.className = "suggestion";
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

    function replace(arr, i, toReplace) {
        arr[i] = toReplace;
        let reconstructed = arr.join(" ");
        Office.context.document.setSelectedDataAsync(reconstructed.substring(0, reconstructed.length - 1));
    }
})();