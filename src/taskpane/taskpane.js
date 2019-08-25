/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady(info => {
    if (info.host === Office.HostType.Word) {
        readTemplateData();

        document.getElementById("scan").onclick = readDocumentContent;
    }
});
//var { SimilarSearch } = require("node-nlp");
var natural = require("natural");
const url = "https://raw.githubusercontent.com/hms-bcl/Templates/master/templates.json";

var content = "",
    templates = {},
    templateData = {};

console.log("yo")

function readDocumentContent() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function(asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed)
            content = asyncResult.error.message;
        else
            content = asyncResult.value;
        scan(content)
    });
}

async function scan(con) {
    $("#suggestions .suggestion").remove();
    //var sim = new SimilarSearch();
    var best = [],
        i = 0;

    let data = await runJasonsBackend(con);
    data.forEach(datum => {
        addSuggestion(con, datum["start"], datum["end"], datum["suggestion"])
    });
    /* console.log(templates["Hemoglobin"])
    console.log(natural.LevenshteinDistance(templates["Hemoglobin"][1], con, { search: true }))
    Object.keys(templates).forEach(key => {
        let comp = natural.LevenshteinDistance(templates[key][0], con, { search: true });
        //console.log(comp);
        console.log(comp["substring"].length / comp["distance"]);
        if (comp["distance"] == 0 || comp["substring"].length / comp["distance"] > 1.2)
            templates[key].forEach(poss => {
                var data = natural.LevenshteinDistance(poss, con, { search: true });
                if (data["distance"] != 0 && data["substring"].length / data["distance"] > 3)
                    best.push({ substring: data["substring"], distance: data["distance"], suggestion: poss });
            })
    })

    console.log(best);

    if (best.length > 0) {
        best.forEach(sugg => {
            addSuggestion(con, sugg.substring, sugg.suggestion)
        });
    } else {
        suggestions.innerHTML = "No matches."
    } */
    //Office.context.document.setSelectedDataAsync(response["suggestion"])
    /*let words = con.split(" ");
    for (var i = 0; i < words.length; i++) {
        if (words[i] == "demo")
            addSuggestion(words, i, "demonstration");
    }*/
}

function addSuggestion(con, start, end, sugg) { //con, prev, sugg) {
    let prev = con.substring(start, end);
    var previewing = false;
    var suggestion = document.createElement("div");
    suggestion.className = "suggestion";
    suggestion.innerHTML = "<span>" + prev + " -> " + sugg + "</span>";
    let suggestions = $("#suggestions");
    suggestions.append(suggestion);
    suggestion.onclick = function() {
        replace(con, prev, sugg);
        readDocumentContent();
    }
    suggestion.onmouseenter = function() {
        replace(con, prev, sugg);
        previewing = true;
    }
    suggestion.onmouseleave = function() {
        if (previewing)
            replace(con, prev, prev)
        previewing = false;
    }
}

function replace(con, prev, toReplace) {
    con = con.substring(0, con.indexOf(prev)) + toReplace + con.substring(con.indexOf(prev) + prev.length);
    Office.context.document.setSelectedDataAsync(con);
}

function iterate(arr, obj, str) {
    if (str == null)
        str = "";

    //console.log(obj)
    obj.content.forEach(element => {
        if (element.type == "text") {
            if (obj.type == "dropdown" && !arr[str + element.content])
                arr.push(str + element.content);
            else
                str += element.content;
        } else {
            iterate(arr, element, str);
        }
    });

    if (!arr[str]) arr.push(str);
}

function findEveryPossibility() {
    templateData.forEach(template => {
        console.log(template)
        var possibilites = [];
        iterate(possibilites, template.English);
        templates[template.name] = possibilites;
    });
}

function readTemplateData() {
    let req = new XMLHttpRequest();
    req.open("GET", url, true);
    req.send(null);
    req.onreadystatechange = () => {
        if (req.readyState == 4) {
            templateData = JSON.parse(req.responseText);
            console.log("data found. beginning fep...");
            findEveryPossibility();
            console.log("fep complete.");
        }
    };
}

async function runJasonsBackend(data) {
    console.log(data);
    let returnData;
    await fetch("https://guncolony.com/template?querystring=" + data)
        .then(res => {
            //console.log(res.json());
            returnData = res.json()
        }) //)
        .catch(err => console.error(err))
    console.log(returnData)
    return returnData;
}