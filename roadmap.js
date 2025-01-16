// Basically the methos used in ff-efd53b.js, reworked for node.js

var XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;
var roadmapData;

// load roadmap json data
function loadRoadmapFeatureData() {
    try {
        var apiEndpoint = "https://www.microsoft.com/msonecloudapi/roadmap/features/";
        //var apiEndpoint = "/en-us/microsoft-365/roadmap/assets/test/mocks/features.json";
        const request = new XMLHttpRequest();
        request.open("GET", apiEndpoint, false);
        request.send(null);
        if (request.status === 200) {
            roadmapData = JSON.parse(request.responseText);
        }

        receivedRoadmapData = true;

    } catch (e) {
        console.log("Unable to get Roadmap data from DOM: " + e);
    }
}

loadRoadmapFeatureData();
console.log(roadmapData[0].tagsContainer.products[0].tagName);