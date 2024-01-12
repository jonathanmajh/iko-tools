<!doctype html>
<html lang="en">
<head>
    <!-- Required meta tags -->
    <!-- Remember to refer to SOP before working on file -->
    <!-- P:\MRO Items\CORPORATE RELIABILITY ENGINEERING\Co-op Files\CO-OP TOOLS AND PROGRAMS\Asset Redirect - DNA Enhanced Page\AssetRedirectSOP.docx -->
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <meta http-equiv="X-UA-Compatible" content="IE=Edge">

    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css" />
    <link rel="stylesheet"
        href="https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.5.3/css/bootstrap.min.css" />
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/select2/4.0.13/css/select2.min.css" />
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.5.1/jquery.slim.min.js"
        type="text/javascript"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.5.3/js/bootstrap.bundle.min.js"
        type="text/javascript"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.3.0/papaparse.min.js"
        type="text/javascript"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/select2/4.0.13/js/select2.min.js"
        type="text/javascript"></script>

    <title>Recommendations</title>

</head>

<body>
    <nav class="navbar navbar-expand-md navbar-light bg-white navbar-border">
        <a class="navbar-brand" href="http://connect.na.local/Pages/Connect.aspx"><img
                src="/support/Reliability/ReliabilityShared/Pages/IKO_Logo-nobg.png" alt="IKO" height=50px /></a>
        <!--a class="navbar-brand" href="http://connect.na.local/Pages/Connect.aspx"><img
                src="IKO_Logo-nobg.png" alt="IKO" height=50px /></a-->
        <button class="navbar-toggler" type="button" data-toggle="collapse" data-target="#navbarsExample03"
            aria-controls="navbarsExample03" aria-expanded="false" aria-label="Toggle navigation">
            <span class="navbar-toggler-icon"></span>
        </button>

        <div class="collapse navbar-collapse" id="navbarsExample03">
            <ul class="navbar-nav mr-auto">
                <li class="nav-item nav-iko active">
                    <a class="nav-link"
                        href="http://operations.connect.na.local/support/Reliability/SitePages/Home.aspx">Reliability
                        SharePoint Homepage</a>
                </li>
                <li class="nav-item nav-iko active">
                    <a class="nav-link"
                        href="http://operations.connect.na.local/support/Reliability/ReliabilityShared/Pages/DNASitemap.aspx?mobile=0">DNA
                        Map</a>
                </li>
                <li class="nav-item nav-iko active">
                    <a class="nav-link"
                        href="http://operations.connect.na.local/support/Reliability/ReliabilityShared/Pages/AssetFinder2.aspx?mobile=0">DNA
                        Finder</a>
                </li>
                <li class="nav-item nav-iko active">
                    <a class="nav-link" href="http://operations.connect.na.local/support/Reliability/ReliabilityShared/Pages/AssetRedirect.aspx?mobile=0">Asset Redirect</a>
                </li>
                <li class="nav-item nav-iko active">
                    <a class="nav-link"
                        href="http://operations.connect.na.local/support/Reliability/ReliabilityPublished/TrainingMaterial/AssetVideos/Forms/AllItems.aspx">Technical
                        Training Videos</a>
                </li>
                <li class="nav-item nav-iko active">
                    </liclass>
                    <a class="nav-link"
                        href="http://operations.connect.na.local/support/Reliability/ReliabilityShared/Pages/AssetRedirectSOP.pdf">Help
                        (SOP)</a>
                </li>
            </ul>
            <ul class="navbar-nav  justify-content-start">
                <li class="nav-item nav-iko">
                    <a class="nav-link"
                        href="http://operations.connect.na.local/support/Reliability/ReliabilityShared/Pages/AssetRedirectTesting.html?cheese=GH=&cheeseNum=C0641=&mobile=0"><i
                            class="fa fa-flask"></i></a>
                </li>
            </ul>
        </div>
    </nav>
    <div class="container-fluid">
        <div class="row">
            <div class="table-responsive  text-nowrap">
                <tbody>
                    <tr>
                        <td>
                            <table class="table" id="information">
                                <tbody >
                                    <tr>
                                        <td rowspan="2" style="vertical-align: middle;">
                                            <h1 id="recommendations-title-main"> <strong>Recommendations</strong>
                                            </h1>
                                            <h4 class="title-theme" id="recommendations-title-sub"> from Corporate FMECA Analysis
                                            </h4>
                                        </td>
                                        <td>
                                            <div class="row">
                                                <h6>Site ID:</h6>
                                            </div>
                                            <div class="row">
                                                <select id="siteName" onchange="ChangeSite()">
						    <option value="ANT">ANT: Antwerp</option>
                                                    <option value="RAM">RAM: Alconbury</option>
                                                    <option value="CAM">CAM: Appley Bridge</option>
                                                    <option value="GE">GE: Ashcroft</option>
                                                    <option value="GR">GR: BramCal</option>
                                                    <option value="BA">BA: Calgary</option>
                                                    <option value="COM">COM: Combronde</option>
						    <option value="GP">GP: CRC Brampton</option>
                                                    <option value="GJ">GJ: CRC Toronto</option>
                                                    <option value="BL">BL: Hagerstown </option>
                                                    <option selected value="GH">GH: Hawkesbury</option>
                                                    <option value="GV">GV: Hillsboro (Southwest)</option>
                                                    <option value="GK">GK: IG Brampton</option>
						    <option value="GM">GM: IG High River</option>
                                                    <option value="AA">AA: IKO Brampton</option>
                                                    <option value="CA">CA: Kankakee</option>
                                                    <option value="GI">GI: Madoc </option>
						    <option value="GX">GX: Maxi-Mix </option>
                                                    <option value="PBM">PBM: Senica/Sloviakia</option>
                                                    <option value="GC">GC: Sumas</option>
                                                    <option value="GS">GS: Sylacauga</option>
                                                    <option value="GA">GA: Wilmington</option>
                                                    <option value="KLU">KLU: Klundert</option>
                                                </select>
                                            </div>
                                            <div class="row" style="padding-top: 15px;">
                                                <h6>Failure Class Description:</h6>
                                            </div>
                                            <div class="row text-wrap width-250">
                                                <p id="failure-description">Please enter asset number</p>
                                            </div>
                                        </td>
                                        <td>
                                            <div class="row">
                                                <h6>Asset Number:</h6>
                                            </div>
                                            <div class="row">
                                                <select id="assetNum" class="select2-asset-id"></select>
                                            </div>
                                            <div class="row" style="padding-top: 10px;">
                                                <h6>Asset Description:</h6>
                                            </div>
                                            <div class="row text-wrap">
                                                <p id="description">Please enter asset number</p>
                                            </div>
                                        </td>
                                        <td style="padding-top: 40px;" >
                                            <button id="submitNum" onclick="checkData()">Generate</button>
                                        </td>
                                        <td style="position:fixed; min-width: 295px;">
                                            <video id="dna-render" autoplay muted loop playsinline 
                                                poster="/support/Reliability/ReliabilityShared/Pages/RotatingDNA_Cover.jpg">
                                                <source
                                                    src="/support/Reliability/ReliabilityShared/Pages/RotatingDNA.mp4?mobile=0"
                                                    type="video/mp4" />
                                            </video>
                                            <!--video id="dna-render" autoplay muted loop playsinline width='110' height="110" controls
                                                poster = RotatingDNA_Cover.jpg>
                                                <source
                                                    src="RotatingDNA.mp4"
                                                    type="video/mp4" />
                                            </video-->

                                        </td>
                                    </tr>
                                </tbody>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <table class="table">
                                <thead>
                                    <h3 id="table-heading">
                                    </h3>
                                </thead>
                                <tbody id='dynamic-table-body'>
                                    <tr id="placeholder">
                                        <td>
                                            <h4><i class="fa fa-ban"></i> Please refresh the page if you still see this
                                                message after 5 seconds</h4>
                                        </td>
                                    </tr>
                                </tbody>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <h4 style="padding-top: 30px">Navigation links</h4>
                            <table class="table table-fit">
                                <tbody id='navlinks-table-body'>
                                    <tr>
                                        <td><a class="btn btn-primary btn-light-gray"
                                                href="http://operations.connect.na.local/support/Reliability/ReliabilityShared/Pages/AssetRedirectSOP.pdf">Help</a>
                                        </td>
                                        <td>
                                            <h4><i class="fa fa-question-circle"></i> Help (SOP)</h4>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td><a class="btn btn-primary btn-light-gray"
                                                href="http://operations.connect.na.local/support/Reliability/ReliabilityPublished/TrainingMaterial/AssetVideos/Forms/AllItems.aspx">Videos</a>
                                        </td>
                                        <td>
                                            <h4><i class="fa fa-youtube-play"></i> Technical Training Videos (In
                                                Developement)</h4>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td><a class="btn btn-primary btn-light-gray"
                                                href="http://operations.connect.na.local/support/Reliability/SitePages/Home.aspx">Home</a>
                                        </td>
                                        <td>
                                            <h4><i class="fa fa-gears"></i> Reliability Sharepoint Homepage</h4>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td><a class="btn btn-primary btn-light-gray"
                                                href="http://operations.connect.na.local/support/Reliability/ReliabilityShared/Pages/DNASitemap.aspx?mobile=0">Map</a>
                                        </td>
                                        <td>
                                            <h4><i class="fa fa-map"></i> DNA Site Map</h4>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td><a class="btn btn-primary btn-light-gray"
                                                href="http://operations.connect.na.local/support/Reliability/ReliabilityShared/Pages/AssetFinder2.aspx?mobile=0">Finder</a>
                                        </td>
                                        <td>
                                            <h4><i class="fa fa-search"></i> DNA Asset Finder</h4>
                                        </td>
                                    </tr>
                                </tbody>
                            </table>
                        </td>
                    </tr>
                </tbody>
                </table>
            </div>
        </div>
    </div>
</body>

</html>

<script type="application/javascript">
    "use strict";
    console.log("Starting JS");
    var currentAssetListSite = "N/A";
    var dataURL = "/support/Reliability/ReliabilityShared/Pages/";
    var maximoReportUrl = "http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/";
    var sharepointReliabilityUrl = "http://operations.connect.na.local/support/Reliability/";
    var url_query = { siteID: "None", assetID: "None", assetDescription: "Loading Asset Description" };
    var failureClassDataSheet;
    var failureClassNameSheet;
    var failureCode = "N/A";
    var reliabilityAlerts;
    var validAssets
    var attribute;
    var taskByAsset;
    var site_names = {
	"ANT": "ANT:%20Antwerp",
        "GE": "GE:%20Ashcroft",
        "GR": "GR:%20BramCal",
        "AA": "AA:IKO%20Brampton",
        "BA": "BA:%20Calgary",
        "GJ": "GJ:%20Canroof/CRC%20Toronto",
        "GP": "GP:%20CRC%20Insulation",
        "GG": "GG:%20Hamilton",
        "GH": "GH:%20Hawkesbury",
        "GM": "GM:%20IG%20West%20High%20River",
        "GK": "GK:%20IG%20Machine%20%26%20Fibers%20Brampton",
        "GT": "GT:%20Ingersoll",
        "CA": "CA:%20Kankakee",
        "GI": "GI:%20Madoc",
        "GX": "GX:%20Maxi-Mix",
        "PBM": "PBM:%20Slovakia",
        "GS": "GS:%20Southeast%20(Sylacauga)",
        "GV": "GV:%20Southwest%20(Hillsboro)",
        "GC": "GC:%20Sumas",
        "GA": "GA:%20Wilmington",
        "KLU": "KLU:%20Klundert",
        "COM": "COM:%20Combronde",
        "BL": "BL:%20Hagerstown",
    };
    
    var frequency = {
    "1D": "Daily",   
    "1W": "Weekly",
    "1W \nbefore startup": "A week before start-up",
    "2W": "Bi-Weekly",
    "3W": "Every 3 Weeks",
    "4": "Every 4 Weeks",
    "1M": "Monthly",
    "2M": "Bi-Monthly",
    "3M": "Every 3 Months",
    "6M": "Semi-Annually",
    "1Y": "Annually",
    "":" ",
    };


    function checkData() {
        console.log("Running checkData");
        // while using a predefined asset list already makes sure only existing assets can be selected
        // this checking function is left here in case we need to open up custom user inputs (see 'tags' for select2)
        var patt = new RegExp("[A-Z][0-9]{4}"); //check if input is right format
        var data = $('.select2-asset-id').select2('data')[0];
        var userInput = data.id.toUpperCase();
        if (userInput == "TEST VERSION") {
            var paraURL = window.location.href;
            window.open(paraURL.replace(".aspx/", "Testing.html"));
        }
        var result = patt.test(userInput);
        console.log(result)
        if (result == true) {
            assetNum.value = userInput;
            genTasks(data);
        } else {
            return false; //TODO some kinda of error message
        }
    }

    function ChangeSite() {
        console.log("Running ChangeSite");
        var siteID = document.getElementById("siteName").value;
        if (!(currentAssetListSite == siteID)) {
            var url = dataURL.concat(siteID, ".csv");
            readCSV(url, PopulateAssetIds);
        }
        else {
            console.log("Asset List for ".concat(siteID, " is already loaded"));
        }
    }

    function PopulateAssetIds(csvData) {
        console.log("Running PopulateAssetIds");
        // console.log(csvData);
        var j = csvData.length;
        var assetIds = [];
        var assetIdTracker = {};
        for (var i = 0; i < j; i++) {
            if (csvData[i][0] == url_query.assetID) {
                url_query.assetDescription = csvData[i][0];
            }
            if (!(csvData[i][0] in assetIdTracker) && CheckRecommendations(csvData[i][0])) {    //Ensures that only assets that exist and have recommendations are populated
                if (csvData[i][5]) {
                    assetIds.push({ id: csvData[i][0], text: csvData[i][0].concat(": ", csvData[i][1], " : ", csvData[i][5]) });
                } else {
                    assetIds.push({ id: csvData[i][0], text: csvData[i][0].concat(": ", csvData[i][1] )});
                }
                
                assetIdTracker[csvData[i][0]] = "";
            }
        }
        $(".select2-asset-id").empty();
        $(".select2-asset-id").select2({
            // create and set options for the asset id dropdown
            // tags: true, uncomment for custom inputs
            data: assetIds,
            templateSelection: formatDisplay, //uses formatDisplay function to only display id in input field
            dropdownAutoWidth: true, //make the dropdown wider than the input field
        });
        currentAssetListSite = document.getElementById("siteName").value;
        
        if (currentAssetListSite == url_query.siteID) {
            // check if selected site is the same as the one specified in the url
            // this would mean we should use the initial asset id as well
            $(".select2-asset-id").val(url_query.assetID);
            $(".select2-asset-id").trigger('change');
            $(".select2-asset-id").trigger('select2:select');
        }
        checkData();
    }

    function formatDisplay(asset) {
        return asset.id;
    }

    function genTasks(selected_data) {
        console.log("Running genTasks");
        var description = selected_data.text;
        var assetN = selected_data.id;
        var siteID = document.getElementById("siteName").value;
        var tasks = [];
        var relatedTasks = [];
        var taskTypes;

        //get asset description
        document.getElementById("description").textContent = description;
        
        //get failure class
        var j = failureClassDataSheet.length;
        for (var i = 0; i < j - 1; i++) {
            if (assetN == failureClassDataSheet[i][0] && siteID == failureClassDataSheet[i][2]) {
                failureCode = failureClassDataSheet[i][1];
            }
        }
        //get failure class description
        var failureClassDescription = "No Failure Class Found for This Asset";
        j = failureClassNameSheet.length;
        for (i = 0; i < j - 1; i++) {
            if (failureCode == failureClassNameSheet[i][0]) {
                failureClassDescription = failureClassNameSheet[i][1];
            }
        }
        document.getElementById("failure-description").textContent = failureClassDescription;
        
        var n = taskByAsset.length;
        var relatedTasks;

        for (i = 0; i < n; i++){
            if (`${siteID}_${assetN}` == taskByAsset[i][0]) {
                relatedTasks.push(taskByAsset[i][1]);
            }
        }

        //get tasks for assets
        var k = attribute.length;
        var tasks;
        for (j = 0; j < k ; j++){
            //console.log(taskNumAsset[i][1].concat(" ",attribute[j][0]));
            if (relatedTasks.includes(attribute[j][0])){
                console.log("Found Task ID");
                tasks.push(attribute[j]);
            }
        }


        var uniqueTaskType = taskType(tasks,tasks.length); //gets task types for individual asset
        var typeCount = uniqueTaskType.slice(-1);          //gets count of unique tasks types for individual asset

        //sets text and look for location/asset description
        document.getElementById("table-heading").textContent  = site_names[siteID].replace(/%20/g," ").replace(/%26/g," ").split(": ")[1].concat(" - ",description.split(": ")[0]," - ",description.split(": ")[1]);
        document.getElementById("table-heading").setAttribute("class","asset-title-theme");
        document.getElementById("table-heading").setAttribute("style","padding-left: 1%");
       
        var icons = {
            "Design Change": "fa fa-pencil-square-o", 
            "Installation Procedure": "fa fa-wrench",
            "SOC/SOP": "fa fa-tasks",
            "Pre start Checklist": "fa fa-check-square-o",
            "Theme":{
                "Design Change": "design-theme", 
                "Installation Procedure": "install-theme",
                "Pre start Checklist": "prestart-theme",
                "SOC/SOP": "soc-theme",
            },
            "Color":{
                "Design Change": "#d2d3db", 
                "Installation Procedure": "#cce4e1",
                "Pre start Checklist": "#c4c0d4",
                "SOC/SOP": "#bcd0e9",
            } 
        };

        var theme = ["design-theme","install-theme","prestart-theme","soc-theme"];
  
        var parent = document.getElementById("dynamic-table-body");
        parent.textContent = ''; //remove placeholder or old data

        for (i = 0; i < typeCount; i++){

            var countEach = 0;

            var taskTypeRow = document.createElement("tr");
            taskTypeRow.setAttribute("id", "head-dynamic-row-".concat(i));
            taskTypeRow.setAttribute("class",icons["Theme"][uniqueTaskType[i]]);
            //taskTypeRow.setAttribute("class","border-radius-test");
            
            
            var taskTypeCell = document.createElement("td");
            taskTypeCell.setAttribute("id", "cell-".concat(i));
            taskTypeCell.setAttribute("colspan","3");
            taskTypeCell.setAttribute("style","padding-bottom: 15px; padding-left: 40px;");
            //taskTypeCell.setAttribute("class","mw-collapsible");
            
            
            var taskTypeHeading = document.createElement("h3");
            taskTypeHeading.setAttribute("id", "type-heading-".concat(i));
            taskTypeHeading.setAttribute("style","padding-left: 20px;");
            taskTypeHeading.setAttribute("style","padding-top: 15px;");
            
            var taskTypeIcon = document.createElement("i");
            taskTypeIcon.setAttribute("class", icons[uniqueTaskType[i]]);
            taskTypeIcon.setAttribute("style", "padding-right: 20px;");
            taskTypeHeading.appendChild(taskTypeIcon);
            
            var taskTypeText = document.createTextNode(uniqueTaskType[i]);
            taskTypeHeading.appendChild(taskTypeText);
            
            taskTypeCell.appendChild(taskTypeHeading);
            taskTypeRow.appendChild(taskTypeCell);
            parent.appendChild(taskTypeRow);
           
            for (j = 0; j < tasks.length; j++){
                
                if (tasks[j][1] == document.getElementById("type-heading-".concat(i)).textContent){

                    countEach++;

                    var taskTitleRow = document.createElement("tr");
                    taskTitleRow.setAttribute("class", "task-title-theme");

                    var taskTitleCell = document.createElement("td");
                    taskTitleCell.setAttribute("colspan","3");
                    taskTitleCell.setAttribute("style","padding-left: 40px;")

                    
                    
                    var taskTitleHeading = document.createElement("h5");
                    taskTitleHeading.setAttribute("style","padding-top: 10px;");
                    taskTitleHeading.setAttribute("style", "font-size: 20px;");
                    taskTitleHeading.setAttribute("class", "strong");

                    var taskTitleText = document.createTextNode(tasks[j][2]);

                    taskTitleHeading.appendChild(taskTitleText);
                    taskTitleCell.appendChild(taskTitleHeading);
                    taskTitleRow.appendChild(taskTitleCell);
                    parent.appendChild(taskTitleRow);
                    
                    var taskDetailsRow = document.createElement("tr");
                    
                    for (let m = 0; m < 3; m++){
                      
                        var taskDetailsCell = document.createElement("td");
                        taskDetailsCell.setAttribute("style","padding-left: 40px;");
                        taskDetailsCell.setAttribute("id","cell_".concat(countEach,"_",m));
                        

                        if (m == 0){
                            taskDetailsCell.setAttribute("class","break-it");
                        }
                        else{
                            taskDetailsCell.setAttribute("style","text-align: center");
                            taskDetailsCell.setAttribute("class","width-20-percent"); 
                        }
                        
                        if(m == 2){
                            var frequencyAbrv = tasks[j][5];
                            tasks[j][5] = frequency[frequencyAbrv];
                        }
                        
                        var detailsTxt = document.createTextNode(tasks[j][m+3]);
                        taskDetailsCell.appendChild(detailsTxt);

                        taskDetailsRow.appendChild(taskDetailsCell);
                        parent.appendChild(taskDetailsRow);   
                    }
                }
            }

            
            var div = document.createElement("div");
            div.setAttribute("style", "width: 166.7%; border-bottom: 3px solid ".concat(icons["Color"][uniqueTaskType[i]]));
            
            parent.appendChild(div); 
             

            var space = document.createElement("br");
            space.setAttribute("class", "break-it");
            parent.appendChild(space);
            parent.appendChild(space);
            parent.appendChild(space);
            parent.appendChild(space);
           

        }
    }

    // Checks if specific asset has recommendation tasks
    function CheckRecommendations(assetN){
        var siteID = document.getElementById("siteName").value;
        var checkForRecommendations = siteID.concat("_",assetN);
        var recFound = false;
        for (let i = 0; i < validAssets.length; i++){
            if (validAssets[i][0] == checkForRecommendations){
                recFound = true;
                break;
            }
        } 
        return (recFound);
    }

    // Gets task types for specific asset
    function taskType(assetTasks,length){
    // Sort the array
        assetTasks.sort(sortFunction);
        var uniqueTaskType = [];
  
        // Traverse the sorted array
        var count = 0;
        for (let i = 0; i < length; i++){
  
            // Move the index ahead while there are duplicates
            while (i < length - 1 && assetTasks[i][1] == assetTasks[i + 1][1]) {
                i++;
            }
            count++;

            uniqueTaskType.push(assetTasks[i][1]); 
        }

        //var order = {"Pre start Checklist": 1, "SOC/SOP": 2, "Installation Procedure": 3, "Design Change": 4};
        var order = ["Pre start Checklist", "SOC/SOP", "Installation Procedure", "Design Change"];

        uniqueTaskType.sort(function(a,b){      // Sorts task type by desired order
            if ( a == b ) return a - b;
            return order.indexOf(a) - order.indexOf( b);
        });

        console.log(uniqueTaskType);


        uniqueTaskType.push(count); //Adds count of unique task types to array
        
        
        return uniqueTaskType;
    }

    // Alphabetically sorts array
    function sortFunction(a, b) {
        if (a[1] === b[1]) {
            return 0;
        }
        else {
            return (a[1] < b[1]) ? -1 : 1;
        }
    }


    function readCSV(url, callback) { //reads csv into an array using papaparse libary
        console.log("Running readCSV: " + url);
        Papa.parse(url, {
            download: true,
            complete: function (results) {
                console.log("Finished Parsing Array: " + url);
                callback(results.data.slice(0));
            },
            error: function (err, file, inputElem, reason) {
                console.log("Error parsing Array: " + url);
                console.log(err);
                console.log(reason);
            }
        });
    }

    function startgenTasks(data) {
        failureClassDataSheet = data;
        delayStart();
    }

    function getFailureCodeNames(data) {
        failureClassNameSheet = data;
        delayStart();
    }

    function getReliabilityAlerts(data) {
        reliabilityAlerts = data;
        delayStart();
    }

    function getRecommendations(data){
        validAssets = data;
        delayStart();
    }

    function getTaskAttribute(data){
        attribute = data;
        delayStart();
    }

    function getTaskByAsset(data){
        taskByAsset = data;
        delayStart();
    }

    function delayStart() {
        if (validAssets && attribute && taskByAsset ) {
            console.log("FC and RA loaded");
            ChangeSite();
        } else {
            console.log("FC & RA not loaded")
        }
    }

    document.addEventListener('DOMContentLoaded', function () { //set to siteID equal to the variable in the URL ...?cheese=GK case insensitive
        //this runs as soon as the site has been loaded
        console.log("Running Script after loading");
        var paraURL = window.location.search;

        var url = dataURL.concat("FailureClassesDescriptions.csv");
        readCSV(url, getFailureCodeNames);

        url = dataURL.concat("FailureClasses.csv");
        readCSV(url, startgenTasks);

        url = dataURL.concat("TaskAssets.csv");
        readCSV(url, getRecommendations);

        url = dataURL.concat("TaskAttribute.csv");
        readCSV(url, getTaskAttribute);
        
        url = dataURL.concat("TaskByAsset.csv");
        readCSV(url, getTaskByAsset);

        if (paraURL.indexOf("cheese") != -1) {
            var cheese = paraURL.split('=')[1]; //get siteID
            siteName.value = cheese.toUpperCase();
            url_query.siteID = cheese.toUpperCase();
        }
        if (paraURL.indexOf("cheeseNum") != -1) {
            var cheeseNum = paraURL.split('=')[3]; //get assetNum
            url_query.assetID = cheeseNum.toUpperCase();
        }
    }, false);


</script>


<style>
    select {
        width: 200px;
    }

    .asset-title-theme {
        border-bottom: 3.5mm ridge #607d8b;
    }

    .btn-light-gray {
        background-color: #f1f1f1;
        color: black;
    }

    .design-theme {
        border: none;
        border-collapse: separate;
        overflow: hidden;
        perspective: 5px;
        display: flex;
        width: 166.7%;
        border-radius: 15px 15px 0px 0px;
        background-color: #d2d3db;
        color: #34364b;
        border-color: #bcd0e9;
        column-span: all;
        flex-direction: row;
    }

    .install-theme {
        border: none;
        border-collapse: separate;
        overflow: hidden;
        perspective: 5px;
        display: flex;
        width: 166.7%;
        border-radius: 15px 15px 0px 0px;
        background-color: #cce4e1;
        color: #004238;
        border-color: #bcd0e9;
        column-span: all;
        flex-direction: row;
    }
    
    .prestart-theme {
        border: none;
        border-collapse: separate;
        overflow: hidden;
        perspective: 5px;
        display: flex;
        width: 166.7%;
        border-radius: 15px 15px 0px 0px;
        background-color: #c4c0d4; 
        color: #1f1544;
        border-color: #bcd0e9;
        column-span: all;
        flex-direction: row;
    }

    .soc-theme {
        border: none;
        border-collapse: separate;
        overflow: hidden;
        perspective: 5px;
        display: flex;
        width: 166.7%;
        border-radius: 15px 15px 0px 0px;
        background: #bcd0e9;
        color: #070c41;
        border-color: #bcd0e9;
        column-span: all;
        flex-direction: row;
    }

    .task-title-theme {
        background-color: #fafafa;
        color: black;
    }

    .btn-baby-blue {
        background-color: #484b6a;
        color: black;
    }

    .title-theme {
        color: #37474f;
    }

    .btn-light-green {
        background-color: #ddffdd;
        color: black;
    }

    .nav-iko :hover {
        background-color: #f0a1a1;
    }

    .navbar-border {
        border-bottom: 4px solid black;
    }

    .border-radius-test {
        border-radius: 10% 30% 50% 70%;
        border-color: black;
    }

    a.btn.disabled {
        text-decoration: line-through;
        background-color: #fc2c55;
    }

    .width-250 {
        min-width: 250px;
    }

    #dna-render {
        position: absolute;
        right: 0px;
        top: 0px;
        height: 250px;
    }

    tbody#navlinks-table-body tr td a {
        min-width: 100px;
    }

    .table-fit {
        width: 1%;
    }

    h1#DNA {
        /* letter-spacing: 0.03em; */
        font-weight: 300;
        font-size: 250%;
        color: #3f3f3f;
    }


    table#information td {
        padding-right: 30px;
    }

    .table {
        margin-bottom: 0px;
        margin: 1%;
        max-width: 98%;
        
    }

    table.rounded-corners {
        border-spacing: 0;
        border-collapse: separate;
        border-radius: 100px;
    }

    table.rounded-corners th, table.rounded-corners td {
        border: 1px solid black;
    }

    .table-dynamic-theme{
        margin: 1%;

    }

    .width-20-percent{
        width: 20%;
    }

    .break-it{
        width: 60%;
        white-space: pre-line;
    }

    .width-100-percent{
        width: 100;
    }
</style>