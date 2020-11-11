<!doctype html>
<html lang="en">

<head>
    <!-- Required meta tags -->
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <meta http-equiv="X-UA-Compatible" content="IE=Edge">
    <!-- Bootstrap CSS -->
    <script src="https://code.jquery.com/jquery-3.3.1.slim.min.js" type="text/javascript"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.7/umd/popper.min.js"
        type="text/javascript"></script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/js/bootstrap.min.js"
        type="text/javascript"></script>
    <script src="https://unpkg.com/papaparse@5.3.0/papaparse.min.js" type="text/javascript"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css"
        type="text/css">
    <title>IKO Asset Redirect</title>
</head>

<body>
    <nav class="navbar navbar-expand-md navbar-light bg-white navbar-border">
        <a class="navbar-brand" href="http://connect.na.local/Pages/Connect.aspx"><img
                src="/support/Reliability/ReliabilityShared/Pages/Assets/IKO_Logo-nobg.png" alt="IKO" height=50px /></a>
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
                        href="http://operations.connect.na.local/support/Reliability/ReliabilityShared/Pages/AssetFinder2.aspx/?mobile=0">DNA
                        Finder</a>
                </li>
                <li class="nav-item nav-iko active">
                    <a class="nav-link" href="#">Lockout and Tagout</a>
                </li>
                <li class="nav-item nav-iko active">
                    <a class="nav-link"
                        href="http://operations.connect.na.local/support/Reliability/ReliabilityPublished/TrainingMaterial/AssetVideos/Forms/AllItems.aspx">Technical
                        Training Videos</a>
                </li>
                <!-- <li class="nav-item nav-iko active">
                    <a id="plant-layout-link" class="nav-link disabled" href="#" aria-disabled="true">[Plant] Layout</a>
                    pain in the butt since plant layouts dont follow a naming scheme
                </li> -->
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
                                <tbody>
                                    <tr>
                                        <td>
                                            <h1 id="DNA"><strong>D</strong>ocument <br><strong>N</strong>avigation
                                                <br><strong>A</strong>ccelerator</h1>
                                        </td>
                                        <td>
                                            <div class="row">
                                                <h6>Site ID:</h6>
                                            </div>
                                            <div class="row">
                                                <select id="siteName">
                                                    <option value="RAM">RAM: Alconbury</option>
                                                    <option value="CAM">CAM: Appley Bridge</option>
                                                    <option value="GE">GE: Ashcroft</option>
                                                    <option value="GR">GR: BramCal</option>
                                                    <option value="BA">BA: Calgary</option>
                                                    <option value="GJ">GJ: CRC Toronto</option>
                                                    <option value="BL">BL: Hagerstown </option>
                                                    <option value="GH">GH: Hawkesbury</option>
                                                    <option value="GV">GV: Hillsboro (Southwest)</option>
                                                    <option value="GK">GK: IG Brampton</option>
                                                    <option value="CA">CA: Kankakee</option>
                                                    <option value="GI">GI: Madoc </option>
                                                    <option value="PBM">PBM: Senica/Sloviakia</option>
                                                    <option value="GC">GC: Sumas</option>
                                                    <option value="GS">GS: Sylacauga</option>
                                                    <option value="GA">GA: Wilmington</option>
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
                                                <div id="AssetNum"><input type="text" id="assetNum"
                                                        oninput="buttonEnable()"></div>
                                            </div>
                                            <div class="row" style="padding-top: 10px;">
                                                <h6>Asset Description:</h6>
                                            </div>
                                            <div class="row text-wrap">
                                                <p id="description">Please enter asset number</p>
                                            </div>
                                        </td>
                                        <td style="padding-top: 40px;">
                                            <button id="submitNum" disabled=true
                                                onclick="checkSiteCSV()">Generate</button>
                                        </td>
                                        <td style="position: relative; min-width: 295px;">
                                            <video id="dna-render" playsinline autoplay muted loop>
                                                <source
                                                    src="/support/Reliability/ReliabilityShared/Pages/Assets/RotatingDNA.mp4"
                                                    type="video/mp4" />
                                                Your browser does not support this video</video>
                                        </td>
                                    </tr>
                                </tbody>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <h4>Asset & Failure Classes Report Links</h4>
                            <td>
                            <table class="table">
                                <thead>
                                    <tr>
                                        <th scope="col">Asset<br />(Site Specific)</th>
                                        <th scope="col">Failure Class<br />(Site Specific)</th>
                                        <th scope="col">Asset<br />(All Sites)</th>
                                        <th scope="col">Failure Class<br />(All Sites)</th>
                                        <th scope="col">Description</th>
                                    </tr>
                                </thead>
                                <tbody id='dynamic-table-body'>
                                    <tr id="placeholder">
                                        <td><a class="btn btn-primary btn-light-gray">Please Enter</a></td>
                                        <td><a class="btn btn-primary btn-light-yellow">Asset ID</a></td>
                                        <td><a class="btn btn-primary btn-baby-blue">And Select </a></td>
                                        <td><a class="btn btn-primary btn-light-green">A Site</a></td>
                                        <td>
                                            <h4><i class="fa fa-ban"></i> Placeholder Text</h4>
                                        </td>
                                    </tr>
                                </tbody>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <h4>Navigation links</h4>
                            <table class="table table-fit">
                                <tbody id='navlinks-table-body'>
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
                                        <td><a class="btn btn-primary btn-light-gray" hidden href="#"></a></td>
                                        <td>
                                            <h4><i class="fa fa-chain"></i> Lockout and Tagout (In Developement)</h4>
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
                                                href="http://operations.connect.na.local/support/Reliability/ReliabilityShared/Pages/AssetFinder2.aspx/?mobile=0">Finder</a>
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
    var loadedCSV = "";
    var failureClassDataSheet;
    var failureClassNameSheet;
    var siteCVSReadIn;
    var failureCode = "N/A";
    var assetN;
    var siteID;
    var reliabilityAlerts;
    var site_names = {
        "GE": "GJ:%20Ashcroft",
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
    }

    function checkValid() {
        var patt = new RegExp("[A-Z][0-9]{4}"); //check if input is right format
        var str = assetNum.value.toUpperCase();
        var result = patt.test(str);
        if (result == true && str.length == 5) {
            assetNum.value = str;
            return true;
        } else {
            return false;
        }
    }

    function buttonEnable() { //enable the button when a valid asset number is inputted

        if (checkValid()) {
            document.getElementById("submitNum").disabled = false;
            console.log("enable");
        } else {
            document.getElementById("submitNum").disabled = true;
            console.log("disable");
        }
    }

    function checkSiteCSV() {
        assetN = assetNum.value;
        siteID = siteName.value;
        if (siteName.value == loadedCSV) {
            genLinks(siteCVSReadIn);
        } else {
            loadcsv()
        }
    }

    function loadcsv() {
        console.log("load csv");
        //when the generate button is clicked this function runs, it gets the siteID and asset num,
        //and if the sites are equal it imports that CSV then delays before the rest of the function
        //is read because reading the file takes time but running is instant causing a crash if no delay
        assetN = assetNum.value;
        siteID = siteName.value;
        loadedCSV = siteID
        var failureCode = "N/A";
        var desc = "No description available";

        //get asset list using site id
        var url = "/support/Reliability/ReliabilityShared/Pages/Assets/" + siteID + ".csv";
        readCSV(url, genLinks);

    }
    function genLinks(data) {
        console.log("start processing data");
        siteCVSReadIn = data;
        var description = "No Description Found For This Asset";
        var j = siteCVSReadIn.length;
        for (var i = 0; i < j - 1; i++) {
            if (siteCVSReadIn[i][0] == assetN) {
                description = siteCVSReadIn[i][1];
                i = j;
            } //consider adding asset location links ie replace 'Asset Finder' in nav bar with this
        }
        document.getElementById("description").textContent = description
        //get failure class
        var j = failureClassDataSheet.length;
        for (var i = 0; i < j - 1; i++) {
            if (assetN == failureClassDataSheet[i][0] && siteID == failureClassDataSheet[i][2]) {
                failureCode = failureClassDataSheet[i][1];
            };
        }
        //get failure class description
        var failureClassDescription = "No Failure Class Found for This Asset";
        var j = failureClassNameSheet.length;
        for (var i = 0; i < j - 1; i++) {
            if (failureCode == failureClassNameSheet[i][0]) {
                failureClassDescription = failureClassNameSheet[i][1];
            };
        }
        document.getElementById("failure-description").textContent = failureClassDescription

        var urls = [
            [
                "http://operations.connect.na.local/support/Reliability/ReliabilityShared/Pages/SymptomDatabase.html?cheese=" + siteID + "=&cheeseNum=" + assetN + "=&asset=1=&site=1",
                "http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/Asset%20Spare%20Parts&paramSiteID=" + siteID + "&paramAssetNum=" + assetN + "&paramFailureClass=",
                "http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/Failure%20Modes%20Problem%20Asset&paramSiteID=" + siteID + "&paramAssetNum=" + assetN,
                "http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/PMProgram&paramSiteID=" + siteID + "&paramAssetNum=" + assetN + "&paramFailureClass=",
                "http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/WorkOrders&paramSiteID=" + siteID + "&paramAssetNum=" + assetN + "&paramFailureClass=",
                "http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/Downtime%20MTBF&paramSiteID=" + siteID + "&paramAssetNum=" + assetN + "&paramFailureClass=",
                "http://operations.connect.na.local/support/Reliability/ReliabilityPublished/ReliabilityCriticality-RCA-FMECA/ReliabilityAlerts/Forms/AllItems.aspx?FilterField1=RelatedAssetNumber_x0028_s_x0029_2&FilterValue1=" + assetN + "&FilterField2=Site_x0020_Descriptions2&FilterValue2=" + site_names[siteID],
            ],
            [
                "http://operations.connect.na.local/support/Reliability/ReliabilityShared/Pages/SymptomDatabase.html?cheese=" + siteID + "=&cheeseNum=" + assetN + "=&asset=0=&site=1",
                "",
                "http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/Failure%20Modes%20Problem&paramSiteID=" + siteID + "&paramFailureClass=" + failureCode,
                "http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/PMProgram&paramSiteID=" + siteID + "&paramAssetNum=" + "&paramFailureClass=" + failureCode,
                "http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/WorkOrders&paramSiteID=" + siteID + "&paramAssetNum=" + "&paramFailureClass=" + failureCode,
                "http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/Downtime%20MTBF&paramSiteID=" + siteID + "&paramAssetNum=" + "&paramFailureClass=" + failureCode,
                "http://operations.connect.na.local/support/Reliability/ReliabilityPublished/ReliabilityCriticality-RCA-FMECA/ReliabilityAlerts/Forms/AllItems.aspx?FilterField1=Originated_x0020_Asset_x0023_&FilterValue1=" + failureCode + "&FilterField2=Site_x0020_Descriptions2&FilterValue2=" + site_names[siteID],
            ],
            [
                "http://operations.connect.na.local/support/Reliability/ReliabilityShared/Pages/SymptomDatabase.html?cheese=" + siteID + "=&cheeseNum=" + assetN + "=&asset=1=&site=0",
                "http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/Asset%20Spare%20Parts&paramSiteID=&paramAssetNum=" + assetN + "&paramFailureClass=",
                "http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/Failure%20Modes%20Problem%20Asset&paramSiteID=" + "&paramAssetNum=" + assetN,
                "http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/PMProgram&paramSiteID=" + "&paramAssetNum=" + assetN + "&paramFailureClass=",
                "http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/WorkOrders&paramSiteID=" + "&paramAssetNum=" + assetN + "&paramFailureClass=",
                "http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/Downtime%20MTBF&paramSiteID=" + "&paramAssetNum=" + assetN + "&paramFailureClass=",
                "http://operations.connect.na.local/support/Reliability/ReliabilityPublished/ReliabilityCriticality-RCA-FMECA/ReliabilityAlerts/Forms/AllItems.aspx?FilterField1=RelatedAssetNumber_x0028_s_x0029_2&FilterValue1=" + assetN,
            ],
            [
                "http://operations.connect.na.local/support/Reliability/ReliabilityShared/Pages/SymptomDatabase.html?cheese=" + siteID + "=&cheeseNum=" + assetN + "=&asset=0=&site=0",
                "",
                "http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/Failure%20Modes%20Problem&paramSiteID=" + "&paramFailureClass=" + failureCode,
                "",
                "http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/WorkOrders&paramSiteID=" + "&paramAssetNum=" + "&paramFailureClass=" + failureCode,
                "http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/Downtime%20MTBF&paramSiteID=" + "&paramAssetNum=" + "&paramFailureClass=" + failureCode,
                "http://operations.connect.na.local/support/Reliability/ReliabilityPublished/ReliabilityCriticality-RCA-FMECA/ReliabilityAlerts/Forms/AllItems.aspx?FilterField1=Originated_x0020_Asset_x0023_&FilterValue1=" + failureCode,
            ],
        ]
        var cellFormat = ["btn-light-gray", "btn-light-yellow", "btn-baby-blue", "btn-light-green"]
        var cellText = [assetN + "(" + siteID + ")", failureCode + "(" + siteID + ")", assetN + "(All Sites)", failureCode + "(All Sites)"]
        var descriptions = [
            ["fa fa-th-list", " Symptom Database"],
            ["fa fa-file-text-o", " Spare Parts List"],
            ["fa fa-chain-broken", " Failure Modes & Causes (In Progress)"],
            ["fa fa-check-square-o", " PMs For The Asset (In Progress)"],
            ["fa fa-bullhorn", " Work Orders (In Progress)"],
            ["fa fa-hourglass-2", " Downtime Report with MTBF (In Progress)"],
            ["fa fa-bell-o", " Reliability Alerts (Disabled if there are no alerts)"],
        ]
        var parent = document.getElementById("dynamic-table-body");
        parent.textContent = ''; //remove placeholder or old data
        for (var i = 0; i < 7; i++) {
            var row = document.createElement("tr");
            row.setAttribute("id", "dynamic-row-".concat(i));
            for (var j = 0; j < 5; j++) {
                var cellTD = document.createElement("td");
                cellTD.setAttribute("id", "cellR".concat(i, "C", j));
                var button = document.createElement("a");
                if (j == 4) {
                    var button = document.createElement("h4");
                    var icon = document.createElement("i");
                    icon.setAttribute("class", descriptions[i][0]);
                    button.appendChild(icon);
                    var buttonText = document.createTextNode(descriptions[i][1]);
                } else {
                    if (urls[j][i] == "") {
                        button.setAttribute("hidden", true);
                    } else {
                        button.setAttribute("href", urls[j][i]);
                    }
                    button.setAttribute("class", "btn btn-primary " + [cellFormat[j]]);
                    var buttonText = document.createTextNode(cellText[j]);
                }
                button.appendChild(buttonText);
                cellTD.appendChild(button);
                row.appendChild(cellTD);
            }
            parent.appendChild(row);
        }
        //disable buttons as required
        var checkForValues = [siteID.concat("_", assetN), siteID.concat("_", failureCode), assetN, failureCode];
        for (var i = 0; i < 4; i++) {
            var found = false;
            var k = reliabilityAlerts[i].length;
            for (j = 0; j < k; j++) {
                if (reliabilityAlerts[i][j] == checkForValues[i]) {
                    found = true;
                    break //if value is found no need to disable button, so go check next value
                }
            }
            if (!found) {
                const div = document.getElementById("cellR6C".concat(i)).firstElementChild
                div.classList.add("disabled")
            }
        }
    }


    function readCSV(url, callback) { //reads csv into an array using papaparse libary
        console.log("Start CSV");
        Papa.parse(url, {
            download: true,
            complete: function (results) {
                console.log("Finished Parsing Array");
                callback(results.data.slice(0))
            }
        });
    }
    function startGenLinks(data) {
        failureClassDataSheet = data;
        if (checkValid()) {
            document.getElementById("submitNum").disabled = false;
            loadcsv();
        }
    }

    function getFailureCodeNames(data) {
        failureClassNameSheet = data;
    }

    function getReliabilityAlerts(data) {
        reliabilityAlerts = data;
    }

    document.addEventListener('DOMContentLoaded', function () { //set to siteID equal to the variable in the URL ...?cheese=GK case insensitive
        //this runs as soon as the site has been loaded
        console.log("finished loading")
        var paraURL = window.location.search

        var url = "/support/Reliability/ReliabilityShared/Pages/Assets/FailureClassesDescriptions.csv";
        readCSV(url, getFailureCodeNames);

        url = "/support/Reliability/ReliabilityShared/Pages/Assets/FailureClasses.csv";//reads in failure classes as soon as program starts
        readCSV(url, startGenLinks);

        url = "/support/Reliability/ReliabilityShared/Pages/Assets/ReliabilityAlerts.csv";
        readCSV(url, getReliabilityAlerts);

        if (paraURL.indexOf("cheese") != -1) {
            var cheese = paraURL.split('=')[1]; //get siteID
            siteName.value = cheese.toUpperCase();
        }
        if (paraURL.indexOf("cheeseNum") != -1) {
            var cheeseNum = window.location.search.split('=')[3]; //get assetNum
            assetNum.value = cheeseNum.toUpperCase();
        }

    }, false);


</script>


<style>
    .btn-lilac {
        background-color: #C8A2C8;
        color: black;
    }

    .btn-light-gray {
        background-color: #f1f1f1;
        color: black;
    }

    .btn-light-yellow {
        background-color: #fdf5e6;
        color: black;
    }

    .btn-baby-blue {
        background-color: #ddffff;
        color: black;
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
        height: 360px;
        /* using absolute position makes the element 
        float above other content without moving them
        Note: a relative element must be a parent of 
        the absolute element */
    }

    tbody#navlinks-table-body tr td a {
        min-width: 100px;
    }

    .table-fit {
        width: 1%;
    }

    h1#DNA {
        letter-spacing: 0.03em;
        font-weight: 300;
        font-size: 350%;
        color: #3f3f3f;
    }

    table#information td {
        padding-right: 30px;
    }
</style>