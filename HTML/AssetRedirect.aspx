<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<!--
Design by TEMPLATED
http://templated.co
Released for free under the Creative Commons Attribution License

Name       : Swarming
Description: A two-column, fixed-width design with dark color scheme.
Version    : 1.0
Released   : 20131201

-->
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <!--tag to make IE use its modern rendering engine THE FIRST LINE IS VERY IMPORTANT FOR ALL CODES SAMSON AND MOST OF IKO USE INTERNET EXPLORER-->
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta name="keywords" content="" />
    <meta name="description" content="" />
    <link href="http://fonts.googleapis.com/css?family=Source+Sans+Pro:200,300,400,600,700,900" rel="stylesheet" />
    <link rel="stylesheet" href="https://www.w3schools.com/w3css/4/w3.css">
    <script src="https://unpkg.com/papaparse@5.3.0/papaparse.min.js" type="text/javascript"></script>
    <meta name="viewport" content="width=1600, initial-scale=1">
    <link rel="stylesheet" href="https://www.w3schools.com/w3css/4/w3.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
    <link href="https://fonts.googleapis.com/icon?family=Material+Icons" rel="stylesheet">

</head>


<body>

    <div id="page-wrapper">
        <div id="page" class="container">
            <div id="content">
                <div class="title">
                    <h1><strong>D</strong>ocument <br><strong>N</strong>avigation <br><strong>A</strong>ccelerator </h1>
                </div>

                <div id="searchTable">
                    <table>
                        <tr>
                            <td>
                                <as>Site ID:</as>
                            </td>
                            <td>
                                <!--    drop down menu for site selector        -->
                                <!-- TODO change to load from js dictionary -->
                                <select id="siteName">
                                    <option value="RAM">RAM: Alconbury</option>
                                    <option value="CAM">CAM: Appley Bridge</option>
                                    <option value="GE" >GE: Ashcroft</option>
                                    <option value="GR" >GR: BramCal</option>
                                    <option value="BA" >BA: Calgary</option>
                                    <option value="GJ" >GJ: CRC Toronto</option>
                                    <option value="BL" >BL: Hagerstown </option>
                                    <option value="GH" >GH: Hawkesbury</option>
                                    <option value="GV" >GV: Hillsboro (Southwest)</option>
                                    <option value="GK" >GK: IG Brampton</option>
                                    <option value="CA" >CA: Kankakee</option>
                                    <option value="GI" >GI: Madoc </option>
                                    <option value="PBM">PBM: Senica/Sloviakia</option>
                                    <option value="GC" >GC: Sumas</option>
                                    <option value="GS" >GS: Sylacauga</option>
                                    <option value="GA" >GA: Wilmington</option>
                                </select>
                            </td>
                        </tr>
                        <tr>
                            <td id="col1">
                                <as2>
                                    <p>Asset Number:</p>
                                </as2>
                            </td>
                            <td id="col2">
                                <!--Input box & dropdown for asset number-->
                                <as3>
                                    <div id="AssetNum"><input type="text" id="assetNum" oninput="buttonEnable()"></div>
                                </as3>
                            </td>
                            <td id="col4">
                                <as3> <button id="submitNum" disabled=true onclick="genLinks()">Generate</button> </as3>
                            </td>
                        </tr>
                    </table>
                </div>

                <dsa>Description:</dsa>


            </div>
            <div id="sidebar"><a href="#" class="image image-full"><img
                        src="http://operations.connect.na.local/support/Reliability/PublishingImages/IKO_Logo.png"
                        alt="" /></a></div>
            <video class="title" id="dsad" width="200" height="290" autoplay="autoplay" loop="true" alt="DNA Render">
                            <source src="/support/Reliability/ReliabilityShared/Pages/Assets/RotatingDNA.mp4" type="video/mp4" />
                        Your browser does not support this video</video>
            <!-- <div id="dsad" class="title"> <img
                    src="http://operations.connect.na.local/support/Reliability/PublishingImages/RotatingDNA.gif"
                    alt="HTML5 Icon" width="200" height="290"> </div> -->

        </div>
        <sad id="result6"></sad>

    </div>



    <div id="featured-wrapper">
        <div id="featured" class="container">
            <!-- New code with table. If there's a better way to code this (via JS) then do so -->
            <!-- Yeah your way is pretty bonobo -->

            <table id="fTable">
                <tr height="65%">
                    <td class="left">
                        <div class="boxy">
                            <p id="result1_1"></p>
                        </div>
                    </td>
                    <td class="all">
                        <div class="boxy">
                            <p id="result2_1"></p>
                        </div>
                    </td>
                    <td class="all">
                        <div class="boxy">
                            <p id="result3_1"></p>
                        </div>
                    </td>
                    <td class="right">
                        <div class="boxy">
                            <p id="result4_1"></p>
                            <div class="boxy">
                    </td>
                    <td>
                        <div class="descript">
                            <p id="result5_1"></p>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td class="left">
                        <div class="boxy">
                            <p id="result1_2"></p>
                        </div>
                    </td>
                    <td class="all">
                        <div class="boxy">
                            <p id="result2_2"></p>
                        </div>
                    </td>
                    <td class="all">
                        <div class="boxy">
                            <p id="result3_2"></p>
                        </div>
                    </td>
                    <td class="all">
                        <div class="boxy">
                            <p id="result4_2"></p>
                        </div>
                    </td>
                    <td class="right">
                        <div class="descript">
                            <p id="result5_2"></p>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td class="left">
                        <div class="boxy">
                            <p id="result1_3"></p>
                        </div>
                    </td>
                    <td class="all">
                        <div class="boxy">
                            <p id="result2_3"></p>
                        </div>
                    </td>
                    <td class="all">
                        <div class="boxy">
                            <p id="result3_3"></p>
                        </div>
                    </td>
                    <td class="all">
                        <div class="boxy">
                            <p id="result4_3"></p>
                        </div>
                    </td>
                    <td class="right">
                        <div class="descript">
                            <p id="result5_3"></p>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td class="left">
                        <div class="boxy">
                            <p id="result1_4"></p>
                        </div>
                    </td>
                    <td class="all">
                        <div class="boxy">
                            <p id="result2_4"></p>
                        </div>
                    </td>
                    <td class="all">
                        <div class="boxy">
                            <p id="result3_4"></p>
                        </div>
                    </td>
                    <td class="all">
                        <div class="boxy">
                            <p id="result4_4"></p>
                        </div>
                    </td>
                    <td class="right">
                        <div class="descript">
                            <p id="result5_4"></p>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td class="left">
                        <div class="boxy">
                            <p id="result1_5"></p>
                        </div>
                    </td>
                    <td class="all">
                        <div class="boxy">
                            <p id="result2_5"></p>
                        </div>
                    </td>
                    <td class="all">
                        <div class="boxy">
                            <p id="result3_5"></p>
                        </div>
                    </td>
                    <td class="all">
                        <div class="boxy">
                            <p id="result4_5"></p>
                        </div>
                    </td>
                    <td class="right">
                        <div class="descript">
                            <p id="result5_5"></p>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td class="left">
                        <div class="boxy">
                            <p id="result1_6"></p>
                        </div>
                    </td>
                    <td class="all">
                        <div class="boxy">
                            <p id="result2_6"></p>
                        </div>
                    </td>
                    <td class="all">
                        <div class="boxy">
                            <p id="result3_6"></p>
                        </div>
                    </td>
                    <td class="all">
                        <div class="boxy">
                            <p id="result4_6"></p>
                        </div>
                    </td>
                    <td class="right">
                        <div class="descript">
                            <p id="result5_6"></p>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td class="left">
                        <div class="boxy">
                            <p id="result1_7"></p>
                        </div>
                    </td>
                    <td class="all">
                        <div class="boxy">
                            <p id="result2_7"></p>
                        </div>
                    </td>
                    <td class="all">
                        <div class="boxy">
                            <p id="result3_7"></p>
                        </div>
                    </td>
                    <td class="all">
                        <div class="boxy">
                            <p id="result4_7"></p>
                        </div>
                    </td>
                    <td class="right">
                        <div class="descript">
                            <p id="result5_7"></p>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td class="left">
                        <div class="boxy">
                            <p id="result1_8"></p>
                        </div>
                    </td>
                    <td class="all">
                        <div class="boxy">
                            <p id="result2_8"></p>
                        </div>
                    </td>
                    <td class="all">
                        <div class="boxy">
                            <p id="result3_8"></p>
                        </div>
                    </td>
                    <td class="all">
                        <div class="boxy">
                            <p id="result4_8"></p>
                        </div>
                    </td>
                    <td class="right">
                        <div class="descript">
                            <p id="result5_8"></p>
                        </div>
                    </td>
                </tr>
            </table>

            <table id="sTable">
                <!-- sTable for second table -->
                <tr height="65%">
                    <td>
                        <div class="boxy">
                            <p><a class="w3-button w3-light-grey w3-border w3-border-black w3-round-large"
                                    href="http://operations.connect.na.local/support/Reliability/ReliabilityPublished/TrainingMaterial/AssetVideos/Forms/AllItems.aspx">Videos</a>
                            </p>
                        </div>
                    </td>
                    <td>
                        <div class="descript">
                            <p class="text"><a><i class="fa fa-youtube-play"> Technical Training Videos (In Progress)
                                    </i></a></p>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td>
                        <div class="boxy">
                            <p><a class="w3-button w3-text-white w3-hover-white w3-hover-text-white w3-round-large w3-disabled"
                                    href="#">Blank</a></p>
                        </div>
                    </td>
                    <td>
                        <div class="descript">
                            <p class="text"><a><i class="fa fa-chain"> Lockout and Tagout (In Progress) </i></a></p>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td>
                        <div class="boxy">
                            <p><a class="w3-button w3-light-grey w3-border w3-border-black w3-round-large"
                                    href="http://operations.connect.na.local/support/Reliability/SitePages/Home.aspx">Homepage</a>
                            </p>
                        </div>
                    </td>
                    <td>
                        <div class="descript">
                            <p class="text"><a><i class="fa fa-gears"> Reliability Sharepoint Homepage </i></a></p>
                        </div>
                    </td> <!-- back option at bottom of the page -->
                </tr>
            </table>

        </div>
    </div>

</body>

</html>

<script type="application/javascript">
    var start;
    var arrayExist = false;
    var failureClassDataSheet;
    var siteCVSReadIn;
    var convertCase = false;
    var sites = ["BA", "BL", "CA", "GA", "GC", "GE", "GH", "GI", "GJ", "GK", "GR", "GS", "GV", "PBM", "CAM", "RAM"];
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

    function genLinks() {
        console.log("generating links1");
        //when the generate button is clicked this function runs, it gets the siteID and asset num,
        //and if the sites are equal it imports that CSV then delays before the rest of the function
        //is read because reading the file takes time but running is instant causing a crash if no delay
        var assetN = assetNum.value;
        var siteID = siteName.value;

        for (i = 0; i < sites.length; i++) {
            if (siteID == sites[i]) {
                url = "/support/Reliability/ReliabilityShared/Pages/" + sites[i] + ".csv";
                readCSV1(url);
            }
        }

        var delayInMilliseconds = 500;
        setTimeout(function () {
            console.log("delay");
            genLinksPart2();
        }, delayInMilliseconds);

    }

    function genLinksPart2() {

        console.log("generating links2");
        var assetN = assetNum.value;
        var siteID = siteName.value;
        var FC = "N/A";
        //	var textRow = ["","","","","","","","",""];
        var text = ["", "", "", "", ""];
        var desc = "No description available";


        var a = siteCVSReadIn.length;
        for (i = 0; i < a - 1; i++) {
            if (siteCVSReadIn[i][0] == assetN) {
                desc = siteCVSReadIn[i][1];
            }
        }

        var n = failureClassDataSheet.length;
        for (i = 0; i < n - 1; i++) {
            if (assetN == failureClassDataSheet[i][0] && siteID == failureClassDataSheet[i][2]) {
                FC = failureClassDataSheet[i][1];
            } else if (FC == "" || FC == "N/A") {
                FC = "N/A";
            }
        }


        //  var divLine ="__________________"
        //if you need quotes in the middle of block of text use /" which makes the quotation mark not end your quote but still read by HTML
        console.log("before text");

        var reliability_alert_url = "http://operations.connect.na.local/support/Reliability/ReliabilityPublished/ReliabilityCriticality-RCA-FMECA/ReliabilityAlerts/Forms/AllItems.aspx?FilterField1=RelatedAssetNumber_x0028_s_x0029_2&FilterValue1=" + assetN + "&FilterField2=Site_x0020_Descriptions2&FilterValue2=" + site_names[siteID];
        text[0] = new Array("<div class=\"box1\"><d><a>Asset <br/>(Site Specific)</a></d></div>",
            "<p><a class=\"w3-button w3-light-grey w3-border w3-border-black w3-round-large \" href=\"http://operations.connect.na.local/support/Reliability/ReliabilityShared/Pages/SymptomDatabase.html?cheese=" + siteID + "=&cheeseNum=" + assetN + "=&asset=1=&site=1" + "\">" + assetN + "(" + siteID + ")" + "</a></p>",
            "<p><a class=\"w3-button w3-light-grey w3-border w3-border-black w3-round-large \" href=\"http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/Asset%20Spare%20Parts&paramSiteID=" + siteID + "&paramAssetNum=" + assetN + "&paramFailureClass=" + "\">" + assetN + "(" + siteID + ")" + "</a></p>",
            "<p><a class=\"w3-button w3-light-grey w3-border w3-border-black w3-round-large \" href=\"http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/Failure%20Modes%20Problem%20Asset&paramSiteID=" + siteID + "&paramAssetNum=" + assetN + "\">" + assetN + "(" + siteID + ")" + "</a></p>",
            "<p><a class=\"w3-button w3-light-grey w3-border w3-border-black w3-round-large \" href=\"http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/PMProgram&paramSiteID=" + siteID + "&paramAssetNum=" + assetN + "&paramFailureClass=" + "\">" + assetN + "(" + siteID + ")" + "</a></p>",
            "<p><a class=\"w3-button w3-light-grey w3-border w3-border-black w3-round-large \" href=\"http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/WorkOrders&paramSiteID=" + siteID + "&paramAssetNum=" + assetN + "&paramFailureClass=" + "\">" + assetN + "(" + siteID + ")" + "</a></p>",
            "<p><a class=\"w3-button w3-light-grey w3-border w3-border-black w3-round-large \" href=\"http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/Downtime%20MTBF&paramSiteID=" + siteID + "&paramAssetNum=" + assetN + "&paramFailureClass=" + "\">" + assetN + "(" + siteID + ")" + "</a></p>",
            "<p><a class=\"w3-button w3-light-grey w3-border w3-border-black w3-round-large \" href=" + reliability_alert_url +">" + assetN + "(" + siteID + ")" + "</a></p>");

        var reliability_alert_url = "http://operations.connect.na.local/support/Reliability/ReliabilityPublished/ReliabilityCriticality-RCA-FMECA/ReliabilityAlerts/Forms/AllItems.aspx?FilterField1=Originated_x0020_Asset_x0023_&FilterValue1=" + FC + "&FilterField2=Site_x0020_Descriptions2&FilterValue2=" + site_names[siteID];
        text[1] = new Array("<div class=\"box2\"><d><a>Failure Class <br/>(Site Specific)</a></d></div>",
            "<p><a class=\"w3-button w3-sand w3-border w3-border-black w3-round-large \" href=\"http://operations.connect.na.local/support/Reliability/ReliabilityShared/Pages/SymptomDatabase.html?cheese=" + siteID + "=&cheeseNum=" + assetN + "=&asset=0=&site=1" + "\">" + FC + "(" + siteID + ")" + "</a></p>",
            "<p><a class=\"w3-button w3-text-white w3-hover-white w3-hover-text-white w3-round-large w3-disabled\" href=\"#\">Blank</a></p>",
            "<p><a class=\"w3-button w3-sand w3-border w3-border-black w3-round-large \" href=\"http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/Failure%20Modes%20Problem&paramSiteID=" + siteID + "&paramFailureClass=" + FC + "\">" + FC + "(" + siteID + ")" + "</a></p>",
            "<p><a class=\"w3-button w3-sand w3-border w3-border-black w3-round-large \" href=\"http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/PMProgram&paramSiteID=" + siteID + "&paramAssetNum=" + "&paramFailureClass=" + FC + "\">" + FC + "(" + siteID + ")" + "</a></p>",
            "<p><a class=\"w3-button w3-sand w3-border w3-border-black w3-round-large \" href=\"http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/WorkOrders&paramSiteID=" + siteID + "&paramAssetNum=" + "&paramFailureClass=" + FC + "\">" + FC + "(" + siteID + ")" + "</a></p>",
            "<p><a class=\"w3-button w3-sand w3-border w3-border-black w3-round-large \" href=\"http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/Downtime%20MTBF&paramSiteID=" + siteID + "&paramAssetNum=" + "&paramFailureClass=" + FC + "\">" + FC + "(" + siteID + ")" + "</a></p>",
            "<p><a class=\"w3-button w3-sand w3-border w3-border-black w3-round-large \" href=" + reliability_alert_url + ">" + FC + "(" + siteID + ")" + "</a></p>");

            reliability_alert_url = "http://operations.connect.na.local/support/Reliability/ReliabilityPublished/ReliabilityCriticality-RCA-FMECA/ReliabilityAlerts/Forms/AllItems.aspx?FilterField1=RelatedAssetNumber_x0028_s_x0029_2&FilterValue1=" + assetN
        text[2] = new Array("<div class=\"box3\"><d><a><a>Asset <br/>(All Sites)</a></d></div>",
            "<p><a class=\"w3-button w3-pale-blue w3-border w3-border-black w3-round-large \" href=\"http://operations.connect.na.local/support/Reliability/ReliabilityShared/Pages/SymptomDatabase.html?cheese=" + siteID + "=&cheeseNum=" + assetN + "=&asset=1=&site=0" + "\">" + assetN + "(All Sites)" + "</a></p>",
            "<p><a class=\"w3-button w3-pale-blue w3-border w3-border-black w3-round-large \" href=\"http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/Asset%20Spare%20Parts&paramSiteID=&paramAssetNum=" + assetN + "&paramFailureClass=" + "\">" + assetN + "(All Sites)" + "</a></p>",
            "<p><a class=\"w3-button w3-pale-blue w3-border w3-border-black w3-round-large \" href=\"http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/Failure%20Modes%20Problem%20Asset&paramSiteID=" + "&paramAssetNum=" + assetN + "\">" + assetN + "(All Sites)" + "</a></p>",
            "<p><a class=\"w3-button w3-pale-blue w3-border w3-border-black w3-round-large \" href=\"http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/PMProgram&paramSiteID=" + "&paramAssetNum=" + assetN + "&paramFailureClass=" + "\">" + assetN + "(All Sites)" + "</a></p>",
            "<p><a class=\"w3-button w3-pale-blue w3-border w3-border-black w3-round-large \" href=\"http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/WorkOrders&paramSiteID=" + "&paramAssetNum=" + assetN + "&paramFailureClass=" + "\">" + assetN + "(All Sites)" + "</a></p>",
            "<p><a class=\"w3-button w3-pale-blue w3-border w3-border-black w3-round-large \" href=\"http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/Downtime%20MTBF&paramSiteID=" + "&paramAssetNum=" + assetN + "&paramFailureClass=" + "\">" + assetN + "(All Sites)" + "</a></p>",
            "<p><a class=\"w3-button w3-pale-blue w3-border w3-border-black w3-round-large \" href=" + reliability_alert_url +">" + assetN + "(All Sites)" + "</a></p>");

            reliability_alert_url = "http://operations.connect.na.local/support/Reliability/ReliabilityPublished/ReliabilityCriticality-RCA-FMECA/ReliabilityAlerts/Forms/AllItems.aspx?FilterField1=Originated_x0020_Asset_x0023_&FilterValue1=" + FC;
        text[3] = new Array("<div class=\"box4\"><d><a><a>Failure Class <br/>(All Sites)</a></d></div>",
            "<p><a class=\"w3-button w3-pale-green w3-border w3-border-black w3-round-large \" href=\"http://operations.connect.na.local/support/Reliability/ReliabilityShared/Pages/SymptomDatabase.html?cheese=" + siteID + "=&cheeseNum=" + assetN + "=&asset=0=&site=0" + "\">" + FC + "(All Sites)" + "</a></p>",
            "<p><a class=\"w3-button w3-text-white w3-hover-white w3-hover-text-white w3-round-large w3-disabled\" href=\"#\">Blank</a></p>",
            "<p><a class=\"w3-button w3-pale-green w3-border w3-border-black w3-round-large \" href=\"http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/Failure%20Modes%20Problem&paramSiteID=" + "&paramFailureClass=" + FC + "\">" + FC + "(All Sites)" + "</a></p>",
            "<p><a class=\"w3-button w3-text-white w3-hover-white w3-hover-text-white w3-round-large w3-disabled\" href=\"#\">Blank</a></p>",
            "<p><a class=\"w3-button w3-pale-green w3-border w3-border-black w3-round-large \" href=\"http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/WorkOrders&paramSiteID=" + "&paramAssetNum=" + "&paramFailureClass=" + FC + "\">" + FC + "(All Sites)" + "</a></p>",
            "<p><a class=\"w3-button w3-pale-green w3-border w3-border-black w3-round-large \" href=\"http://nscandacssrs1/ReportServer/Pages/ReportViewer.aspx?/Maximo/Downtime%20MTBF&paramSiteID=" + "&paramAssetNum=" + "&paramFailureClass=" + FC + "\">" + FC + "(All Sites)" + "</a></p>",
            "<p><a class=\"w3-button w3-pale-green w3-border w3-border-black w3-round-large \" href=" + reliability_alert_url + ">" + FC + "(All Sites)" + "</a></p>");

        text[4] = new Array("",
            "<p class= \"text\"><a><i class=\"fa fa-th-list\"> Symptom Database</i></a></p>",
            "<p class= \"text\"><a><i class=\"fa fa-file-text-o\"> Spare Parts List</i></a></p>",
            "<p class= \"text\"><a><i class=\"fa fa-chain-broken\"> Failure Modes & Causes (In Progress)</i></a></p>",
            "<p class= \"text\"><a><i class=\"fa fa-check-square-o\"> PMs For The Asset (In Progress)</i></a></p>",
            "<p class= \"text\"><a><i class=\"fa fa-bullhorn\"> Work Orders (In Progress)</i></a></p>",
            "<p class= \"text\"><a><i class=\"fa fa-hourglass-2\"> Downtime Report with MTBF (In Progress)</i></a></p>",
            "<p class= \"text\"><a><i class=\"fa fa-bell-o\"> Reliability Alerts </i></a></p>");
        console.log("after text")

        // var color = "<div class=\"box1\"></div><div class=\"box2\"></div>"
        var numCol = 5; //number of columns in Failure Case table.
        var numRow = 8; //number of rows in Failure Case table.

        for (var col = 0; col < numCol; col++) {
            var stringCol = col + 1;
            var sCol = stringCol.toString();
            for (var row = 0; row < numRow; row++) {
                var stringRow = row + 1;
                var sRow = stringRow.toString();
                document.getElementById("result" + sCol + "_" + sRow).innerHTML = text[col][row];
            }
        }
        console.log("created buttons")
        //this functions takes the text and spits it out in the html code as if it was html text where result(number) is seen
        document.getElementById("result6").innerHTML = desc; //outputs text where result(number) is in the html code as if it was written in
        //	document.getElementById("result7").innerHTML = color; //colored box

        // change reliability alert url to be asset + site specific
        // var reliability_alert_url = "http://operations.connect.na.local/support/Reliability/ReliabilityPublished/ReliabilityCriticality-RCA-FMECA/ReliabilityAlerts/Forms/AllItems.aspx?FilterField1=RelatedAssetNumber_x0028_s_x0029_2&FilterValue1=" + assetN + "&FilterField2=Site_x0020_Descriptions2&FilterValue2=" + site_names[siteID]
        // var reliability_alert = document.getElementById("reliability_alert")
        // reliability_alert.setAttribute('href', reliability_alert_url)
    }

    function readCSV(url) { //reads csv into an array using papaparse libary
        console.log("Start CSV");
        Papa.parse(url, {
            download: true,
            complete: function (results) {
                arrayExist = true;
                console.log("Finished Parsing Array");
                failureClassDataSheet = results.data.slice(0);

            }
        });
    }

    function readCSV1(url) { //This function is for parsing the array for the descriptions of the asset, it reruns everytime the site changes and the data gets overwritten
        console.log("Start CSV");
        Papa.parse(url, {
            download: true,
            complete: function (results) {
                arrayExist = true;//not really needed anymore
                console.log("Finished Parsing Array");
                siteCVSReadIn = results.data.slice(0);

            }
        });
    }

    document.addEventListener('DOMContentLoaded', function () { //set to siteID equal to the variable in the URL ...?cheese=GK case insensitive
        //this runs as soon as the site has been loaded
        var paraURL = window.location.search

        var url = "/support/Reliability/ReliabilityShared/Pages/Assets/FailureClasses.csv";//reads in failure classes as soon as program starts
        readCSV(url);
        document.body.style.zoom = "100%"


        if (paraURL.indexOf("cheese") != -1) {
            var cheese = paraURL.split('=')[1]; //get siteID
            siteName.value = cheese.toUpperCase();
        }
        if (paraURL.indexOf("cheeseNum") != -1) {
            var cheeseNum = window.location.search.split('=')[3]; //get assetNum
            assetNum.value = cheeseNum.toUpperCase();
        }

        //call function for generating links
        if (checkValid()) {
            document.getElementById("submitNum").disabled = false;
            genLinks();
        }
    }, false);
</script>


<style>
    html,

    body {
        margin: 0px;
        padding: 0px;
        background: #191919;
        font-family: 'Source Sans Pro', sans-serif;
        font-size: 12pt;
        font-weight: 400;
        color: #4E4D4D;

    }


    p {
        margin-top: 0;
    }


    p {
        line-height: 180%;
    }

    a {

        width: 170px
    }

    .container {
        width: 1300px;
        margin: 0px 40px;
        position: relative;
        top: 50px;
    }


    .image {
        display: inline-block;
    }

    .image img {
        display: block;
        width: 100%;
    }

    .image-full {
        display: block;
        width: 100%;
        margin: 0 0 0 0;
    }

    .image-left {
        float: left;
        margin: 0 2em 2em 0;
    }

    .image-centered {
        display: block;
        margin: 0 0 2em 0;
    }

    .image-centered img {
        margin: 0 auto;
        width: auto;
    }

    .title {
        text-align: center;
        position: relative;
        top: 0px;
        left: 100px;
    }

    .title h1 {
        letter-spacing: 0.03em;
        font-weight: 300;
        font-size: 420%;
        color: #3f3f3f;
        position: absolute;
        left: 690px;
        top: 20px;
        z-index: 1;
    }

    h3 {

        font-weight: 800;
    }

    .title .byline {
        display: block;
        padding-top: 1em;
        font-weight: 300;
        font-size: 1.1em;
    }

    .title2 h2 {
        text-align: left;
    }

    #page-wrapper {
        background: #FFFFFF;
        height: 330px;

    }

    #page {
        padding: 2em 0em;
        height: 330px;
        position: relative;
        top: 0px;
    }


    #content {
        float: left;
        width: 588px;
        text-align: center;

    }

    #content .title {
        padding: 1em 2em 2em 1em;
        text-align: left;

    }

    #content h2 {
        font-size: 2em;
    }

    #content .byline {
        padding-top: 0;
        font-size: 1.5em;
        padding-bottom: 0;
    }


    #sidebar {
        float: left;
        width: 500px;
        position: relative;
        top: 0px;
        left: -550px;
        transform: scale(0.68, 0.68);
    }

    /*
	#instruct {
		position: relative;
		left: 0px;
		top: 140px;
		z-index: 2;
	}
*/
    #searchTable {
        position: relative;
        left: -40px;
        top: 100px;
        z-index: 3;
    }

    #searchTable td {
        position: relative;
        left: 90px;
        font-size: 95%;
    }

    #searchTable p {
        font-size: 95%;
    }

    #copyright {
        overflow: hidden;
        padding: 5em 0em;
        height: 60px;
    }

    #copyright p {
        letter-spacing: 0.20em;
        text-align: center;
        text-transform: uppercase;
        font-size: 0.80em;
        color: #6F6F6F;
        position: relative;
        bottom: 85px;
    }

    #copyright a {
        text-decoration: none;
        color: #8C8C8C;
    }

    #featured-wrapper {
        overflow: auto;
        padding: 1em 1em;
        background: #FFFFFF;
        height: 800px;
        z-index: 4;
    }

    #featured {
        position: relative;
        top: 3px;
    }

    #featured h2 {
        text-align: left;
        position: relative;
        bottom: 45px;
    }

    .boxy {
        position: relative;
        top: 8px;
    }

    .descript {
        position: relative;
        top: 15px;
        padding-left: 25px;
    }

    #fTable {
        border-collapse: collapse;
        position: relative;
        top: -18px;
    }

    .all {
        border-style: solid none solid none;
        border-color: #999999;
    }

    .left {
        border-style: solid none solid solid;
        border-color: #999999;
    }

    .right {
        border-style: solid solid solid none;
        border-color: #999999;
    }


    #sTable {
        border-collapse: collapse;
        border-style: solid;
        border-color: #999999;
        position: relative;
        top: 35px;
    }

    #sTable td {
        border-style: none none solid none;
        border-color: #999999;
    }

    p {
        text-align: center;
    }

    p.text {
        font-size: 31.7px;
        text-align: left;
        line-height: 19.5px;
    }

    as {
        font-size: 19px;
        position: relative;
        top: 0px;
        left: 9px;
    }

    as2 {
        font-size: 20px;
        position: relative;
        top: 7px;
        right: 20px;
    }

    as3 {
        position: relative;
        left: 5px;
    }

    li {
        font-size: 12px;
    }

    d {

        font-size: 18px;
        font-weight: 450;
        position: relative;
        left: 0px;
        bottom: 0px;
        color: #E3EBF2;


    }

    .box1 {
        height: 60px;
        width: 137.5px;
        background-color: #696969;
        position: relative;
        top: 0px;
        left: 0px;
        border-radius: 10px;
    }

    .box2 {
        height: 60px;
        width: 137.5px;
        background-color: #B8860B;
        position: relative;
        top: 0px;
        left: 0px;
        border-radius: 10px;
    }

    .box3 {
        height: 60px;
        width: 137.5px;
        background-color: #3770C6;
        position: relative;
        top: 0px;
        left: 0px;
        border-radius: 10px;
    }

    .box4 {
        height: 60px;
        width: 137.5px;
        background-color: #008000;
        position: relative;
        top: 0px;
        left: 0px;
        border-radius: 10px;
    }

    .box5 {
        height: 60px;
        width: 137.5px;
        background-color: #008000;
        position: relative;
        top: 0px;
        left: 0px;
        border-radius: 10px;
    }

    #dsad {
        transform: scale(1.15, 1.15);
        position: absolute;
        left: 570px;
        top: 70px;
    }

    dsa {
        position: relative;
        font-size: 20px;
        left: -189px;
        top: 100px;
        font-size: 110%;
    }

    sad {
        font-size: 20px;
        position: relative;
        text-align: left;
        left: 220px;
        top: -60px;
        font-size: 110%;
    }
</style>