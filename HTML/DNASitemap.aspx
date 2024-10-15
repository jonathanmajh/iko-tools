<!DOCTYPE html>

<head>
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" integrity="sha256-p4NxAoJBhIIN+hmNHrzRCf9tD/miZyoHS5obTRR9BMY=" crossorigin="" />
    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js" integrity="sha256-20nQCchB9co0qIjJZRGuk2/Z9VM+kNiyxNV1lvTlZBo=" crossorigin=""></script>

    <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.css" />
    <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.Default.css" />
    <script src="https://unpkg.com/leaflet.markercluster@1.5.3/dist/leaflet.markercluster-src.js"></script>

    <title>DNA Map</title>

    <style>
        body {
            /*            hides the scroll bars*/
            overflow-y: hidden;
            overflow-x: hidden;
            /*background-color: #d3d3d3;*/
            background-color: white;
        }

        img {
            width: 250px;
            height: 75px;
        }

        body,
        html {
            height: 98%;
            /*            99% because 100 makes the bottom slightly cut off*/
        }

        h2 {
            color: #466995;
            font-family: 'Oswald', sans-serif;
            font-size: 13px;
            text-align: left;
            font-weight: 700;
            text-transform: uppercase;
            padding: 1px;
        }

        h3 {
            color: #466995;
            font-family: 'Oswald', sans-serif;
            font-size: 13px;
            text-align: left;
            font-weight: 700;
            text-transform: uppercase;
            padding: 1px;
            margin: 1px;
        }

        #map {
            height: 96%;
            /* take as much as possible */
            border-radius: 10px;
            /*border: 2px black solid;*/
            margin: 1px;
            display: flex;
        }

        .title {
            text-align: center;
            font-size: 2em;
            margin-top: 0px;
            margin-bottom: 0px;
            font-family: sans-serif;
            letter-spacing: .05em;
            font-weight: 400;
            /*border: 2px dodgerblue solid;*/
            display: flex;
            justify-content: center;
            align-items: center;
            position: relative;
            top: -75px;
            /*left: -110px;*/
        }

        .ikologo {
            font-size: 2em;
            font-family: sans-serif;
            letter-spacing: .05em;
            font-weight: 400;
            /*border: 2px dodgerblue solid;*/
            display: flex;
            justify-content: left;
            align-items: left;
        }

        .description {
            text-align: center;
            font-size: 1.5em;
            margin-top: 0px;
            margin-bottom: 0px;
            font-family: sans-serif;
            letter-spacing: .025em;
            font-weight: 200;
            /*border: 2px dodgerblue solid;*/
            display: flex;
            justify-content: center;
            align-items: center;
            position: relative;
            top: -75px
                /* deleted container */
        }

        .container {
            display: grid;
            width: 100%;
            height: 100%;
            /* this will be treated as a Minimum!
                                     It will stretch to fit content */
            /*background-color: #d3d3d3;*/
            background-color: white;
            grid-template-columns: 1fr 3fr 1fr;
            grid-template-rows: 1fr 1fr 1fr 90%;
            grid-row-gap: 2px;
            /*border: 2px blue solid;*/
        }


        .legendGreen {
            text-align: center;
            /*border: 2px #466995 solid ;*/
            margin: 1px;
            height: 25%;
            background-color: #31ba2e;
            padding: 5px 10px 6px 10px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            position: relative;
        }

        .legendYellow {
            text-align: center;
            /*border: 2px #466995 solid ;*/
            margin: 1px;
            height: 25%;
            background-color: #cac428;
            padding: 5px 10px 6px 10px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            position: relative;
        }

        .legendGrey {
            text-align: center;
            /*border: 2px #466995 solid ;*/
            margin: 1px;
            height: 25%;
            background-color: #7a7a7a;
            padding: 5px 10px 6px 10px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            position: relative;
        }

        .a {
            grid-area: 1 / 1 / span 3 /span 1;
        }

        .b {
            grid-area: 1 / 2 / span 2 /span 1;
        }

        .c {
            grid-area: 3 / 2 / span 1 / span 1;
        }

        .da {
            grid-area: 1 / 3 / span 1 / span 1;
            display: flex;
            justify-content: flex-end;
            align-items: center;
            position: relative;
            top: -130px;
            /*left: -275px; */
        }

        .db {
            grid-area: 2 / 3 / span 1 / span 1;
            display: flex;
            justify-content: flex-end;
            align-items: center;
            position: relative;
            top: -150px;
            /*left: -150px */
        }

        .dc {
            grid-area: 3 / 3 / span 1 / span 1;
            display: flex;
            justify-content: flex-end;
            align-items: center;
            position: relative;
            top: -170px
        }

        .e {
            grid-area: 4 / 1 / span 1 / span 3;
            position: relative;
            top: -180px
        }

        .leaflet-popup-tip,
        .leaflet-popup-content-wrapper {
            font-size: 20px;
            /*          double the size of the font so its more visible on mobile devices
                    background: #e93434;
                    color: #ffffff;
        */
        }
    </style>
</head>

<body>

    <div class="ikologo a">
        <img src="/support/Reliability/ReliabilityShared/PublishingImages/SitePages/Home/IKOLogo.png">
        <!--image file path -->
    </div>
    <div class="title b"><b>DNA Map</b></div>
    <div class="description c">Map of All Manufacturing Sites with DNA Implemented</div>
    <div class="da">
        <div>
            <h2>DNA Completed</h2>
        </div>
        <div>
            <h3 class="legendGreen"></h3>
        </div>
    </div>
    <div class="db">
        <div>
            <h2>DNA In Progress</h2>
        </div>
        <div>
            <h3 class="legendYellow"></h3>
        </div>
    </div>
    <div class="dc">
        <div>
            <h2>DNA Not Implemented</h2>
        </div>
        <div>
            <h3 class="legendGrey"></h3>
        </div>

    </div>

    <div id="map" class="e"></div>

    <script>

        var greenIcon = new L.Icon({
            iconUrl: '/support/Reliability/ReliabilityShared/Pages/green.png',
            shadowUrl: 'https://cdnjs.cloudflare.com/ajax/libs/leaflet/0.7.7/images/marker-shadow.png',
            iconSize: [25, 41],
            iconAnchor: [12, 41],
            popupAnchor: [1, -34],
            shadowSize: [41, 41]
        });

        var greyIcon = new L.Icon({
            iconUrl: '/support/Reliability/ReliabilityShared/Pages/grey.png',
            shadowUrl: 'https://cdnjs.cloudflare.com/ajax/libs/leaflet/0.7.7/images/marker-shadow.png',
            iconSize: [25, 41],
            iconAnchor: [12, 41],
            popupAnchor: [1, -34],
            shadowSize: [41, 41]
        });

        var yellowIcon = new L.Icon({
            iconUrl: '/support/Reliability/ReliabilityShared/Pages/yellow.png',
            shadowUrl: 'https://cdnjs.cloudflare.com/ajax/libs/leaflet/0.7.7/images/marker-shadow.png',
            iconSize: [25, 41],
            iconAnchor: [12, 41],
            popupAnchor: [1, -34],
            shadowSize: [41, 41]
        });

        var siteURL = {
            "IKO Calgary": ["http://operations.connect.na.local/support/Reliability/IKOCalgary/CalgaryAssetDocuments/IKO_CALGARY_DNA.pdf", 51.0157632, -114.0270971, greenIcon], //1600-42nd Ave. S.E. Calgary,Alberta T2G 5B5
            "IKO Kankakee": ["http://operations.connect.na.local/support/Reliability/IKOKankakee/KankakeeAssetDocuments/Kankakee_Plant_Layout.pdf", 41.087923, -87.8753309, greenIcon], //235 West South Tec Drive, Kankakee, IL. 60901
            "IKO Wilmington": ["http://operations.connect.na.local/support/Reliability/IKOWilmington/PlantAssetDocs/Wilmington_Plant_Layout.pdf", 39.7530151, -75.5001767, greenIcon], //120 Hay Road Wilmington, DE. 19809
            "IKO Sumas": ["http://operations.connect.na.local/support/Reliability/IKOSumas/Sumas%20Asset%20Documents/Sumas%20Main%20Plant%20Overview.pdf", 48.9939495, -122.2834137, greenIcon], //850 West Front Street, Sumas, WA 98295
            "IKO Hawkesbury": ["http://operations.connect.na.local/support/Reliability/IKOHawkesbury/HawkesburyAssetDocuments/Plant_Layout.pdf", 45.5955576, -74.5947497, greenIcon], //1451 Spence Avenue Hawkesbury, Ontario K6A 3T4
            "IKO Hillsboro (Southwest)": ["http://operations.connect.na.local/support/Reliability/IKOSouthwest/PlantAssetDocs/Plant_Layout.pdf", 32.0350965, -97.0995177, greenIcon], //1001 IKO Way, Hillsboro, TX  76645
            "IKO Slovakia": ["http://operations.connect.na.local/support/Reliability/IKOSlovakia/PlantAssetDocs/SE-Level-01.pdf", 48.6670758, 17.3565852, greenIcon], //KaplinskÃ¯Â¿Â½ Pole 16, 905 01 Senica, Slovakia
            "IG Brampton": ["http://operations.connect.na.local/support/Reliability/IGBrampton/IGBramptonAssetDocuments/IG_Plant_Layout_First_Level.pdf", 43.6936032, -79.7383604, greenIcon], //87 Orenda Road, Brampton, ON L6W 1V7
            "IKO Sylacauga": ["http://operations.connect.na.local/support/Reliability/IKOSylacauga/SylacaugaAssetDocuments/IKO_Southeast_Plant_Layout.pdf", 33.1648712, -86.307571, greenIcon], //1708 Sylacauga Fayetteville Highway, Sylacauga, AL 35151
            "Bramcal": ["http://operations.connect.na.local/support/Reliability/Bramcal/Plant%20Asset%20Documents/Plant_Layout.pdf", 44.3520829, -79.6680303, greenIcon], //400 Huronia Road, Barrie, Ontario, L49 8Y9
            "IKO Ashcroft": ["http://operations.connect.na.local/support/Reliability/IKOAshcroft/PlantAssetDocs/Ashcroft_Plant_Layout.pdf", 50.7453664, -121.2514044, greenIcon], //P.O. Box 1000, Ashcroft, B.C., V0K 1A0
            "IKO Madoc": ["http://operations.connect.na.local/support/Reliability/IKOMadoc/PlantAssetDocs/Madoc_Plant_Layout.pdf", 44.4985398, -77.5371204, greenIcon], //105084 Hwy 7, Madoc, ON K0K 2K0
            "Maxi-Mix": ["http://operations.connect.na.local/support/Reliability/Maxi-Mix/Plant%20Asset%20Documents/Mortar_and_Grout_Mix_Manufacturing_System.pdf", 43.543959, -79.8798467, greenIcon], //8105 Esquesing Line, Milton, ON, L9T 2X9
            "CRC Toronto": ["http://operations.connect.na.local/support/Reliability/CRCToronto/PlantAssetDocs/CRC_Toronto_PLant_Layout.pdf", 43.6558681, -79.334133, greenIcon], //560 Commissioners St., Toronto, Ontario M4M 1A7
            "IKO Brampton": ["http://operations.connect.na.local/support/Reliability/IKOBrampton/BramptonAssetDocuments/B2_Plant_Overview.pdf", 43.6925284, -79.7444545, greenIcon], //71 Orenda Road Brampton, Ontario L6W 1V8
            "CRC Brampton": ["http://operations.connect.na.local/support/Reliability/CRCBrampton/PlantAssetDocuments/Plant_Overview.pdf", 43.6825432, -79.7258757, greenIcon], //309 Rutherford Road South, Brampton, Ontario, L6W 3R4
            //["IKO Ingersoll", "https://www.iko.com", 43.0975945, -80.8928448], //355120 35th Line, Ingersoll ON N0J 1J0 *DELETED
            "IG High River": ["http://operations.connect.na.local/support/Reliability/IGHighRiver/Plant%20Asset%20Documents/H-R-DNA_OVERVIEW.pdf", 50.577680567664146, -113.85059714634718, greenIcon], //H592+JH High River, Alberta
            "Blair Rubber": ["http://blairrubber.com/", 41.027703, -81.8729277, yellowIcon], //5020 Enterprise Parkway Seville, OH 44273 
            //["Hyload", "http://hyload.com/", 41.0017634, -81.7514049], //9976 Rittman Rd, Wadsworth, OH 44281 *DELETED
            "Appley Bridge": ["http://operations.connect.na.local/support/Reliability/AppleyBridge/PlantAssetDocs/Appley_Plant_Layout.pdf", 53.584967, -2.7232832, greenIcon], //Appley Lane North, Appley Bridge, Wigan,  WN6 9AB
            "IKO PLC (Grangemill)": ["http://www.ikogroup.co.uk", 53.1137388, -1.6367977, greyIcon], //IKO PLC, Water Lane,Grangemill, Matlock, Derbyshire, DE4 4BW
            //["Pure Asphalt Company Limited", "http://www.ikogroup.co.uk", 53.568467, -2.4137798], //Pure Asphalt Company Limited, Burnden Works, Burnden Road, Bolton, Lancashire, BL3 2RD *DELETED
            "IKO Polymeric (Clay Cross)": ["http://www.ikogroup.co.uk", 53.1679258, -1.3989654, greyIcon], //Coney Green Road, Clay Cross, Chesterfield, S45 9HZ
            "IKO Alconbury": ["http://operations.connect.na.local/support/Reliability/Alconbury/PlantAssetDocuments/Alconbury_Plant_Layout.pdf", 52.3794161, -0.2414008, greenIcon],
            "Tenco": ["http://www.tenco.nl", 52.4548094, 4.8126184, greyIcon], //Touwen \u0026 Co B.V.  Oostzijde 300 1508 ET Zaandam
            "IKO Enertherm [Klundert,Netherlands]": ["http://operations.connect.na.local/support/Reliability/Klundert/PlantAssetDocuments/home.pdf", 51.6759272, 4.5433087, greenIcon], //Wielewaalweg 1, 4791 PD Klundert  Postbus 45, 4780 AA Moerdijk
            "IKO NV [Antwerp,Belgium]": ["http://operations.connect.na.local/support/Reliability/Antwerp/PlantAssetDocuments/Antwerp_Plant_Overview.pdf", 51.2000273, 4.362123, greenIcon], //IKO NV, u0027Herbouvillekaai 80, 2020 Antwerp, Belgium
            //["AWA Produktions GmbH", "http://www.iko.eu/", 50.7370443, 7.133369], //Maarstr. 48, Bonn, 53227, Germany *DELETED
            "Meple": ["http://www.meple.com/", 49.3171085, 1.0583233, greyIcon],
            "IKO Enertherm [Combronde,France]": ["http://operations.connect.na.local/support/Reliability/Combronde/PlantAssetDocuments/Combronde_Plant_Layout.pdf", 45.997945, 3.0876591, greenIcon], //Rue Allemagne, 63460 Combronde, Frankrijk
            "IKO Hagerstown": ["http://operations.connect.na.local/support/Reliability/Snowman/PlantAssetDocs/Hagerstown_Site_Layout.pdf", 39.6389, -77.7576, greenIcon], //100 Tandy Dr, Hagerstown, MD 21740
            "IKO Chester Mat": ["http://operations.connect.na.local/support/Reliability/IKOChesterMat/Plant%20Asset%20Documents/IKO_Chester_Mat_Plant_Layout.pdf", 34.73433576989333, -81.12535595634944, greenIcon], //SC Highway 9 and Cedarhurst Road, Chester, SC 29706
            "Blair Rubber": ["http://operations.connect.na.local/support/Reliability/BlairRubber/Plant%20Asset%20Documents/Blair_Rubber_Plant_Layout.pdf", 41.0278354504909, -81.87009376719323, yellowIcon], //5020 Enterprise Pkwy, Seville, OH 44273
            //rip hamilton 628 Victoria Ave N, Hamilton, ON L8L 8B3
            "IKO Metals (Ennis)": ["http://operations.connect.na.local/support/Reliability/IKOMetals/Plant%20Asset%20Documents/Ennis_Plant_Layout%20_top_level.pdf", 32.3005685,-96.5952441, yellowIcon], //4400 Sterilite Dr, Ennis, TX 75119, United States
        };

        // create map and set center and zoom level

        var map = L.map('map', {
            center: [53.96, -50.88], //set default zoom and location
            zoom: 3.5
        });

        L.tileLayer('http://{s}.tile.osm.org/{z}/{x}/{y}.png', {
            attribution: '&copy; <a href="http://osm.org/copyright">OpenStreetMap</a> contributors'
        }).addTo(map); //this is the box at the bottom that gives credit, !!DO NOT REMOVE

        var tooltipPopup;
        // console.log(size)

        var markers = L.markerClusterGroup();

        for (var i = 0; i < Object.keys(siteURL).length; i++) {
            key = Object.keys(siteURL)[i]
            marker = L.marker([siteURL[key][1], siteURL[key][2]], { icon: siteURL[key][3] }); // x, y location from array for marker position, also icon color
            marker.bindPopup(key);
            marker.on("click", function (e) { // get name from popup to use in dictionary to lookup plant layout drawing url
                window.open(siteURL[e.target.getPopup().getContent()][0])
            });
            marker.on('mouseout', function (e) {
                e.target.closePopup();
            });
            marker.on('mouseover', function (e) {
                e.target.openPopup();
            });
            markers.addLayer(marker)
        }
        map.addLayer(markers);
    </script>
</body>

</html>