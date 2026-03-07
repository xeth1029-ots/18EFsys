<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="VisualChart.aspx.vb" Inherits="WDAIIP.VisualChart" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>首頁</title>
    <meta charset="utf-8" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="css/style.css" type="text/css" rel="stylesheet" />
    <link href="css/homebase.css" type="text/css" rel="stylesheet" />
    <link href="css/jquery-confirm.min.css" rel="stylesheet" />
    <link href="css/bootstrap3-3-6.min.css" rel="stylesheet" />
    <link href="css/bootstrap-treeview.css" rel="stylesheet" />
    <link href="css/font-awesome-4.7.0.min.css" rel="stylesheet" />
    <link href="css/font-awesome.css" rel="stylesheet" />
    <link href="css/font-awesome.min.css" rel="stylesheet" />
    <script type="text/javascript" src="Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript" src="Scripts/highcharts.js"></script>
    <script type="text/javascript" src="Scripts/exporting.js"></script>
    <script type="text/javascript" src="Scripts/export-data.js"></script>
    <script type="text/javascript" src="Scripts/accessibility.js"></script>
    <script type="text/javascript" src="js/SetVisualChart.js"></script>

    <script type="text/javascript">
        $(document).ready(function () {

            //jquery array push
            //var Hid_data1= $('#Hid_data1');
            //fruits.push("Kiwi");
            var vdata1 = $('#Hid_data1').val().split(",");
            var vdata2 = $('#Hid_data2').val().split(",");
            var vdata3 = $('#Hid_data3').val().split(",");
            var vdata4 = $('#Hid_data4').val().split(",");
            var vdata5 = $('#Hid_data5').val().split(",");
            vdata1 = vdata1.map(function (element, index, array) { return parseInt(element, 10); });
            vdata2 = vdata2.map(function (element, index, array) { return parseInt(element, 10); });
            vdata3 = vdata3.map(function (element, index, array) { return parseInt(element, 10); });
            vdata4 = vdata4.map(function (element, index, array) { return parseFloat(element); });
            //debugger;
            //自辦在職：指標1_各分署辦理訓練人次統計
            creatChartB_1('VisualChartB_1', vdata1, vdata2, vdata3, vdata4, vdata5);

            //自辦在職：指標1_各分署辦理訓練人次統計
            //creatChartB_1('VisualChartB_1', 
            //[100, 150, 120, 110, 100, 200, 150, 130, 100, 140], //訓練目標人數
            //[60, 70, 62, 55, 45, 115, 72, 68, 92, 88],	//開訓人數
            //[52, 38, 45, 29, 45, 96, 63, 67, 82, 54],	//結訓人數
            //[0.22, 0.33, 0.66, 0.44, 0.36, 0.68, 0.99, 0.86, 0.57, 0.66],	//訓練達成率
            ////[0.12, 0.23, 0.56, 0.34, 0.26, 0.77, 0.88, 0.66, 0.47, 0.56]	//錄取率
            //);

            var vdata6 = $('#Hid_data6').val().split(",");
            var vdata7 = $('#Hid_data7').val().split(",");
            var vdata8 = $('#Hid_data8').val().split(",");
            var vdata9 = $('#Hid_data4').val().split(",");
            var vdata5 = $('#Hid_data5').val().split(",");
            vdata1 = vdata1.map(function (element, index, array) { return parseInt(element, 10); });
            vdata2 = vdata2.map(function (element, index, array) { return parseInt(element, 10); });
            vdata3 = vdata3.map(function (element, index, array) { return parseInt(element, 10); });
            vdata4 = vdata4.map(function (element, index, array) { return parseFloat(element); });


            //自辦在職：指標2_各分署辦理訓練班次統計
            creatChartB_1('VisualChartB_2', vdata1, vdata2, vdata3, vdata4, vdata5);

            creatChart('container1',
           [15000, 7300, 5000, 3000, 4000], //目標人數
           [14000, 6000, 4000, 2000, 3000],	//開訓人數
           [12000, 5000, 3000, 1000, 2000],	//結訓人數
           [0.499, 0.715, 0.80, 0.902, 0.95],	//開訓率
           [0.309, 0.605, 0.704, 0.802, 0.85]	//結訓率
           );

            createMFChart('container2',
                        [								//男 14歲以下~65歲以上
                            -2.2, -2.1, -2.2, -2.4,
                            -2.7, -3.0, -3.3, -3.2,
                            -2.9, -3.5, -4.4, -4.1
                        ], [								//女 14歲以下~65歲以上
                            2.1, 2.0, 2.1, 2.3,
                            2.9, 3.2, 3.1, 2.9,
                            4.3, 4.0, 3.5, 2.9
                        ]
            );

            creatDountChart('container3',
                        [[0.1, 1.3, 53.02, 1.4, 0.88], //第一次使用 			北基~高屏 百分比
                        [1.02, 7.36, 0.35, 0.11, 0.1],	//第二~五次使用			北基~高屏 百分比
                        [6.2, 0.29, 0.27, 0.47, 0.47]]	//第六次(含)以上使用	北基~高屏 百分比
            );

            creatStackChart('container4',
            [5, 3, 4, 7, 2],	//三年內使用經費已達八成以上的學員數
            [3, 4, 4, 2, 5],	//三年內使用經費未達八成的學員數
            [2, 5, 6, 2, 1],	//一年內使用經費已達八成以上的學員數
            [3, 0, 4, 4, 3]		//一年內使用經費未達八成的學員數
            );

        }
       );
    </script>

</head>
<body>
    <form id="form1" method="post" runat="server">
        <input id="Hid_data1" type="hidden" runat="server" />
        <input id="Hid_data2" type="hidden" runat="server" />
        <input id="Hid_data3" type="hidden" runat="server" />
        <input id="Hid_data4" type="hidden" runat="server" />
        <input id="Hid_data5" type="hidden" runat="server" />
        <input id="Hid_data6" type="hidden" runat="server" />
        <input id="Hid_data7" type="hidden" runat="server" />
        <input id="Hid_data8" type="hidden" runat="server" />
        <input id="Hid_data9" type="hidden" runat="server" />
        <input id="Hid_data10" type="hidden" runat="server" />
        <input id="Hid_data11" type="hidden" runat="server" />
        <input id="Hid_data12" type="hidden" runat="server" />

        <div class="wrap">
            <div id="VisualChartB_1"></div>

            <div id="container1"></div>

            <div id="container2"></div>

            <div id="container3"></div>

            <div id="container4"></div>
        </div>
    </form>
</body>
</html>
