<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Chart28_1.aspx.vb" Inherits="WDAIIP.Chart28_1" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>首頁</title>
    <meta charset="utf-8" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../css/style.css" type="text/css" rel="stylesheet" />
    <link href="../css/homebase.css" type="text/css" rel="stylesheet" />
    <link href="../css/jquery-confirm.min.css" rel="stylesheet" />
    <link href="../css/bootstrap3-3-6.min.css" rel="stylesheet" />
    <link href="../css/bootstrap-treeview.css" rel="stylesheet" />
    <link href="../css/font-awesome-4.7.0.min.css" rel="stylesheet" />
    <link href="../css/font-awesome.css" rel="stylesheet" />
    <link href="../css/font-awesome.min.css" rel="stylesheet" />
    <script type="text/javascript" src="../Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="../Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript" src="../Scripts/highcharts.js"></script>
    <script type="text/javascript" src="../Scripts/exporting.js"></script>
    <script type="text/javascript" src="../Scripts/export-data.js"></script>
    <script type="text/javascript" src="../Scripts/accessibility.js"></script>
    <script type="text/javascript" src="../js/SetVisualChart.js"></script>

    <script type="text/javascript">
        $(document).ready(function () {

            var vdata1 = $('#Hid_data1').val().split(",");
            var vdata2 = $('#Hid_data2').val().split(",");
            var vdata3 = $('#Hid_data3').val().split(",");
            var vdata4 = $('#Hid_data4').val().split(",");
            var vdata5 = $('#Hid_data5').val().split(",");
            var vdata6 = $('#Hid_data6').val().split(",");
            var vdata7 = $('#Hid_data7').val().split(",");
            var vdata8 = $('#Hid_data8').val().split(",");

            vdata2 = vdata2.map(function (element, index, array) { return parseInt(element, 10); });
            vdata3 = vdata3.map(function (element, index, array) { return parseInt(element, 10); });
            vdata4 = vdata4.map(function (element, index, array) { return parseInt(element, 10); });
            vdata5 = vdata5.map(function (element, index, array) { return parseInt(element, 10); });
            vdata6 = vdata6.map(function (element, index, array) { return parseInt(element, 10); });
            vdata7 = vdata7.map(function (element, index, array) { return parseFloat(element); });
            vdata8 = vdata8.map(function (element, index, array) { return parseFloat(element); });

            //debugger;
            //產投：指標1_總體參訓人數指標
            creatChartA_1('VisualChartA_1', vdata1, vdata2, vdata3, vdata4, vdata5, vdata6, vdata7, vdata8);
            //creatChartA_1('VisualChartA_1', vdata1,
            //    [200, 150, 80, 120, 100, 200, 150, 80, 120, 100], //核定人數
            //    [180, 130, 65, 88, 92, 180, 130, 65, 88, 92],	//報名人數
            //    [160, 100, 33, 55, 63, 160, 100, 33, 55, 63],	//開訓人數
            //    [120, 80, 26, 43, 35, 130, 72, 26, 33, 43],	//結訓人數
            //    [80, 70, 13, 26, 24, 76, 56, 22, 15, 20],	//撥款人數
            //    [0.499, 0.715, 0.80, 0.902, 0.95, 0.499, 0.715, 0.80, 0.902, 0.95],	//錄取率
            //    [0.309, 0.605, 0.704, 0.802, 0.85, 0.309, 0.605, 0.704, 0.802, 0.85]	//結訓率
            //    );

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

        <div class="wrap">
            <div id="VisualChartA_1"></div>
        </div>
    </form>
</body>
</html>
