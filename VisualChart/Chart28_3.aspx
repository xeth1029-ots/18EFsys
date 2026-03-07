<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Chart28_3.aspx.vb" Inherits="WDAIIP.Chart28_3" %>

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

            vdata2 = vdata2.map(function (element, index, array) { return parseInt(element, 10); });
            vdata3 = vdata3.map(function (element, index, array) { return parseInt(element, 10); });
            vdata4 = vdata4.map(function (element, index, array) { return parseInt(element, 10); });
            vdata5 = vdata5.map(function (element, index, array) { return parseInt(element, 10); });
            vdata6 = vdata6.map(function (element, index, array) { return parseFloat(element); });

            //debugger;
            //產投：指標3_總體補助費指標
            creatChartA_3('VisualChartA_3', vdata1, vdata2, vdata3, vdata4, vdata5, vdata6);
            //creatChartA_3('VisualChartA_3', vdata1,
            //    [1200, 2500, 2200, 1890, 1900, 2000, 2160, 1840, 1860, 2100], //申請補助費
            //    [1100, 2300, 2000, 1780, 1600, 1800, 2000, 1640, 1560, 1600],	//核定補助費
            //    [980, 1800, 1680, 1620, 1350, 1670, 1900, 1350, 1260, 1160],	//預估補助費
            //    [900, 1600, 1600, 1600, 1300, 1520, 1800, 1250, 1160, 1000],	//核撥補助費
            //    [0.499, 0.715, 0.80, 0.902, 0.95, 0.499, 0.515, 0.60, 0.77, 0.95]	//執行率
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

        <div class="wrap">
            <div id="VisualChartA_3"></div>
        </div>
    </form>
</body>
</html>
