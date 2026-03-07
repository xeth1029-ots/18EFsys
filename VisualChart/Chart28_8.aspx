<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Chart28_8.aspx.vb" Inherits="WDAIIP.Chart28_8" %>

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
            vdata1 = vdata1.map(function (element, index, array) { return parseInt(element, 10); });
            vdata2 = vdata2.map(function (element, index, array) { return parseInt(element, 10); });

            //debugger;
            //自辦在職：指標8_參訓年齡/性別分佈
            creatChartA_8('VisualChartA_8', vdata1, vdata2, vdata3);
            //createMFChartB_3('VisualChartB_3',
            //    [-10, -20, -30, -50, -90, -70, -60, -33, -26, -10, -6, -2],
            //    [8, 26, 22, 96, 102, 156, 88, 64, 24, 15, 3, 0]);

        }
       );
    </script>

</head>
<body>
    <form id="form1" method="post" runat="server">
        <input id="Hid_data1" type="hidden" runat="server" />
        <input id="Hid_data2" type="hidden" runat="server" />
        <input id="Hid_data3" type="hidden" runat="server" />
        <div class="wrap">
            <div id="VisualChartA_8"></div>
        </div>
    </form>
</body>
</html>
