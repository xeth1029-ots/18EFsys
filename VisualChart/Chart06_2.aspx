<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Chart06_2.aspx.vb" Inherits="WDAIIP.Chart06_2" %>

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
            vdata1 = vdata1.map(function (element, index, array) { return parseInt(element, 10); });
            vdata2 = vdata2.map(function (element, index, array) { return parseInt(element, 10); });
            vdata3 = vdata3.map(function (element, index, array) { return parseFloat(element); });
            //vdata4 = vdata4.map(function (element, index, array) { return parseFloat(element); });
            //debugger;
            //自辦在職：指標2_各分署辦理訓練班次統計
            creatChartB_2('VisualChartB_2', vdata1,vdata2,vdata3,vdata4,vdata5);
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
        <div class="wrap">
            <div id="VisualChartB_2"></div>
        </div>
    </form>
</body>
</html>
