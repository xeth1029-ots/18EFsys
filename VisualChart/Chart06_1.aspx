<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Chart06_1.aspx.vb" Inherits="WDAIIP.Chart06_1" %>

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
    <%--<script>
        function resize() {
            parent.document.parentElement.height = document.body.scrollHeight;
        }  //將子頁面高度傳到父頁面框架}
    </script>--%>

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

        <div id="VisualChartB_1" style="height: 100%;"></div>
    </form>
</body>
</html>
