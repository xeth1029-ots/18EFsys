<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Chart54_9.aspx.vb" Inherits="WDAIIP.Chart54_9" %>

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

            vdata2 = vdata2.map(function (element, index, array) { return parseFloat(element); });
            vdata3 = vdata3.map(function (element, index, array) { return parseFloat(element); });
            vdata4 = vdata4.map(function (element, index, array) { return parseFloat(element); });

            //debugger;
            //產投：指標9_三年補助使用情形
            //creatChartA_9('VisualChartA_9', vdata1, vdata2, vdata3, vdata4, vdata5, vdata6, vdata7, vdata8);
            creatChartA_9('VisualChartA_9', vdata1, vdata2, vdata3, vdata4);
            /*
            creatDountChartA_9('VisualChartA_9',
                //["北基宜花金馬分署","桃竹苗分署","中彰投分署","雲嘉南分署","高屏澎東分署"],
                [[0.1, 1.3, 53.02, 1.4, 0.88], //三年內第一次使用
                [1.02, 7.36, 0.35, 0.11, 0.1],	//第二~五次使用
                [6.2, 0.29, 0.27, 0.47, 0.47]]	//第六次(含)以上使用
                );
            */
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
        <div class="wrap">
            <div id="VisualChartA_9"></div>
        </div>
    </form>
</body>
</html>
