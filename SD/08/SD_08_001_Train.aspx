<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_08_001_Train.aspx.vb" Inherits="WDAIIP.SD_08_001_Train" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_08_001_Train</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../style.css" type="text/css" rel="stylesheet">
    <script src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script>
        function ReturnMyValue(MyValue, MyText) {
            ReStartSec();
            parent.document.getElementById(getParamValue('TextField')).value = MyValue;
            parent.document.getElementById(getParamValue('FrameId')).style.display = 'none';
        }

        var TimerID1;
        var sec;

        function HidFrame() {
            if (parent.document.getElementById(getParamValue('FrameId')).style.display == 'inline') {
                ReStartSec();
                TimerID1 = setInterval("GetHidTime()", 100)		//啟動計時器
            }
        }

        function GetHidTime(num) {
            //程式內容
            if (sec < 8) {
                sec++;
            }
            else {
                parent.document.getElementById(getParamValue('FrameId')).style.display = 'none';
                ReStartSec();
            }
        }

        function ReStartSec() {
            sec = 0;
            clearInterval(TimerID1);
        }
    </script>
</head>
<body topmargin="2" leftmargin="2" onmouseover="ReStartSec();" onmouseout="HidFrame();">
    <form id="form1" method="post" runat="server">
        <asp:DataGrid ID="DataGrid1" runat="server" AutoGenerateColumns="False" ShowHeader="False" Width="100%"
            CssClass="font" BorderColor="Red">
            <Columns>
                <asp:BoundColumn DataField="LTCName"></asp:BoundColumn>
            </Columns>
        </asp:DataGrid>
    </form>
</body>
</html>
