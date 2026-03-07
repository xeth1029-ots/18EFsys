<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_08_001_Unit.aspx.vb" Inherits="WDAIIP.SD_08_001_Unit" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_08_001_Unit</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../style.css" type="text/css" rel="stylesheet">
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
            parent.document.getElementById(getParamValue('TextField')).value = MyValue;
            parent.document.getElementById(getParamValue('FrameId')).style.display = 'none';
            ReStartSec();
        }

        var TimerID1;
        var sec;

        function HidFrame() {
            ReStartSec();
            TimerID1 = setInterval("GetHidTime()", 800)		//啟動計時器
        }

        function GetHidTime(num) {
            //程式內容
            if (sec < 5) {
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
<body onmouseover="ReStartSec();" leftmargin="2" topmargin="2" onload="HidFrame();" onmouseout="HidFrame();">
    <form id="form1" method="post" runat="server">
        <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" BorderColor="Red" CssClass="font" ShowHeader="False"
            AutoGenerateColumns="False">
            <Columns>
                <asp:BoundColumn DataField="LUName"></asp:BoundColumn>
            </Columns>
        </asp:DataGrid><input id="NowValue" type="hidden" runat="server">
        <asp:Button ID="Button1" runat="server" Text="快速搜尋(隱藏)"></asp:Button>
    </form>
</body>
</html>
