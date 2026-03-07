<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_08_004_R.aspx.vb" Inherits="WDAIIP.SD_08_004_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>職業訓練生活津貼統計明細表</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script language="javascript">
        function choose_class() {
            var RID = document.form1.RIDValue.value;
            document.getElementById('OCIDValue').value = '';
            openClass('GetClass.aspx?RID=' + RID + '&OCIDField=OCIDValue');
        }
        function ReportPrint() {
            if (document.getElementById('OCIDValue').value == '') {
                alert('請選擇班級');
                return false;
            }
            return true;
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="Table1" cellspacing="1" cellpadding="1" width="600" border="0">
            <tr>
                <td>
                    <table id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
                        <tr>
                            <td>首頁&gt;&gt;學員動態管理&gt;&gt;職業訓練生活津貼&gt;&gt;<font color="#990000">職業訓練生活津貼統計明細表</font></td>
                        </tr>
                    </table>
                    <table id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
                        <tr>
                            <td bgcolor="#2aafc0" width="100">
                                <font color="#ffffff">&nbsp;&nbsp;&nbsp; 訓練機構</font></td>
                            <td bgcolor="#ecf7ff">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="310px"></asp:TextBox><input type="button" value="..." id="Button1" name="Button1" runat="server"><input id="RIDValue" type="hidden" name="Hidden3" runat="server" size="1"><br>
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="310px"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td bgcolor="#2aafc0">
                                <font color="#ffffff">&nbsp;&nbsp;&nbsp; 職類/班別</font></td>
                            <td bgcolor="#ecf7ff">
                                <input id="OCIDValue" type="hidden" name="OCIDValue" runat="server" size="1"><input onclick="choose_class()" type="button" value="挑選班級"></td>
                        </tr>
                        <tr>
                            <td style="width: 84px" width="84" bgcolor="#ecf7ff"></td>
                            <td bgcolor="#ecf7ff">點選"挑選班級"可複選班別</td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <p align="center">
                                    <asp:Button ID="Button_print" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                                </p>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
