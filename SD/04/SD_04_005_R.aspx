<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_04_005_R.aspx.vb" Inherits="WDAIIP.SD_04_005_R" %>
 

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>列印時間配當即預定進度表</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        function GETvalue() {
            document.getElementById('Button3').click();
        }

        function choose_class() {
            var RID = document.form1.RIDValue.value;
            openClass('../02/SD_02_ch.aspx?RID=' + RID, 'Class');
        }

        function ReportPrint() {
            if (document.form1.OCIDValue1.value == '') {
                alert('請先選擇班級職類!');
                return false;
            }
            return true;
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;甄試及錄取&gt;&gt;列印時間配當及預定進度表</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table id="Table2" width="100%" class="table_nw" cellpadding="1" cellspacing="1">
                        <tr>
                            <td class="bluecol" width="20%">訓練機構</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="64%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" size="1">
                                <input type="button" value="..." id="Button2" name="Button2" runat="server">
                                <asp:Button ID="Button3" Style="display: none" runat="server" Text="Button3"></asp:Button>
                                <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()"><asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="20%">班別/職類</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="32%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="32%"></asp:TextBox>
                                <input type="button" value="..." onclick="choose_class()">
                                <input id="OCIDValue1" type="hidden" name="Hidden2" runat="server" size="1">
                                <input id="TMIDValue1" type="hidden" name="Hidden1" runat="server" size="1">
                                <span id="HistoryList" style="position: absolute; display: none; left: 32%"><asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table></span>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td colspan="2" class="whitecol">
                                <p align="center">
                                    <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>&nbsp;
                                    <asp:Button ID="Button4" runat="server" Text="列印封面" CssClass="asp_Export_M"></asp:Button>
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