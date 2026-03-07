<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_04_015.aspx.vb" Inherits="WDAIIP.SD_04_015" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>班級課程師資授課時數統計</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        function GETvalue() {
            document.getElementById('Button3').click();
        }
        function choose_class() {
            var RIDValue = document.getElementById('RIDValue');
            openClass('../02/SD_02_ch.aspx?RID=' + RIDValue.value, 'Class');
        }

        function check_data() {
            var OCIDValue1 = document.getElementById('OCIDValue1');
            var msg = '';
            if (OCIDValue1.value == '') msg += '請選擇職類班別\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;教師資料管理&gt;&gt;班級課程師資授課時數統計</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_nw" id="Table3" cellpadding="1" cellspacing="1">
                        <tr>
                            <td class="bluecol" style="width: 20%">訓練機構
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                <input id="Button8" type="button" value="..." name="Button8" runat="server">
                                <asp:Button ID="Button3" Style="display: none" runat="server"></asp:Button>
                                <asp:Button ID="Button13" Style="display: none" runat="server" Text="Button13"></asp:Button>
                                <span id="HistoryList2" style="display: none; position: absolute" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">職類/班級
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input id="Button5" type="button" value="..." name="Button5" runat="server" onclick="choose_class()">
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                <span id="HistoryList" style="display: none; left: 270px; position: absolute">
                                    <asp:Table ID="HistoryTable" runat="server" Width="310">
                                    </asp:Table>
                                </span>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td colspan="2" class="whitecol" align="center">
                                <asp:Button ID="btnPrint1" runat="server" Text="列印教師" CssClass="asp_Export_M"></asp:Button>
                                <asp:Button ID="btnPrint2" runat="server" Text="列印第二教師1" CssClass="asp_Export_M"></asp:Button>
                                <asp:Button ID="btnPrint3" runat="server" Text="列印第二教師2" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
