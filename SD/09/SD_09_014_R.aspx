<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_09_014_R.aspx.vb" Inherits="WDAIIP.SD_09_014_R" %>
 

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>授課教師學歷統計表</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function GETvalue() {
            document.getElementById('Button3').click();
        }
        function choose_class() {
            var RID = document.form1.RIDValue.value;
            openClass('../02/SD_02_ch.aspx?RID=' + RID, 'Class');
        }

        function check_data() {
            var msg = '';
            if (document.form1.OCIDValue1.value == '') msg += '請選擇職類班別\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;教務資料管理&gt;&gt;授課教師學歷統計報表</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch" id="Table3" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" width="20%">訓練機構</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini" />
                                <asp:Button ID="Button3" Style="display: none" runat="server" Text="Button3"></asp:Button>
                                <span id="HistoryList2" style="display: none; position: absolute" onclick="GETvalue()"><asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="20%">班別/職類
                            </td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="25%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input type="button" value="..." onclick="choose_class()" class="button_b_Mini" />
                                <input id="OCIDValue1" type="hidden" name="Hidden2" runat="server" />
                                <input id="TMIDValue1" type="hidden" name="Hidden1" runat="server" />
                                <span id="HistoryList" style="display: none; left: 30%; position: absolute"><asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table></span>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <br />
        <div align="center" class="whitecol" style="width: 100%"><asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button></div>
    </form>
</body>
</html>
