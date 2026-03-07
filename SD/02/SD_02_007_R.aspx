<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_02_007_R.aspx.vb" Inherits="WDAIIP.SD_02_007_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>列印報到名冊</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        function GETvalue() { document.getElementById('Button13').click(); }

        function choose_class() { openClass('SD_02_ch.aspx'); }

        function ReportPrint() {
            var msg = ''
            if (document.form1.OCIDValue1.value == '') msg += '請選擇職類!\n'
            if (!isChecked(document.form1.SelResult)) msg += '請選擇錄取總類!\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
            return true;
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <%--<asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;甄試及錄取&gt;&gt;列印報到名冊</asp:Label>--%>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;列印報到名冊</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch" id="Table3" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" style="width: 20%">訓練機構</td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="Button8" type="button" value="..." name="Button8" runat="server" class="button_b_Mini" />
                                <asp:Button ID="Button13" Style="display: none" runat="server" Text="Button13"></asp:Button>
                                <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">職類/班別</td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="button_b_Mini" />
                                <input id="OCIDValue1" type="hidden" name="Hidden1" runat="server" />
                                <input id="TMIDValue1" type="hidden" name="Hidden2" runat="server" />
                                <span id="HistoryList" style="position: absolute; display: none; left: 28%">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">錄取</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="SelResult" runat="server" CssClass="font" RepeatDirection="Horizontal"></asp:RadioButtonList></td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <br />
        <div class="whitecol" align="center" style="width: 100%">
            <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button></div>
    </form>
</body>
</html>
