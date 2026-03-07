<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_14_026.aspx.vb" Inherits="WDAIIP.SD_14_026" %>

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>送件檢核表</title>
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
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;表單列印&gt;&gt;送件檢核表</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch">
                        <tr>
                            <td class="bluecol" width="20%">訓練機構</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="center" runat="server" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="Hidden2" runat="server">
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr id="tr_rbl_AppStage_TP28" runat="server">
                            <td class="bluecol_need" width="20%">申請階段</td>
                            <td class="whitecol" width="80%">
                                <asp:RadioButtonList ID="rbl_AppStage" runat="server" Width="100%" CssClass="font" CellSpacing="0" CellPadding="0" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="1">1:上半年</asp:ListItem>
                                    <asp:ListItem Value="2">2:下半年</asp:ListItem>
                                    <asp:ListItem Value="3">3:政策性產業</asp:ListItem>
                                </asp:RadioButtonList></td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td align="center" class="whitecol">
                    <asp:Button ID="btnPrint1" runat="server" Text="列印" CssClass="asp_Export_M" /></td>
            </tr>
        </table>
    </form>
</body>
</html>
