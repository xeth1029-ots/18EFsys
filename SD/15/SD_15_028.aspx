<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_028.aspx.vb" Inherits="WDAIIP.SD_15_028" %>

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>訓練單位辦訓情形</title>
    <link rel="stylesheet" type="text/css" href="../../css/style.css" />
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;統計表&gt;&gt;訓練單位辦訓情形</asp:Label>
                </td>
            </tr>
        </table>
        <table width="100%" cellpadding="1" cellspacing="1" class="table_sch">
            <tr>
                <td class="bluecol_need" width="20%">訓練計畫</td>
                <td colspan="3" class="whitecol" width="80%">
                    <asp:DropDownList ID="TPlanlist1" runat="server"></asp:DropDownList>
                </td>
            </tr>

            <tr>
                <td class="bluecol_need" width="20%">計畫年度</td>
                <td colspan="3" class="whitecol" width="80%">
                    <asp:DropDownList ID="yearlist" runat="server"></asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">匯出檔案格式</td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                        <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                        <asp:ListItem Value="ODS">ODS</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
        </table>
        <table width="100%">
            <tr>
                <td class="whitecol" align="center">
                    <asp:Button ID="bt_EXPORT" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                    <%--<asp:Button ID="bt_DOWNLOADFILE" runat="server" Text="檔案下載" CssClass="asp_Export_M"></asp:Button>--%>
                    <div align="center"></div>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                    </div>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="Hid_YEAR_ROC" runat="server" />
    </form>
</body>
</html>
