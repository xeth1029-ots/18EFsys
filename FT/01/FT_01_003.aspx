<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="FT_01_003.aspx.vb" Inherits="WDAIIP.FT_01_003" %>

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>綜合查詢統計表(定版)</title>
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;定版數據統計表&gt;&gt;綜合查詢統計表(定版)</asp:Label>
                </td>
            </tr>
        </table>
        <table width="100%" cellpadding="1" cellspacing="1" class="table_sch">
            <tr>
                <td class="bluecol_need" width="20%">訓練計畫</td>
                <td colspan="3" class="whitecol" width="80%">
                    <asp:DropDownList ID="TPlanlist1" AutoPostBack="True" runat="server"></asp:DropDownList>
                </td>
            </tr>
            <%--<tr>
                <td class="bluecol_need">分署</td>
                <td colspan="3" class="whitecol">
                    <asp:DropDownList ID="ddlDISTID" runat="server"></asp:DropDownList>
                </td>
            </tr>--%>
            <tr>
                <td class="bluecol_need" width="20%">年度</td>
                <td colspan="3" class="whitecol" width="80%">
                    <asp:DropDownList ID="yearlist" AutoPostBack="True" runat="server"></asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need" width="20%">資料版本</td>
                <td colspan="3" class="whitecol" width="80%">
                    <asp:DropDownList ID="monthlist" AutoPostBack="True" runat="server"></asp:DropDownList>
                    <asp:RadioButtonList ID="rbl_BDATAVER" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                        <asp:ListItem Value="N" Selected="True">當年度</asp:ListItem>
                        <asp:ListItem Value="B">前一年度版</asp:ListItem>
                    </asp:RadioButtonList>版
                </td>
            </tr>

            <%-- <tr>
                <td class="bluecol_need">資料版本</td>
                <td colspan="3" class="whitecol">
                    <asp:RadioButtonList ID="rbl_BDATAVER" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                        <asp:ListItem Value="0301" Selected="True">3月1日</asp:ListItem>
                        <asp:ListItem Value="0401">4月1日</asp:ListItem>
                    </asp:RadioButtonList>版
                </td>
            </tr>--%>
            <%-- <tr>
                <td class="bluecol">匯出檔案格式</td>
                <td colspan="3" class="whitecol">
                    <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                        <asp:ListItem Value="XLSX" Selected="True">EXCEL</asp:ListItem>
                        <asp:ListItem Value="ODS">ODS</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>--%>
        </table>
        <table width="100%">
            <tr>
                <td class="whitecol" align="center">
                    <%--<asp:Button ID="bt_EXPORT" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>--%>
                    <asp:Button ID="bt_DOWNLOADFILE" runat="server" Text="檔案下載" CssClass="asp_Export_M"></asp:Button>
                    <div align="center"></div>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                    </div>
                </td>
            </tr>
        </table>

    </form>
</body>
</html>
