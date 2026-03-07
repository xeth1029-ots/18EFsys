<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_003_import.aspx.vb"
    Inherits="WDAIIP.TC_01_003_import" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>班別代碼設定-匯入</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="table_sch" cellpadding="1" cellspacing="1" width="100%">
            <tr>
                <td class="bluecol" width="25%">來源年度
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="Year" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">目的年度
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="Year2" runat="server">
                    </asp:DropDownList>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td align="center" class="whitecol" colspan="2">
                    <asp:Button ID="Button1" runat="server" Text="執行匯入" CssClass="asp_Export_M"></asp:Button>
                    &nbsp;&nbsp;<asp:Button ID="Btn_Back1" runat="server" Text="回上一頁" CssClass="asp_button_M" />
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
