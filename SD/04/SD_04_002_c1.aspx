<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_04_002_c1.aspx.vb" Inherits="WDAIIP.SD_04_002_c1" %>

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title></title>
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
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
</head>
<body>
    <form id="Form1" method="post" runat="server">
        <table cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid4" runat="server" CssClass="font" BorderColor="Black" AutoGenerateColumns="True" Width="100%">
                                    <ItemStyle BackColor="#ECF7FF" HorizontalAlign="Center"></ItemStyle>
                                    <HeaderStyle CssClass="head_navy" />
                                </asp:DataGrid>
                            </td>
                        </tr>
                    </table>
                    <br />
                    <center>
                        <asp:Button Style="z-index: 0" ID="btnReNew" runat="server" Text="重新查詢" CssClass="asp_button_M"></asp:Button></center>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
