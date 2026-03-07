<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="TC_01_003_print.aspx.vb" Inherits="WDAIIP.TC_01_003_print" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>TC_01_003_print</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
    <%--<table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
        <tr>
            <td>
                <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                            <asp:Label ID="TitleLab2" runat="server">
                            首頁&gt;&gt;訓練機構管理&gt;&gt;基本資料設定&gt;&gt;<font color="#990000"> 班別代碼列印</font>
                            </asp:Label>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>--%>    
    <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
        <tr>
            <td id="td6" runat="server" class="bluecol" style="width:20%">
                班別代碼
            </td>
            <td class="whitecol">
                <asp:TextBox ID="TB_classid" runat="server" Width="20%"></asp:TextBox>
            </td>
        </tr>
    </table>
    <table width="100%">
        <tr>
            <td class="whitecol" align="center">
                <asp:Button ID="print" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>&nbsp;
                <asp:Button ID="Button1" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button></FONT>
            </td>
        </tr>
    </table>
    <br>
    <asp:Panel ID="Panel" runat="server" Visible="False">
        <table class="font" id="search_tbl" cellspacing="0" cellpadding="0" width="100%" border="1" runat="server">
        </table>
    </asp:Panel>
    </form>
</body>
</html>
