<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_027_imp.aspx.vb" Inherits="WDAIIP.TC_01_027_imp" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>師資資料維護-年度複製</title>
    <meta name="ProgId" content="SharePoint.WebPartPage.Document">
    <meta name="WebPartPageExpansion" content="full">
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function chkdata() {
            var year1 = document.getElementById('Fromyear');
            var year2 = document.getElementById('hidToyear');
            var msg = '';
            //if (document.getElementById('Table3').style.display == 'inline' && document.form1.DistID.selectedIndex == 0) msg += '請選擇轄區\n';
            if (document.form1.DistID.selectedIndex == 0) msg += '請選擇轄區\n';
            if (year1.selectedIndex == 0) msg += '請選擇來源年度\n';
            //if (year2.selectedIndex == 0) msg += '請選擇目的年度\n';
            if (year1.value == year2.value) msg += '不能匯入相同的年度!\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <%-- 
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tbody>
                <tr>
                    <td>
                        <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                        <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;開班資料設定&gt;&gt;<font color="#990000">師資資料維護-年度複製</font></asp:Label>
                    </td>
                </tr>
            </tbody>
        </table>
        --%>
        <%--<table class="table_sch" id="Table3" runat="server" cellpadding="1" cellspacing="1"></table>--%>
        <table class="table_sch" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td class="bluecol" width="20%">轄區 </td>
                <td class="whitecol"><asp:DropDownList ID="DistID" runat="server" cellpadding="1" cellspacing="1"></asp:DropDownList></td>
            </tr>
            <tr>
                <td class="bluecol">來源年度 </td>
                <td class="whitecol"><asp:DropDownList ID="Fromyear" runat="server"></asp:DropDownList></td>
            </tr>
            <%--
            <tr>
                <td width="100" class="bluecol">目的年度 </td>
                <td class="whitecol"><asp:DropDownList ID="Toyear" runat="server"></asp:DropDownList></td>
            </tr>
            --%>
            <tr>
                <td class="whitecol" align="center" colspan="2">
                    <asp:Button ID="btnImpYear" runat="server" Text="年度複製" CssClass="asp_button_M"></asp:Button>
                    <asp:Button ID="btnBack1" runat="server" Text="回上頁" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="hidToyear" runat="server" />
    </form>
</body>
</html>