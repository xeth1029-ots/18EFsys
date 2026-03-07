<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_004_unit.aspx.vb" Inherits="WDAIIP.TC_01_004_unit" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>班級名稱</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function checkvalue() {
            if (parseInt(getCheckBoxListValue("CheckBoxList1")) == 0) {
                alert('請選擇單元!!');
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <asp:Panel ID="Panel1" runat="server" CssClass="table_nw" Width="100%">
            <p style="background: #f1f9fc">
                班級名稱:
				<asp:Label ID="Label1" runat="server" CssClass="font" Width="100%"></asp:Label>
            </p>
            <p>
                <input id="tb_class_unit" type="hidden" name="tb_class_unit" runat="server"><input id="tb_class_name" type="hidden" name="tb_class_name" runat="server">
            </p>
            <asp:CheckBoxList ID="CheckBoxList1" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="1">
                <asp:ListItem Value="1">1.電腦基本操作</asp:ListItem>
                <asp:ListItem Value="2">2.文書處理及問題演練</asp:ListItem>
                <asp:ListItem Value="3">3.網際網路應用與問題演練</asp:ListItem>
                <asp:ListItem Value="4">4.心理輔導</asp:ListItem>
            </asp:CheckBoxList>
            <p class="whitecol">
                <asp:Button ID="bt_send" runat="server" Text="送出" CssClass="asp_button_M"></asp:Button>
            </p>
        </asp:Panel>
    </form>
</body>
</html>
