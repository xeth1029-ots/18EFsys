<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_03_Adm2.aspx.vb" Inherits="WDAIIP.TC_03_Adm2" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>新增營業稅</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function check_date() {
            var msg = '';
            if (document.form1.TaxGrant.value == '') msg += '請輸入百分比\n';
            else if (!isUnsignedInt(document.form1.TaxGrant.value)) msg += '百分比必須為整數\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="table_nw" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td class="head_navy">項目</td>
                <td class="head_navy">百分比</td>
                <td class="head_navy">功能</td>
            </tr>
            <tr>
                <td class="whitecol">
                    <asp:CheckBoxList ID="TaxFlag" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow" RepeatColumns="3"></asp:CheckBoxList></td>
                <td class="whitecol">*<asp:TextBox ID="TaxGrant" runat="server" Columns="5" Width="50%">5</asp:TextBox>%</td>
                <td class="whitecol">
                    <asp:Button ID="Button1" runat="server" Text="送出" CssClass="asp_button_M"></asp:Button></td>
            </tr>
        </table>
        <asp:HiddenField ID="Hid_CostItem_GUID1" runat="server" />
    </form>
</body>
</html>
