<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_01_002_att.aspx.vb" Inherits="WDAIIP.SYS_01_002_att" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>帳號計畫賦予-委訓單位歸屬</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function SelectAll(flag, num) {
            //var mytable=document.getElementById('OrgList');
            for (var i = 0; i < num; i++) {
                var mycheckbox = document.getElementById('OrgList_' + i);
                mycheckbox.checked = flag;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">

        <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr class="head_navy" align="left">
                <td style="width: 20%">分署</td>
                <td style="width: 80%">委訓單位
						<asp:CheckBox ID="CheckBox1" runat="server" Text="全選"></asp:CheckBox></td>
            </tr>
            <tr>
                <td></td>
                <td>
                    <asp:CheckBoxList ID="OrgList" runat="server" CssClass="font" RepeatColumns="3"></asp:CheckBoxList>
                    <asp:Label ID="msg2" runat="server"></asp:Label></td>
            </tr>
            <tr>
                <td colspan="2" align="center" class="whitecol">
                    <input id="Button1" type="button" value="回上一頁" name="Button1" runat="server" class="asp_button_M">
                    <asp:Button ID="but_add" runat="server" Text="儲存" class="asp_button_M"></asp:Button></td>
            </tr>
        </table>
    </form>
</body>
</html>
