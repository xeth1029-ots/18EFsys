<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_02_001.aspx.vb" Inherits="WDAIIP.SYS_02_001" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>帳號組織設定</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;帳號維護管理&gt;&gt;帳號組織設定</asp:Label>
                </td>
            </tr>
        </table>
        <%--<table class="font" width="100%">
		<tr>
			<td class="font">
				首頁&gt;&gt;系統管理&gt;&gt;帳號維護管理&gt;&gt;帳號組織設定
			</td>
		</tr>
	</table>--%>
        <br>
        <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td class="bluecol" style="width: 20%">階層
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="ShowLev" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">角色
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="ShowRole" runat="server" AutoPostBack="True">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">帳號
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="ShowAcc" runat="server" AutoPostBack="True">
                    </asp:DropDownList>
                    ※說明:帳號組織建立,須由職務低層往上建立
                </td>
            </tr>
            <tr>
                <td class="bluecol">&nbsp;
                </td>
                <td class="whitecol" align="left">帳號
                </td>
            </tr>
            <tr>
                <td class="bluecol">&nbsp;
                </td>
                <td class="whitecol">
                    <asp:CheckBoxList ID="acclist" runat="server" CssClass="font">
                    </asp:CheckBoxList>
                    &nbsp;
                </td>
            </tr>
        </table>
        <table width="100%">
            <tr>
                <td class="whitecol" align="center">
                    <asp:Button ID="but_sub" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
