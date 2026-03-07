<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_02_002.aspx.vb" Inherits="WDAIIP.SYS_02_002_" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>職務更換</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;帳號維護管理&gt;&gt;職務更換</asp:Label>
                </td>
            </tr>
        </table>
        <%--<table class="font" width="100%">
				<tr>
					<td class="font">首頁&gt;&gt;系統管理&gt;&gt;帳號維護管理&gt;&gt;職務更換</td>
				</tr>
			</table>--%>
        <br>
        <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td class="bluecol" style="width: 20%">角色</td>
                <td colspan="2" class="whitecol">
                    <asp:DropDownList ID="BefLev" AutoPostBack="True" runat="server"></asp:DropDownList>
                    <asp:RequiredFieldValidator ID="mustBefLev" runat="server" ErrorMessage="請選擇角色" ControlToValidate="BefLev" Display="None"></asp:RequiredFieldValidator></td>
            </tr>
            <tr>
                <td class="bluecol">帳號</td>
                <td colspan="2" class="whitecol">
                    <asp:DropDownList ID="BefAcc" runat="server" AutoPostBack="True"></asp:DropDownList>
                    <asp:RequiredFieldValidator ID="mustBefAcc" runat="server" Display="None" ControlToValidate="BefAcc" ErrorMessage="請選擇帳號"></asp:RequiredFieldValidator></td>
            </tr>
            <tr>
                <td class="bluecol">選擇計畫</td>
                <td colspan="2" class="whitecol">
                    <asp:RadioButtonList ID="Plan_lst" runat="server" AutoPostBack="True" CssClass="font"></asp:RadioButtonList>
                    <asp:RequiredFieldValidator ID="mustPlan" runat="server" Display="None" ControlToValidate="Plan_lst" ErrorMessage="請選擇計畫"></asp:RequiredFieldValidator></td>
            </tr>
            <tr>
                <td class="bluecol" align="left" colspan="3">轉換給</td>
            </tr>
            <tr>
                <td class="bluecol">角色</td>
                <td colspan="2" class="whitecol">
                    <asp:DropDownList ID="AftLev" AutoPostBack="True" runat="server"></asp:DropDownList>
                    <asp:RequiredFieldValidator ID="mustAftLev" runat="server" ErrorMessage="請選擇角色" ControlToValidate="AftLev" Display="None"></asp:RequiredFieldValidator></td>
            </tr>
            <tr>
                <td class="bluecol">帳號</td>
                <td colspan="2" class="whitecol">
                    <asp:DropDownList ID="AftAcc" runat="server" AutoPostBack="True"></asp:DropDownList>
                    <asp:RequiredFieldValidator ID="mustAftAcc" runat="server" ErrorMessage="請選擇帳號" ControlToValidate="AftAcc" Display="None"></asp:RequiredFieldValidator></td>
            </tr>

        </table>
        <table width="100%">
            <tr>
                <td colspan="3" align="center" class="whitecol">
                    <asp:ValidationSummary ID="totalmsg" runat="server" DisplayMode="List" ShowMessageBox="True" ShowSummary="False"></asp:ValidationSummary>
                    <asp:Button ID="btu_sub" runat="server" Text="轉換" CssClass="asp_button_M"></asp:Button></td>
            </tr>
        </table>
    </form>
</body>
</html>
