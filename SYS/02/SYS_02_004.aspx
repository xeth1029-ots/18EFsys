<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_02_004.aspx.vb" Inherits="WDAIIP.SYS_02_004" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>職務機構調動</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;帳號維護管理&gt;&gt;職務機構調動</asp:Label>
                </td>
            </tr>
        </table>
        <%--<table class="font" width="100%">
				<tr>
					<td class="font">首頁&gt;&gt;系統管理&gt;&gt;帳號維護管理&gt;&gt;職務機構調動</td>
				</tr>
			</table>--%>
        <br>
        <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td class="bluecol" style="width: 20%">角色</td>
                <td colspan="2" class="whitecol">承辦人</td>
            </tr>
            <tr>
                <td class="bluecol">帳號</td>
                <td class="whitecol">
                    <asp:DropDownList ID="BefAcc" AutoPostBack="True" runat="server"></asp:DropDownList>
                    <asp:RequiredFieldValidator ID="mustBefAcc" runat="server" ErrorMessage="請選擇帳號" ControlToValidate="BefAcc" Display="None"></asp:RequiredFieldValidator></td>
                <td class="whitecol" width="200">
                    <asp:RadioButtonList ID="Plan_lst" runat="server" CssClass="font" AutoPostBack="True"></asp:RadioButtonList>
                    <asp:RequiredFieldValidator ID="mustPlan" runat="server" ErrorMessage="請選擇計畫" ControlToValidate="Plan_lst" Display="None"></asp:RequiredFieldValidator></td>
            </tr>
            <tr>
                <td class="bluecol">機構</td>
                <td colspan="2" class="whitecol">
                    <asp:CheckBoxList ID="Org_lis" runat="server" AutoPostBack="True" CssClass="font"></asp:CheckBoxList></td>
            </tr>
            <tr>
                <td class="bluecol" colspan="3">調給</td>
            </tr>
            <tr>
                <td class="bluecol">角色</td>
                <td colspan="2" class="whitecol">承辦人</td>
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
                <td align="center" class="whitecol">
                    <asp:ValidationSummary ID="totalmsg" runat="server" DisplayMode="List" ShowMessageBox="True" ShowSummary="False"></asp:ValidationSummary>
                    <asp:Button ID="btu_sub" runat="server" Text="轉換" CssClass="asp_button_M"></asp:Button></td>
            </tr>
        </table>
    </form>
</body>
</html>
