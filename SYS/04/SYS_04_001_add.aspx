<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_04_001_add.aspx.vb" Inherits="WDAIIP.SYS_04_001_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>年度行事曆(新增)</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
</head>
<body>
    <form id="form1" runat="server">
<%--880FE13F7EEC3837AEF83751BB45F8856B7A1E382FC44F1AAB10441D870CB1F2EEF6C41E7A9C9A9F5C5AFA518D186ACFB50DB9BA80C96592F6E1C827
CA90A0CD57C12EC577C2FD32FE3F23272CE3A073A108307080A3D99C7C0FB31E40905BE40C2A98C4BFC0AE191F4AACE6BED9928986B6012E59BCE8C5
DE882A8B5305A049EA5A87D9151EC7A3352DC52A396593CE1B7BFCBBF52285093A06FAFE3C1FDC29--%>
        <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td class="bluecol_need" style="width: 20%">日期</td>
                <td class="whitecol">&nbsp;&nbsp;&nbsp;&nbsp;
				<asp:TextBox ID="Start_Date" runat="server" onfocus="this.blur()" Width="15%" MaxLength="10"></asp:TextBox>
                    <img style="cursor: pointer" onclick="javascript:show_calendar('Start_Date','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">～
				<asp:TextBox ID="End_Date" runat="server" onfocus="this.blur()" Width="15%" MaxLength="10"></asp:TextBox>
                    <img style="cursor: pointer" onclick="javascript:show_calendar('End_Date','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">事由</td>
                <td class="whitecol">&nbsp;&nbsp;&nbsp;&nbsp;<asp:TextBox ID="Reason" runat="server" Width="50%" MaxLength="50"></asp:TextBox></td>
            </tr>
        </table>
        <table width="100%">
            <tr>
                <td colspan="2" align="center" class="whitecol">
                    <%--<asp:RequiredFieldValidator ID="MustSDate" runat="server" ErrorMessage="請輸入起日期" Display="None" ControlToValidate="Start_Date"></asp:RequiredFieldValidator>
				<asp:ValidationSummary ID="Totalmsg" runat="server" DisplayMode="List" ShowMessageBox="True" ShowSummary="False"></asp:ValidationSummary>
				<asp:CompareValidator ID="CpareDate" runat="server" ErrorMessage="迄日不可比起日小" Display="None" ControlToValidate="End_Date" Type="Date" ControlToCompare="Start_Date" Operator="GreaterThanEqual"></asp:CompareValidator>
                    --%>
                    <asp:Button ID="but_save" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                    <asp:Button ID="BtnBack1" runat="server" Text="回上一頁" CssClass="asp_button_M" />
                    <%--<input id="back" type="button" value="回上一頁" name="back" runat="server" class="button_b_S" onclick="return back_onclick()">--%>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
