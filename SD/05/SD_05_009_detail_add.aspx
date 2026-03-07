<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_009_detail_add.aspx.vb" Inherits="WDAIIP.SD_05_009_detail_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>新增講師上課明細</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<link href="../../style.css" type="text/css" rel="stylesheet">
	<script language="javascript" src="../../js/date-picker.js"></script>
	<script language="javascript" src="../../js/common.js"></script>
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<script language="javascript">
		function window_onload() {
			if (document.getElementById("BypassCheck")) {
				if (window.confirm('新增鍾點費重複,是否合併此筆資料?')) {
					document.getElementById('save_Button').click();
				} else {
					document.getElementById('BypassCheck').value = '0';
				}
			}
		}

		//'檢查日期格式-Melody(2005/3/28)
		function check_date() {
			if (!checkDate(form1.start_date.value)) {
				alert('請輸入正確的日期格式,YYYY/MM/DD!!\n');
			}
		}
	</script>
</head>
<body onload="window_onload();">
	<form id="form1" method="post" runat="server">
	<asp:Panel ID="Panel1" Style="z-index: 101; left: 24px; position: absolute; top: 48px" runat="server" CssClass="font" Height="24px" Width="312px">
		講師:&nbsp;
		<asp:Label ID="Label1" runat="server" CssClass="font"></asp:Label>
	</asp:Panel>
	<asp:Panel ID="Panel" Style="z-index: 102; left: 24px; position: absolute; top: 72px" runat="server" Height="56px" Width="312px" Visible="true">
		<table class="font" id="search_tbl" cellspacing="0" cellpadding="0" width="265" border="1" runat="server">
		</table>
		<table class="font" id="Table1" cellspacing="1" cellpadding="1" width="300" border="1">
			<tr>
				<td style="width: 42px; height: 23px" bgcolor="#2aafc0">
					<font color="#ffffff">日期</font><font color="#ff0000">*</font>
				</td>
				<td style="height: 23px" bgcolor="#ecf7ff" colspan="2">
					<asp:TextBox ID="start_date" runat="server"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" width="30" height="30">
					<asp:RequiredFieldValidator ID="TB_date" runat="server" Display="None" ErrorMessage="請輸入授課日期" ControlToValidate="start_date"></asp:RequiredFieldValidator>
				</td>
			</tr>
			<tr>
				<td style="width: 42px; height: 20px" bgcolor="#2aafc0">
					<font color="#ffffff">單價</font><font color="#ff0000">*</font>
				</td>
				<td style="height: 20px" bgcolor="#ecf7ff" colspan="2">
					<asp:TextBox ID="TB_Price" runat="server"></asp:TextBox>
					<asp:RequiredFieldValidator ID="TB_price_check" runat="server" Display="None" ErrorMessage="請輸入單價" ControlToValidate="TB_Price"></asp:RequiredFieldValidator>
				</td>
			</tr>
			<tr>
				<td style="width: 42px" bgcolor="#2aafc0">
					<font color="#ffffff">時數</font><font color="#ff0000">*</font>
				</td>
				<td bgcolor="#ecf7ff" colspan="2">
					<asp:TextBox ID="TB_hour" runat="server"></asp:TextBox>
					<asp:RequiredFieldValidator ID="TB_hour_check" runat="server" Display="None" ErrorMessage="請輸入授課時數" ControlToValidate="TB_hour"></asp:RequiredFieldValidator>
				</td>
			</tr>
		</table>
		<p>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<asp:Button ID="save_Button" runat="server" Text="新增"></asp:Button></p>
		<p>
		</p>
		<p>
			<asp:ValidationSummary ID="Summary" runat="server" Width="200px" CssClass="font" DisplayMode="List" ShowSummary="False" ShowMessageBox="True"></asp:ValidationSummary>
		</p>
	</asp:Panel>
	<asp:CustomValidator ID="CustomValidator1" Style="z-index: 103; left: 168px; position: absolute; top: 16px" runat="server" Height="8px" Display="None" ErrorMessage="CustomValidator" CssClass="font"></asp:CustomValidator>
	<%--<asp:literal id="clientscript" runat="server"></asp:literal>--%>
	</form>
</body>
</html>
