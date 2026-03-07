<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_07_002_add.aspx.vb" Inherits="WDAIIP.CP_07_002_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>CP_07_002_add</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<style type="text/css">
		.Tb_01 { width: 600px; font-size: 13px; }
		.Td_02 { background-color: #4067b2; width: 600px; }
		.FontColor { font-size: 13px; }
	</style>
	<script language="javascript" src="../../js/date-picker.js"></script>
	<script type="text/javascript">
		function ChgFillDate() {
			__doPostBack('LinkButton1', '');
		}
	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<table cellspacing="1" cellpadding="1" width="740" border="0">
		<tr>
			<td>
				<table class="font" width="100%" border="0">
					<tr>
						<td>
							首頁&gt;&gt;查核/績效管理&gt;&gt;<font color="#990000">受訓學員座談紀錄表</font>
						</td>
					</tr>
				</table>
				<table cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>
							<p align="center">
								<font style="text-align: center; font-variant: normal; font-family: 新細明體; font-size: 16px; vertical-align: baseline; font-weight: bold">受訓學員座談紀錄表</font></p>
						</td>
					</tr>
				</table>
				<table class="table_sch" cellpadding="1" cellspacing="1">
					<tr>
						<td class="bluecol">
							計畫名稱：
						</td>
						<td class="whitecol">
							<asp:Label ID="lb_PlanName" runat="server" ForeColor="Black"></asp:Label>
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							培訓單位：
						</td>
						<td class="whitecol">
							<asp:Label ID="lb_OrgName" runat="server" ForeColor="Black"></asp:Label>
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							訓練班別：
						</td>
						<td class="whitecol">
							<asp:Label ID="lb_OCID" runat="server" ForeColor="Black"></asp:Label>
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							訓練期間：
						</td>
						<td style="height: 17px" class="whitecol">
							<asp:Label ID="lb_STDate" runat="server" ForeColor="Black" Width="90px"></asp:Label>~
							<asp:Label ID="lb_FTDate" runat="server" Width="90px"></asp:Label>
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							座談日期：
						</td>
						<td class="whitecol">
							<asp:TextBox ID="lb_FillDate" runat="server" ForeColor="Black" Width="100px" onfocus="this.blur()"></asp:TextBox>(星期
							<asp:Label ID="WeekDay" runat="server" ForeColor="Black" Width="20px"></asp:Label>)
							<input id="txtFillDate" style="width: 88px; height: 22px" type="hidden" size="9" name="txtFillDate" runat="server" onpropertychange="ChgFillDate();">
							<img style="cursor: pointer" onclick="javascript:show_calendar('<%= txtFillDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" width="30" height="30" align="top">
							<asp:LinkButton ID="LinkButton1" runat="server"></asp:LinkButton>
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							學號：
						</td>
						<td class="whitecol">
							<asp:Label ID="studID" runat="server"></asp:Label>
							<input id="center" type="hidden" name="center" runat="server">
							<input id="RIDValue" type="hidden" name="RIDValue" runat="server">
							<input id="TMID1" type="hidden" name="TMID1" runat="server">
							<input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
							<input id="OCID1" type="hidden" name="OCID1" runat="server">
							<input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
							<input id="STDate1" type="hidden" name="STDate1" runat="server">
							<input id="STDate2" type="hidden" name="STDate2" runat="server">
							<input id="hid_socid" type="hidden" name="hid_socid" runat="server">
							<input id="hid_planid" type="hidden" name="hid_planid" runat="server">
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							學員：
						</td>
						<td class="whitecol">
							<font face="新細明體">&nbsp;
								<asp:DropDownList ID="ddl_SOCID" runat="server" Width="144px" AutoPostBack="True">
								</asp:DropDownList>
								<asp:Label ID="msg" runat="server" ForeColor="Red" Width="160px"></asp:Label></font>
						</td>
					</tr>
				</table>
				<table class="font" width="600">
					<tr>
						<td width="10%">
						</td>
						<td align="left" colspan="5" height="25">
							&nbsp;
							<asp:CheckBox ID="ChkBox01" runat="server" Width="120px" Text="是否參考簽到表"></asp:CheckBox>
						</td>
					</tr>
					<tr>
						<td align="center" colspan="6">
							座談反應意見：
						</td>
					</tr>
					<tr>
						<td align="center" colspan="6">
							<asp:TextBox ID="txt_Content" runat="server" Width="500px" TextMode="MultiLine" Height="200px"></asp:TextBox>
						</td>
					</tr>
					<tr>
						<td align="center" colspan="6">
							<asp:Button ID="bt_back" runat="server" Text="回上頁" CssClass="asp_button_S"></asp:Button>&nbsp;<asp:Button ID="bt_save" runat="server" Text="儲存" CssClass="asp_button_S"></asp:Button>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
	<input id="hidQstatus" runat="server" type="hidden" />
	</form>
</body>
</html>
