 

<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_003.aspx.vb" Inherits="WDAIIP.SD_05_003" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>SD_05_003</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<script language="javascript" src="../../js/date-picker.js"></script>
	<script language="javascript" src="../../js/openwin/openwin.js"></script>
	<script language="javascript" src="../../js/common.js"></script>
	<script language="javascript">
		function search() {
			if (document.form1.start_date.value == '' || document.form1.end_date.value == '') {
				alert('請選擇日期範圍!')
				return false;
			}
		}
	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<table id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0" runat="server">
		<tr>
			<td>
				<table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">轉班作業</font> </td>
					</tr>
				</table>
				<table class="table_sch" id="Table3" cellpadding="1" cellspacing="1">
					<tr>
						<td width="100" class="bluecol_need">轉班時間 </td>
						<td class="whitecol">
							<asp:TextBox ID="start_date" runat="server" Width="75px"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">～
							<asp:TextBox ID="end_date" runat="server" Width="75px"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
						</td>
					</tr>
				</table>
				<table width="100%">
					<tr>
						<td class="whitecol">
							<p align="center">
								<asp:Button ID="Button2" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>&nbsp;&nbsp;<asp:Button ID="Button3" runat="server" Text="新增" CssClass="asp_button_S"></asp:Button>
							</p>
						</td>
					</tr>
				</table>
				<table id="Table4" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
					<tr>
						<td>
							<asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True">
								<AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
								<HeaderStyle CssClass="head_navy"></HeaderStyle>
								<Columns>
									<asp:BoundColumn HeaderText="序號">
										<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
										<ItemStyle HorizontalAlign="Center"></ItemStyle>
									</asp:BoundColumn>
									<asp:BoundColumn DataField="Name" HeaderText="學員姓名">
										<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
										<ItemStyle HorizontalAlign="Center"></ItemStyle>
									</asp:BoundColumn>
									<asp:BoundColumn DataField="ClassCName1" HeaderText="原班級">
										<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
									</asp:BoundColumn>
									<asp:BoundColumn DataField="ApplyDate" HeaderText="轉班日期" DataFormatString="{0:d}">
										<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
										<ItemStyle HorizontalAlign="Center"></ItemStyle>
									</asp:BoundColumn>
									<asp:BoundColumn DataField="ClassCName2" HeaderText="轉至班級">
										<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
									</asp:BoundColumn>
									<asp:BoundColumn DataField="Reason" HeaderText="轉班原因">
										<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
									</asp:BoundColumn>
									<asp:TemplateColumn HeaderText="功能">
										<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
										<ItemStyle HorizontalAlign="Center"></ItemStyle>
										<ItemTemplate>
											<asp:Button ID="Button1" runat="server" Text="回復" CommandName="back"></asp:Button>
										</ItemTemplate>
									</asp:TemplateColumn>
									<asp:BoundColumn Visible="False" DataField="CyclType1" HeaderText="CyclType1"></asp:BoundColumn>
									<asp:BoundColumn Visible="False" DataField="LevelType1" HeaderText="LevelType1"></asp:BoundColumn>
									<asp:BoundColumn Visible="False" DataField="CyclType2" HeaderText="CyclType2"></asp:BoundColumn>
									<asp:BoundColumn Visible="False" DataField="LevelType2" HeaderText="LevelType2"></asp:BoundColumn>
								</Columns>
								<PagerStyle Visible="False"></PagerStyle>
							</asp:DataGrid>
						</td>
					</tr>
					<tr>
						<td>
							<p align="center">
								<uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
							</p>
						</td>
					</tr>
				</table>
				<p align="center">
					<asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label></p>
			</td>
		</tr>
	</table>
	</form>
</body>
</html>
