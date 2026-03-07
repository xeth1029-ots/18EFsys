 

<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SV_01_003.aspx.vb" Inherits="WDAIIP.SV_01_003" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>問卷題目設定</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<script src="../../js/common.js"></script>
	<script language="javascript" type="text/javascript">
	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
		<tr>
			<td>
				首頁&gt;&gt;系統管理&gt;&gt;問卷管理&gt;&gt;問卷題目設定
			</td>
		</tr>
	</table>
	<table class="table_nw" id="table_F" width="740" runat="server">
		<tr>
			<td class="bluecol">
				<label>
					<span class="style3">問卷名稱</span></label>
			</td>
			<td class="whitecol">
				<input id="Ipt_Name" style="width: 373px; height: 22px" type="text" maxlength="100" size="56" name="Ipt_Name" runat="server" height="18">
			</td>
		</tr>
		<tr align="center">
			<td colspan="2" class="whitecol">
				<asp:Button ID="btnSearch" runat="server" Text="查詢" CssClass="button_b_S" />
				<%--<input id="search" type="button" value="查詢" name="search" runat="server" class="button_b_S" onclick="return search_onclick()">--%>
			</td>
		</tr>
	</table>
	<table id="Table2" width="740" runat="server">
		<tr width="100%">
			<td align="center" width="100%">
				<asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AllowPaging="True" AutoGenerateColumns="False" CssClass="font" runat="server">
					<AlternatingItemStyle HorizontalAlign="Center" BackColor="#F5F5F5"></AlternatingItemStyle>
					<HeaderStyle CssClass="head_navy"></HeaderStyle>
					<Columns>
						<asp:BoundColumn HeaderText="序號">
							<HeaderStyle Width="10%"></HeaderStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Name" HeaderText="問卷名稱">
							<HeaderStyle Width="65%"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left" VerticalAlign="Middle"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Avail" HeaderText="狀態">
							<HeaderStyle Width="15%"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center" VerticalAlign="Middle"></ItemStyle>
						</asp:BoundColumn>
						<asp:TemplateColumn HeaderText="功能">
							<HeaderStyle Width="10%"></HeaderStyle>
							<ItemTemplate>
								<asp:Button ID="Btn_edit" runat="server" Text="問卷題目設定" CommandName="edit"></asp:Button>
							</ItemTemplate>
						</asp:TemplateColumn>
						<asp:BoundColumn Visible="False" DataField="SVID" HeaderText="SVID"></asp:BoundColumn>
					</Columns>
					<PagerStyle Visible="False"></PagerStyle>
				</asp:DataGrid>
				<uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
			</td>
		</tr>
		<tr width="100%">
			<td width="100%">
				<p align="center">
					<asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label></p>
			</td>
		</tr>
	</table>
	</form>
</body>
</html>
