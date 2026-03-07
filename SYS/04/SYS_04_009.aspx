<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_04_009.aspx.vb" Inherits="WDAIIP.SYS_04_009" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>計畫核銷方式設定</title>
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;系統參數管理&gt;&gt;計畫核銷方式設定</asp:Label>
                </td>
            </tr>
        </table>	
		<table id="FrameTable1" cellspacing="1" cellpadding="1" width="100%" border="0">
			<tr>
				<td align="center" class="whitecol">
					<%--<table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
						<tr>
							<td>
								<font face="新細明體">首頁&gt;&gt;系統管理&gt;&gt;系統參數管理&gt;&gt;<font color="#990000">計畫核銷方式設定</font></font>
							</td>
						</tr>
					</table>--%>
					<asp:DataGrid ID="DataGrid1" runat="server" AutoGenerateColumns="False" CssClass="font" Width="100%" CellPadding="8">
						<AlternatingItemStyle BackColor="#F5F5F5" />
						<HeaderStyle CssClass="head_navy" />
						<Columns>
							<asp:BoundColumn DataField="PlanName" HeaderText="訓練計畫" HeaderStyle-Width="60%"></asp:BoundColumn>
							<asp:TemplateColumn HeaderText="成本加工法">
								<HeaderStyle Width="8%"></HeaderStyle>
								<ItemStyle HorizontalAlign="Center"></ItemStyle>
								<ItemTemplate>
									<input id="CancelID1" type="radio" value="1" runat="server">
								</ItemTemplate>
							</asp:TemplateColumn>
							<asp:TemplateColumn HeaderText="多期核銷金額">
								<HeaderStyle Width="8%"></HeaderStyle>
								<ItemStyle HorizontalAlign="Center"></ItemStyle>
								<ItemTemplate>
									<input id="CancelID2" type="radio" value="2" runat="server">
								</ItemTemplate>
							</asp:TemplateColumn>
							<asp:TemplateColumn HeaderText="依總平均單價及預算別核銷">
								<HeaderStyle Width="8%"></HeaderStyle>
								<ItemStyle HorizontalAlign="Center"></ItemStyle>
								<ItemTemplate>
									<input id="CancelID3" type="radio" value="3" runat="server">
								</ItemTemplate>
							</asp:TemplateColumn>
							<asp:TemplateColumn HeaderText="依學習單元與預算別核銷">
								<HeaderStyle Width="8%"></HeaderStyle>
								<ItemStyle HorizontalAlign="Center"></ItemStyle>
								<ItemTemplate>
									<input id="CancelID4" type="radio" value="4" runat="server">
								</ItemTemplate>
							</asp:TemplateColumn>
							<asp:TemplateColumn HeaderText="依主要參訓身分別核銷">
								<HeaderStyle Width="8%"></HeaderStyle>
								<ItemStyle HorizontalAlign="Center"></ItemStyle>
								<ItemTemplate>
									<input id="CancelID5" type="radio" value="5" runat="server">
								</ItemTemplate>
							</asp:TemplateColumn>
						</Columns>
					</asp:DataGrid>
					<asp:Button ID="Button1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
				</td>
			</tr>
		</table>	
	</form>
</body>
</html>
