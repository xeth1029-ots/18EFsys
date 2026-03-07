<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SYS_04_012.aspx.vb" Inherits="TIMS.SYS_04_012" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>計畫核銷方式設定</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<LINK href="../../css/style.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body MS_POSITIONING="FlowLayout">
		<form id="form1" method="post" runat="server">
			<FONT face="新細明體">
				<TABLE id="FrameTable" cellSpacing="1" cellPadding="1" width="600" border="0">
					<TR>
						<TD align="center">
							<TABLE class="font" id="Table1" cellSpacing="1" cellPadding="1" width="100%" border="0">
								<TR>
									<TD><FONT face="新細明體">首頁&gt;&gt;系統管理&gt;&gt;系統參數管理&gt;&gt;<FONT color="#990000">計畫訓練性質設定</FONT></FONT></TD>
								</TR>
							</TABLE>
							<asp:DataGrid id="DataGrid1" runat="server" AutoGenerateColumns="False" CssClass="font" Width="100%">
								<AlternatingItemStyle BackColor="#F5F5F5" />
                                <HeaderStyle CssClass="head_navy" />
                                <Columns>
									<asp:BoundColumn DataField="PlanName" HeaderText="訓練計畫"></asp:BoundColumn>
									<asp:TemplateColumn HeaderText="職前">
										<HeaderStyle Width="80px" HorizontalAlign="Center"></HeaderStyle>
										<ItemStyle HorizontalAlign="Center"></ItemStyle>
										<ItemTemplate>
											<INPUT id="PropertyID0" type="radio" value="0" runat="server" NAME="RadioGroup">
										</ItemTemplate>
									</asp:TemplateColumn>
									<asp:TemplateColumn HeaderText="在職">
										<HeaderStyle Width="80px" HorizontalAlign="Center"></HeaderStyle>
										<ItemStyle HorizontalAlign="Center"></ItemStyle>
										<ItemTemplate>
											<INPUT id="PropertyID1" type="radio" value="1" runat="server" NAME="RadioGroup">
										</ItemTemplate>
									</asp:TemplateColumn>
									<asp:TemplateColumn HeaderText="其他(停用)">
										<HeaderStyle Width="80px" HorizontalAlign="Center"></HeaderStyle>
										<ItemStyle HorizontalAlign="Center"></ItemStyle>
										<ItemTemplate>
											<INPUT id="PropertyIDX" type="radio" value="3" runat="server" NAME="RadioGroup">
										</ItemTemplate>
									</asp:TemplateColumn>
								</Columns>
							</asp:DataGrid>
							<asp:Button id="Button1" runat="server" Text="儲存" CssClass="asp_button_S"></asp:Button></TD>
					</TR>
				</TABLE>
			</FONT>
		</form>
	</body>
</HTML>
