<%@ Register TagPrefix="uc1" TagName="PageControler" Src="../../PageControler.ascx" %>
<%@ Page Language="vb" AutoEventWireup="false" Codebehind="QuestionSearch.aspx.vb" Inherits="TIMS.QuestionSearch" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>QuestionSearch</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../style.css" type="text/css" rel="stylesheet">
		<script src="../../js/common.js"></script>
	</HEAD>
	<body MS_POSITIONING="FlowLayout">
		<form id="form1" method="post" runat="server">
			<table class="font" id="table_F" width="100%" runat="server">
				<TR>
					<TD class="SD_TD1" style="WIDTH: 82px" align="center"><label><span class="style3">問卷名稱</span></label></TD>
					<TD class="SD_TD2"><input id="Ipt_Name" style="WIDTH: 373px; HEIGHT: 22px" type="text" maxLength="100" size="56"
							name="Ipt_Name" runat="server" height="18"></TD>
				</TR>
				<TR align="center">
					<TD colSpan="2"><input id="search" type="button" value="查詢" name="search" runat="server"></TD>
				</TR>
			</table>
			<TABLE id="Table2" width="100%" runat="server">
				<TR Width="100%">
					<TD align="center" width="100%"><FONT face="新細明體"><asp:datagrid id="DataGrid1" runat="server" Width="100%" Runat="server" CssClass="font" AutoGenerateColumns="False"
								AllowPaging="True">
								<AlternatingItemStyle HorizontalAlign="Center" BackColor="White"></AlternatingItemStyle>
								<ItemStyle HorizontalAlign="Center" BackColor="#EBF8FF"></ItemStyle>
								<HeaderStyle HorizontalAlign="Center" ForeColor="White" BackColor="#2AAFC0"></HeaderStyle>
								<Columns>
									<asp:TemplateColumn>
										<ItemTemplate>
											<INPUT id="Check1" value='<%# DataBinder.Eval(Container.DataItem,"SVID")%>' type="radio" name="Check1">
										</ItemTemplate>
									</asp:TemplateColumn>
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
									<asp:BoundColumn Visible="False" DataField="SVID" HeaderText="SVID"></asp:BoundColumn>
								</Columns>
								<PagerStyle Visible="False"></PagerStyle>
							</asp:datagrid></FONT><uc1:pagecontroler id="PageControler1" runat="server"></uc1:pagecontroler></TD>
				</TR>
				<TR Width="100%">
					<TD width="100%">
						<P align="center"><asp:label id="msg" runat="server" CssClass="font" ForeColor="Red"></asp:label></P>
					</TD>
				</TR>
				<TR>
					<TD>
						<P align="center"><asp:button id="send" runat="server" Text="送出"></asp:button></P>
					</TD>
				</TR>
			</TABLE>
		</form>
	</body>
</HTML>
