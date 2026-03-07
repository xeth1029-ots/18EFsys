<%@ Control Language="vb" AutoEventWireup="false" Codebehind="DataGridPage.ascx.vb" Inherits="WDAIIP.DataGridPage" TargetSchema="http://schemas.microsoft.com/intellisense/ie5" %>
<TABLE id="Table1" cellSpacing="1" cellPadding="1" border="0">
	<TR>
		<TD>
			<asp:ImageButton id="ImageButton1" runat="server" ImageUrl="images/BackBack.gif" CausesValidation="False"></asp:ImageButton></TD>
		<TD>
			<asp:ImageButton id="ImageButton2" runat="server" ImageUrl="images/Back2.gif" CausesValidation="False"></asp:ImageButton></TD>
		<TD>
			<asp:TextBox id="TextBox1" runat="server" Width="30px" AutoPostBack="True"></asp:TextBox>/
			<asp:TextBox id="TextBox2" runat="server" Width="30px" ReadOnly="True"></asp:TextBox></TD>
		<TD>
			<asp:ImageButton id="ImageButton3" runat="server" ImageUrl="images/Next.gif" CausesValidation="False"></asp:ImageButton></TD>
		<TD>
			<asp:ImageButton id="ImageButton4" runat="server" ImageUrl="images/NextNext.gif" CausesValidation="False"></asp:ImageButton></TD>
	</TR>
</TABLE>
