<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SV_01_004_Insert.aspx.vb" Inherits="TIMS.SV_01_004_Insert" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>F</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
</head>
<body ms_positioning="FlowLayout">
	<form id="form1" method="post" runat="server">
	<table id="Title" width="100%" runat="server">
		<tr>
			<td>
				<asp:Label ID="Label1" runat="server">學號 :</asp:Label>
				<asp:Label ID="StudentIDL" runat="server"></asp:Label>
			</td>
			<td>
				<asp:Label ID="Label2" runat="server">學員姓名 :</asp:Label>
				<asp:Label ID="SnameL" runat="server"></asp:Label>
			</td>
		</tr>
	</table>
	<asp:Panel ID="Panel1" runat="server">
		<asp:PlaceHolder ID="PlaceHolder1" runat="server"></asp:PlaceHolder>
	</asp:Panel>
	<table id="btntable" width="100%" runat="server">
		<tr>
			<td align="center">
				<asp:Button ID="Save" runat="server" Text="儲存"></asp:Button>
				<asp:Button ID="returnQ" runat="server" Text="回上一頁"></asp:Button>
				<asp:Button ID="reset" runat="server" Text="重填"></asp:Button>
			</td>
		</tr>
	</table>
	<input id="Type" type="hidden" runat="server">
	<input id="SOCID" type="hidden" runat="server">
	<input id="SVID" type="hidden" runat="server">
	</form>
</body>
</html>
