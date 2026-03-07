<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_007_copy.aspx.vb" Inherits="WDAIIP.TC_01_007_copy" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>師資複製</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<link rel="stylesheet" type="text/css" href="../../css/style.css" />
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
	<script type="text/javascript">
		function ReturnMyValue(TechID) {
			opener.document.getElementById('TechID').value = TechID;
			opener.document.getElementById('Button4').click();
			window.close();
		}
	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<asp:DataGrid ID="DataGrid1" runat="server" AutoGenerateColumns="False" CssClass="font" Width="100%" CellPadding="8">
		<Columns>
			<asp:TemplateColumn>
				<HeaderStyle CssClass="head_navy" Width="10%"></HeaderStyle>
                <ItemStyle HorizontalAlign="Center"/>
				<ItemTemplate>
					<input id="Radio1" type="radio" value="Radio1" runat="server">
				</ItemTemplate>
			</asp:TemplateColumn>
			<asp:BoundColumn HeaderText="所在計畫" HeaderStyle-CssClass="head_navy">
                <HeaderStyle Width="30%"></HeaderStyle>
			</asp:BoundColumn>
			<asp:BoundColumn DataField="TeacherID" HeaderText="講師代碼" HeaderStyle-CssClass="head_navy">
				<HeaderStyle Width="30%"></HeaderStyle>
			</asp:BoundColumn>
			<asp:BoundColumn DataField="TeachCName" HeaderText="姓名" HeaderStyle-CssClass="head_navy">
				<HeaderStyle Width="30%"></HeaderStyle>
			</asp:BoundColumn>
		</Columns>
	</asp:DataGrid>
	</form>
</body>
</html>
