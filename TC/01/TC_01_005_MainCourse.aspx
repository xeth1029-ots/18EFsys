<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_005_MainCourse.aspx.vb" Inherits="WDAIIP.TC_01_005_MainCourse" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>主課程</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<script language="javascript" src="../../js/date-picker.js"></script>
	<script language="javascript" src="../../js/openwin/openwin.js"></script>
	<script language="javascript">
		function returnValue(id, name) {
			opener.document.form1.TB_CourName.value = name;
			opener.document.form1.TB_CourName.style.cssText = "COLOR: black";
			opener.document.form1.courid.value = id;
			window.close();
		}
	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<table class="table_nw" style="width: 100%">
		<tr>
			<td class="bluecol" style="width:20%">				
				訓練職類				
			</td>
			<td Class="whitecol">
				<p>					
					<asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Width="30%" ></asp:TextBox><input id="career" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." name="career" runat="server" class="asp_button_M"><input id="trainValue" style="width: 32px; height: 22px" type="hidden" name="trainValue" runat="server">
					<asp:Button ID="Button4" runat="server" CausesValidation="False" Text="清除" CssClass="asp_button_M"></asp:Button>
				</p>
			</td>
		</tr>
		<tr>
			<td class="bluecol">				
				課程代碼
			</td>
			<td Class="whitecol">
				<p>					
			        <asp:TextBox ID="ClassID" runat="server" Width="20%" MaxLength="12" CssClass="whitecol"></asp:TextBox>
				</p>
			</td>
		</tr>
		<tr>
			<td colspan="2" align="center" class="whitecol">
				<asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
			</td>
		</tr>
	</table>	
	<asp:Panel ID="Panel" runat="server" Visible="true" Width="100%">
		<table class="font" id="search_tbl" cellspacing="0" cellpadding="0" width="100%" border="1" runat="server">
		</table>
	</asp:Panel>
	</form>
</body>
</html>
