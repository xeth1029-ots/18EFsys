<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_04_001_R.aspx.vb" Inherits="WDAIIP.TR_04_001_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>TR_04_001_R</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
	<script>
		function print() {
			var msg = '';
			if (document.form1.SYear.selectedIndex == 0) msg += '請選擇開訓期間-年度起\n';
			if (document.form1.SMonth.selectedIndex == 0) msg += '請選擇開訓期間-月起\n';
			if (document.form1.FYear.selectedIndex == 0) msg += '請選擇開訓期間-年度迄\n';
			if (document.form1.FMonth.selectedIndex == 0) msg += '請選擇開訓期間-月迄\n';
			if (msg != '') {
				alert(msg);
				return false;
			}
		}
	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<table id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
		<tr>
			<td>
				<table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>
							<asp:Label ID="TitleLab1" runat="server"></asp:Label>
							<asp:Label ID="TitleLab2" runat="server">
									首頁&gt;&gt;訓練與就業需求管理&gt;&gt;就業追蹤報表&gt;&gt;<FONT color="#990000">研習計畫參加人數統計月報表</FONT>
							</asp:Label>
						</td>
					</tr>
				</table>
				<table class="table_sch" id="Table2" cellspacing="1" cellpadding="1">
					<tr>
						<td class="bluecol_need" width="80">
							開訓期間
						</td>
						<td class="whitecol">
							<font face="新細明體" color="#000000">
								<asp:DropDownList ID="SYear" runat="server">
								</asp:DropDownList>
								年
								<asp:DropDownList ID="SMonth" runat="server">
								</asp:DropDownList>
								月～
								<asp:DropDownList ID="FYear" runat="server">
								</asp:DropDownList>
								年
								<asp:DropDownList ID="FMonth" runat="server">
								</asp:DropDownList>
								月</font>
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							轄區
						</td>
						<td class="whitecol">
							<font face="新細明體">
								<asp:RadioButtonList ID="DistID" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
								</asp:RadioButtonList>
							</font>
						</td>
					</tr>
				</table>
				<p align="center">
					<asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
				</p>
			</td>
		</tr>
	</table>
	</form>
</body>
</html>
