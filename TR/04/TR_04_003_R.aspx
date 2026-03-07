<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_04_003_R.aspx.vb" Inherits="WDAIIP.TR_04_003_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>TR_04_003_R</title>
	<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
	<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
	<meta name="vs_defaultClientScript" content="JavaScript">
	<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
	<script type="text/javascript" language="javascript">
		function print() {
			var msg = '';
			if (isEmpty(document.form1.CPoint)) {
				msg += '請選擇就業查核點!\n';
			}
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
							首頁&gt;&gt;訓練與就業需求管理&gt;&gt;就業追蹤報表&gt;&gt;<font color="#990000">推介成果統計表二</font>
						</td>
					</tr>
				</table>
				<table class="table_sch" id="Table2" cellspacing="1" cellpadding="1">
					<tr>
						<td class="bluecol" width="80">
							結訓期間
						</td>
						<td class="whitecol">
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
							月
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							就服中心
						</td>
						<td class="whitecol">
							<asp:RadioButtonList ID="Station" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow" RepeatColumns="3">
							</asp:RadioButtonList>
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							<font face="新細明體">&nbsp;&nbsp;&nbsp; 就業率查核點<font color="red">*</font></font>
						</td>
						<td class="whitecol">
							<asp:RadioButtonList ID="CPoint" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
								<asp:ListItem Value="0">結訓後一個月內</asp:ListItem>
								<asp:ListItem Value="1">結訓後三個月內</asp:ListItem>
								<asp:ListItem Value="2">結訓後六個月內</asp:ListItem>
							</asp:RadioButtonList>
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
