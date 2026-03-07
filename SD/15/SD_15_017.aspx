<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_017.aspx.vb" Inherits="WDAIIP.SD_15_017" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
	<title>項目金額查詢</title>
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<script language="javascript" src="../../js/date-picker.js"></script>
	<script language="javascript" src="../../js/openwin/openwin.js"></script>
	<script language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
	<script language="javascript" type="text/javascript">
		function SelectAll(obj, hidobj) {
			var num = getCheckBoxListValue(obj).length; //長度
			var myallcheck = document.getElementById(obj + '_' + 0); //第1個
			//alert(getCheckBoxListValue(obj));
			if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
				document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0); //記憶
				for (var i = 1; i < num; i++) {
					var mycheck = document.getElementById(obj + '_' + i);
					mycheck.checked = myallcheck.checked;
				}
			}
			else {
				//若有全選
				if (getCheckBoxListValue(obj).charAt(0) == '1') {
					myallcheck.checked = false; //全選改為false
					document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0); //記憶
				}

			}
		}
	</script>
</head>
<body>
	<form id="form1" runat="server">
	<table class="font" cellspacing="1" cellpadding="1" width="740" border="0">
		<tr>
			<td>
				<table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>
							<asp:Label ID="TitleLab1" runat="server"></asp:Label>
							<asp:Label ID="TitleLab2" runat="server">
							首頁&gt;&gt;學員動態管理&gt;&gt;統計表&gt;&gt;<FONT color="#990000">項目金額查詢</FONT>
							</asp:Label>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td class="bluecol" width="20%">
							年度
						</td>
						<td class="whitecol">
							<asp:DropDownList ID="yearlist1" runat="server" AutoPostBack="True">
							</asp:DropDownList>
							～
							<asp:DropDownList ID="yearlist2" runat="server" AutoPostBack="True">
							</asp:DropDownList>
						</td>
					</tr>
					<tr>
						<td class="bluecol" width="20%">
							轄區
						</td>
						<td class="whitecol">
							<asp:CheckBoxList ID="Distid" runat="server" RepeatDirection="Horizontal" RepeatColumns="3">
							</asp:CheckBoxList>
							<input id="DistHidden" type="hidden" value="0" runat="server">
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							訓練業別1
						</td>
						<td class="whitecol">
							<asp:CheckBoxList ID="GovClassName1" runat="server" RepeatDirection="Horizontal" RepeatColumns="9">
							</asp:CheckBoxList>
							<input id="HidGovClass1" type="hidden" value="0" runat="server">
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							訓練業別2
						</td>
						<td class="whitecol">
							<asp:CheckBoxList ID="GovClassName2" runat="server" RepeatDirection="Horizontal" RepeatColumns="9">
							</asp:CheckBoxList>
							<input id="HidGovClass2" type="hidden" value="0" runat="server">
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							課程分類
						</td>
						<td class="whitecol">
							<asp:CheckBoxList ID="cblDepot12" runat="server" RepeatDirection="Horizontal" RepeatColumns="5">
							</asp:CheckBoxList>
							<input id="HidcblDepot12" type="hidden" value="0" runat="server">
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							品名
						</td>
						<td class="whitecol">
							<asp:TextBox ID="txtCName" runat="server" MaxLength="30" Height="19px"></asp:TextBox>
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							比對模式
						</td>
						<td class="whitecol">
							<asp:RadioButtonList ID="rblCompareMode1" runat="server" RepeatDirection="Horizontal">
								<asp:ListItem Selected="True" Value="1">模糊比對</asp:ListItem>
								<asp:ListItem Value="2">完整比對</asp:ListItem>
							</asp:RadioButtonList>
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							是否核定
						</td>
						<td class="whitecol">
							<asp:RadioButtonList ID="rblIsSuccess" runat="server" RepeatDirection="Horizontal">
								<asp:ListItem Value="A">不區分</asp:ListItem>
								<asp:ListItem Selected="True" Value="Y">是</asp:ListItem>
								<asp:ListItem Value="N">否</asp:ListItem>
							</asp:RadioButtonList>
						</td>
					</tr>
					<tr>
						<td align="center" colspan="4">
							<asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center">
				<div>
					<asp:Button ID="BtnExp" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
				</div>
			</td>
		</tr>
	</table>
	</form>
</body>
</html>
