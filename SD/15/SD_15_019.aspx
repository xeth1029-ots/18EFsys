<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_019.aspx.vb" Inherits="WDAIIP.SD_15_019" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
	<title>重複參訓統計表</title>
	<link rel="stylesheet" type="text/css" href="../../css/style.css">
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
	<form id="form1" method="post" runat="server">
	<table class="font" cellspacing="1" cellpadding="1" width="740" border="0">
		<tr>
			<td>
				<table id="Table1" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
					<tr>
						<td>
							<asp:Label ID="TitleLab1" runat="server"></asp:Label>
							<asp:Label ID="TitleLab2" runat="server">
							首頁&gt;&gt;學員動態管理&gt;&gt;統計表&gt;&gt;<FONT color="#990000">重複參訓統計表</FONT>
							</asp:Label>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table id="Table2" class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td class="bluecol_need">
							年度
						</td>
						<td class="whitecol">
							<asp:DropDownList ID="ddlYears" runat="server">
							</asp:DropDownList>
						</td>
					</tr>
					<tr>
						<td class="bluecol_need">
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
							開訓期間
						</td>
						<td class="whitecol">
							<asp:TextBox ID="STDate1" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">～
							<asp:TextBox ID="STDate2" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">
						</td>
					</tr>
					<tr>
						<td class="bluecol_need">
							相同訓練業別
						</td>
						<td class="whitecol">
							<asp:RadioButtonList ID="sameJOB1" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
								<asp:ListItem Value="N" Selected="True">否</asp:ListItem>
								<asp:ListItem Value="Y">是</asp:ListItem>
							</asp:RadioButtonList>
						</td>
					</tr>
					<tr>
						<td class="bluecol_need">
							產出格式
						</td>
						<td class="whitecol">
							<asp:RadioButtonList ID="expType1" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
								<asp:ListItem Value="1" Selected="True">統計表</asp:ListItem>
								<asp:ListItem Value="2">明細表</asp:ListItem>
							</asp:RadioButtonList>
						</td>
					</tr>
					<tr id="trPlanKind" runat="server">
						<td class="bluecol">
							計畫範圍
						</td>
						<td class="whitecol">
							<asp:RadioButtonList ID="OrgKind2" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
								<asp:ListItem Value="A">不區分</asp:ListItem>
								<asp:ListItem Value="G" Selected="True">產業人才投資計畫</asp:ListItem>
								<asp:ListItem Value="W">提升勞工自主學習計畫</asp:ListItem>
							</asp:RadioButtonList>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center">
				<div align="center">
					<asp:Button ID="btnPrint1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
				</div>
			</td>
		</tr>
	</table>
	</form>
</body>
</html>
