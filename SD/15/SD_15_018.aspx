<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_018.aspx.vb" Inherits="WDAIIP.SD_15_018" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
	<title>訓練統計月報總表</title>
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
		function CheckSearch1() {
			var ddlYears = document.getElementById('ddlYears');
			var ddlMonths = document.getElementById('ddlMonths');
			var ddlDistID = document.getElementById('ddlDistID');
			//var rblSearchPlan = document.getElementById('rblSearchPlan');
			var msg = '';

			if (ddlYears && ddlYears.selectedIndex == 0) msg += '請選擇年度\n';
			if (ddlMonths && ddlMonths.selectedIndex == 0) msg += '請選擇月份\n';
			if (ddlDistID && ddlDistID.selectedIndex == 0) msg += '請選擇轄區\n';

			var SearchPlan1 = getRadioValue(document.form1.rblSearchPlan);
			if (SearchPlan1 == 'W' || SearchPlan1 == 'G') {
				//SearchPlan1 = SearchPlan; alert(SearchPlan1);
			}

			if (msg != '') {
				alert(msg);
				return false;
			}
		}

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
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;統計表&gt;&gt;訓練統計月報總表</asp:Label>
                </td>
            </tr>
        </table>
	<table class="table_sch" cellspacing="1" cellpadding="1" width="100%" border="0">
		<%--<tr>
			<td>
				<table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>
							<asp:Label ID="TitleLab1" runat="server"></asp:Label>
							<asp:Label ID="TitleLab2" runat="server">
							首頁&gt;&gt;學員動態管理&gt;&gt;統計表&gt;&gt;<FONT color="#990000">訓練統計月報總表</FONT>
							</asp:Label>
						</td>
					</tr>
				</table>
			</td>
		</tr>--%>
		<tr>
			<td>
				<table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td class="bluecol_need" style="width:20%">
							年度
						</td>
						<td class="whitecol" colspan="3">
							<asp:DropDownList ID="ddlYears" runat="server">
							</asp:DropDownList>
						</td>
					</tr>
					<tr>
						<td class="bluecol_need">
							月份
						</td>
						<td class="whitecol" colspan="3">
							<asp:DropDownList ID="ddlMonths" runat="server">
							</asp:DropDownList>
						</td>
					</tr>
					<tr>
						<td class="bluecol_need">
							轄區
						</td>
						<td class="whitecol" colspan="3">
							<asp:DropDownList ID="ddlDistID" runat="server">
							</asp:DropDownList>
						</td>
					</tr>
					<tr>
						<td class="bluecol_need">
							計畫範圍
						</td>
						<td class="whitecol" colspan="3">
							<asp:RadioButtonList ID="rblSearchPlan" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
								<asp:ListItem Value="G" Selected="True">產業人才投資計畫</asp:ListItem>
								<asp:ListItem Value="W">提升勞工自主學習計畫</asp:ListItem>
							</asp:RadioButtonList>
						</td>
					</tr>
					<tr>
						<td align="center" colspan="4">
							<%--<asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>--%>
						</td>
					</tr>
					<tr>
						<td align="center" colspan="4" class="whitecol">
							<asp:Button ID="btnPrint1" runat="server" Text="列印" Style="z-index: 0" CssClass="asp_Export_M"></asp:Button>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
	</form>
</body>
</html>
