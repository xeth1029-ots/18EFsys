<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_04_024_R.aspx.vb" Inherits="WDAIIP.TR_04_024_R" %>

 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head runat="server">
	<title>職前訓練年度就業率分析</title>
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<script type="text/javascript" language="javascript" src="../../js/common.js"></script>
	<script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
	<script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
	<script type="text/javascript" language="javascript">

		function chk() {
			var msg = '';
			var STDate1 = document.getElementById("STDate1");
			var STDate2 = document.getElementById("STDate2");
			var FTDate1 = document.getElementById("FTDate1");
			var FTDate2 = document.getElementById("FTDate2");

			if (STDate1.value != '') {
				if (!IsDate(STDate1.value)) msg += '開訓日期的起始日 不是正確的日期格式\n';
			}
			if (STDate2.value != '') {
				if (!IsDate(STDate2.value)) msg += '開訓日期的迄日 不是正確的日期格式\n';
			}
			if (FTDate1.value != '') {
				if (!IsDate(FTDate1.value)) msg += '結訓日期的起始日 不是正確的日期格式\n';
			}
			if (FTDate2.value != '') {
				if (!IsDate(FTDate2.value)) msg += '結訓日期的迄日 不是正確的日期格式\n';
			}
			if (msg != '') {
				alert(msg);
				return false;
			}
		}

		function IsDate(MyDate) {
			if (MyDate != '') {
				if (!checkDate(MyDate))
					return false;
			}
			return true;
		}


		//選擇全部
		function SelectAll(obj, hidobj) {
			var num = getCheckBoxListValue(obj).length; //長度
			var myallcheck = document.getElementById(obj + '_' + 0); //第1個

			if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
				document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
				for (var i = 1; i < num; i++) {
					var mycheck = document.getElementById(obj + '_' + i);
					mycheck.checked = myallcheck.checked;
				}
			}
			else {
				for (var i = 1; i < num; i++) {
					if ('0' == getCheckBoxListValue(obj).charAt(i)) {
						document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(i);
						var mycheck = document.getElementById(obj + '_' + i);
						myallcheck.checked = mycheck.checked;
						break;
					}
				}
			}
		}
	</script>
</head>
<body>
	<form id="form1" runat="server">
	<table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
		<tr>
			<td>
				<table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>
							<asp:Label ID="TitleLab1" runat="server"></asp:Label>
							<asp:Label ID="TitleLab2" runat="server">
							首頁&gt;&gt;訓練需求管理&gt;&gt;就業追蹤報表&gt;&gt;<font color="#990000">職前訓練年度就業率分析</font>
							</asp:Label>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table class="table_sch" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr id="Year_TR" runat="server">
						<td class="bluecol" width="100">
							年度
						</td>
						<td class="whitecol" colspan="3">
							<asp:DropDownList ID="yearlist" runat="server">
							</asp:DropDownList>
						</td>
					</tr>
					<tr id="DistID_TR" runat="server">
						<td class="bluecol">
							轄區
						</td>
						<td class="whitecol" colspan="3">
							<asp:CheckBoxList ID="DistID" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
							</asp:CheckBoxList>
							<input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server" size="1">
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							開訓區間
						</td>
						<td class="whitecol" colspan="3">
							<asp:TextBox ID="STDate1" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">~<asp:TextBox ID="STDate2" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							結訓區間
						</td>
						<td class="whitecol" colspan="3">
							<asp:TextBox ID="FTDate1" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">~<asp:TextBox ID="FTDate2" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
						</td>
					</tr>
					<tr id="TPlanID0_TR" runat="server">
						<td class="bluecol">
							訓練計畫(職前)
						</td>
						<td class="whitecol" colspan="3">
							<asp:CheckBoxList ID="chkTPlanID0" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3" CellPadding="0" CellSpacing="0">
							</asp:CheckBoxList>
							<input id="TPlanID0HID" type="hidden" value="0" name="TPlanID0HID" runat="server" size="1">
						</td>
					</tr>
					<tr id="TPlanID1_TR" runat="server">
						<td class="bluecol">
							訓練計畫(在職)
						</td>
						<td class="whitecol" colspan="3">
							<asp:CheckBoxList ID="chkTPlanID1" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3" CellPadding="0" CellSpacing="0">
							</asp:CheckBoxList>
							<input id="TPlanID1HID" type="hidden" value="0" name="TPlanID1HID" runat="server" size="1">
						</td>
					</tr>
					<tr id="TPlanIDX_TR" runat="server">
						<td class="bluecol">
							訓練計畫(其他)
						</td>
						<td class="whitecol" colspan="3">
							<asp:CheckBoxList ID="chkTPlanIDX" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3" CellPadding="0" CellSpacing="0">
							</asp:CheckBoxList>
							<input id="TPlanIDXHID" type="hidden" value="0" name="TPlanIDXHID" runat="server" size="1">
						</td>
					</tr>
				</table>
				<p align="center">
					<asp:Button ID="btnPrint1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
					<%--<asp:Button ID="Query" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>--%>
				</p>
			</td>
		</tr>
		 
	</table>
	</form>
</body>
</html>
