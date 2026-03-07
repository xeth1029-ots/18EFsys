<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_04_009_R.aspx.vb" Inherits="WDAIIP.TR_04_009_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>就業追蹤統計表_依轄區</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<script language="javascript" src="../../js/date-picker.js"></script>
	<script language="javascript" src="../../js/openwin/openwin.js"></script>
	<script language="javascript" src="../../js/common.js"></script>
	<script language="javascript">
		function showHide(type) {
			//var DistID = document.getElementById('DistID');
			//var OCID = document.getElementById('OCID');
			var nxlayer_01 = document.getElementById('nxlayer_01');
			var TPlanID = document.getElementById('TPlanID');
			nxlayer_01.style.visibility = 'hidden';
			TPlanID.style.visibility = 'visible';
			if (type == 1) {
				//document.all.nxlayer_01.style.visibility = 'visible'
				//document.all.DistID.style.visibility = 'hidden'
				//document.all.OCID.style.visibility = 'hidden'
				//document.all.TPlanID.style.visibility = 'hidden'
				nxlayer_01.style.visibility = 'visible';
				TPlanID.style.visibility = 'hidden';
			}
			//else {
			//	document.all.nxlayer_01.style.visibility = 'hidden'
			//	document.all.DistID.style.visibility = 'visible'
			//	document.all.OCID.style.visibility = 'visible'
			//	document.all.TPlanID.style.visibility = 'visible'
			//}
		}

		function search() {
			var msg = '';
			var CPoint = document.getElementsByName('CPoint');
			var STDate1 = document.getElementById('STDate1');
			var STDate2 = document.getElementById('STDate2');
			var FTDate1 = document.getElementById('FTDate1');
			var FTDate2 = document.getElementById('FTDate2');

			if (!isChecked(CPoint)) msg += '請選擇就業查核點\n';

			if (STDate1.value != '') {
				if (!IsDate(STDate1.value)) msg += '開訓日期的起始日不是正確的日期格式\n';
			}

			if (STDate2.value != '') {
				if (!IsDate(STDate2.value)) msg += '開訓日期的迄日不是正確的日期格式\n';
			}

			if (FTDate1.value != '') {
				if (!IsDate(FTDate1.value)) msg += '結訓日期的起始日不是正確的日期格式\n';
			}

			if (FTDate2.value != '') {
				if (!IsDate(FTDate2.value)) msg += '結訓日期的迄日不是正確的日期格式\n';
			}

			if (msg != '') {
				alert(msg);
				return false;
			}
		}

		function SelectAll(obj, hidobj) {
			var num = getCheckBoxListValue(obj).length;
			var myallcheck = document.getElementById(obj + '_' + 0);

			if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
				document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
				for (var i = 1; i < num; i++) {
					var mycheck = document.getElementById(obj + '_' + i);
					mycheck.checked = myallcheck.checked;
				}
			}
		}

		function IsDate(MyDate) {
			if (MyDate != '') {
				if (!checkDate(MyDate))
					return false;
			}
			return true;
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
						<td title="點選可以查看說明" style="cursor: pointer" onclick="showHide(1)">
							<asp:Label ID="Label1" runat="server"></asp:Label>
							<asp:Label ID="Label2" runat="server">
									首頁&gt;&gt;訓練與就業需求管理&gt;&gt;就業追蹤報表&gt;&gt;<FONT color="#990000">就業追蹤統計表_依轄區</FONT>
							</asp:Label>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<%--
				--%>
				<table class="font" id="nxlayer_01" style="border-right: #9eb5cd 1px solid; border-top: #9eb5cd 1px solid; filter: progid: DXImageTransform.Microsoft.Shadow(Color=#919899, Strength=4, Direction=135); visibility: hidden; border-left: #9eb5cd 1px solid; border-bottom: #9eb5cd 1px solid; position: absolute" cellspacing="0" cellpadding="0" width="100%" border="0">
					<tbody>
						<tr>
							<td width="90%" bgcolor="#ffffff" align="center"><a onclick="showHide(0)" href="#"><font color="black">關閉[X]</font></a> </td>
						</tr>
						<tr>
							<td class="dashline" style="height: 1px" height="1"><u></u></td>
						</tr>
						<tr>
							<td style="padding-right: 8px; padding-left: 8px; padding-bottom: 6px; padding-top: 8px" width="100%" bgcolor="#f1faff" colspan="2">說明：<br>
								若輸入開訓期間:2006/01/01~2006/12/31 則會查出2006年開訓的班級 (含2006年開訓但結訓日跨2007年的班級)
								<br>
								若輸入結訓期間:2006/01/01~2006/12/31 則會查出2006年結訓的班級 (含2005年開訓,但是結訓日期是在2006年的班級)
								<br>
								若輸入開訓期間:2006/01/01~ 不填
								<br>
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 結訓期間: &nbsp;&nbsp;&nbsp; 不填 &nbsp; &nbsp; ~ 2006/12/31 則會查出在2006年開訓及結訓的班級 (不含跨年度)
								<br>
								以上情況的前提是:[年度] 需選擇 [====請選擇====] </td>
						</tr>
					</tbody>
				</table>
				<table class="table_sch" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
					<%--	
							<TR>
								<TD class="CM_TD1" width="100">&nbsp;&nbsp;&nbsp; 年度</TD>
								<TD class="CM_TD2"><asp:dropdownlist id="Syear" runat="server"></asp:dropdownlist></TD>
							</TR>
					--%>
					<tr>
						<td class="bluecol_need">訓練計畫 </td>
						<td class="whitecol">
							<asp:CheckBoxList ID="TPlanID" runat="server" RepeatColumns="3" CssClass="font" RepeatDirection="Horizontal" CellSpacing="0" CellPadding="0">
							</asp:CheckBoxList>
							<input id="TPlanHidden" type="hidden" value="0" name="TPlanHidden" runat="server">
						</td>
					</tr>
					<tr>
						<td class="bluecol_need">就業查核點 </td>
						<td class="whitecol">
							<asp:RadioButtonList ID="CPoint" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
								<asp:ListItem Value="0">已結訓(未達三個月)</asp:ListItem>
								<asp:ListItem Value="1">結訓三個月</asp:ListItem>
								<asp:ListItem Value="2">結訓六個月</asp:ListItem>
								<asp:ListItem Value="3">結訓九個月</asp:ListItem>
								<asp:ListItem Value="4">結訓12個月</asp:ListItem>
							</asp:RadioButtonList>
						</td>
					</tr>
					<tr>
						<td class="bluecol">開訓期間 </td>
						<td class="whitecol">
							<asp:TextBox ID="STDate1" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"><font color="#000000">～
								<asp:TextBox ID="STDate2" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></font> </td>
					</tr>
					<tr>
						<td class="bluecol">結訓期間 </td>
						<td class="whitecol">
							<asp:TextBox ID="FTDate1" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"><font color="#000000">～
								<asp:TextBox ID="FTDate2" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></font> </td>
					</tr>
					<tr>
						<td class="bluecol">&nbsp;&nbsp;身分別 </td>
						<td class="whitecol">
							<asp:CheckBoxList ID="Identity" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="4" RepeatLayout="Flow">
							</asp:CheckBoxList>
							<input id="hidIdentity" value="0" type="hidden" name="hidIdentity" runat="server">
						</td>
					</tr>
					<tr>
						<td class="bluecol">&nbsp;&nbsp;年齡 </td>
						<td class="whitecol">
							<asp:CheckBoxList ID="ddlyearsOld" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="4" RepeatLayout="Flow">
							</asp:CheckBoxList>
							<input id="hidddlyearsOld" value="0" type="hidden" name="hidddlyearsOld" runat="server">
						</td>
					</tr>
					<tr>
						<td class="bluecol">&nbsp;&nbsp;教育程度 </td>
						<td class="whitecol">
							<asp:CheckBoxList ID="ddlDegreeID" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="4" RepeatLayout="Flow">
							</asp:CheckBoxList>
							<input id="hidddlDegreeID" value="0" type="hidden" name="hidDegreeID" runat="server">
						</td>
					</tr>
					<tr>
						<td class="bluecol">&nbsp;&nbsp;性別 </td>
						<td class="whitecol">
							<asp:CheckBoxList ID="ddlSex" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="4" RepeatLayout="Flow">
							</asp:CheckBoxList>
							<input id="hidddlSex" value="0" type="hidden" name="hidddlSex" runat="server">
						</td>
					</tr>
					<%--<tr>
						<td class="bluecol">
							報表格式
						</td>
						<td class="whitecol" colspan="3">
							<asp:RadioButtonList ID="PrintStyle" runat="server" Width="65%" RepeatDirection="Horizontal" CssClass="font">
								<asp:ListItem Value="2006">舊格式(2006年度含之前)</asp:ListItem>
								<asp:ListItem Value="2007" Selected="True">新格式(2007年度含之後)</asp:ListItem>
							</asp:RadioButtonList>
						</td>
					</tr>--%>
				</table>
				<%--	<p align="center">
				</p>--%>
				<table width="740">
					<tr>
						<td align="center" colspan="4">
							<asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
							&nbsp;
							<asp:Button ID="btnExport1" runat="server" Text="匯出班級明細" CssClass="asp_Export_M"></asp:Button>
						</td>
					</tr>
					<tr>
						<td align="center" colspan="4">
							<asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
						</td>
					</tr>
				</table>
				<div id="Div1" runat="server">
					<asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Width="100%">
						<AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
						<HeaderStyle CssClass="head_navy"></HeaderStyle>
					</asp:DataGrid>
				</div>
			</td>
		</tr>
	</table>
	</form>
</body>
</html>
