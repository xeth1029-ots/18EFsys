<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_04_008_R.aspx.vb" Inherits="WDAIIP.TR_04_008_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>就業追蹤統計表_依班別</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<script language="javascript" src="../../js/date-picker.js"></script>
	<script language="javascript" src="../../js/openwin/openwin.js"></script>
	<script language="javascript" src="../../js/common.js"></script>
	<script language="javascript">
		function SelectAll(strobj, hidobj) {
			var MyValue = getCheckBoxListValue(strobj);
			var MyAllCheck = document.getElementById(strobj + '_0');
			var HidObj = document.getElementById(hidobj);

			//比對第1個值不相等
			if (HidObj.value != MyAllCheck.checked.toString()) {
				HidObj.value = MyAllCheck.checked;
				for (var i = 1; i < MyValue.length; i++) {
					var MyCheck = document.getElementById(strobj + '_' + i);
					MyCheck.checked = MyAllCheck.checked;
				}
			}
			else {
				//比對第1個值相等
				for (var i = 1; i < MyValue.length; i++) {
					var MyCheck = document.getElementById(strobj + '_' + i);
					if (!MyCheck.checked) { MyAllCheck.checked = false; HidObj.value = MyAllCheck.checked; break; }
				}
			}
		}

		function showHide(type) {
			var nxlayer_01 = document.getElementById('nxlayer_01');
			var DistID = document.getElementById('DistID');
			var OCIDList = document.getElementById('OCIDList');
			var TPlanID = document.getElementById('TPlanID');

			if (type == 1) {
				DistID.style.visibility = 'visible';
				OCIDList.style.visibility = 'visible';
				TPlanID.style.visibility = 'visible';
				if (nxlayer_01) {
					nxlayer_01.style.visibility = 'visible';
					DistID.style.visibility = 'hidden';
					OCIDList.style.visibility = 'hidden';
					TPlanID.style.visibility = 'hidden';
				}
			}
			else {
				//document.all.Syear.style.visibility='visible';
				//document.all.OCID.style.visibility='visible';
				if (nxlayer_01) {
					nxlayer_01.style.visibility = 'hidden';
					DistID.style.visibility = 'visible';
					OCIDList.style.visibility = 'visible';
					TPlanID.style.visibility = 'visible';
				}
			}
		}

		function GetMode() {
			var ocidObj = document.getElementById('OCIDList');
			var center = document.getElementById('center');
			var RIDValue = document.getElementById('RIDValue');
			var OCIDValue = document.getElementById('OCIDValue');

			center.value = '';
			RIDValue.value = '';
			OCIDValue.value = '';
			for (var i = 0; i < ocidObj.options.length; i++) {
				ocidObj.options[i] = null;
			}
			ocidObj.options[0] = new Option('請選擇機構');

			document.form1.Button3.disabled = true; //不啟用
			if (document.form1.DistID.selectedIndex != 0
				&& document.form1.TPlanID.selectedIndex != 0) {
				document.form1.Button3.disabled = false; //啟用
			}
		}

		function print() {
			var msg = '';
			var ocidObj = document.getElementById('OCIDList');
			var ocidVal = '';
			for (var i = 0; i < ocidObj.options.length; i++) {
				if (ocidObj.options[i].selected && ocidObj.options[i].value != '') {
					if (ocidVal != '') ocidVal += ',';
					ocidVal += ocidObj.options[i].value;
				}
			}
			//debugger;
			//if(document.form1.Syear.selectedIndex==0) msg+='請選擇年度\n';
		    //if (document.form1.DistID.selectedIndex == 0) msg += '請選擇轄區中心\n';
			if (document.form1.DistID.selectedIndex == 0) msg += '請選擇轄區分署\n';
			if (document.form1.TPlanID.selectedIndex == 0) msg += '請選擇訓練計畫\n';

			if (isEmpty(document.form1.CPoint)) {
				msg += '請選擇就業查核點!\n';
			}
			if (ocidObj.selectedIndex == -1 || ocidVal == '') {
				if (document.form1.STDate1.value == '' && document.form1.FTDate1.value == '')
					msg += '開訓或結訓起日不可空白!\n';
			}

			if (document.form1.STDate1.value != '') {
				if (!IsDate(document.form1.STDate1.value)) msg += '開訓日期的起始日不是正確的日期格式\n';
			}

			if (document.form1.STDate2.value != '') {
				if (!IsDate(document.form1.STDate2.value)) msg += '開訓日期的迄日不是正確的日期格式\n';
			}

			if (document.form1.FTDate1.value != '') {
				if (!IsDate(document.form1.FTDate1.value)) msg += '結訓日期的起始日不是正確的日期格式\n';
			}

			if (document.form1.FTDate2.value != '') {
				if (!IsDate(document.form1.FTDate2.value)) msg += '結訓日期的迄日不是正確的日期格式\n';
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

	
	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<table id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
		<tr>
			<td>
				<table class="font" id="nxlayer_01" style="border-bottom: #9eb5cd 1px solid; position: absolute; filter: progid: DXImageTransform.Microsoft.Shadow(Color=#919899, Strength=4, Direction=135); border-left: #9eb5cd 1px solid; visibility: hidden; border-top: #9eb5cd 1px solid; border-right: #9eb5cd 1px solid" cellspacing="0" cellpadding="0" width="100%" border="0">
					<tbody>
						<tr>
							<td width="90%" bgcolor="#ffffff">首頁&gt;&gt;訓練與就業需求管理&gt;&gt;就業追蹤報表&gt;&gt;<font color="#990000">就業追蹤統計表_依班別</font> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a onclick="showHide(0)" href="#"><font color="black">關閉</font></a> </td>
						</tr>
						<tr>
							<td class="dashline" style="height: 1px" height="1"><u></u></td>
						</tr>
						<tr>
							<td style="padding-bottom: 6px; padding-left: 8px; padding-right: 8px; padding-top: 8px" width="100%" bgcolor="#f1faff" colspan="2">說明：<br>
								若輸入開訓期間:2006/01/01~2006/12/31 則會查出2006年開訓的班級 (含2006年開訓但結訓日跨2007年的班級)
								<br>
								若輸入結訓期間:2006/01/01~2006/12/31 則會查出2006年結訓的班級 (含2005年開訓,但是結訓日期是在2006年的班級)
								<br>
								若輸入開訓期間:2006/01/01~ 不填
								<br>
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 結訓期間: &nbsp;&nbsp;&nbsp; 不填 &nbsp; &nbsp; ~ 2006/12/31 則會查出在2006年開訓及結訓的班級 (不含跨年度) </td>
						</tr>
					</tbody>
				</table>
				<table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td title="點選可以查看說明" style="cursor: pointer" onclick="showHide(1)">首頁&gt;&gt;訓練與就業需求管理&gt;&gt;就業追蹤報表&gt;&gt;<font color="#990000">就業追蹤統計表_依班別</font> </td>
					</tr>
				</table>
				<table class="table_sch" id="SearchTable" runat="server" cellspacing="1" cellpadding="1">
					<tr>
						<%--<td class="bluecol_need">轄區中心 </td>--%>
                        <td class="bluecol_need">轄區分署 </td>
						<td class="whitecol" colspan="3">
							<asp:DropDownList ID="DistID" runat="server">
							</asp:DropDownList>
						</td>
					</tr>
					<tr>
						<td class="bluecol_need">訓練計畫 </td>
						<td class="whitecol" colspan="3">
							<asp:DropDownList ID="TPlanID" runat="server">
							</asp:DropDownList>
						</td>
					</tr>
					<tr>
						<td class="bluecol">轄區縣市 </td>
						<td class="whitecol" colspan="3">
							<asp:CheckBoxList ID="SCTID1" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="8">
							</asp:CheckBoxList>
							<input id="HidSCTID1" type="hidden" runat="server">
						</td>
					</tr>
					<tr>
						<td class="bluecol">上課縣市 </td>
						<td class="whitecol" colspan="3">
							<asp:CheckBoxList ID="SCTID2" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="8">
							</asp:CheckBoxList>
							<input id="HidSCTID2" type="hidden" runat="server">
						</td>
					</tr>
					<tr>
						<td class="bluecol">訓練機構 </td>
						<td class="whitecol" colspan="3">
							<asp:TextBox ID="center" runat="server" Width="410px"></asp:TextBox>
							<input id="RIDValue" type="hidden" runat="server">
							<input id="Button3" onclick="javascript:wopen('../../Common/MainOrg.aspx?DistID='+document.form1.DistID.value+'&amp;TPlanID='+document.form1.TPlanID.value,'訓練機構',400,400,1)" type="button" value="..." name="Button3" runat="server" class="button_b_Mini">
						</td>
					</tr>
					<tr>
						<td class="bluecol">班別 </td>
						<td class="whitecol" colspan="3">
							<input id="OCIDValue" type="hidden" name="OCIDValue" runat="server" size="1">
							<asp:Button ID="Button2" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>
							(當有指定班別時(含全選)，系統將會忽略開、結訓期間)
							<br>
							<asp:ListBox ID="OCIDList" runat="server" Width="410px" SelectionMode="Multiple" Rows="6"></asp:ListBox>
						</td>
					</tr>
					<tr>
						<td class="bluecol">班別關鍵字 </td>
						<td class="whitecol" colspan="3">
							<asp:TextBox Style="z-index: 0" ID="txtSearchClassName" runat="server" Width="210px"></asp:TextBox>
						</td>
					</tr>
					<tr>
						<td class="bluecol">就業查核點 </td>
						<td class="whitecol" colspan="3">
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
					<%--<tr>
						<td class="bluecol">
							報表格式
						</td>
						<td bgcolor="#ecf7ff" colspan="3" class="whitecol">
							<asp:RadioButtonList ID="PrintStyle" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
								<asp:ListItem Value="2006">舊格式(2006年度含之前)</asp:ListItem>
								<asp:ListItem Value="2007" Selected="True">新格式(2007年度含之後)</asp:ListItem>
							</asp:RadioButtonList>
						</td>
					</tr>--%>
				</table>
				<table width="740">
					<tr>
						<td align="center" colspan="4">
							<asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>&nbsp;
							<asp:Button ID="btnExport1" runat="server" Text="匯出Excel" CssClass="asp_Export_M"></asp:Button>
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
				<input id="PlanID" type="hidden" name="PlanID" runat="server">
			</td>
		</tr>
	</table>
	</form>
</body>
</html>
