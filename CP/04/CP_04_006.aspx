<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_04_006.aspx.vb" Inherits="WDAIIP.CP_04_006" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>CP_04_006</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<style type="text/css">
		.class_link A { color: #000000; }
		.class_link A:link { color: #0000ff; }
		.class_link A:hover { color: #0000ff; }
		A:visited { color: #0000ff; }
		A:active { color: #0000ff; }
	</style>
	<script type="text/javascript" src="../../js/common.js"></script>
	<script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
	<script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
	<script type="text/javascript" language="javascript">

		//選擇全部
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


		//檢查日期格式
		function check_date() {
			if (!checkDate(form1.STDate1.value) || !checkDate(form1.STDate2.value)) {
				document.form1.STDate1.value = '';
				document.form1.STDate2.value = '';
				alert('請輸入正確的日期格式,YYYY/MM/DD!!\n');
			}

			if (!checkDate(form1.FTDate1.value) || !checkDate(form1.FTDate2.value)) {
				document.form1.FTDate1.value = '';
				document.form1.FTDate2.value = '';
				alert('請輸入正確的日期格式,YYYY/MM/DD!!\n');
			}

		}

		//檢查開訓及結訓日期為
		function search() {

			var msg = '';

			if (document.form1.STDate1.value != '') {
				if ((document.form1.FTDate1.value == '') && (document.form1.FTDate2.value == '') && (document.form1.STDate2.value == '')) msg += '請選擇開訓迄日!\n';

			}

			if (document.form1.STDate2.value != '') {

				if ((document.form1.FTDate1.value == '') && (document.form1.FTDate2.value == '') && (document.form1.STDate1.value == '')) msg += '請選擇開訓起日!\n';

			}

			if (document.form1.FTDate1.value != '') {

				if ((document.form1.STDate1.value == '') && (document.form1.STDate2.value == '') && (document.form1.FTDate2.value == '')) msg += '請選擇結訓訖日!\n';

			}

			if (document.form1.FTDate2.value != '') {
				if ((document.form1.STDate1.value == '') && (document.form1.STDate2.value == '') && (document.form1.FTDate1.value == '')) msg += '請選擇結訓起日!\n';
			}

			if ((document.form1.STDate1.value == '') && (document.form1.STDate2.value == '') && (document.form1.FTDate1.value == '') && (document.form1.FTDate2.value == '')) msg += '開訓日期及結訓日期請擇一輸入!\n';

			if (msg != '') {
				alert(msg);
				return false;
			}
		}		    
	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<table cellspacing="1" cellpadding="1" width="740" border="0">
		<tr>
			<td>
				<table class="font" id="Table9" width="100%">
					<tr>
						<td>
							<font class="font" size="2">首頁&gt;&gt;訓練查核與績效管理&gt;&gt;訓練資料查詢&gt;&gt;<font color="#800000">各縣市政府班級統計表</font></font>
						</td>
					</tr>
				</table>
				<table class="table_sch" id="Table1" cellspacing="1" cellpadding="1">
					<tbody>
						<tr>
							<td style="width: 12%; height: 14px" class="bluecol">
								年度
							</td>
							<td style="height: 14px" class="whitecol">
								<asp:DropDownList ID="yearlist" runat="server">
								</asp:DropDownList>
							</td>
						</tr>
						<tr id="Dist" runat="server">
							<td style="width: 10%" class="bluecol">
								轄區
							</td>
							<td class="whitecol">
								<asp:CheckBoxList ID="DistrictList" runat="server" CssClass="font" RepeatDirection="Horizontal" Width="512px" Height="11px">
								</asp:CheckBoxList>
								<input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server">
							</td>
						</tr>
						<tr>
							<td style="width: 10%" class="bluecol">
								訓練計畫
							</td>
							<td class="whitecol">
								<asp:CheckBoxList ID="PlanList" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3">
								</asp:CheckBoxList>
								<input id="TPlanHidden" type="hidden" value="0" name="TPlanHidden" runat="server"><font face="新細明體">&nbsp;</font>
							</td>
							<tr>
								<td class="bluecol">
									開訓期間
								</td>
								<td class="whitecol">
									<asp:TextBox ID="STDate1" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"><font color="#000000">～
										<asp:TextBox ID="STDate2" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></font>
								</td>
							</tr>
							<tr>
								<td class="bluecol">
									結訓期間
								</td>
								<td class="whitecol">
									<asp:TextBox ID="FTDate1" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"><font color="#000000">～
										<asp:TextBox ID="FTDate2" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></font>
								</td>
							</tr>
							<%--
						            <TR>
						            <TD style="WIDTH: 10%" bgColor="#cc6666"><FONT class="font" face="新細明體" color="#ffffff" size="2">報表類別</FONT></TD>
						            <TD bgColor="#ffecec">
							            <TABLE id="Table5" style="WIDTH: 100%; HEIGHT: 52px" cellSpacing="1" cellPadding="1" width="536"
								            border="0">
								            <TR>
								            </TR>
							            </TABLE>
							            <asp:radiobuttonlist id="PrintStaus" runat="server" CssClass="font" RepeatDirection="Horizontal">
								            <asp:ListItem Value="1" Selected="True">縣市</asp:ListItem>
								            <asp:ListItem Value="2">鄉鎮市區</asp:ListItem>
							            </asp:radiobuttonlist></TD>
					            </TR>--%>
						</tr>
					</tbody>
				</table>
			</td>
		</tr>
	</table>
	<table class="font" id="Table6" cellspacing="0" cellpadding="0" width="600" border="0">
		<tr align="center">
			<td>
				<font face="新細明體">
					<asp:Button ID="print" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button></font>
			</td>
		</tr>
	</table>
	</form>
</body>
</html>
