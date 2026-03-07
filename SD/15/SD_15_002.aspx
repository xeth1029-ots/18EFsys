<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_002.aspx.vb" Inherits="WDAIIP.SD_15_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>全國開課明細表</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
	<script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
	<script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
	<script type="text/javascript" language="javascript">
		function GETvalue() {
			document.getElementById('Button5').click();
		}
		/**
		function OpenOrg(vTPlanID){
		if(document.getElementById('DistID').selectedIndex==0){
		alert('請先選擇轄區');
		return false;
		}
		else{
		wopen('../../common/MainOrg.aspx?DistID='+document.getElementById('DistID').value+'&TPlanID='+vTPlanID,'',400,400,'yes');
		}
		}
		**/
		function choose_class() {
			document.getElementById('OCID1').value = '';
			document.getElementById('TMID1').value = '';
			document.getElementById('OCIDValue1').value = '';
			document.getElementById('TMIDValue1').value = '';

			openClass('../02/SD_02_ch.aspx?&RID=' + document.getElementById('RIDValue').value);
		}
		function CheckPrint() {
			var STDate1 = document.getElementById('STDate1').value;
			var STDate2 = document.getElementById('STDate2').value;
			var DistID = document.getElementById('DistID').value;
			//var PlanID=document.getElementById('PlanID').value;
			var RID = document.getElementById('RIDValue').value;

			var msg = '';
			if (!checkDate(STDate1) && STDate1 != '') msg += '開訓起始日期必須為正確日期格式\n';
			if (!checkDate(STDate2) && STDate2 != '') msg += '開訓結束日期必須為正確日期格式\n';

			if (msg != '') {
				alert(msg);
				return false;
			}

			//SD_15_002
			openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=BussinessTrain&filename=SD_15_002&path=TIMS&STDate1=' + STDate1 + '&STDate2=' + STDate2 + '&DistID=' + DistID + '&RID=' + RID + '&OCID=' + document.getElementById('OCIDValue1').value, '', '');
		}
	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<table class="font" id="Table1" cellspacing="1" cellpadding="1" width="740" border="0">
		<tr>
			<td>
				<table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>
							<asp:Label ID="TitleLab1" runat="server"></asp:Label>
							<asp:Label ID="TitleLab2" runat="server">
							首頁&gt;&gt;學員動態管理&gt;&gt;產學訓統計表&gt;&gt;<FONT color="#990000">全國開課明細表</FONT>
							</asp:Label>
						</td>
					</tr>
				</table>
				<table class="table_sch">
					<tr>
						<td class="bluecol" width="100">開訓期間 </td>
						<td class="whitecol">
							<asp:TextBox ID="STDate1" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">～
							<asp:TextBox ID="STDate2" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
						</td>
					</tr>
					<tr>
						<td class="bluecol">轄區 </td>
						<td class="whitecol"><font face="新細明體">
							<asp:DropDownList ID="DistID" runat="server">
							</asp:DropDownList>
						</font></td>
					</tr>
					<tr>
						<td class="bluecol" width="100">訓練機構 </td>
						<td class="whitecol" colspan="3">
							<asp:TextBox ID="center" runat="server" Width="410px"></asp:TextBox>
							<input id="RIDValue" type="hidden" name="Hidden2" runat="server">
							<input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
							<asp:Button ID="Button5" Style="display: none" runat="server" Text="Button5"></asp:Button>
							<span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
								<asp:Table ID="HistoryRID" runat="server" Width="310px">
								</asp:Table>
							</span></td>
					</tr>
					<tr>
						<td class="bluecol">職類/班別 </td>
						<td class="whitecol" colspan="3">
							<asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="210px"></asp:TextBox>
							<asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="200px"></asp:TextBox>
							<input onclick="choose_class()" type="button" value="..." class="button_b_Mini">
							<input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
							<input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
							<span id="HistoryList" style="position: absolute; width: 312px; display: none; height: 32px; left: 270px">
								<asp:Table ID="HistoryTable" runat="server" Width="310">
								</asp:Table>
							</span></td>
					</tr>
				</table>
				<p align="center">
					<input onclick="CheckPrint();" type="button" value="列印" class="asp_Export_M">
				</p>
			</td>
		</tr>
	</table>
	<input id="Years" type="hidden" name="Years" runat="server">
	</form>
</body>
</html>
