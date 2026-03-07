<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_003_add.aspx.vb" Inherits="WDAIIP.SD_05_003_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>SD_05_003_add</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<script language="javascript" src="../../js/date-picker.js"></script>
	<script language="javascript" src="../../js/openwin/openwin.js"></script>
	<script language="javascript" src="../../js/common.js"></script>
	<script language="javascript">
		function search() {
			if (document.form1.OCIDValue1.value == '') {
				alert('請選擇職類班別!');
				return false;
			}
		}

		function choose_class(num) {
			if (num == 1)
			//window.open('../02/SD_02_ch.aspx?special=2&&sort='+num,'','width=550,height=520,location=0,status=0,menubar=0,scrollbars=1,resizable=0');
				window.open('../02/SD_02_ch.aspx?sort=' + num + '&BtnName=Button5', '', 'width=550,height=520,location=0,status=0,menubar=0,scrollbars=1,resizable=0');
			else
				window.open('../02/SD_02_ch.aspx?sort=' + num, '', 'width=550,height=520,location=0,status=0,menubar=0,scrollbars=1,resizable=0');
		}
		function chkdata() {
			var msg = '';
			if (document.form1.ApplyDate.value == '') msg += '請輸入轉班日期\n';
			if (document.form1.OCIDValue1.value == '') msg += '請選擇原始職類班別\n';
			if (document.form1.SOCID.selectedIndex == 0) msg += '請選擇學員\n';
			if (document.form1.OCIDValue2.value == '') msg += '請選擇要轉到的職類班別\n';
			if (document.form1.OCIDValue1.value == document.form1.OCIDValue2.value) msg += '不可以轉到同一個班級\n';

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
				<table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">轉班作業</font> </td>
					</tr>
				</table>
				<div align="center">
					<table class="table_sch" id="Table3" cellpadding="1" cellspacing="1">
						<tr>
							<td width="100" class="bluecol_need">轉班日期 </td>
							<td colspan="3" class="whitecol">
								<asp:TextBox ID="ApplyDate" runat="server" Width="75px"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('ApplyDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
								&nbsp;轉班日期超過開訓日兩週不能轉班 </td>
						</tr>
						<tr>
							<td width="100" class="bluecol_need">原職類/班級 </td>
							<td colspan="3" class="whitecol">
								<asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="210px"></asp:TextBox><asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="200px"></asp:TextBox>
								<input id="Button2" onclick="choose_class(1)" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
								<input id="TMIDValue1" type="hidden" name="Hidden2" runat="server"><input id="OCIDValue1" type="hidden" name="Hidden1" runat="server">
								<span id="HistoryList" style="display: none; left: 270px; position: absolute">
									<asp:Table ID="HistoryTable" runat="server" Width="310">
									</asp:Table>
								</span></td>
						</tr>
						<tr>
							<td width="100" class="bluecol_need">學員姓名 </td>
							<td colspan="3" class="whitecol">
								<asp:DropDownList ID="SOCID" runat="server">
								</asp:DropDownList>
							</td>
						</tr>
						<tr>
							<td width="100" class="bluecol_need">轉到職類/班級 </td>
							<td colspan="3" class="whitecol">
								<asp:TextBox ID="TMID2" runat="server" onfocus="this.blur()" Width="210px"></asp:TextBox>
								<asp:TextBox ID="OCID2" runat="server" onfocus="this.blur()" Width="200px"></asp:TextBox>
								<input id="Button3" onclick="choose_class(2)" type="button" value="..." name="Button3" runat="server" class="button_b_Mini">
								<input id="TMIDValue2" type="hidden" name="Hidden3" runat="server">
								<input id="OCIDValue2" type="hidden" name="Hidden4" runat="server">
								<span id="HistoryList2" style="display: none; left: 270px; position: absolute">
									<asp:Table ID="HistoryTable2" runat="server" Width="310">
									</asp:Table>
								</span></td>
						</tr>
						<tr>
							<td class="bluecol">轉班原因 </td>
							<td colspan="3" class="whitecol">
								<asp:TextBox ID="Reason" runat="server" Width="400px" TextMode="MultiLine" Height="80px"></asp:TextBox>
							</td>
						</tr>
					</table>
				</div>
				<p align="center">
					<asp:Button ID="Button1" runat="server" Text="儲存" CssClass="asp_button_S"></asp:Button>
					&nbsp;<input id="Button4" type="button" value="回上一頁" name="Button4" runat="server" class="button_b_S">&nbsp;
					<asp:Button ID="Button5" runat="server" Text="取得班級學員" CssClass="asp_button_M"></asp:Button></p>
			</td>
		</tr>
	</table>
	</form>
</body>
</html>
