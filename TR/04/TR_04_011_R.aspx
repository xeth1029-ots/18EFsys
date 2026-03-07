<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_04_011_R.aspx.vb" Inherits="WDAIIP.TR_04_011_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>學員輔導就業成果名冊</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<script language="javascript" src="../../js/date-picker.js"></script>
	<script language="javascript" src="../../js/openwin/openwin.js"></script>
	<script language="javascript" src="../../js/common.js"></script>
	<script language="javascript">
		function GetMode() {
			document.form1.center.value = '';
			document.form1.RIDValue.value = '';
			document.form1.PlanID.value = '';
			document.form1.OCIDValue.value = '';
			for (var i = document.form1.OCID.options.length - 1; i >= 0; i--) {
				document.form1.OCID.options[i] = null;
			}

			//	document.form1.OCID.options[0]=new Option('請選擇機構');
			if (document.form1.DistID.selectedIndex != 0 && document.form1.TPlanID.selectedIndex != 0) {
				document.form1.Button3.disabled = false; //選擇機構
			}
			else {
				document.form1.Button3.disabled = true; //選擇機構
			}
			document.form1.Button2.click();
		}

		function IsDate(MyDate) {
			if (MyDate != '') {
				if (!checkDate(MyDate))
					return false;
			}
			return true;
		}

		function print() {
			var msg = '';
			if (document.form1.Syear.selectedIndex == 0) msg += '請選擇年度\n';
		    //if (document.form1.DistID.selectedIndex == 0) msg += '請選擇轄區中心\n';
			if (document.form1.DistID.selectedIndex == 0) msg += '請選擇轄區分署\n';
			if (document.form1.TPlanID.selectedIndex == 0) msg += '請選擇訓練計畫\n';
			if (document.form1.RIDValue.value == '') msg += '請選擇訓練機構\n';

			if (!IsDate(document.form1.STDate1.value)) msg += '開訓日期的起始日不是正確的日期格式\n';
			if (!IsDate(document.form1.STDate2.value)) msg += '開訓日期的結束日不是正確的日期格式\n';
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
							<asp:Label ID="TitleLab1" runat="server"></asp:Label>
							<asp:Label ID="TitleLab2" runat="server">
									首頁&gt;&gt;學員動態管理&gt;&gt;就業媒合管理&gt;&gt;<font color="#990000">學員輔導就業成果名冊</font>
							</asp:Label>
						</td>
					</tr>
				</table>
				<table class="table_sch" id="SearchTable" runat="server" cellspacing="1" cellpadding="1">
					<tr>
						<td width="100" class="bluecol_need">
							年度
						</td>
						<td colspan="3" class="whitecol">
							<asp:DropDownList ID="Syear" runat="server">
							</asp:DropDownList>
						</td>
					</tr>
					<tr>
						<%--<td width="100" class="bluecol_need">轄區中心</td>--%>
                        <td width="100" class="bluecol_need">轄區分署</td>
						<td colspan="3" class="whitecol">
							<asp:DropDownList ID="DistID" runat="server">
							</asp:DropDownList>
						</td>
					</tr>
					<tr>
						<td class="bluecol_need">
							訓練計畫
						</td>
						<td colspan="3" class="whitecol">
							<asp:DropDownList ID="TPlanID" runat="server">
							</asp:DropDownList>
						</td>
					</tr>
					<tr>
						<td class="bluecol_need">
							訓練機構
						</td>
						<td colspan="3" class="whitecol">
							<asp:TextBox ID="center" runat="server" Width="410px" onfocus="this.blur()"></asp:TextBox>
							<input id="RIDValue" type="hidden" name="RIDValue" runat="server">
							<input id="Button3" onclick="javascript:" type="button" value="..." name="Button3" runat="server" class="button_b_Mini">
							<input id="PlanID" type="hidden" name="RIDValue" runat="server">
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							開訓期間
						</td>
						<td colspan="3" class="whitecol">
							<asp:TextBox ID="STDate1" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30"><font color="#000000">～</font><font color="#ffffff"> </font>
							<asp:TextBox ID="STDate2" runat="server" Columns="10"></asp:TextBox><font color="#ffffff"><img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">
							</font>
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							班別
						</td>
						<td colspan="3" class="whitecol">
							<asp:DropDownList ID="OCID" runat="server">
							</asp:DropDownList>
							<input id="OCIDValue" type="hidden" name="OCIDValue" runat="server">
							<asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Visible="False"></asp:TextBox>
							<asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Visible="False"></asp:TextBox>
							<input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
							<input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
							<asp:Button ID="Button2" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							報名管道
						</td>
						<td colspan="3" class="whitecol">
							<asp:CheckBoxList ID="CkEnterChannel" runat="server" RepeatColumns="4" RepeatDirection="Horizontal" Font-Size="Small">
								<asp:ListItem Value="1">網路</asp:ListItem>
								<asp:ListItem Value="2">現場</asp:ListItem>
								<asp:ListItem Value="3">通訊</asp:ListItem>
								<asp:ListItem Value="4">推介</asp:ListItem>
							</asp:CheckBoxList>
						</td>
					</tr>
				</table>
				<p align="center">
					&nbsp;
					<asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
					&nbsp;&nbsp;
					<asp:Button ID="btnExport" runat="server" Text="匯出EXCEL" CssClass="asp_Export_M"></asp:Button>
				</p>
				<p align="center">
					<asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
				</p>
			</td>
		</tr>
	</table>
	<asp:HiddenField ID="hid_CHKJOBRELATE_NG" runat="server" />
	</form>
</body>
</html>
