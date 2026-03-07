<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="CP_02_012_R.aspx.vb" Inherits="WDAIIP.CP_02_012_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>辦理職業訓練累計週報表-按性別</title>
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
			var msg = '';
			if (isEmpty(document.form1.RIDValue)) {
				msg += '請選擇訓練機構!\n';
			}

			if (document.form1.Tplan.selectedIndex == 0) {
				msg += '請選擇訓練計畫!\n';
			}

			if (msg != '') {
				alert(msg);
				return false;
			}
		}

		function ShowFrame() {
			document.getElementById('FrameObj').style.display = document.getElementById('HistoryList2').style.display
		}
	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
		<tr>
			<td>
				<table id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
					<tr>
						<td>
							<asp:Label ID="TitleLab1" runat="server"></asp:Label>
							<asp:Label ID="TitleLab2" runat="server">
									首頁&gt;&gt;訓練查核與績效管理&gt;&gt;公務統計報表&gt;&gt;<font color="#990000">辦理職業訓練累計週報表-按性別</font>
							</asp:Label>
						</td>
					</tr>
				</table>
				<table id="Table3" class="table_sch" cellspacing="1" cellpadding="1">
					<tr>
						<td width="100" class="bluecol">
							年度
						</td>
						<td class="whitecol">
							<asp:DropDownList ID="syear" runat="server">
							</asp:DropDownList>
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							開訓日期
						</td>
						<td class="whitecol">
							<asp:TextBox ID="start_date" runat="server" Width="100px"></asp:TextBox>
							<img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
							～<asp:TextBox ID="end_date" runat="server" Width="100px"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							結訓日期
						</td>
						<td class="whitecol">
							<asp:TextBox ID="start_date1" runat="server" Width="100px"></asp:TextBox>
							<img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
							～<asp:TextBox ID="end_date1" runat="server" Width="100px"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
						</td>
					</tr>
					<tr>
						<td width="100" class="bluecol_need">
							訓練機構
						</td>
						<td class="whitecol">
							<asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox><input type="button" value="..." id="Button1" name="Button1" runat="server" class="button_b_Mini"><input id="RIDValue" type="hidden" name="Hidden1" runat="server" size="1"><br>
							<span id="HistoryList2" style="display: none; z-index: 1; position: absolute">
								<asp:Table ID="HistoryRID" runat="server" Width="310px">
								</asp:Table>
							</span>
						</td>
					</tr>
					<tr>
						<td bgcolor="#cc6666" width="100" class="bluecol_need">
							訓練計畫
						</td>
						<td class="whitecol">
							<asp:DropDownList ID="Tplan" runat="server">
							</asp:DropDownList>
							<iframe id="FrameObj" style="display: none; left: 115px; width: 310px; position: absolute; height: 23px" scrolling="no" frameborder="0"></iframe>
						</td>
					</tr>
				</table>
				<p align="center">
					<asp:Button ID="Button2" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
				</p>
			</td>
		</tr>
	</table>
	</form>
</body>
</html>
