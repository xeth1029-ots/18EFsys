<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_09_003_R.aspx.vb" Inherits="TIMS.SD_09_003_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>列印巡查課紀錄表</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
	<meta content="JavaScript" name="vs_defaultClientScript" />
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
	<link href="../../css/style.css" type="text/css" rel="stylesheet" />
	<script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
	<script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
	<script type="text/javascript" src="../../js/common.js"></script>
	<script type="text/javascript" language="javascript">
		function choose_class() {
			wopen('../02/SD_02_ch.aspx?RID=' + document.form1.RIDValue.value, 'Class', 540, 520, 1);
		}
		function print() {
			if (document.form1.SDate.value == '') {
				alert('請選擇查課日期!!');
				return false;
			}
		}
		function search() {
			if (document.form1.OCIDValue1.value == '') {
				alert('請選擇班別!!');
				return false;
			}
		}		
	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<table id="Table1" cellspacing="1" cellpadding="1" width="600" border="0">
		<tr>
			<td>
				<table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>
							<asp:Label ID="TitleLab1" runat="server"></asp:Label>
							<asp:Label ID="TitleLab2" runat="server">
									首頁&gt;&gt;學員動態管理&gt;&gt;教務報表管理&gt;&gt;<font color="#990000">列印巡查課紀錄表</font>
							</asp:Label>
						</td>
					</tr>
				</table>
				<table class="table_sch" id="Table3">
					<tr>
						<td class="bluecol" width="100">
							訓練機構
						</td>
						<td class="whitecol">
							<asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="410px"></asp:TextBox>
							<input type="button" value="..." id="Button3" name="Button3" runat="server" class="button_b_Mini" />
							<input id="RIDValue" type="hidden" name="Hidden2" runat="server" size="1" />
							<span id="HistoryList2" style="display: none; position: absolute">
								<asp:Table ID="HistoryRID" runat="server" Width="310px">
								</asp:Table>
							</span>
						</td>
					</tr>
					<tr>
						<td class="bluecol" width="100">
							班別/職類
						</td>
						<td class="whitecol">
							<asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="210px"></asp:TextBox>
							<asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="200px"></asp:TextBox>
							<input onclick="choose_class()" type="button" value="..." class="button_b_Mini" />
							<input id="TMIDValue1" style="width: 40px; height: 22px" type="hidden" size="1" name="Hidden1" runat="server" />
							<input id="OCIDValue1" style="width: 32px; height: 22px" type="hidden" size="1" name="Hidden3" runat="server" />
							<asp:Button ID="Button2" runat="server" Text="查詢課程" Visible="False" CssClass="asp_button_S"></asp:Button>
						</td>
					</tr>
					<tr>
						<td class="bluecol_need" width="100">
							查課日期
						</td>
						<td class="whitecol">
							<asp:TextBox ID="SDate" runat="server" Width="80px"></asp:TextBox>
							<img style="cursor: hand" onclick="javascript:show_calendar('SDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="24" height="24" />
						</td>
					</tr>
					<tr>
					</tr>
				</table>
			</td>
		</tr>
	</table>
	<br />
	<div style="width: 600" align="center" class="whitecol">
		<asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_button_S"></asp:Button>
	</div>
	</form>
</body>
</html>
