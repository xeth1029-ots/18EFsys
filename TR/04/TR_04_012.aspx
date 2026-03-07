<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_04_012.aspx.vb" Inherits="WDAIIP.TR_04_012" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>學員就業成果統計表</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<script language="javascript" src="../../js/date-picker.js"></script>
	<script language="javascript" src="../../js/openwin/openwin.js"></script>
	<script language="javascript" src="../../js/common.js"></script>
	<script>
		function search() {
			var msg = '';
			if (!IsDate(document.form1.STDate1.value)) msg += '開訓日期的起始日不是正確的日期格式\n';
			if (!IsDate(document.form1.STDate2.value)) msg += '開訓日期的結束日不是正確的日期格式\n';
			if (!IsDate(document.form1.FTDate1.value)) msg += '結訓日期的起始日不是正確的日期格式\n';
			if (!IsDate(document.form1.FTDate2.value)) msg += '結訓日期的結束日不是正確的日期格式\n';
			if (document.form1.Range1.value == '') msg += '必須輸入失業週期1\n'
			else if (!isUnsignedInt(document.form1.Range1.value)) msg += '失業週期1不是正確的數字\n';
			if (document.form1.Range2.value == '') msg += '必須輸入失業週期2\n'
			else if (!isUnsignedInt(document.form1.Range2.value)) msg += '失業週期2不是正確的數字\n';
			if (document.form1.Range3.value == '') msg += '必須輸入失業週期3\n'
			else if (!isUnsignedInt(document.form1.Range3.value)) msg += '失業週期3不是正確的數字\n';
			if (document.form1.Range4.value == '') msg += '必須輸入失業週期4\n'
			else if (!isUnsignedInt(document.form1.Range4.value)) msg += '失業週期4不是正確的數字\n';

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
		function SelectAll(Flag) {
			var ItemCount = getCheckBoxListValue('TPlanID').length;
			for (i = 0; i < ItemCount; i++) {
				document.getElementById('TPlanID_' + i).checked = Flag;
			}
		}
	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
		<tr>
			<td>
				<table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>
							<asp:Label ID="TitleLab1" runat="server"></asp:Label>
							<asp:Label ID="TitleLab2" runat="server">
							首頁&gt;&gt;訓練與就業需求管理&gt;&gt;就業追蹤報表&gt;&gt;<font color="#990000">學員就業成果統計表</font>
							</asp:Label>
						</td>
					</tr>
				</table>
				<table class="table_sch" id="SearchTable" runat="server" cellspacing="1" cellpadding="1">
					<tr>
						<td class="bluecol" width="80">
							開訓期間
						</td>
						<td class="whitecol" colspan="3">
							<asp:TextBox ID="STDate1" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"><font color="#000000">～</font><font color="#ffffff"> </font>
							<asp:TextBox ID="STDate2" runat="server" Columns="10"></asp:TextBox><font color="#ffffff"><img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
							</font>
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							結訓日期
						</td>
						<td colspan="3" class="whitecol">
							<asp:TextBox ID="FTDate1" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"><font color="#000000">～</font><font color="#ffffff"> </font>
							<asp:TextBox ID="FTDate2" runat="server" Columns="10"></asp:TextBox><font color="#ffffff"><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
							</font>
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							訓練計畫
						</td>
						<td colspan="3" class="whitecol">
							<asp:CheckBox ID="SelectAllItem" runat="server" Text="全選"></asp:CheckBox><asp:CheckBoxList ID="TPlanID" runat="server" CellPadding="0" CellSpacing="0" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="3">
							</asp:CheckBoxList>
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							失業週數區間
						</td>
						<td bgcolor="#ecf7ff" colspan="3">
							<font color="#ffffff">
								<table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
									<tr>
										<td class="whitecol">
											<asp:TextBox ID="Range1" runat="server" Columns="3" MaxLength="2">30</asp:TextBox>週(含)以下
										</td>
									</tr>
									<tr>
										<td class="whitecol">
											<asp:TextBox ID="Range2" runat="server" Columns="3" MaxLength="2">31</asp:TextBox>週~
											<asp:TextBox ID="Range3" runat="server" Columns="3" MaxLength="2">52</asp:TextBox>週
										</td>
									</tr>
									<tr>
										<td class="whitecol">
											<asp:TextBox ID="Range4" runat="server" Columns="3" MaxLength="2">53</asp:TextBox>週(含)以上
										</td>
									</tr>
								</table>
							</font>
						</td>
					</tr>
				</table>
				<p align="center">
					<asp:Button ID="Button1" runat="server" Text="查詢"></asp:Button></p>
				<p align="center">
					<asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></p>
				<asp:Table ID="ShowDataTable" runat="server" CellPadding="3" CellSpacing="1" CssClass="font">
				</asp:Table>
			</td>
		</tr>
	</table>
	</form>
</body>
</html>
