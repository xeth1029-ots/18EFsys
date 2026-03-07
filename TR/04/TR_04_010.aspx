<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_04_010.aspx.vb" Inherits="WDAIIP.TR_04_010" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>學員就業成果統計表</title>
	<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
	<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
	<meta name="vs_defaultClientScript" content="JavaScript">
	<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	<link rel="stylesheet" type="text/css" href="../../CSS/style.css">
	<script src="../../js/date-picker.js" type="text/javascript"></script>
	<script language="javascript" src="../../js/openwin/openwin.js"></script>
	<script language="javascript" src="../../js/common.js"></script>
	<script language="javascript">
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

		/*
		*/
		function GetMode() {
			document.form1.center.value = '';
			document.form1.RIDValue.value = '';
			document.form1.OCIDValue.value = '';
			document.form1.PlanID.value = '';
			for (var i = document.form1.OCID.options.length - 1; i >= 0; i--) {
				document.form1.OCID.options[i] = null;
			}
			document.form1.OCID.options[0] = new Option('請選擇機構');
			if (document.form1.DistID.selectedIndex != 0 && document.form1.TPlanID.selectedIndex != 0) {
				document.form1.Button3.disabled = false;
			}
			else {
				document.form1.Button3.disabled = true;
			}
		}

		function ChangeMode(num) {
			document.getElementById('ShowMode').value = num;

			if (num == 1) {
				document.getElementById('ShowDataTable').style.display = 'inline';
				document.getElementById('ShowDataTable2').style.display = 'none';
				document.getElementById('ShowDataTable3').style.display = 'none';
			}
			else if (num == 2) {
				document.getElementById('ShowDataTable').style.display = 'none';
				document.getElementById('ShowDataTable2').style.display = 'inline';
				document.getElementById('ShowDataTable3').style.display = 'none';
			}
			else {
				document.getElementById('ShowDataTable').style.display = 'none';
				document.getElementById('ShowDataTable2').style.display = 'none';
				document.getElementById('ShowDataTable3').style.display = 'inline';
			}
		}
	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<table id="FrameTable" class="font" border="0" cellspacing="1" cellpadding="1" width="740">
		<tr>
			<td>
				<table id="Table2" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
					<tr>
						<td>
							<asp:Label ID="TitleLab1" runat="server"></asp:Label>
							<asp:Label ID="TitleLab2" runat="server">
									首頁&gt;&gt;訓練與就業需求管理&gt;&gt;就業追蹤報表&gt;&gt;<font color="#990000">學員就業成果統計表</font>
							</asp:Label>
						</td>
					</tr>
				</table>
				<table id="SearchTable" class="table_sch" border="0" cellspacing="1" cellpadding="1" width="740" runat="server">
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
							結訓日期
						</td>
						<td colspan="3" class="whitecol">
							<asp:TextBox ID="FTDate1" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30"><font color="#000000">～</font><font color="#ffffff"> </font>
							<asp:TextBox ID="FTDate2" runat="server" Columns="10"></asp:TextBox><font color="#ffffff"><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">
							</font>
						</td>
					</tr>
					<tr>
						<%--<td class="bluecol">轄區中心</td>--%>
                        <td class="bluecol">轄區分署</td>
						<td colspan="3" class="whitecol">
							<asp:DropDownList ID="DistID" runat="server">
							</asp:DropDownList>
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							訓練計畫
						</td>
						<td colspan="3" class="whitecol">
							<asp:DropDownList ID="TPlanID" runat="server">
							</asp:DropDownList>
							<font color="#ff0000">
								<br>
								(在職進修訓練、接受企業委託訓練、學習券不列入統計數據)</font>
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							訓練機構
						</td>
						<td colspan="3" class="whitecol">
							<asp:TextBox ID="center" runat="server" Width="300px"></asp:TextBox><input id="RIDValue" type="hidden" name="RIDValue" runat="server">
							<input id="Button3" onclick="javascript:wopen('../../Common/MainOrg.aspx?DistID='+document.form1.DistID.value+'&amp;TPlanID='+document.form1.TPlanID.value,'訓練機構',400,400,1)" value="..." type="button" name="Button3" runat="server" class="button_b_Mini">
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							班別關鍵字
						</td>
						<td colspan="3" class="whitecol">
							<asp:TextBox ID="classname" runat="server" Width="200px"></asp:TextBox>
							<asp:Button ID="Button6" runat="server" Text="班別查詢" CssClass="asp_button_M"></asp:Button>
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
							<input id="PlanID" type="hidden" name="PlanID" runat="server">
							<asp:Button ID="Button2" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button><br>
							(當有指定班別時，系統將會忽略開訓期間)
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							失業週數區間
						</td>
						<td bgcolor="#ecf7ff" colspan="3">
							<font color="#ffffff">
								<table id="Table1" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
									<tr>
										<td class="whitecol">
											<asp:TextBox ID="Range1" runat="server" Columns="3" MaxLength="2"></asp:TextBox>週(含)以下
										</td>
									</tr>
									<tr>
										<td class="whitecol">
											<asp:TextBox ID="Range2" runat="server" Columns="3" MaxLength="2"></asp:TextBox>週~
											<asp:TextBox ID="Range3" runat="server" Columns="3" MaxLength="2"></asp:TextBox>週
										</td>
									</tr>
									<tr>
										<td class="whitecol">
											<asp:TextBox ID="Range4" runat="server" Columns="3" MaxLength="2"></asp:TextBox>週(含)以上
										</td>
									</tr>
								</table>
							</font>
						</td>
					</tr>
				</table>
				<p align="center">
					<asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>&nbsp;
					<input id="Button4" value="列印就業率" type="button" name="Button4" runat="server" class="asp_Export_M">&nbsp;
					<input id="Button5" value="列印參考資料" type="button" name="Button4" runat="server" class="asp_Export_M">
				</p>
				<p align="center">
					<asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
				</p>
			</td>
		</tr>
	</table>
	<table id="MenuTable" class="font" border="0" cellspacing="0" cellpadding="0" height="20" runat="server">
		<tr>
			<td style="cursor: pointer" background="..\..\images\BookMark_01.gif" width="1">
			</td>
			<td style="cursor: pointer" onclick="ChangeMode(1);" background="..\..\images\BookMark_02.gif" width="113" align="center">
				就&nbsp;&nbsp; 業&nbsp;&nbsp; 率
			</td>
			<td style="cursor: pointer" background="..\..\images\BookMark_03.gif" width="11">
			</td>
			<td style="cursor: pointer" background="..\..\images\BookMark_01.gif" width="1">
			</td>
			<td style="cursor: pointer" onclick="ChangeMode(2);" background="..\..\images\BookMark_02.gif" width="113" align="center">
				參考資料
			</td>
			<td style="cursor: pointer" background="..\..\images\BookMark_03.gif" width="11">
			</td>
			<td style="cursor: pointer" id="TranId1" background="..\..\images\BookMark_01.gif" width="1" runat="server">
			</td>
			<td style="cursor: pointer" id="TranId2" onclick="ChangeMode(3);" background="..\..\images\BookMark_02.gif" width="113" align="center" runat="server">
				訓用合一專用
			</td>
			<td style="cursor: pointer" id="TranId3" background="..\..\images\BookMark_03.gif" width="11" runat="server">
			</td>
			<td style="cursor: pointer" id="TranId4" onclick="ChangeMode(3);" background="..\..\images\BookMark_02.gif" width="130" align="center" runat="server">
				與企業合作辦訓專用
			</td>
			<td style="cursor: pointer" id="TranId5" background="..\..\images\BookMark_03.gif" width="11" runat="server">
			</td>
			<td style="cursor: pointer" id="TranId6" onclick="ChangeMode(3);" background="..\..\images\BookMark_02.gif" width="240" align="center" runat="server">
				推動營造業事業單位辦理職前培訓專用
			</td>
			<td style="cursor: pointer" id="TranId7" background="..\..\images\BookMark_03.gif" width="11" runat="server">
			</td>
		</tr>
	</table>
	<asp:Table ID="ShowDataTable" runat="server" CellPadding="3" CellSpacing="1" CssClass="font">
	</asp:Table>
	<asp:Table ID="ShowDataTable2" runat="server" CellPadding="3" CellSpacing="1" CssClass="font">
	</asp:Table>
	<asp:Table ID="ShowDataTable3" runat="server" CellPadding="3" CellSpacing="1" CssClass="font">
	</asp:Table>
	<input id="ShowMode" value="1" type="hidden" runat="server">
	</form>
</body>
</html>
