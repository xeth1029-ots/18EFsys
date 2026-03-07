<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_008_C.aspx.vb" Inherits="WDAIIP.SD_05_008_C" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>結訓學員資料卡登錄</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
	<script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
	<script type="text/javascript" src="../../js/common.js"></script>
	<script type="text/javascript" language="javascript">
		function change() {
			//alert(document.form1.Train1.checked);
			var myobj = document.form1.elements['TrainCommend2'];

			//第一次讀取狀態
			if (!document.form1.Train1.checked && !document.form1.Train2.checked) {
				for (var i = 0; i < myobj.length; i++)
					myobj[i].disabled = true;
				document.form1.TrainComd2Other.disabled = true;
			}
			else {
				for (var i = 0; i < myobj.length; i++)
					myobj[i].disabled = document.form1.Train1.checked;
				document.form1.TrainComd2Other.disabled = document.form1.Train1.checked;
			}
		}

		function chkdata(num) {
			var msg = '';

			if (num == 2) {
				if (document.form1.UnitCode1.selectedIndex == 0) msg += '必須選擇訓練機構\n';
				if (document.form1.TPlanID1.selectedIndex == 0) msg += '必須選擇訓練計畫\n';
				if (document.form1.ClassName1.value == '') msg += '必須填寫班別\n';
			}

			if (!isChecked(document.form1.Trainice)) msg += '請選擇訓練性質\n';
			if (!document.form1.Train1.checked && !document.form1.Train2.checked) msg += '請選擇是否委託訓練\n';
			if (document.form1.Train2.checked) {
				if (!isChecked(document.form1.TrainCommend2)) msg += '請選擇委託訓練的單位\n';
				var rvalue = getRadioValue(document.form1.TrainCommend2);
				if (rvalue == '4' && document.form1.TrainComd2Other.value == '') msg += '請填寫委託訓練的單位\n';
			}
			if (!isChecked(document.form1.SchoolTime)) msg += '請選擇上課時段\n';

			if (num == 2) {
				if (document.form1.ResultCount1.value == '') msg += '必須填入總人數\n';
				if (document.form1.ResultCount1.value != '' && !isUnsignedInt(document.form1.ResultCount1.value)) msg += '總人數必須為數字\n';
				if (document.form1.ResultDate1.value == '') msg += '必須填入結訓日期\n';
				if (document.form1.ResultDate1.value != '' && !checkDate(document.form1.ResultDate1.value)) msg += '結訓日期格式不正確\n';
				if (document.form1.TrainingTHour1.value == '') msg += '必須填入訓練總時數\n';
				if (document.form1.TrainingTHour1.value != '' && !isUnsignedInt(document.form1.TrainingTHour1.value)) msg += '訓練總時數必須為數字\n';
			}

			if (msg != '') {
				alert(msg);
				return false;
			}
		}
	</script>
</head>
<body onload="change()">
	<form id="form1" method="post" runat="server">	
		<table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
			<tr>
				<td>
					<%--<table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
						<tr>
							<td>首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">結訓學員資料卡登錄</font> </td>
						</tr>
					</table>--%>
					<table class="table_sch" id="Table3" cellpadding="1" cellspacing="1">
						<tr>
							<td class="bluecol" style="width:20%">訓練機構 </td>
							<td class="whitecol">
								<asp:Label ID="UnitCode" runat="server">Label</asp:Label>
								<asp:DropDownList ID="UnitCode1" runat="server">
									<asp:ListItem Value="0">--請選擇--</asp:ListItem>
									<asp:ListItem Value="007">退輔會訓練中心</asp:ListItem>
									<asp:ListItem Value="008">青輔會青年分署</asp:ListItem>
									<asp:ListItem Value="009">農委會漁業署遠洋漁業開發中心</asp:ListItem>
									<asp:ListItem Value="010">台北市分署</asp:ListItem>
									<asp:ListItem Value="011">高雄市訓練就業中心</asp:ListItem>
								</asp:DropDownList>
								<%--
												<asp:ListItem Value="0">--請選擇--</asp:ListItem>
												<asp:ListItem Value="007">退輔會訓練中心</asp:ListItem>
												<asp:ListItem Value="008">青輔會青年分署</asp:ListItem>
												<asp:ListItem Value="009">農委會漁業署遠洋漁業開發中心</asp:ListItem>
												<asp:ListItem Value="010">台北市分署</asp:ListItem>
												<asp:ListItem Value="011">高雄市訓練就業中心</asp:ListItem>
												<asp:ListItem Value="014">新北市政府職業訓練中心</asp:ListItem>
								--%>
								<input id="UnitCodeValue" type="hidden" name="Hidden1" runat="server" size="1">
								<input id="RIDValue" type="hidden" name="Hidden1" runat="server" size="1">
							</td>
						</tr>
						<tr>
							<td class="bluecol">訓練計畫 </td>
							<td class="whitecol">
								<asp:Label ID="TPlanID" runat="server">Label</asp:Label>
								<asp:DropDownList ID="TPlanID1" runat="server">
								</asp:DropDownList>
								<input id="TPlanIDValue" type="hidden" name="Hidden2" runat="server">
							</td>
						</tr>
						<tr>
							<td class="bluecol">班別 </td>
							<td class="whitecol">
								<asp:Label ID="ClassName" runat="server">Label</asp:Label>
								<input id="OCID" type="hidden" name="Hidden1" runat="server" size="1">
								<asp:TextBox ID="ClassName1" runat="server" Width="30%"></asp:TextBox>
							</td>
						</tr>
						<tr>
							<td class="bluecol">訓練職類 </td>
							<td class="whitecol">
								<asp:Label ID="TrainName" runat="server">Label</asp:Label>
								<asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
								<input id="Button2" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." name="Button2" runat="server">
								<input id="trainValue" type="hidden" name="Hidden1" runat="server" size="1">
							</td>
						</tr>
						<tr>
							<td class="bluecol">訓練性質 </td>
							<td class="whitecol">
								<asp:RadioButtonList ID="Trainice" runat="server" RepeatDirection="Horizontal" CssClass="font">
									<asp:ListItem Value="1">職前</asp:ListItem>
									<asp:ListItem Value="2">進修</asp:ListItem>
								</asp:RadioButtonList>
							</td>
						</tr>
						<tr>
							<td class="bluecol">委託訓練 </td>
							<td class="whitecol">
								<table class="font" id="Table4" cellspacing="1" cellpadding="1" width="100%" border="0">
									<tr>
										<td>
											<asp:RadioButton ID="Train1" runat="server" CssClass="font" Text="否" GroupName="R1"></asp:RadioButton>
										</td>
										<td></td>
									</tr>
									<tr>
										<td>
											<asp:RadioButton ID="Train2" runat="server" CssClass="font" Text="是" GroupName="R1"></asp:RadioButton>
										</td>
										<td>
											<asp:RadioButtonList ID="TrainCommend2" runat="server" CssClass="font">
												<asp:ListItem Value="1">學校</asp:ListItem>
												<asp:ListItem Value="2">民間企業或法人團體</asp:ListItem>
												<asp:ListItem Value="3">公營企業</asp:ListItem>
												<asp:ListItem Value="4">其他：</asp:ListItem>
											</asp:RadioButtonList>
											<asp:TextBox ID="TrainComd2Other" runat="server" Width="40%"></asp:TextBox>
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td class="bluecol">上課時段 </td>
							<td class="whitecol">
								<asp:RadioButtonList ID="SchoolTime" runat="server" RepeatDirection="Horizontal" CssClass="font">
									<asp:ListItem Value="1">日間</asp:ListItem>
									<asp:ListItem Value="2">夜間</asp:ListItem>
								</asp:RadioButtonList>
							</td>
						</tr>
						<tr>
							<td class="bluecol">結訓人數 </td>
							<td class="whitecol">
								<asp:Label ID="ResultCount" runat="server">Label</asp:Label><asp:TextBox ID="ResultCount1" runat="server" Width="10%"></asp:TextBox>
								<asp:Label ID="Person" runat="server">人</asp:Label>
								(執行<font color="#993333">班級結訓作業</font>後，可求得正確結訓人數) </td>
						</tr>
						<tr>
							<td class="bluecol">結訓日期 </td>
							<td class="whitecol">
								<asp:Label ID="ResultDate" runat="server">Label</asp:Label><asp:TextBox ID="ResultDate1" runat="server" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('ResultDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" id="IMG1" runat="server" width="30" height="30">
							</td>
						</tr>
						<tr>
							<td class="bluecol">訓練總時數 </td>
							<td class="whitecol">
								<asp:Label ID="TrainingTHour" runat="server">Label</asp:Label><asp:TextBox ID="TrainingTHour1" runat="server" Width="10%"></asp:TextBox>
							</td>
						</tr>
						<tr>
							<%--<td width="100" class="bluecol">匯入非局屬<br />學員名冊 </td>--%>
                            <td width="100" class="bluecol">匯入非署屬<br />學員名冊 </td>
							<td class="whitecol" id="trImport1" runat="server">
								<p>
									<input id="File1" type="file" size="40" name="File1" runat="server"  accept=".csv" />
									<asp:Button ID="btnImport1" runat="server" Text="匯入學員名冊" CssClass="asp_button_M"></asp:Button>(必須為csv格式)
									<br />
									<asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="../../Doc/STUD_DATALID.zip" ForeColor="#8080FF" CssClass="font">下載整批上載格式檔</asp:HyperLink>
								</p>
							</td>
						</tr>
					</table>
					<p align="center" class="whitecol">
						<asp:Button ID="btnSaveData1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
						<asp:Button ID="btnWriteStud" runat="server" Text="填寫學員資料" CssClass="asp_button_M"></asp:Button>
						<input id="BtnBack1" type="button" value="回查詢頁面" runat="server" class="button_b_M"></p>
				</td>
			</tr>
		</table>	
	<input id="HidDLID" type="hidden" name="HidDLID" runat="server" size="1">
	</form>
</body>
</html>
