<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CM_03_011.aspx.vb" Inherits="WDAIIP.CM_03_011" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>交叉分析統計表</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<script  type="text/javascript"language="javascript" src="../../js/date-picker.js"></script>
	<script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
	<script  type="text/javascript"language="javascript" src="../../js/common.js"></script>
	<script type="text/javascript">
		//function OpenOrg(){
		//if(document.getElementById('DistID').selectedIndex==0){
		//alert('請先選擇轄區');
		//return false;
		//}
		//else{
		//wopen('../../common/MainOrg.aspx?DistID='+document.getElementById('DistID').value+'&TPlanID=28','',400,400,'Yes');
		//}
		//}
		function CheckSearch() {
			var STDate1 = document.getElementById('STDate1').value;
			var STDate2 = document.getElementById('STDate2').value;

			var msg = '';
			if (!checkDate(STDate1) && STDate1 != '') msg += '開訓起始日期必須為正確日期格式\n';
			if (!checkDate(STDate2) && STDate2 != '') msg += '開訓結束日期必須為正確日期格式\n';

			if (msg != '') {
				alert(msg);
				return false;
			}

			if (!isChecked(document.form1.StudStatus)) {
				alert('請選擇統計範圍');
				return false;
			}

			if (!isChecked(document.form1.XRoll) || !isChecked(document.form1.YRoll)) {
				alert('請選擇XY軸分析項目');
				return false;
			}
		}

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
			
	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
		<tr>
			<td>
				<table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>
							<font face="新細明體">首頁&gt;&gt;訓練與就業需求管理&gt;&gt;統計分析&gt;&gt;<font color="#990000">交叉分析統計表</font></font>
						</td>
					</tr>
				</table>
				<table class="table_sch" id="Table2" cellspacing="1" cellpadding="1" >
					<tr>
						<td class="bluecol" width="100">
							計畫年度
						</td>
						<td class="whitecol">
							<asp:DropDownList ID="ddlYear" runat="server">
							</asp:DropDownList>
						</td>
					</tr>
					<tr>
						<td class="bluecol" width="100">
							轄區
						</td>
						<td class="whitecol">
							<asp:DropDownList ID="DistID" runat="server" Width="310px">
							</asp:DropDownList>
						</td>
					</tr>
					<tr>
						<td class="bluecol" width="100">
							開訓期間
						</td>
						<td class="whitecol">
							<font face="新細明體">
								<asp:TextBox ID="STDate1" runat="server" Columns="10"></asp:TextBox>
                                <span runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>～
								<asp:TextBox ID="STDate2" runat="server" Columns="10"></asp:TextBox>
                                <span runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span> </font>
						</td>
					</tr>
					<tr>
						<td class="bluecol" width="100">
							結訓期間
						</td>
						<td class="whitecol">
							<font face="新細明體">
								<asp:TextBox ID="FTDate1" runat="server" Columns="10"></asp:TextBox>
                                <span runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>～
								<asp:TextBox ID="FTDate2" runat="server" Columns="10"></asp:TextBox>
                                <span runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span> </font>
						</td>
					</tr>
					<tr>
						<td class="bluecol" style="width: 109px" width="109">
							計畫範圍
						</td>
						<td class="whitecol" colspan="4">
							<asp:CheckBoxList ID="TPlanID" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="3" CellSpacing="0" CellPadding="0">
							</asp:CheckBoxList>
							<input id="TPlanHidden" type="hidden" value="0" name="TPlanHidden" runat="server">
						</td>
					</tr>
					<tr>
						<td class="bluecol" width="100">
							訓練機構
						</td>
						<td class="whitecol">
							<asp:TextBox ID="center" runat="server" Width="310px"></asp:TextBox>
							<input id="Button3" type="button" value="..." name="Button1" runat="server" class="button_b_Mini"><input id="RIDValue" type="hidden" name="RIDValue" runat="server"><input id="PlanID" type="hidden" name="PlanID" runat="server">
						</td>
					</tr>
					<tr>
						<td class="bluecol_need" style="height: 24px">
							統計範圍
						</td>
						<td class="whitecol" style="height: 24px">
							<asp:RadioButtonList ID="StudStatus" runat="server" RepeatDirection="Horizontal" CssClass="font" CellSpacing="1" CellPadding="1" AutoPostBack="True">
								<asp:ListItem Value="11">報名人數</asp:ListItem>
								<asp:ListItem Value="12">開訓人數</asp:ListItem>
								<asp:ListItem Value="13">結訓人數</asp:ListItem>
								<asp:ListItem Value="14">就業人數</asp:ListItem>
								<asp:ListItem Value="15">在職者(托育及照服員計畫)</asp:ListItem>
							</asp:RadioButtonList>
						</td>
					</tr>
					<tr>
						<td class="bluecol_need" >
							X軸
						</td>
						<td class="whitecol" >
							<asp:RadioButtonList ID="XRoll" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="4" CellSpacing="1" CellPadding="1" Width="100%">
								<asp:ListItem Value="1">性別</asp:ListItem>
								<asp:ListItem Value="2">年齡</asp:ListItem>
								<asp:ListItem Value="3">教育程度</asp:ListItem>
								<asp:ListItem Value="4">身份別</asp:ListItem>
								<asp:ListItem Value="5">受訓學員(通訊)地理分佈</asp:ListItem>
								<%--<asp:ListItem Value="6">受訓學員(戶籍)地理分佈</asp:ListItem>--%>
								<asp:ListItem Value="7">參訓單位類別</asp:ListItem>
								<asp:ListItem Value="8">開班縣市</asp:ListItem>
								<asp:ListItem Value="9">訓練時數</asp:ListItem>
								<asp:ListItem Value="21">訓練職類(大類)</asp:ListItem>
								<asp:ListItem Value="22">就職狀況</asp:ListItem>
							</asp:RadioButtonList>
						</td>
					</tr>
					<tr>
						<td class="bluecol_need">
							Y軸
						</td>
						<td class="whitecol">
							<asp:RadioButtonList ID="YRoll" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="4" CellSpacing="1" CellPadding="1" Width="100%">
								<asp:ListItem Value="1">性別</asp:ListItem>
								<asp:ListItem Value="2">年齡</asp:ListItem>
								<asp:ListItem Value="3">教育程度</asp:ListItem>
								<asp:ListItem Value="4">身份別</asp:ListItem>
								<asp:ListItem Value="5">受訓學員(通訊)地理分佈</asp:ListItem>
								<%--<asp:ListItem Value="6">受訓學員(戶籍)地理分佈</asp:ListItem>--%>
								<asp:ListItem Value="7">參加單位類別</asp:ListItem>
								<asp:ListItem Value="8">開班縣市</asp:ListItem>
								<asp:ListItem Value="9">訓練時數</asp:ListItem>
								<asp:ListItem Value="21">訓練職類(大類)</asp:ListItem>
								<asp:ListItem Value="22">就職狀況</asp:ListItem>
							</asp:RadioButtonList>
						</td>
					</tr>
				</table>
				<p align="center">
					<asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button></p>
				<table id="DataGroupTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
					<tr>
						<td style="height: 40px">
							<div id="Div1" runat="server">
								<asp:Table ID="DataTable1" runat="server" CssClass="font" CellSpacing="0" CellPadding="2" Width="100%">
								</asp:Table>
							</div>
						</td>
					</tr>
					<tr>
						<td align="center">
							<asp:Button ID="Button2" runat="server" Text="列印" CssClass="asp_button_M"></asp:Button>
							&nbsp;
							<asp:Button ID="btnExport" runat="server" Text="匯出Excel" CssClass="asp_button_M"></asp:Button>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
	</form>
</body>
</html>
