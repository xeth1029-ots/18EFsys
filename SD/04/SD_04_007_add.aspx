<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_04_007_add.aspx.vb" Inherits="WDAIIP.SD_04_007_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>作息時間設定</title>
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<script type="text/javascript" language="javascript" src="../../js/OpenWin/openwin.js"></script>
	<script type="text/javascript" language="javascript" src="../../js/common.js"></script>
	<script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
	<script language="javascript" type="text/javascript">
		function check() {
			var msg = '';

			if (document.form1.C11.value == '') {
				msg += '第一節[起始小時]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C11.value)) msg += '第一節[起始小時]必須為數字\n';
				if (document.form1.C11.value.length != 2) msg += '第一節[起始小時]需為二位數\n';
			}
			if (document.form1.C12.value == '') {
				msg += '第一節[起始分鐘]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C12.value)) msg += '第一節[起始分鐘]必須為數字\n';
				if (document.form1.C12.value.length != 2) msg += '第一節[起始分鐘]需為二位數\n';
			}
			if (document.form1.C13.value == '') {
				msg += '第一節[結束小時]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C13.value)) msg += '第一節[結束小時]必須為數字\n';
				if (document.form1.C13.value.length != 2) msg += '第一節[結束小時]需為二位數\n';
			}
			if (document.form1.C14.value == '') {
				msg += '第一節[結束分鐘]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C14.value)) msg += '第一節[結束分鐘]必須為數字\n';
				if (document.form1.C14.value.length != 2) msg += '第一節[結束分鐘]需為二位數\n';
			}

			if (document.form1.C21.value == '') {
				msg += '第二節[起始小時]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C21.value)) msg += '第二節[起始小時]必須為數字\n';
				if (document.form1.C21.value.length != 2) msg += '第二節[起始小時]需為二位數\n';
			}
			if (document.form1.C22.value == '') {
				msg += '第二節[起始分鐘]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C22.value)) msg += '第二節[起始分鐘]必須為數字\n';
				if (document.form1.C22.value.length != 2) msg += '第二節[起始分鐘]需為二位數\n';
			}
			if (document.form1.C23.value == '') {
				msg += '第二節[結束小時]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C23.value)) msg += '第二節[結束小時]必須為數字\n';
				if (document.form1.C23.value.length != 2) msg += '第二節[結束小時]需為二位數\n';
			}
			if (document.form1.C24.value == '') {
				msg += '第二節[結束分鐘]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C24.value)) msg += '第二節[結束分鐘]必須為數字\n';
				if (document.form1.C24.value.length != 2) msg += '第二節[結束分鐘]需為二位數\n';
			}

			if (document.form1.C31.value == '') {
				msg += '第三節[起始小時]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C31.value)) msg += '第三節[起始小時]必須為數字\n';
				if (document.form1.C31.value.length != 2) msg += '第三節[起始小時]需為二位數\n';
			}
			if (document.form1.C32.value == '') {
				msg += '第三節[起始分鐘]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C32.value)) msg += '第三節[起始分鐘]必須為數字\n';
				if (document.form1.C32.value.length != 2) msg += '第三節[起始分鐘]需為二位數\n';
			}
			if (document.form1.C33.value == '') {
				msg += '第三節[結束小時]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C33.value)) msg += '第三節[結束小時]必須為數字\n';
				if (document.form1.C33.value.length != 2) msg += '第三節[結束小時]需為二位數\n';
			}
			if (document.form1.C34.value == '') {
				msg += '第三節[結束分鐘]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C34.value)) msg += '第三節[結束分鐘]必須為數字\n';
				if (document.form1.C34.value.length != 2) msg += '第三節[結束分鐘]需為二位數\n';
			}

			if (document.form1.C41.value == '') {
				msg += '第四節[起始小時]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C41.value)) msg += '第四節[起始小時]必須為數字\n';
				if (document.form1.C41.value.length != 2) msg += '第四節[起始小時]需為二位數\n';
			}
			if (document.form1.C42.value == '') {
				msg += '第四節[起始分鐘]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C42.value)) msg += '第四節[起始分鐘]必須為數字\n';
				if (document.form1.C42.value.length != 2) msg += '第四節[起始分鐘]需為二位數\n';
			}
			if (document.form1.C43.value == '') {
				msg += '第四節[結束小時]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C43.value)) msg += '第四節[結束小時]必須為數字\n';
				if (document.form1.C43.value.length != 2) msg += '第四節[結束小時]需為二位數\n';
			}
			if (document.form1.C44.value == '') {
				msg += '第四節[結束分鐘]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C44.value)) msg += '第四節[結束分鐘]必須為數字\n';
				if (document.form1.C44.value.length != 2) msg += '第四節[結束分鐘]需為二位數\n';
			}

			if (document.form1.C51.value == '') {
				msg += '第五節[起始小時]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C51.value)) msg += '第五節[起始小時]必須為數字\n';
				if (document.form1.C51.value.length != 2) msg += '第五節[起始小時]需為二位數\n';
			}
			if (document.form1.C52.value == '') {
				msg += '第五節[起始分鐘]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C52.value)) msg += '第五節[起始分鐘]必須為數字\n';
				if (document.form1.C52.value.length != 2) msg += '第五節[起始分鐘]需為二位數\n';
			}
			if (document.form1.C53.value == '') {
				msg += '第五節[結束小時]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C53.value)) msg += '第五節[結束小時]必須為數字\n';
				if (document.form1.C53.value.length != 2) msg += '第五節[結束小時]需為二位數\n';
			}
			if (document.form1.C54.value == '') {
				msg += '第五節[結束分鐘]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C54.value)) msg += '第五節[結束分鐘]必須為數字\n';
				if (document.form1.C54.value.length != 2) msg += '第五節[結束分鐘]需為二位數\n';
			}

			if (document.form1.C61.value == '') {
				msg += '第六節[起始小時]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C61.value)) msg += '第六節[起始小時]必須為數字\n';
				if (document.form1.C61.value.length != 2) msg += '第六節[起始小時]需為二位數\n';
			}
			if (document.form1.C62.value == '') {
				msg += '第六節[起始分鐘]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C62.value)) msg += '第六節[起始分鐘]必須為數字\n';
				if (document.form1.C62.value.length != 2) msg += '第六節[起始分鐘]需為二位數\n';
			}
			if (document.form1.C63.value == '') {
				msg += '第六節[結束小時]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C63.value)) msg += '第六節[結束小時]必須為數字\n';
				if (document.form1.C63.value.length != 2) msg += '第六節[結束小時]需為二位數\n';
			}
			if (document.form1.C64.value == '') {
				msg += '第六節[結束分鐘]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C64.value)) msg += '第六節[結束分鐘]必須為數字\n';
				if (document.form1.C64.value.length != 2) msg += '第六節[結束分鐘]需為二位數\n';
			}

			if (document.form1.C71.value == '') {
				msg += '第七節[起始小時]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C71.value)) msg += '第七節[起始小時]必須為數字\n';
				if (document.form1.C71.value.length != 2) msg += '第七節[起始小時]需為二位數\n';
			}
			if (document.form1.C72.value == '') {
				msg += '第七節[起始分鐘]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C72.value)) msg += '第七節[起始分鐘]必須為數字\n';
				if (document.form1.C72.value.length != 2) msg += '第七節[起始分鐘]需為二位數\n';
			}
			if (document.form1.C73.value == '') {
				msg += '第七節[結束小時]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C73.value)) msg += '第七節[結束小時]必須為數字\n';
				if (document.form1.C73.value.length != 2) msg += '第七節[結束小時]需為二位數\n';
			}
			if (document.form1.C74.value == '') {
				msg += '第七節[結束分鐘]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C74.value)) msg += '第七節[結束分鐘]必須為數字\n';
				if (document.form1.C74.value.length != 2) msg += '第七節[結束分鐘]需為二位數\n';
			}

			if (document.form1.C81.value == '') {
				msg += '第八節[起始小時]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C81.value)) msg += '第八節[起始小時]必須為數字\n';
				if (document.form1.C81.value.length != 2) msg += '第八節[起始小時]需為二位數\n';
			}
			if (document.form1.C82.value == '') {
				msg += '第八節[起始分鐘]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C82.value)) msg += '第八節[起始分鐘]必須為數字\n';
				if (document.form1.C82.value.length != 2) msg += '第八節[起始分鐘]需為二位數\n';
			}
			if (document.form1.C83.value == '') {
				msg += '第八節[結束小時]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C83.value)) msg += '第八節[結束小時]必須為數字\n';
				if (document.form1.C83.value.length != 2) msg += '第八節[結束小時]需為二位數\n';
			}
			if (document.form1.C84.value == '') {
				msg += '第八節[結束分鐘]為空值\n';
			} else {
				if (!isUnsignedInt(document.form1.C84.value)) msg += '第八節[結束分鐘]必須為數字\n';
				if (document.form1.C84.value.length != 2) msg += '第八節[結束分鐘]需為二位數\n';
			}

			if (document.form1.C91.value != '') {
				if (!isUnsignedInt(document.form1.C91.value)) msg += '第九節[起始小時]必須為數字\n';
				if (document.form1.C91.value.length != 2) msg += '第九節[起始小時]需為二位數\n';
			}
			if (document.form1.C92.value != '') {
				if (!isUnsignedInt(document.form1.C92.value)) msg += '第九節[起始分鐘]必須為數字\n';
				if (document.form1.C92.value.length != 2) msg += '第九節[起始分鐘]需為二位數\n';
			}
			if (document.form1.C93.value != '') {
				if (!isUnsignedInt(document.form1.C93.value)) msg += '第九節[結束小時]必須為數字\n';
				if (document.form1.C93.value.length != 2) msg += '第九節[結束小時]需為二位數\n';
			}
			if (document.form1.C94.value != '') {
				if (!isUnsignedInt(document.form1.C94.value)) msg += '第九節[結束分鐘]必須為數字\n';
				if (document.form1.C94.value.length != 2) msg += '第九節[結束分鐘]需為二位數\n';
			}

			if (document.form1.C101.value != '') {
				if (!isUnsignedInt(document.form1.C101.value)) msg += '第十節[起始小時]必須為數字\n';
				if (document.form1.C101.value.length != 2) msg += '第十節[起始小時]需為二位數\n';
			}
			if (document.form1.C102.value != '') {
				if (!isUnsignedInt(document.form1.C102.value)) msg += '第十節[起始分鐘]必須為數字\n';
				if (document.form1.C102.value.length != 2) msg += '第十節[起始分鐘]需為二位數\n';
			}
			if (document.form1.C103.value != '') {
				if (!isUnsignedInt(document.form1.C103.value)) msg += '第十節[結束小時]必須為數字\n';
				if (document.form1.C103.value.length != 2) msg += '第十節[結束小時]需為二位數\n';
			}
			if (document.form1.C104.value != '') {
				if (!isUnsignedInt(document.form1.C104.value)) msg += '第十節[結束分鐘]必須為數字\n';
				if (document.form1.C104.value.length != 2) msg += '第十節[結束分鐘]需為二位數\n';
			}

			if (document.form1.C111.value != '') {
				if (!isUnsignedInt(document.form1.C111.value)) msg += '第十一節[起始小時]必須為數字\n';
				if (document.form1.C111.value.length != 2) msg += '第十一節[起始小時]需為二位數\n';
			}
			if (document.form1.C112.value != '') {
				if (!isUnsignedInt(document.form1.C112.value)) msg += '第十一節[起始分鐘]必須為數字\n';
				if (document.form1.C112.value.length != 2) msg += '第十一節[起始分鐘]需為二位數\n';
			}
			if (document.form1.C113.value != '') {
				if (!isUnsignedInt(document.form1.C113.value)) msg += '第十一節[結束小時]必須為數字\n';
				if (document.form1.C113.value.length != 2) msg += '第十一節[結束小時]需為二位數\n';
			}
			if (document.form1.C114.value != '') {
				if (!isUnsignedInt(document.form1.C114.value)) msg += '第十一節[結束分鐘]必須為數字\n';
				if (document.form1.C114.value.length != 2) msg += '第十一節[結束分鐘]需為二位數\n';
			}

			if (document.form1.C121.value != '') {
				if (!isUnsignedInt(document.form1.C121.value)) msg += '第十二節[起始小時]必須為數字\n';
				if (document.form1.C121.value.length != 2) msg += '第十二節[起始小時]需為二位數\n';
			}
			if (document.form1.C122.value != '') {
				if (!isUnsignedInt(document.form1.C122.value)) msg += '第十二節[起始分鐘]必須為數字\n';
				if (document.form1.C122.value.length != 2) msg += '第十二節[起始分鐘]需為二位數\n';
			}
			if (document.form1.C123.value != '') {
				if (!isUnsignedInt(document.form1.C123.value)) msg += '第十二節[結束小時]必須為數字\n';
				if (document.form1.C123.value.length != 2) msg += '第十二節[結束小時]需為二位數\n';
			}
			if (document.form1.C124.value != '') {
				if (!isUnsignedInt(document.form1.C124.value)) msg += '第十二節[結束分鐘]必須為數字\n';
				if (document.form1.C124.value.length != 2) msg += '第十二節[結束分鐘]需為二位數\n';
			}

			if (msg != '') {
				alert(msg);
				return false;
			}
		}

		function xBtnSave1() {
			var Hid_saved = document.getElementById('Hid_saved');
			var But1 = document.getElementById('But1');
			var lnk_save = document.getElementById('lnk_save');
			But1.disabled = true;
			if (Hid_saved.value == "") {
				Hid_saved.value = "Y";
			}
			if (check() == false) {
				But1.disabled = false;
				return false;
			}
			else {
				lnk_save.click();
			}
		}

	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
		<tbody>
			<tr>
				<td align="center">
					<table class="table_nw" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
						<%--<tr>
							<td colspan="4">
								首頁&gt;&gt;學員動態管理&gt;&gt;課程管理&gt;&gt;<font color="#990000">作息時間設定</font>
							</td>
						</tr>--%>
						<tr>
							<td id="TD_Dist" colspan="4">
								轄區：<asp:Label ID="DistID" runat="server" CssClass="font"></asp:Label>
							</td>
						</tr>
						<tr id="TR_PlanYear" runat="server">
							<td id="Td_PlanYear" colspan="4" runat="server">
								年度：<asp:Label ID="PlanYear" runat="server" CssClass="font"></asp:Label>
							</td>
						</tr>
						<tr id="TR_CtrlOrg" runat="server">
							<td id="TD_CtrlOrg" colspan="4" runat="server">
								管控單位：<asp:Label ID="CtrlOrg" runat="server" CssClass="font"></asp:Label>
							</td>
						</tr>
						<tr>
							<td id="TD_org" colspan="4" runat="server">
								訓練機構：<asp:Label ID="OrgName" runat="server" CssClass="font"></asp:Label>
								<input id="OrgID" type="hidden" size="4" name="OrgID" runat="server">
								<input id="RIDValue" type="hidden" size="4" runat="server">
								<input id="OCID" type="hidden" size="4" runat="server">
							</td>
						</tr>
						<tr id="TR_Class" runat="server">
							<td id="TD_Class" colspan="4" runat="server">
								班級名稱：<asp:Label ID="ClassName" runat="server"></asp:Label>
							</td>
						</tr>
						<tr>
							<td class="bluecol" style="width:20%">
								第一節
							</td>
							<td class="whitecol">
								<asp:TextBox ID="C11" runat="server" MaxLength="2" Width="20%"></asp:TextBox>：
								<asp:TextBox ID="C12" runat="server" MaxLength="2" Width="20%"></asp:TextBox>～
								<asp:TextBox ID="C13" runat="server" MaxLength="2" Width="20%"></asp:TextBox>：
								<asp:TextBox ID="C14" runat="server" MaxLength="2" Width="20%"></asp:TextBox>
							</td>
							<td class="bluecol" style="width:20%">
								第七節
							</td>
							<td class="whitecol">
								<asp:TextBox ID="C71" runat="server" MaxLength="2" Width="20%"></asp:TextBox>：
								<asp:TextBox ID="C72" runat="server" MaxLength="2" Width="20%"></asp:TextBox>～
								<asp:TextBox ID="C73" runat="server" MaxLength="2" Width="20%"></asp:TextBox>：
								<asp:TextBox ID="C74" runat="server" MaxLength="2" Width="20%"></asp:TextBox>
							</td>
						</tr>
						<tr>
							<td class="bluecol" style="width:20%">
								第二節
							</td>
							<td class="whitecol">
								<asp:TextBox ID="C21" runat="server" MaxLength="2" Width="20%"></asp:TextBox>：
								<asp:TextBox ID="C22" runat="server" MaxLength="2" Width="20%"></asp:TextBox>～
								<asp:TextBox ID="C23" runat="server" MaxLength="2" Width="20%"></asp:TextBox>：
								<asp:TextBox ID="C24" runat="server" MaxLength="2" Width="20%"></asp:TextBox>
							</td>
							<td class="bluecol" style="width:20%">
								第八節
							</td>
							<td class="whitecol">
								<asp:TextBox ID="C81" runat="server" MaxLength="2" Width="20%"></asp:TextBox>：
								<asp:TextBox ID="C82" runat="server" MaxLength="2" Width="20%"></asp:TextBox>～
								<asp:TextBox ID="C83" runat="server" MaxLength="2" Width="20%"></asp:TextBox>：
								<asp:TextBox ID="C84" runat="server" MaxLength="2" Width="20%"></asp:TextBox>
							</td>
						</tr>
						<tr>
							<td class="bluecol" style="width:20%">
								第三節
							</td>
							<td class="whitecol">
								<asp:TextBox ID="C31" runat="server" MaxLength="2" Width="20%"></asp:TextBox>：
								<asp:TextBox ID="C32" runat="server" MaxLength="2" Width="20%"></asp:TextBox>～
								<asp:TextBox ID="C33" runat="server" MaxLength="2" Width="20%"></asp:TextBox>：
								<asp:TextBox ID="C34" runat="server" MaxLength="2" Width="20%"></asp:TextBox>
							</td>
							<td class="bluecol" style="width:20%">
								第九節
							</td>
							<td class="whitecol">
								<asp:TextBox ID="C91" runat="server" MaxLength="2" Width="20%"></asp:TextBox>：
								<asp:TextBox ID="C92" runat="server" MaxLength="2" Width="20%"></asp:TextBox>～
								<asp:TextBox ID="C93" runat="server" MaxLength="2" Width="20%"></asp:TextBox>：
								<asp:TextBox ID="C94" runat="server" MaxLength="2" Width="20%"></asp:TextBox>
							</td>
						</tr>
						<tr>
							<td class="bluecol" style="width:20%">
								第四節
							</td>
							<td class="whitecol">
								<asp:TextBox ID="C41" runat="server" MaxLength="2" Width="20%"></asp:TextBox>：
								<asp:TextBox ID="C42" runat="server" MaxLength="2" Width="20%"></asp:TextBox>～
								<asp:TextBox ID="C43" runat="server" MaxLength="2" Width="20%"></asp:TextBox>：
								<asp:TextBox ID="C44" runat="server" MaxLength="2" Width="20%"></asp:TextBox>
							</td>
							<td class="bluecol" style="width:20%">
								第十節
							</td>
							<td class="whitecol">
								<asp:TextBox ID="C101" runat="server" MaxLength="2" Width="20%"></asp:TextBox>：
								<asp:TextBox ID="C102" runat="server" MaxLength="2" Width="20%"></asp:TextBox>～
								<asp:TextBox ID="C103" runat="server" MaxLength="2" Width="20%"></asp:TextBox>：
								<asp:TextBox ID="C104" runat="server" MaxLength="2" Width="20%"></asp:TextBox>
							</td>
						</tr>
						<tr>
							<td class="bluecol" style="width:20%">
								第五節
							</td>
							<td class="whitecol">
								<asp:TextBox ID="C51" runat="server" MaxLength="2" Width="20%"></asp:TextBox>：
								<asp:TextBox ID="C52" runat="server" MaxLength="2" Width="20%"></asp:TextBox>～
								<asp:TextBox ID="C53" runat="server" MaxLength="2" Width="20%"></asp:TextBox>：
								<asp:TextBox ID="C54" runat="server" MaxLength="2" Width="20%"></asp:TextBox>
							</td>
							<td class="bluecol" style="width:20%">
								第十一節
							</td>
							<td class="whitecol">
								<asp:TextBox ID="C111" runat="server" MaxLength="2" Width="20%"></asp:TextBox>：
								<asp:TextBox ID="C112" runat="server" MaxLength="2" Width="20%"></asp:TextBox>～
								<asp:TextBox ID="C113" runat="server" MaxLength="2" Width="20%"></asp:TextBox>：
								<asp:TextBox ID="C114" runat="server" MaxLength="2" Width="20%"></asp:TextBox>
							</td>
						</tr>
						<tr>
							<td class="bluecol" style="width:20%">
								第六節
							</td>
							<td class="whitecol">
								<asp:TextBox ID="C61" runat="server" MaxLength="2" Width="20%"></asp:TextBox>：
								<asp:TextBox ID="C62" runat="server" MaxLength="2" Width="20%"></asp:TextBox>～
								<asp:TextBox ID="C63" runat="server" MaxLength="2" Width="20%"></asp:TextBox>：
								<asp:TextBox ID="C64" runat="server" MaxLength="2" Width="20%"></asp:TextBox>
							</td>
							<td class="bluecol" style="width:20%">
								第十二節
							</td>
							<td class="whitecol">
								<asp:TextBox ID="C121" runat="server" MaxLength="2" Width="20%"></asp:TextBox>：
								<asp:TextBox ID="C122" runat="server" MaxLength="2" Width="20%"></asp:TextBox>～
								<asp:TextBox ID="C123" runat="server" MaxLength="2" Width="20%"></asp:TextBox>：
								<asp:TextBox ID="C124" runat="server" MaxLength="2" Width="20%"></asp:TextBox>
							</td>
						</tr>
						<tr>
							<td colspan="4">
								說明：如果該單位課程的訓練時段是[晚上]，請輸入9~12節的時間。
							</td>
						</tr>
						<tr>
							<td align="center" colspan="4" class="whitecol">
								<asp:Button ID="But1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
								<asp:Button ID="But2" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
							</td>
						</tr>
					</table>
					<div style="display: none">
						<asp:LinkButton ID="lnk_save" runat="server"></asp:LinkButton>
					</div>
				</td>
			</tr>
		</tbody>
	</table>
	<input id="Hid_saved" type="hidden" size="4" runat="server" value="">
	</form>
</body>
</html>
