<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_008_D.aspx.vb" Inherits="WDAIIP.SD_05_008_D" %>

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
	<script type="text/javascript" language="javascript" src="../../js/common.js"></script>
	<script type="text/javascript" language="javascript">
		function GetCheckBoxListValue(objName) {
			var v = new Array();
			var CheckBoxList = document.getElementById(objName);
			if (CheckBoxList.tagName == "TABLE") {
				for (var i = 0; i < CheckBoxList.rows.length; i++) {
					for (var j = 0; j < CheckBoxList.rows[i].cells.length; j++) {
						if (CheckBoxList.rows[i].cells[j].childNodes[0]) {
							if (CheckBoxList.rows[i].cells[j].childNodes[0].checked == true) {
								v.push(CheckBoxList.rows[i].cells[j].childNodes[1].innerText);
							}
						}
					}
				}
			}
			if (CheckBoxList.tagName == "SPAN") {
				for (var i = 0; i < CheckBoxList.childNodes.length; i++) {
					if (CheckBoxList.childNodes[i].tagName == "INPUT") {
						if (CheckBoxList.childNodes[i].checked == true) {
							v.push(CheckBoxList.childNodes[i].innerHTML);
						}
					}
				}
			}
			return v;
		}


		function chkdata() {
			var msg = '';
			var xIDNO = '';
			var DegreeID = document.getElementById('DegreeID');
			var IdentityID = document.getElementById('IdentityID');

			//SOCIDNum 非署(局)屬 為空 無此資料
			if (document.form1.Name.value == '') msg += '請輸入姓名\n';

			if (document.form1.StudentID.value == '') msg += '請輸入學號\n';
			if (document.form1.IDNO.value == '') msg += '請輸入身分證號碼\n';
			//局屬 檢查下列資料 【非署(局)屬不需檢查下列資料】
			if (msg == '') {
				if (document.form1.IDNO.value != '' && !checkId(document.form1.IDNO.value)) {
					xIDNO = 'x';
					if (!confirm('您輸入的身分證號碼錯誤,確定要儲存?')) {
						msg += '身分證號碼不正確\n';
					}
				}
			}

			//只有錯誤才會帶出 xIDNO='x'; (外籍暫不驗證)
			if (msg == '' && xIDNO == '') {
				if (document.form1.IDNO.value.charAt(1) == 1 && getRadioValue(document.form1.Sex) == '2') msg += '性別與身分證號碼不符合\n';
				else if (document.form1.IDNO.value.charAt(1) == 2 && getRadioValue(document.form1.Sex) == '1') msg += '性別與身分證號碼不符合\n';
			}

			//isUnsignedInt
			if (document.form1.Byear.value == '') msg += '請輸入出生年\n';
			if (document.form1.Byear.value != '' && !isUnsignedInt(document.form1.Byear.value)) msg += '出生年必須為整數數字\n';
			if (document.form1.Bmonth.value == '') msg += '請輸入出生月\n';
			if (document.form1.Bmonth.value != '' && !isUnsignedInt(document.form1.Bmonth.value)) msg += '出生月必須為整數數字\n';
			if (document.form1.Bday.value == '') msg += '請輸入出生日\n';
			if (document.form1.Bday.value != '' && !isUnsignedInt(document.form1.Bday.value)) msg += '出生日必須為整數數字\n';

			//不到15歲
			//var Today = new Date();
			//var TYY=Today.getYear();
			//debugger;
			var TYY = document.getElementById('TODAY1YEAR').value;
			var YY;
			if (document.getElementById('Byear').value != '') {
				YY = TYY - document.getElementById('Byear').value;
			}
			if (YY < 15) {
				msg = msg + '該學員小於15歲\n';
			}

			if (!isChecked(document.form1.Sex)) msg += '請選擇性別\n';

			//debugger;
			if (!DegreeID.disabled) {
				if (document.form1.DegreeID.selectedIndex == 0) msg += '請選擇學歷\n';
			}

			// select * from key_Identity
			var Identity = getCheckBoxListValue('IdentityID');
			var IdentityVal = GetCheckBoxListValue('IdentityID');
			var bln_check1 = false; //非中高齡者
			var bln_check2 = false; //負擔家計婦女
			var cst_中高齡者 = '中高齡者';
			var cst_負擔家計婦女 = '負擔家計婦女';

			for (var i = 0; i < IdentityVal.length; i++) {
				if (!bln_check1 && IdentityVal[i] == cst_中高齡者) {
					bln_check1 = true;
				}
				if (!bln_check2 && IdentityVal[i] == cst_負擔家計婦女) {
					bln_check2 = true;
				}
			}

			//debugger;
			if (!IdentityID.disabled) {
				if (parseInt(Identity, 10) == 0) msg += '請選擇學員身分別\n';
			}

			if (msg == '') {
				//不到45歲，檢查是否勾選「中高齡者」
				if (document.getElementById('Byear').value != '') {
					YY = TYY - document.getElementById('Byear').value;
				}
				if (YY < 45 && bln_check1 == true) {
					msg = msg + '該學員身分未達「中高齡者」\n';
				}
				if (getRadioValue(document.form1.Sex) == '1' && bln_check2 == true) {
					msg = msg + '該學員為男性，身分不能為「負擔家計婦女」\n';
				}
			}

			if (document.form1.SOCIDNum.value != '') {
				if (document.form1.Q7.selectedIndex == 0) msg += '請選擇參加職訓動機\n';
				if (document.form1.Q8.selectedIndex == 0) msg += '請選擇結訓後動向\n';
				if (document.form1.Q8.selectedIndex == 5 && document.form1.Q8Other.value == '') msg += '請填寫結訓動後動向的[其他]欄位\n';

				if (!document.form1.Q9b.checked && !document.form1.Q9a.checked) msg += '請回答第9題\n';
				if (document.form1.Q9a.checked && !isChecked(document.form1.Q9Y)) msg += '請選擇報名職業訓練前一個月有沒有工作的工作類型\n';

				var Que10 = getRadioValue(document.form1.Q10);
				if (!isChecked(document.form1.Q10)) {
					msg += '報名職業訓練前一個月有否尋找工作\n';
				}

				//Q11a, Q11b, Q11N
				//if (document.form1.Q8.selectedIndex != 4) {
				//	if (!document.form1.Q11b.checked && !document.form1.Q11a.checked) msg += '請回答第11題\n';
				//	if (document.form1.Q11b.checked && !isChecked(document.form1.Q11N)) msg += '請選擇尚未找到工作，但覺得對日後尋找工作幫助的程度\n';
				//	//if (!isChecked(document.form1.Q11N)) { msg += '請選擇 您覺得參加本次訓練後，對日後尋找工作幫助的程度\n'; }
				//}
				if (Q8.selectedIndex!=4 && !isChecked(document.form1.Q11N)) { msg += '請選擇 您覺得參加本次訓練後，對日後尋找工作幫助的程度\n'; }

				var trq12y2013 = document.getElementById("trq12y2013");
				var trq12y2014 = document.getElementById("trq12y2014");
				if (trq12y2013) {
					if (!document.form1.Q12AY.checked && !document.form1.Q12AN.checked) msg += '請回答第12題(滿意或不滿意)\n';
					var Que12 = getCheckBoxListValue('Q12B');
					if (document.form1.Q12AN.checked) {
						//alert(Que12);
						if (parseInt(Que12, 10) == 0) msg += '第12題回答不滿意，請選擇需要改進的為何?\n';
					}
				}
				if (trq12y2014) {
					if (!isChecked(document.form1.Q12V1)) {
						msg += '請選擇 12.1參訓職類符合就業市場需求\n';
					}
					if (!isChecked(document.form1.Q12V2)) {
						msg += '請選擇 12.2教學課程安排\n';
					}
					if (!isChecked(document.form1.Q12V3)) {
						msg += '請選擇 12.3訓練師專業及熱忱\n';
					}
					if (!isChecked(document.form1.Q12V4)) {
						msg += '請選擇 12.4訓練設備符合產業需求\n';
					}
					if (!isChecked(document.form1.Q12V5)) {
						msg += '請選擇 12.5訓練時數\n';
					}
				}
				//var Que12 = getCheckBoxListValue('Q12');
				//  //if (Que12.toString(10)==0) msg+='請選擇你參加本次訓練後覺得有沒有幫助的原因\n';
				//  if (Que12.charAt(Que12.length - 1) == '1' && document.form1.Q12Other.value == '') msg += '請說明您參加本次訓練後覺得不滿意需要改進的原因\n';
			}

			if (msg != '') {
				alert(msg);
				return false;
			}
		}
        /*function GetIdent(){,for(var i=0;i<17;i++){,document.getElementById('IdentityID'+i),},}*/

		function change() {
			var myobj1 = document.form1.elements['Q9Y'];
			for (var i = 0; i < myobj1.length; i++)
				myobj1[i].disabled = document.form1.Q9b.checked;


			//var myobj2 = document.form1.elements['Q11N'];
			//for (var i = 0; i < myobj2.length; i++)
			//	myobj2[i].disabled = document.form1.Q11a.checked;

			//if (document.form1.Q8.selectedIndex == 4) {
			//	document.form1.Q11a.disabled = true;
			//	document.form1.Q11b.disabled = true;
			//	//alert(document.form1.Q11a.disabled);
			//	var myobj1 = document.form1.elements['Q11N'];
			//	for (var i = 0; i < myobj1.length; i++)
			//		myobj1[i].disabled = true;
			//}
			//else {
			//	document.form1.Q11a.disabled = false;
			//	document.form1.Q11b.disabled = false;
			//	var myobj1 = document.form1.elements['Q11N'];
			//	for (var i = 0; i < myobj1.length; i++)
			//		myobj1[i].disabled = document.form1.Q11a.checked;
			//}
		}

		//function NoAns() {
		//	if (document.form1.Q8.selectedIndex == 4) {
		//		document.form1.Q11a.disabled = true;
		//		document.form1.Q11b.disabled = true;
		//		var myobj1 = document.form1.elements['Q11N'];
		//		for (var i = 0; i < myobj1.length; i++)
		//			myobj1[i].disabled = true;
		//	}
		//	else {
		//		document.form1.Q11a.disabled = false;
		//		document.form1.Q11b.disabled = false;
		//		var myobj1 = document.form1.elements['Q11N'];
		//		for (var i = 0; i < myobj1.length; i++)
		//			myobj1[i].disabled = false;
		//	}
		//}

		//身分證號第2位帶入性別
		function chkidnosex() {
			if (document.form1.IDNO.value.charAt(1) == '1') document.form1.Sex_0.checked = true;
			else if (document.form1.IDNO.value.charAt(1) == '2') document.form1.Sex_1.checked = true;
		}
	</script>
</head>
<body onload="change();">
	<form id="form1" method="post" runat="server">
	<input id="DLID" type="hidden" name="DLID" runat="server" size="1">
	<input id="SubNo" type="hidden" name="SubNo" runat="server" size="1">
	<input id="SOCIDNum" type="hidden" runat="server" name="SOCIDNum" size="1">
	<table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
		<tr>
			<td width="100%">
				<%--<table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">結訓學員資料卡登錄</font> </td>
					</tr>   
				</table>--%>
				<table class="table_nw" id="Table3" width="100%" runat="server" cellpadding="1" cellspacing="1">
					<tr>
						<td class="bluecol" style="width:20%">未填學員 </td>
						<td class="whitecol">
							<asp:DropDownList ID="SOCID" runat="server" AutoPostBack="True">
							</asp:DropDownList>
						</td>
					</tr>
				</table>
				<br>
				<table class="table_sch" id="Table4" runat="server" cellpadding="1" cellspacing="1">
					<tr>
						<td class="bluecol" style="width:20%">姓名： </td>
						<td class="whitecol" style="width:30%">
							<asp:TextBox ID="Name" runat="server" Width="40%" MaxLength="20"></asp:TextBox>
						</td>
						<td class="bluecol" style="width:20%">學號： </td>
						<td class="whitecol" style="width:30%">
							<asp:TextBox ID="StudentID" runat="server" Width="40%" MaxLength="20"></asp:TextBox>
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							一、身分證字號：
						</td>
						<td class="whitecol">
							<asp:TextBox ID="IDNO" runat="server" MaxLength="15" Columns="15" Width="40%"></asp:TextBox>
						</td>
						<td class="bluecol">
							二、性別：
						</td>
						<td class="whitecol">
							<asp:RadioButtonList ID="Sex" runat="server" RepeatDirection="Horizontal" CssClass="font">
								<asp:ListItem Value="1">男</asp:ListItem>
								<asp:ListItem Value="2">女</asp:ListItem>
							</asp:RadioButtonList>
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							三、出生日期：
						</td>
						<td class="whitecol">西元
                            <asp:TextBox ID="Byear" runat="server" Width="15%" MaxLength="4"></asp:TextBox>年
                            <asp:TextBox ID="Bmonth" runat="server" Width="12%" MaxLength="2"></asp:TextBox>月
                            <asp:TextBox ID="Bday" runat="server" Width="12%" MaxLength="2"></asp:TextBox>日
						</td>
						<td class="bluecol">
							四、學歷(含肄業)：
						</td>
						<td class="whitecol">
							<asp:DropDownList ID="DegreeID" runat="server">
							</asp:DropDownList>
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							五、兵役：
						</td>
						<td colspan="3" class="whitecol">
							<asp:DropDownList ID="MilitaryID" runat="server">
							</asp:DropDownList>
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							六、學員身分(可複選,最多5組)：
						</td>
						<td colspan="3" class="whitecol">
							<asp:CheckBoxList ID="IdentityID" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="3">
							</asp:CheckBoxList>
							<input id="IdentityValue" type="hidden" runat="server">
						</td>
					</tr>
					<tr>
						<td class="bluecol_need">
							七、參加職訓動機：
						</td>
						<td colspan="3" class="whitecol">
							<asp:DropDownList ID="Q7" runat="server">
								<asp:ListItem Value="0">---請選擇---</asp:ListItem>
								<asp:ListItem Value="1">想學得一技之長以利就業</asp:ListItem>
								<asp:ListItem Value="2">為學習第二專長以利轉業</asp:ListItem>
								<asp:ListItem Value="3">為進一步學得技能以利升遷</asp:ListItem>
								<asp:ListItem Value="4">為充實實務經驗以利升學</asp:ListItem>
								<asp:ListItem Value="5">為參加技能檢定</asp:ListItem>
								<asp:ListItem Value="6">其他</asp:ListItem>
							</asp:DropDownList>
						</td>
					</tr>
					<tr>
						<td class="bluecol_need">
							八、結訓後動向：
						</td>
						<td colspan="3" class="whitecol">
							<asp:DropDownList ID="Q8" runat="server">
								<asp:ListItem Value="0">---請選擇---</asp:ListItem>
								<asp:ListItem Value="1">尋找工作</asp:ListItem>
								<asp:ListItem Value="2">繼續進階課程或升學</asp:ListItem>
								<asp:ListItem Value="3">等待服兵役</asp:ListItem>
								<asp:ListItem Value="4">留在原場(廠)服務</asp:ListItem>
								<asp:ListItem Value="5">其他</asp:ListItem>
							</asp:DropDownList>
							(選項為"其它"的，請填註)
							<asp:TextBox ID="Q8Other" runat="server" MaxLength="50" Width="40%"></asp:TextBox>
						</td>
					</tr>
					<tr id="trq09y2019" runat="server">
						<td class="bluecol_need">
							九、報名職業訓練前 一個月有沒有工作：
						</td>
						<td colspan="3" class="whitecol">
							<table class="font" id="Table5" cellspacing="1" cellpadding="1" width="100%" border="0">
								<tr>
									<td colspan="2">
										<asp:RadioButton ID="Q9a" runat="server" GroupName="Q9" Text="有"></asp:RadioButton>
									</td>
								</tr>
								<tr>
                                    <td width="14%">工作類型：</td>
									<td>
										<asp:RadioButtonList ID="Q9Y" runat="server" CssClass="font">
											<asp:ListItem Value="1">從事1小時以上有報酬的工作</asp:ListItem>
											<asp:ListItem Value="2">從事每週 15 小時以上無酬家屬工作</asp:ListItem>
											<asp:ListItem Value="3">有工作而未做領有報酬 ( 傷病、季節性關係、例假、事假、特別假、已受僱用或等待恢復工作而領有報酬 )</asp:ListItem>
										</asp:RadioButtonList>
									</td>
								</tr>
								<tr>
									<td colspan="2">
										<asp:RadioButton ID="Q9b" runat="server" GroupName="Q9" Text="沒有"></asp:RadioButton>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr id="trq10y2019" runat="server">
						<td class="bluecol_need">
							十、報名職業訓練前一個月有否尋找工作：
						</td>
						<td colspan="3" class="whitecol">
							<asp:RadioButtonList ID="Q10" runat="server" RepeatDirection="Horizontal" CssClass="font">
								<asp:ListItem Value="Y">有</asp:ListItem>
								<asp:ListItem Value="N">沒有</asp:ListItem>
							</asp:RadioButtonList>
						</td>
					</tr>
					<tr id="trq11y2019" runat="server">
						<td class="bluecol_need">
							十一、您覺得參加本次訓練後，對日後尋找工作幫助的程度：
						</td>
						<td colspan="3" class="whitecol">
							<table class="font" id="Table6" cellspacing="1" cellpadding="1" width="100%" border="0">
								<tr>
									<td>
										<asp:RadioButtonList ID="Q11N" runat="server" CssClass="font" RepeatDirection="Horizontal">
											<asp:ListItem Value="1">1.非常有幫助</asp:ListItem>
											<asp:ListItem Value="2">2.有幫助</asp:ListItem>
											<asp:ListItem Value="3">3.普通</asp:ListItem>
											<asp:ListItem Value="4">4.沒有幫助</asp:ListItem>
											<asp:ListItem Value="5">5.完全沒有幫助</asp:ListItem>
										</asp:RadioButtonList>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr id="trq12y2013" runat="server">
						<td class="bluecol">
							十二、您參加本次訓練後是否覺得滿意，若不滿意需要改進的為何：<br />
								(可複選，若無不滿意之處，則本問項可不填答) 
						</td>
						<td colspan="3" class="whitecol">
							<table class="font" cellspacing="1" cellpadding="1" border="0">
								<tr>
									<td>
										<asp:RadioButton ID="Q12AY" runat="server" CssClass="font" GroupName="Q12A" Text="滿意"></asp:RadioButton>
									</td>
									<td>
										<asp:RadioButton ID="Q12AN" runat="server" CssClass="font" GroupName="Q12A" Text="不滿意(可複選)："></asp:RadioButton>
									</td>
								</tr>
								<tr>
									<td colspan="2">
										<asp:CheckBoxList ID="Q12B" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="3">
											<asp:ListItem Value="1">1.參訓職類不符就業市場需求</asp:ListItem>
											<asp:ListItem Value="2">2.教學課程安排不當</asp:ListItem>
											<asp:ListItem Value="3">3.訓練師專業及熱誠不足</asp:ListItem>
											<asp:ListItem Value="4">4.訓練設備不符產業需求</asp:ListItem>
											<asp:ListItem Value="5">5.訓練時數不足</asp:ListItem>
										</asp:CheckBoxList>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr id="trq12y2014" runat="server">
						<td class="bluecol_need">
							十二、您參加本次訓練後，是否覺得滿意： 
						</td>
						<td colspan="3" class="whitecol">
							<table class="font" cellspacing="1" cellpadding="1" border="0" width="100%">
								<tr>
									<td>1.參訓職類符合就業市場需求 </td>
								</tr>
								<tr>
									<td>
										<asp:RadioButtonList ID="Q12V1" runat="server" RepeatDirection="Horizontal" CssClass="font">
											<asp:ListItem Value="1">1.非常同意</asp:ListItem>
											<asp:ListItem Value="2">2.同意</asp:ListItem>
											<asp:ListItem Value="3">3.普通</asp:ListItem>
											<asp:ListItem Value="4">4.不同意</asp:ListItem>
											<asp:ListItem Value="5">5.非常不同意</asp:ListItem>
										</asp:RadioButtonList>
									</td>
								</tr>
								<tr>
									<td>2.教學課程安排 </td>
								</tr>
								<tr>
									<td>
										<asp:RadioButtonList ID="Q12V2" runat="server" RepeatDirection="Horizontal" CssClass="font">
											<asp:ListItem Value="1">1.非常滿意</asp:ListItem>
											<asp:ListItem Value="2">2.滿意</asp:ListItem>
											<asp:ListItem Value="3">3.普通</asp:ListItem>
											<asp:ListItem Value="4">4.不滿意</asp:ListItem>
											<asp:ListItem Value="5">5.非常不滿意</asp:ListItem>
										</asp:RadioButtonList>
									</td>
								</tr>
								<tr>
									<td>3.訓練師專業及熱忱 </td>
								</tr>
								<tr>
									<td>
										<asp:RadioButtonList ID="Q12V3" runat="server" RepeatDirection="Horizontal" CssClass="font">
											<asp:ListItem Value="1">1.非常滿意</asp:ListItem>
											<asp:ListItem Value="2">2.滿意</asp:ListItem>
											<asp:ListItem Value="3">3.普通</asp:ListItem>
											<asp:ListItem Value="4">4.不滿意</asp:ListItem>
											<asp:ListItem Value="5">5.非常不滿意</asp:ListItem>
										</asp:RadioButtonList>
									</td>
								</tr>
								<tr>
									<td>4.訓練設備符合產業需求 </td>
								</tr>
								<tr>
									<td>
										<asp:RadioButtonList ID="Q12V4" runat="server" RepeatDirection="Horizontal" CssClass="font">
											<asp:ListItem Value="1">1.非常同意</asp:ListItem>
											<asp:ListItem Value="2">2.同意</asp:ListItem>
											<asp:ListItem Value="3">3.普通</asp:ListItem>
											<asp:ListItem Value="4">4.不同意</asp:ListItem>
											<asp:ListItem Value="5">5.非常不同意</asp:ListItem>
										</asp:RadioButtonList>
									</td>
								</tr>
								<tr>
									<td>5.訓練時數 </td>
								</tr>
								<tr>
									<td>
										<asp:RadioButtonList ID="Q12V5" runat="server" RepeatDirection="Horizontal" CssClass="font">
											<asp:ListItem Value="1">1.非常滿意</asp:ListItem>
											<asp:ListItem Value="2">2.滿意</asp:ListItem>
											<asp:ListItem Value="3">3.普通</asp:ListItem>
											<asp:ListItem Value="4">4.不滿意</asp:ListItem>
											<asp:ListItem Value="5">5.非常不滿意</asp:ListItem>
										</asp:RadioButtonList>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td colspan="4" class="whitecol">說明:
							<br />
							1、本資料卡僅供統計用途，請確實逐項填寫。
							<br />
							2、第八問項：「結訓後動向」若勾選4.留在原場(廠)服務(表示係在職人員)，則第十一問項不需填答。 </td>
					</tr>
				</table>
				<p align="center" class="whitecol">
					<asp:Button ID="Button1" runat="server" Text="繼續填寫" CssClass="asp_button_M"></asp:Button>
					<asp:Button ID="Button2" runat="server" Text="填寫完畢" CssClass="asp_button_M"></asp:Button>
					<asp:Button ID="Button4" runat="server" Text="回查詢頁面" CssClass="asp_button_M"></asp:Button></p>
			</td>
		</tr>
	</table>
	<input id="TODAY1YEAR" type="hidden" name="TODAY1YEAR" runat="server">
	<input id="HidOCID" type="hidden" runat="server">
	<input id="HidSTDate" type="hidden" runat="server">
	</form>
</body>
</html>
