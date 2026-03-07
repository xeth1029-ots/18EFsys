<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_01_001_1_add.aspx.vb" Inherits="TIMS.SD01_001_1_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>報名登錄新增/修改(產業人才投資方案)</title>
	<meta content="microsoft visual studio .net 7.1" name="generator" />
	<meta content="visual basic .net 7.1" name="code_language" />
	<meta content="javascript" name="vs_defaultclientscript" />
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetschema" />
	<link href="../../css/style.css" type="text/css" rel="stylesheet" />
	<script type="text/javascript" src="../../js/date-picker.js"></script>
	<script type="text/javascript" src="../../js/openwin/openwin.js"></script>
	<script type="text/javascript" src="../../js/common.js"></script>
	<script type="text/javascript">

		function changemode(num) {
			if (document.getElementById('detailtable') && document.getElementById('backtable')) {
				if (num == 1) {
					document.getElementById('detailtable').style.display = 'inline';
					document.getElementById('backtable').style.display = 'none';
					document.getElementById('historytable').style.display = 'none';
				}
				if (num == 2) {
					document.getElementById('detailtable').style.display = 'none';
					document.getElementById('backtable').style.display = 'inline';
					document.getElementById('historytable').style.display = 'none';
				}
				if (num == 3) {
					document.getElementById('detailtable').style.display = 'none';
					document.getElementById('backtable').style.display = 'none';
					document.getElementById('historytable').style.display = 'inline';
				}
			}
		}
		function hard() {
			if (document.form1.identityid_4.checked) {
				document.form1.handtypeid.disabled = false;
				document.form1.handlevelid.disabled = false;
			}
			else {
				document.form1.handtypeid.disabled = true;
				document.form1.handlevelid.disabled = true;
			}
		}
		function changeacctmode() {
			if (document.form1.acctmode_0.checked) {
				document.getElementById('porttr').style.display = 'inline';
				document.getElementById('banktr1').style.display = 'none';
				document.getElementById('banktr2').style.display = 'none';
				document.getElementById('banktr3').style.display = 'none';
			}
			else {
				document.getElementById('porttr').style.display = 'none';
				document.getElementById('banktr1').style.display = 'inline';
				document.getElementById('banktr2').style.display = 'inline';
				document.getElementById('banktr3').style.display = 'inline';
			}
		}
		function check1() {
			if (document.form1.q3_2.checked) {
				document.form1.q3_other.disabled = false;
			}
			else {
				document.form1.q3_other.disabled = true;
			}
		}
		function choose_class(tm, oc, num) {
			if (num != 1) {
				if (document.form1.tmid1.value == '') {
					window.alert('請先選擇第一志願!');
					return;
				}
			}
			if (num != 2 && num != 1) {
				if (document.form1.tmid2.value == '') {
					window.alert('請先選擇第二志願!');
					return;
				}
			}
			window.open('sd_01_001_ch.aspx?wish=' + num, '', 'width=550,height=400,location=0,status=0,menubar=0,scrollbars=1,resizable=0');
		}
		function clear_wish(num) {
			switch (num) {
				case 1:
					document.form1.tmid1.value = '';
					document.form1.tmidvalue1.value = '';
					document.form1.ocid1.value = '';
					document.form1.ocidvalue1.value = '';
					document.form1.comidno1.value = '';
					document.form1.seqno1.value = '';
					document.form1.cclid.value = '';
					break;
				case 2:
					document.form1.tmid2.value = '';
					document.form1.tmidvalue2.value = '';
					document.form1.ocid2.value = '';
					document.form1.ocidvalue2.value = '';
					document.form1.comidno2.value = '';
					document.form1.seqno2.value = '';
					break;
				case 3:
					document.form1.tmid3.value = '';
					document.form1.tmidvalue3.value = '';
					document.form1.ocid3.value = '';
					document.form1.ocidvalue3.value = '';
					document.form1.comidno3.value = '';
					document.form1.seqno3.value = '';
					break;
			}
		}
		function chkdata() {
			var msg = '';

			if (document.form1.hidstar3.value != '') {
				if (!confirm('本次登錄之學員,仍在訓中,是否儲存,請確認!')) msg += '學員,仍在訓中\n';
			}

			if (isEmpty('relenterdate')) { msg += '請選擇報名日期\n'; }
			if (isEmpty('name')) { msg += '請輸入中文姓名\n'; }
			if (isEmpty('passportno')) { msg += '請選擇身分別\n'; }
			if (isEmpty('sex')) { msg += '請選擇性別\n'; }
			if (isEmpty('degreeid')) { msg += '請選擇最高學歷\n'; }
			if (isEmpty('school')) { msg += '請輸入學校名稱\n'; }
			if (isEmpty('department')) { msg += '請輸入科系\n'; }
			if (isEmpty('graduatestatus')) { msg += '請選擇畢業狀況\n'; }
			//if (isEmpty('MilitaryID')) { msg += '請選擇兵役狀況\n'; }
			if (isEmpty('MaritalStatus')) { msg += '請選擇婚姻狀況\n'; }
			//if (isEmpty('phoned')) { msg += '請輸入聯絡電話(日)\n'; }
			//rblmobil
			if (!ischecked(document.form1.rblmobil)) {
				msg += '請選擇有無行動電話\n';
			}
			else {
				if (getRadioValue(document.form1.rblmobil) == "y") {
					if (document.form1.cellphone.value == '') { msg += '有行動電話 請輸入行動電話\n'; }
				}
				else {
					if (document.form1.phoned.value == '') { msg += '請輸入聯絡電話(日)\n'; }
					if (document.form1.cellphone.value != '') { msg += '有輸入行動電話,請選擇有行動電話\n'; }
				}
			}

			if (isEmpty('zipcode1')) { msg += '請選擇通訊地址-縣市\n'; }
			if (isEmpty('address')) { msg += '請輸入通訊地址\n'; }
			if (!isEmpty('zipcode2')) {
				if (isEmpty('householdaddress')) { msg += '請輸入戶籍地址\n'; }
			}
			if (isEmpty('midentityid')) { msg += '請選擇主要參訓身分別\n'; }
			if (isEmpty('identityid')) {
				msg += '請選擇參訓身分別\n';
			}
			else {
				var i = 0;
				if (document.form1.identityid_0.checked) { i = i + 1; }
				if (document.form1.identityid_1.checked) { i = i + 1; }
				if (document.form1.identityid_2.checked) { i = i + 1; }
				if (document.form1.identityid_3.checked) { i = i + 1; }
				if (document.form1.identityid_4.checked) { i = i + 1; }
				if (document.form1.identityid_5.checked) { i = i + 1; }
				if (document.form1.identityid_6.checked) { i = i + 1; }
				if (i > 3) { msg += '參訓身分別最多只能選擇三項\n'; }
			}
			if (isEmpty('ocidvalue1')) { msg += '請選擇要報名的班級\n'; }
			if (document.form1.identityid_4.checked) {
				if (isEmpty('handtypeid')) { msg += '請選擇障礙類別\n'; }
				if (isEmpty('handlevelid')) { msg += '請選擇障礙等級\n'; }
			}
			if (!isEmpty('priorworkorg1')) {
				if (isEmpty('title1')) { msg += '請輸入職稱1\n'; }
			}
			if (!isEmpty('title1')) {
				if (isEmpty('priorworkorg1')) { msg += '請輸入受訓服務單位1\n'; }
			}
			if (!isEmpty('priorworkorg2')) {
				if (isEmpty('title2')) { msg += '請輸入職稱2\n'; }
			}
			if (!isEmpty('title2')) {
				if (isEmpty('priorworkorg2')) { msg += '請輸入受訓服務單位2\n'; }
			}
			if (!isEmpty('priorworkpay')) {
				if (!isint(getvalue('priorworkpay'))) { msg = '受訓前薪資.請輸入數字!!!\n'; }
			}
			if (!isEmpty('realjobless')) {
				if (!isint(getvalue('realjobless'))) { msg = '受訓前失業周數.請輸入數字!!!\n'; }
			}
			if (isEmpty('showdetail')) { msg += '請選擇是否提供基本資料供廠商查詢\n'; }
			if (isEmpty('isagree')) { msg += '請選擇是否同意個人基本資料供查詢\n'; }
			if (isEmpty('acctmode')) {
				msg += '請選擇郵局帳號或銀行帳號\n';
			}
			else {
				if (document.form1.acctmode_0.checked) {
					if (isEmpty('postno_1')) { msg += '請輸入郵局-局號1\n'; }
					if (isEmpty('postno_2')) { msg += '請輸入郵局-局號2\n'; }
					if (isEmpty('acctno1_1')) { msg += '請輸入郵局-帳號1\n'; }
					if (isEmpty('acctno1_2')) { msg += '請輸入郵局-帳號2\n'; }
				}
				else {
					if (isEmpty('bankname')) { msg += '請輸入銀行名稱\n'; }
					if (isEmpty('acctheadno')) { msg += '請輸入銀行總代號\n'; }
					if (isEmpty('acctno2')) { msg += '請輸入銀行帳號\n'; }
				}
			}
			if (isEmpty('tel')) { msg += '請輸入公司電話\n'; }
			if (isEmpty('zip')) { msg += '請選擇公司地址-縣市\n'; }
			if (isEmpty('addr')) { msg += '請輸入公司地址\n'; }
			if (isEmpty('q1')) { msg += '請選擇參訓資料背景-是否由公司推薦參訓\n'; }
			if (isEmpty('q2')) { msg += '請選擇參訓資料背景-參訓動機\n'; }
			if (isEmpty('q4')) { msg += '請選擇參訓資料背景-服務單位行業別\n'; }
			if (!isEmpty('q61')) {
				if (!isint(getvalue('q61'))) { msg = '個人工作年資.請輸入數字!!!\n'; }
			}
			if (!isEmpty('q62')) {
				if (!isint(getvalue('q62'))) { msg = '在這家公司的年資.請輸入數字!!!\n'; }
			}
			if (!isEmpty('q63')) {
				if (!isint(getvalue('q63'))) { msg = '在這職位的年資.請輸入數字!!!\n'; }
			}
			if (!isEmpty('q64')) {
				if (!isint(getvalue('q64'))) { msg = '最近升遷離本職幾年.請輸入數字!!!\n'; }
			}
			if (msg == '') {
				return true;
			} else {
				alert(msg);
				return false;
			}
		}		
	</script>
</head>
<body ms_positioning="flowlayout">
	<form id="form1" method="post" runat="server">
	<table id="table2" bordercolor="#cccccc" cellspacing="0" cellpadding="0" width="590" align="left" bgcolor="#ffffff" border="1">
		<tr>
			<td class="font">
				<asp:Label ID="titlelab1" runat="server"></asp:Label><asp:Label ID="titlelab2" runat="server"></asp:Label>
			</td>
		</tr>
		<tr id="trbasic1" runat="server">
			<td bordercolor="#ffffff" align="center">
				<table id="table1" style="width: 590px" cellspacing="1" cellpadding="1" width="590" align="center" border="0">
					<tbody>
						<tr>
							<td>
								<table class="table_nw" id="table3" cellspacing="1" cellpadding="1" width="734">
									<tr>
										<td class="bluecol" align="center">
											個人報名基本資料
										</td>
									</tr>
								</table>
								<table class="font" id="menutable" style="cursor: hand" height="20" cellspacing="0" cellpadding="0" border="0" runat="server">
									<tr>
										<td onclick="changemode(1);" width="1" background="../../images/bookmark_01.gif">
										</td>
										<td onclick="changemode(1);" align="center" width="100" background="../../images/bookmark_02.gif">
											<font size="2">個人基本資料</font>
										</td>
										<td onclick="changemode(1);" width="11" background="../../images/bookmark_03.gif">
										</td>
										<td onclick="changemode(2);" width="1" background="../../images/bookmark_01.gif">
										</td>
										<td onclick="changemode(2);" align="center" width="100" background="../../images/bookmark_02.gif">
											<font size="2">參訓背景</font>
										</td>
										<td onclick="changemode(2);" width="11" background="../../images/bookmark_03.gif">
										</td>
										<td onclick="changemode(3);" width="1" background="../../images/bookmark_01.gif">
										</td>
										<td onclick="changemode(3);" align="center" width="100" background="../../images/bookmark_02.gif">
											<font size="2">參訓記錄</font>
										</td>
										<td onclick="changemode(3);" width="11" background="../../images/bookmark_03.gif">
										</td>
									</tr>
								</table>
							</td>
						</tr>
					</tbody>
				</table>
				<table id="table8" class="font" cellspacing="1" cellpadding="1" width="740" align="center">
					<tbody>
						<tr>
							<td>
								<table class="table_nw" id="detailtable" cellspacing="1" cellpadding="1" width="734" runat="server">
									<tbody>
										<tr>
											<td class="bluecol_need" width="137">
												中文姓名
											</td>
											<td class="whitecol">
												<asp:TextBox ID="name" runat="server" Columns="15"></asp:TextBox>
											</td>
											<td class="bluecol_need" width="151">
												報名日期
											</td>
											<td class="whitecol">
												<asp:TextBox ID="relenterdate" runat="server" Columns="15" onfocus="this.blur()" Width="60px"></asp:TextBox>
												<span id="span1" runat="server"><img id="imgrelenterdate" style="cursor: hand" onclick="javascript:show_calendar('<%= relenterdate.clientid %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="24" height="24" /></span>
											</td>
										</tr>
										<tr id="tr1" runat="server">
											<td class="bluecol_need" width="137">
												身分別
											</td>
											<td class="whitecol">
												<asp:RadioButtonList ID="passportno" runat="server" Width="100%" RepeatDirection="horizontal" CssClass="font" Height="4px">
													<asp:ListItem Value="1">本國</asp:ListItem>
													<asp:ListItem Value="2">外籍(含大陸人士)</asp:ListItem>
												</asp:RadioButtonList>
											</td>
											<td class="bluecol_need" width="151">
												身分證號碼
											</td>
											<td class="whitecol">
												<asp:TextBox ID="idno" runat="server" Columns="15" onfocus="this.blur()" Width="84px"></asp:TextBox>
											</td>
										</tr>
										<tr>
											<td class="bluecol_need" width="137">
												性 別
											</td>
											<td class="whitecol">
												<asp:RadioButtonList ID="sex" runat="server" RepeatDirection="horizontal" CssClass="font">
													<asp:ListItem Value="m">男</asp:ListItem>
													<asp:ListItem Value="f">女</asp:ListItem>
												</asp:RadioButtonList>
											</td>
											<td class="bluecol_need" width="151">
												出生日期
											</td>
											<td class="whitecol">
												<asp:TextBox ID="birthday" runat="server" onfocus="this.blur()" Width="60px"></asp:TextBox>
												<span id="span2" runat="server"><img id="imgbirthday" style="cursor: hand" onclick="javascript:show_calendar('<%= birthday.clientid %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="24" height="24" /></span>
											</td>
										</tr>
										<tr>
											<td class="bluecol_need" width="137">
												最高學歷
											</td>
											<td class="whitecol">
												<asp:DropDownList ID="degreeid" runat="server">
												</asp:DropDownList>
											</td>
											<td class="bluecol_need" width="151">
												學校名稱
											</td>
											<td class="whitecol">
												<asp:TextBox ID="school" runat="server"></asp:TextBox>
											</td>
										</tr>
										<tr>
											<td class="bluecol_need" width="137">
												科 系
											</td>
											<td class="whitecol">
												<asp:TextBox ID="department" runat="server"></asp:TextBox>
											</td>
											<td class="bluecol_need" width="151">
												畢業狀況
											</td>
											<td class="whitecol">
												<asp:DropDownList ID="graduatestatus" runat="server" Width="64px" Height="30px">
												</asp:DropDownList>
											</td>
										</tr>
										<tr>
											<td class="bluecol" width="137">
												兵役狀況
											</td>
											<td class="whitecol">
												<asp:DropDownList ID="MilitaryID" runat="server">
												</asp:DropDownList>
											</td>
											<td class="bluecol_need" width="151">
												婚姻狀況
											</td>
											<td class="whitecol">
												<asp:DropDownList ID="MaritalStatus" runat="server" Width="64px" CssClass="font">
													<asp:ListItem Value="">請選擇</asp:ListItem>
													<asp:ListItem Value="1">已婚</asp:ListItem>
													<asp:ListItem Value="2">未婚</asp:ListItem>
												</asp:DropDownList>
											</td>
										</tr>
										<tr>
											<td class="bluecol_need" width="137">
												聯絡電話
											</td>
											<td class="whitecol">
												<table class="font" id="table7" cellspacing="1" cellpadding="1" width="100%" border="0">
													<tr>
														<td>
															(日)
														</td>
														<td>
															<asp:TextBox ID="phoned" runat="server" Columns="13"></asp:TextBox>
														</td>
													</tr>
													<tr>
														<td>
															(夜)
														</td>
														<td>
															<asp:TextBox ID="phonen" runat="server" Columns="13"></asp:TextBox>
														</td>
													</tr>
												</table>
											</td>
											<td class="bluecol_need" width="151">
												行動電話
											</td>
											<td class="whitecol">
												<asp:RadioButtonList ID="rblmobil" runat="server" CssClass="font" RepeatDirection="horizontal" RepeatLayout="flow">
													<asp:ListItem Value="n">無</asp:ListItem>
													<asp:ListItem Value="y">有</asp:ListItem>
												</asp:RadioButtonList>
												<asp:TextBox ID="cellphone" runat="server" Columns="13"></asp:TextBox>
											</td>
										</tr>
										<tr>
											<td class="bluecol_need" width="137">
												通訊地址
											</td>
											<td class="whitecol" colspan="3">
												<asp:TextBox ID="city1" runat="server" onfocus="this.blur()" Width="130px"></asp:TextBox>
												<input id="zipcode1" type="hidden" size="1" name="zipcode1" runat="server" />
												<input id="btnzipcode1" onclick="getzip('../../js/openwin/zipcode.aspx', 'city1', 'zipcode1')" type="button" value="..." name="button1" runat="server" class="asp_button_Mini" />
												<asp:TextBox ID="address" runat="server" Width="352px"></asp:TextBox>
											</td>
										</tr>
										<tr>
											<td class="bluecol" width="137">
												戶籍地址
											</td>
											<td class="whitecol" colspan="3">
												<asp:CheckBox ID="checkbox1" runat="server" CssClass="font" Text="同通訊地址"></asp:CheckBox><br />
												<asp:TextBox ID="city2" runat="server" onfocus="this.blur()" Width="130px"></asp:TextBox>
												<input id="zipcode2" type="hidden" size="1" name="zipcode2" runat="server" />
												<input id="btnzipcode2" onclick="getzip('../../js/openwin/zipcode.aspx', 'city2', 'zipcode2')" type="button" value="..." name="button1" runat="server" class="asp_button_Mini" />
												<asp:TextBox ID="householdaddress" runat="server" Width="352px"></asp:TextBox>
											</td>
										</tr>
										<tr>
											<td class="bluecol" width="137">
												電子郵件
											</td>
											<td class="whitecol" colspan="6">
												<asp:TextBox ID="email" runat="server" Width="200px"></asp:TextBox>
											</td>
										</tr>
										<tr>
											<td class="bluecol_need" width="137">
												主要參訓身分別
											</td>
											<td class="whitecol" colspan="3">
												<asp:DropDownList ID="midentityid" runat="server" Width="122px">
												</asp:DropDownList>
											</td>
										</tr>
										<tr>
											<td class="bluecol">
												<font color="red">參訓身分別</font><br />
												(可複選，最多三項)
											</td>
											<td onclick="hard();" class="whitecol" colspan="3">
												<asp:CheckBoxList ID="identityid" runat="server" RepeatDirection="horizontal" CssClass="font" RepeatColumns="3">
												</asp:CheckBoxList>
											</td>
										</tr>
										<tr>
											<td class="bluecol_need">
												報名班級
											</td>
											<td class="whitecol" colspan="3">
												<table class="font" width="100%" border="0">
													<tr>
														<td>
															班級一：職類：
														</td>
														<td>
															<asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()"></asp:TextBox><input id="TMIDValue1" type="hidden" size="1" name="TMIDValue1" runat="server" />
														</td>
														<td>
															班別：
														</td>
														<td>
															<asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()"></asp:TextBox>
															<input id="OCIDValue1" type="hidden" size="1" name="OCIDValue1" runat="server" />
															<input id="comidno1" type="hidden" size="1" name="comidno1" runat="server" />
															<input id="seqno1" type="hidden" size="1" name="seqno1" runat="server" />
															<input id="cclid" type="hidden" size="1" name="cclid" runat="server" />
															<input id="button5" onclick="choose_class('tmid1','ocid1',1)" type="button" value="..." name="button5" runat="server" class="asp_button_Mini" />
															<input id="btnclear1" onclick="clear_wish(1)" type="button" value="清除" name="button1" runat="server" class="asp_button_S" />
														</td>
													</tr>
													<tr>
														<td>
															班級二：職類：
														</td>
														<td>
															<asp:TextBox ID="tmid2" runat="server" onfocus="this.blur()"></asp:TextBox>
															<input id="tmidvalue2" type="hidden" size="1" name="tmidvalue2" runat="server" />
														</td>
														<td>
															班別：
														</td>
														<td>
															<asp:TextBox ID="ocid2" runat="server" onfocus="this.blur()"></asp:TextBox>
															<input id="ocidvalue2" type="hidden" size="1" name="ocidvalue2" runat="server" />
															<input id="comidno2" type="hidden" size="1" name="comidno2" runat="server" />
															<input id="seqno2" type="hidden" size="1" name="seqno2" runat="server" />
															<input id="button2" onclick="choose_class('tmid2','ocid2',2)" type="button" value="..." name="button2" runat="server" class="asp_button_Mini" />
															<input id="btnclear2" onclick="clear_wish(2)" type="button" value="清除" name="button4" runat="server" class="asp_button_S" />
														</td>
													</tr>
													<tr>
														<td>
															班級三：職類：
														</td>
														<td>
															<asp:TextBox ID="tmid3" runat="server" onfocus="this.blur()"></asp:TextBox>
															<input id="tmidvalue3" type="hidden" size="1" name="tmidvalue3" runat="server" />
														</td>
														<td>
															班別：
														</td>
														<td id="classid">
															<asp:TextBox ID="ocid3" runat="server" onfocus="this.blur()"></asp:TextBox>
															<input id="ocidvalue3" type="hidden" size="1" name="ocidvalue3" runat="server" />
															<input id="comidno3" type="hidden" size="1" name="comidno3" runat="server" />
															<input id="seqno3" type="hidden" size="1" name="seqno3" runat="server" />
															<input id="button3" onclick="choose_class('tmid3','ocid3',3)" type="button" value="..." name="button3" runat="server" class="asp_button_Mini" />
															<input id="btnclear3" onclick="clear_wish(3)" type="button" value="清除" name="button6" runat="server" class="asp_button_S" />
														</td>
													</tr>
												</table>
											</td>
										</tr>
										<tr>
											<td class="bluecol" width="137">
												障礙類別
											</td>
											<td class="whitecol">
												<asp:DropDownList ID="handtypeid" runat="server" Width="240px" Enabled="false">
												</asp:DropDownList>
											</td>
											<td class="bluecol" width="151">
												障礙等級
											</td>
											<td class="whitecol">
												<asp:DropDownList ID="handlevelid" runat="server" Width="164px" Enabled="false">
												</asp:DropDownList>
											</td>
										</tr>
										<tr>
											<td class="bluecol" width="137" rowspan="2">
												受訓服務單位
											</td>
											<td class="whitecol">
												1.<asp:TextBox ID="priorworkorg1" runat="server"></asp:TextBox>
											</td>
											<td class="bluecol" width="151" rowspan="2">
												職稱
											</td>
											<td class="whitecol">
												1.<asp:TextBox ID="title1" runat="server"></asp:TextBox>
											</td>
										</tr>
										<tr>
											<td class="whitecol">
												2.<asp:TextBox ID="priorworkorg2" runat="server"></asp:TextBox>
											</td>
											<td class="whitecol">
												2.<asp:TextBox ID="title2" runat="server"></asp:TextBox>
											</td>
										</tr>
										<tr>
											<td width="137" class="bluecol">
												受訓前任職起<br />
												迄年月
											</td>
											<td class="whitecol" colspan="3">
												<table class="whitecol" id="table6" cellspacing="1" cellpadding="1" border="0">
													<tr>
														<td>
															<font size="2">1.</font>
														</td>
														<td>
                                                            <asp:TextBox ID="sofficeym1" runat="server" onfocus="this.blur()" Width="75px"></asp:TextBox>
															<span id="span3" runat="server"><img id="imgsofficeym1" style="cursor: hand" onclick="javascript:show_calendar('<%= sofficeym1.clientid %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="24" height="24"></span>
														</td>
														<td>
															<font size="2">～</font>
														</td>
														<td>
                                                            <asp:TextBox ID="fofficeym1" runat="server" onfocus="this.blur()" Width="75px"></asp:TextBox>
															<span id="span4" runat="server"><img id="imgfofficeym1" style="cursor: hand" onclick="javascript:show_calendar('<%= fofficeym1.clientid %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="24" height="24"></span>
														</td>
													</tr>
													<tr>
														<td>
															<font size="2">2.</font>
														</td>
														<td>
															<asp:TextBox ID="sofficeym2" runat="server" onfocus="this.blur()" Width="75px"></asp:TextBox>
                                                            <span id="span5" runat="server"><img id="imgsofficeym2" style="cursor: hand" onclick="javascript:show_calendar('<%= sofficeym2.clientid %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="24" height="24"></span>
														</td>
														<td>
															<font size="2">～</font>
														</td>
														<td>
															<asp:TextBox ID="fofficeym2" runat="server" onfocus="this.blur()" Width="75px"></asp:TextBox>
                                                            <span id="span6" runat="server"><img id="imgfofficeym2" style="cursor: hand" onclick="javascript:show_calendar('<%= fofficeym2.clientid %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="24" height="24"></span>
														</td>
													</tr>
												</table>
											</td>
										</tr>
										<tr>
											<td class="bluecol" width="137">
												受訓前薪資
											</td>
											<td class="whitecol">
												<asp:TextBox ID="priorworkpay" runat="server" Width="100px"></asp:TextBox>
											</td>
											<td class="bluecol" width="151" align="center">
												受訓前失業周數
											</td>
											<td class="whitecol">
												<asp:TextBox ID="realjobless" runat="server" Width="50px"></asp:TextBox>
												<asp:DropDownList ID="joblessid" runat="server">
												</asp:DropDownList>
											</td>
										</tr>
										<tr>
											<td class="bluecol" width="137">
												交通方式
											</td>
											<td class="whitecol" colspan="3">
												<asp:DropDownList ID="traffic" runat="server">
													<asp:ListItem Value="">請選擇</asp:ListItem>
													<asp:ListItem Value="1">住宿</asp:ListItem>
													<asp:ListItem Value="2">通勤</asp:ListItem>
												</asp:DropDownList>
											</td>
										</tr>
										<tr>
											<td class="bluecol_need" width="137">
												提供基本資料<br />
												供查詢
											</td>
											<td class="whitecol" colspan="3">
												<asp:DropDownList ID="showdetail" runat="server">
													<asp:ListItem Value="">請選擇</asp:ListItem>
													<asp:ListItem Value="y">是</asp:ListItem>
													<asp:ListItem Value="n">否</asp:ListItem>
												</asp:DropDownList>
												<font size="2">(姓名、出生年月日、性別、學歷、科系、電話、電子郵件帳號、專長)</font>
											</td>
										</tr>
										<tr>
											<td class="bluecol_need" colspan="4">
												本人
												<asp:RadioButtonList ID="isagree" runat="server" RepeatDirection="horizontal" RepeatLayout="flow" CssClass="font">
													<asp:ListItem Value="y">同意</asp:ListItem>
													<asp:ListItem Value="n">不同意</asp:ListItem>
												</asp:RadioButtonList>
												個人基本資料，供 勞動部勞動力發展署 暨所屬機關運用，以從事職業訓練及就業服務
											</td>
										</tr>
									</tbody>
								</table>
							</td>
						</tr>
						<tr>
							<td>
								<table class="table_nw" id="backtable" cellspacing="1" cellpadding="1" width="735" runat="server">
									<tr>
										<td class="bluecol" align="center" colspan="4">
											服務單位資料
										</td>
									</tr>
									<tr>
										<td class="bluecol_need" width="114">
											郵政/銀行帳號
										</td>
										<td class="whitecol" onclick="changeacctmode();" colspan="3">
											<asp:RadioButtonList ID="acctmode" runat="server" RepeatDirection="horizontal" RepeatLayout="flow" CssClass="font">
												<asp:ListItem Value="0">郵局帳號</asp:ListItem>
												<asp:ListItem Value="1">銀行帳號</asp:ListItem>
											</asp:RadioButtonList>
										</td>
									</tr>
									<tr id="porttr" runat="server">
										<td class="bluecol_need" width="114">
											局號
										</td>
										<td class="whitecol" width="209">
											<asp:TextBox ID="postno_1" runat="server" Columns="8"></asp:TextBox>－
											<asp:TextBox ID="postno_2" runat="server" Columns="1"></asp:TextBox>
										</td>
										<td class="bluecol_need" width="80">
											帳號
										</td>
										<td class="whitecol" width="200">
											<asp:TextBox ID="acctno1_1" runat="server" Columns="8"></asp:TextBox>－
											<asp:TextBox ID="acctno1_2" runat="server" Columns="1"></asp:TextBox>
										</td>
									</tr>
									<tr id="banktr1" runat="server">
										<td class="bluecol_need" width="114">
											銀行名稱
										</td>
										<td class="whitecol" width="200" colspan="3">
											<asp:TextBox ID="bankname" runat="server"></asp:TextBox>
										</td>
									</tr>
									<tr id="banktr2" runat="server">
										<td class="bluecol_need" width="114">
											總代號
										</td>
										<td class="whitecol" width="200" colspan="3">
											<asp:TextBox ID="acctheadno" runat="server" Columns="8"></asp:TextBox>
										</td>
									</tr>
									<tr id="banktr3" runat="server">
										<td class="bluecol_need" width="114">
											帳號
										</td>
										<td class="whitecol" colspan="3">
											<asp:TextBox ID="acctno2" runat="server"></asp:TextBox>
										</td>
									</tr>
									<tr>
										<td class="bluecol" width="114">
											第一次投保日
										</td>
										<td class="whitecol" colspan="3">
											<asp:TextBox ID="firdate" runat="server" Columns="10" onfocus="this.blur()"></asp:TextBox>
											<span id="span7" runat="server"><img id="imgfirdate" style="cursor: hand" onclick="javascript:show_calendar('<%= firdate.clientid %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="24" height="24" /></span>
										</td>
									</tr>
									<tr>
										<td class="bluecol" width="114">
											目前任職公司名稱
										</td>
										<td class="whitecol">
											<asp:TextBox ID="uname" runat="server"></asp:TextBox>
										</td>
										<td class="bluecol">
											統一編號
										</td>
										<td class="whitecol">
											<asp:TextBox ID="intaxno" runat="server"></asp:TextBox>
										</td>
									</tr>
									<tr>
										<td class="bluecol_need" width="114">
											公司電話
										</td>
										<td class="whitecol">
											<asp:TextBox ID="tel" runat="server"></asp:TextBox>
										</td>
										<td class="bluecol">
											公司傳真
										</td>
										<td class="whitecol">
											<asp:TextBox ID="fax" runat="server"></asp:TextBox>
										</td>
									</tr>
									<tr>
										<td class="bluecol_need" width="114">
											公司地址
										</td>
										<td class="whitecol" colspan="3">
											<asp:TextBox ID="city5" runat="server" onfocus="this.blur()" Width="130px"></asp:TextBox>
											<input id="btnzip5" onclick="getzip('../../js/openwin/zipcode.aspx', 'city5', 'zip')" type="button" value="..." name="button1" runat="server" class="asp_button_Mini" />
											<input id="zip" type="hidden" size="1" name="zip" runat="server" />
											<asp:TextBox ID="addr" runat="server" Width="250px"></asp:TextBox>
										</td>
									</tr>
									<tr>
										<td class="bluecol" width="114">
											目前任職部門
										</td>
										<td class="whitecol" style="width: 209px">
											<asp:TextBox ID="servdept" runat="server"></asp:TextBox>
										</td>
										<td class="bluecol">
											職稱
										</td>
										<td class="whitecol">
											<asp:TextBox ID="jobtitle" runat="server"></asp:TextBox>
										</td>
									</tr>
									<tr>
										<td class="bluecol" width="114">
											個人到任目前<br />
											任職公司起日
										</td>
										<td class="whitecol">
											<asp:TextBox ID="sdate" runat="server" Columns="10" onfocus="this.blur()"></asp:TextBox>
											<span id="span8" runat="server"><img id="imgsdate" style="cursor: hand" onclick="javascript:show_calendar('<%= sdate.clientid %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="24" height="24" /></span>
										</td>
										<td class="bluecol">
											個人到任目<br />
											前職務起日
										</td>
										<td class="whitecol">
											<asp:TextBox ID="sjdate" runat="server" Columns="10" onfocus="this.blur()"></asp:TextBox>
											<span id="span9" runat="server"><img id="imgsjdate" style="cursor: hand" onclick="javascript:show_calendar('<%= sjdate.clientid %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="24" height="24" /></span>
										</td>
									</tr>
									<tr>
										<td class="bluecol" width="114">
											最近升遷日期
										</td>
										<td class="whitecol" colspan="3">
											<asp:TextBox ID="spdate" runat="server" Columns="10" onfocus="this.blur()"></asp:TextBox>
											<span id="span10" runat="server"><img id="imgspdate" style="cursor: hand" onclick="javascript:show_calendar('<%= spdate.clientid %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="24" height="24" /></span>
										</td>
									</tr>
									<tr>
										<td class="bluecol" align="center" colspan="4">
											參訓背景資料
										</td>
									</tr>
									<tr>
										<td class="bluecol_need" width="114">
											是否由公司<br />
											推薦參訓
										</td>
										<td class="whitecol" colspan="3">
											<asp:RadioButtonList ID="q1" runat="server" RepeatDirection="horizontal" RepeatLayout="flow" CssClass="font">
												<asp:ListItem Value="1">是</asp:ListItem>
												<asp:ListItem Value="0">否</asp:ListItem>
											</asp:RadioButtonList>
										</td>
									</tr>
									<tr>
										<td class="bluecol_need" width="114">
											參訓動機
										</td>
										<td class="whitecol" colspan="3">
											<asp:CheckBoxList ID="q2" runat="server" RepeatDirection="horizontal" CssClass="font" RepeatColumns="2" CellSpacing="0" CellPadding="0">
												<asp:ListItem Value="1">為補充與原專長相關之技能</asp:ListItem>
												<asp:ListItem Value="2">轉換其他行職業所需技能</asp:ListItem>
												<asp:ListItem Value="3">拓展工作領域及視野</asp:ListItem>
												<asp:ListItem Value="4">其他</asp:ListItem>
											</asp:CheckBoxList>
										</td>
									</tr>
									<tr>
										<td class="bluecol" width="114">
											訓後動向
										</td>
										<td class="whitecol" onclick="check1();" colspan="3">
											<asp:RadioButtonList ID="q3" runat="server" RepeatDirection="horizontal" RepeatLayout="flow" CssClass="font">
												<asp:ListItem Value="1">轉換工作</asp:ListItem>
												<asp:ListItem Value="2">留任</asp:ListItem>
												<asp:ListItem Value="3">其他</asp:ListItem>
											</asp:RadioButtonList>
											<asp:TextBox ID="q3_other" runat="server" Enabled="false"></asp:TextBox>
										</td>
									</tr>
									<tr>
										<td class="bluecol_need" width="114">
											服務單位行業別
										</td>
										<td class="whitecol" colspan="3">
											<asp:DropDownList ID="q4" runat="server" Width="344px">
											</asp:DropDownList>
										</td>
									</tr>
									<tr>
										<td class="bluecol_need" width="114">
											服務單位是否<br />
											屬於中小企業
										</td>
										<td class="whitecol" colspan="3">
											<asp:RadioButtonList ID="q5" runat="server" RepeatDirection="horizontal" RepeatLayout="flow" CssClass="font">
												<asp:ListItem Value="1">是</asp:ListItem>
												<asp:ListItem Value="0">否</asp:ListItem>
											</asp:RadioButtonList>
										</td>
									</tr>
									<tr>
										<td class="bluecol" width="114">
											個人工作年資
										</td>
										<td class="whitecol">
											<asp:TextBox ID="q61" runat="server" Columns="5"></asp:TextBox>
										</td>
										<td class="bluecol" align="center">
											在這家公司的年資
										</td>
										<td class="whitecol">
											<asp:TextBox ID="q62" runat="server" Columns="5"></asp:TextBox>
										</td>
									</tr>
									<tr>
										<td class="bluecol" width="114">
											在這職位的年資
										</td>
										<td class="whitecol">
											<asp:TextBox ID="q63" runat="server" Columns="5"></asp:TextBox>
										</td>
										<td class="bluecol" align="center">
											最近升遷離本職幾年
										</td>
										<td class="whitecol">
											<asp:TextBox ID="q64" runat="server" Columns="5"></asp:TextBox>
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td>
								<table class="font" id="historytable" style="width: 735px" cellspacing="1" cellpadding="1" width="735" align="left" border="0" runat="server">
									<tr>
										<td align="right" colspan="4">
											(為避免消耗主機效能，最大搜尋筆數為2000筆)共計：
											<asp:Label ID="recordcount" runat="server">label</asp:Label>筆資料
										</td>
									</tr>
									<tr>
										<td align="center" colspan="4">
											<asp:DataGrid ID="datagrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="false">
												<AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
												<HeaderStyle CssClass="head_navy" />
												<Columns>
													<asp:BoundColumn HeaderText="序號">
														<HeaderStyle HorizontalAlign="Center" Width="15px"></HeaderStyle>
														<ItemStyle HorizontalAlign="Center"></ItemStyle>
													</asp:BoundColumn>
													<asp:BoundColumn DataField="distname" HeaderText="轄區&lt;br&gt;中心">
														<HeaderStyle HorizontalAlign="Center" Width="70px"></HeaderStyle>
														<ItemStyle HorizontalAlign="Center"></ItemStyle>
													</asp:BoundColumn>
													<asp:BoundColumn DataField="years" HeaderText="年度">
														<HeaderStyle Width="25px"></HeaderStyle>
													</asp:BoundColumn>
													<asp:BoundColumn DataField="planname" HeaderText="訓練計畫">
														<HeaderStyle Width="100px"></HeaderStyle>
													</asp:BoundColumn>
													<asp:BoundColumn DataField="orgname" HeaderText="訓練機構">
														<HeaderStyle Width="100px"></HeaderStyle>
													</asp:BoundColumn>
													<asp:BoundColumn DataField="tmid" HeaderText="訓練職類">
														<HeaderStyle Width="100px"></HeaderStyle>
													</asp:BoundColumn>
													<asp:BoundColumn DataField="classname" HeaderText="班別名稱">
														<HeaderStyle Width="100px"></HeaderStyle>
													</asp:BoundColumn>
													<asp:BoundColumn DataField="thours" HeaderText="受訓&lt;br&gt;時數">
														<HeaderStyle Width="25px"></HeaderStyle>
													</asp:BoundColumn>
													<asp:BoundColumn DataField="tround" HeaderText="受訓期間">
														<HeaderStyle HorizontalAlign="Center" Width="50px"></HeaderStyle>
														<ItemStyle HorizontalAlign="Center"></ItemStyle>
													</asp:BoundColumn>
													<asp:BoundColumn DataField="skillname" HeaderText="技能檢定">
														<HeaderStyle Width="60px"></HeaderStyle>
													</asp:BoundColumn>
													<asp:BoundColumn DataField="tflag" HeaderText="訓練&lt;br&gt;狀態">
														<HeaderStyle Width="25px"></HeaderStyle>
													</asp:BoundColumn>
												</Columns>
												<PagerStyle Visible="false"></PagerStyle>
											</asp:DataGrid>
										</td>
									</tr>
									<tr>
										<td align="center" colspan="4">
											<asp:Label ID="msg" runat="server" ForeColor="red"></asp:Label>
										</td>
									</tr>
								</table>
							</td>
						</tr>
					</tbody>
				</table>
				<table id="table4" style="width: 739px; height: 28px" cellspacing="1" cellpadding="1" width="739" align="center" border="0">
					<tr>
						<td align="center">
							<asp:Button ID="send" runat="server" Text="送出" CssClass="asp_button_S"></asp:Button>
							<asp:Button ID="button22" runat="server" Text="回報名登錄" CssClass="asp_button_M"></asp:Button>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
	<input id="hidstar3" type="hidden" size="1" name="hidstar3" runat="server">
	<input id="lociddate" type="hidden" size="1" name="lociddate" runat="server">
	</form>
</body>
</html>
