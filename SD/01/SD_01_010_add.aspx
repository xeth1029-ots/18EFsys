<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_01_010_add.aspx.vb" Inherits="WDAIIP.SD_01_010_add" %>

<!DOCTYPE html PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>報名作業(產業人才投資方案)</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function ChangeAcctMode() {
            if (document.forms["form1"].AcctMode_0.checked) {
                //document.getElementById('Porttr').style.display='inline';
                document.getElementById('Porttr').style.display = '';
                document.getElementById('Banktr1').style.display = 'none';
                document.getElementById('Banktr2').style.display = 'none';
                document.getElementById('Banktr3').style.display = 'none';
            } else if (document.forms["form1"].AcctMode_1.checked) {
                document.getElementById('Porttr').style.display = 'none';
                document.getElementById('Banktr1').style.display = '';
                document.getElementById('Banktr2').style.display = '';
                document.getElementById('Banktr3').style.display = '';
                //document.getElementById('Banktr1').style.display='inline';
                //document.getElementById('Banktr2').style.display='inline';
                //document.getElementById('Banktr3').style.display='inline';	
            } else if (document.forms["form1"].AcctMode_2 != null) {
                if (document.forms["form1"].AcctMode_2.checked) {
                    document.getElementById('Porttr').style.display = 'none';
                    document.getElementById('Banktr1').style.display = 'none';
                    document.getElementById('Banktr2').style.display = 'none';
                    document.getElementById('Banktr3').style.display = 'none';
                }
            }
        }

        function Clear_Zip2() {
            if (document.forms["form1"].CheckBox1.checked == true) {
                document.getElementById('ZipCode2').value = document.getElementById('ZipCode1').value;
                document.getElementById('City2').value = document.getElementById('City1').value;
                document.getElementById('ZipCode2_B3').value = document.getElementById('ZipCode1_B3').value;
                document.getElementById('HouseholdAddress').value = document.getElementById('Address').value;
            }
            else {
                if (isEmpty('ZipCode2')) {
                    document.getElementById('ZipCode2').value = document.getElementById('ZipCode1').value;
                    document.getElementById('City2').value = document.getElementById('City1').value;
                }
                if (isEmpty('ZipCode2_B3')) {
                    document.getElementById('ZipCode2_B3').value = document.getElementById('ZipCode1_B3').value;
                }
                if (isEmpty('HouseholdAddress')) {
                    document.getElementById('HouseholdAddress').value = document.getElementById('Address').value;
                }
            }
        }

        function check1() {
            if (document.forms["form1"].Q3_2.checked) {
                document.forms["form1"].Q3_Other.disabled = false;
            }
            else {
                document.forms["form1"].Q3_Other.disabled = true;
            }
        }

        function ChkData() {
            var msg = '';
            if (isEmpty('Name')) { msg += '請輸入中文姓名\n'; }
            if (isEmpty('PassPortNO')) { msg += '請選擇身分別\n'; }
            if (isEmpty('IDNO')) {
                msg += '請輸入身分證字號!!!\n';
            }
            else {
                if (document.forms["form1"].PassPortNO[0].checked == true) {
                    if (!checkId(document.forms["form1"].IDNO.value)) {
                        msg += '身分證號碼錯誤(如果有此身分證號碼，請聯絡系統管理者)!!!\n';
                    }
                }
            }
            if (isEmpty('Sex')) { msg += '請選擇性別\n'; }
            if (document.forms["form1"].DegreeID.selectedIndex == 0) { msg += '請選擇最高學歷\n'; }
            if (isEmpty('GraduateStatus')) { msg += '請選擇畢業狀況\n'; }
            if (isEmpty('School')) { msg += '請輸入學校名稱\n'; }
            if (isEmpty('Department')) { msg += '請輸入科系名稱\n'; }
            //if (isEmpty('School')) { msg += '請輸入學校名稱\n'; }
            //if (isEmpty('Department')) { msg += '請輸入科系\n'; }
            //if (document.forms["form1"].GraduateStatus.selectedIndex==0) { msg += '請選擇畢業狀況\n'; }
            //if (document.forms["form1"].MilitaryID.value=='0') { msg += '請選擇兵役狀況\n'; }
            //if (document.forms["form1"].MaritalStatus.selectedIndex == 0) { msg += '請選擇婚姻狀況\n'; }
            //rblMobil
            if (!isChecked(document.forms["form1"].rblMobil)) {
                msg += '請選擇有無行動電話\n';
            }
            else {
                if (getRadioValue(document.forms["form1"].rblMobil) == "Y") {
                    if (document.forms["form1"].CellPhone.value == '') { msg += '有行動電話 請輸入行動電話\n'; }
                }
                else {
                    if (isEmpty('PhoneD')) { msg += '請輸入聯絡電話(日)\n'; }
                    //if (document.forms["form1"].PhoneD.value=='') {msg+='請輸入聯絡電話(日)\n';}
                    if (document.forms["form1"].CellPhone.value != '') { msg += '有輸入行動電話,請選擇有行動電話\n'; }
                }
            }
            if (isEmpty('Address')) { msg += '請輸入通訊地址\n'; }
            if (isEmpty('ZipCode1')) { msg += '請選擇通訊地址-縣市\n'; }
            //checkzip23 郵遞區號
            msg += checkzip23(true, '通訊地址', 'ZipCode1_B3');
            //if (!isEmpty('ZipCode1_B3')) {
            //    if (!isInt(getValue('ZipCode1_B3'))) { msg += '通訊地址郵遞區號2碼.請輸入數字!!!\n'; }
            //} else {
            //    msg += '請輸入通訊地址郵遞區號2碼!!!\n';
            //}

            if (isEmpty('HouseholdAddress')) { msg += '請輸入戶籍地址\n'; }
            if (isEmpty('ZipCode2')) { msg += '請選擇戶籍地址-縣市\n'; }
            //checkzip23 郵遞區號
            msg += checkzip23(true, '戶籍地址', 'ZipCode2_B3');
            //if (!isEmpty('ZipCode2_B3')) {
            //    if (!isInt(getValue('ZipCode2_B3'))) { msg += '戶籍地址郵遞區號2碼.請輸入數字!!!\n'; }
            //} else {
            //    msg += '請輸入戶籍地址郵遞區號2碼!!!\n';
            //}

            //rblEmail
            var emailValue = "N";
            if (!isChecked(document.forms["form1"].rblEmail)) {
                //未點選
                msg += '請選擇有無電子郵件\n';
            }
            else {
                if (getRadioValue(document.forms["form1"].rblEmail) == "Y") {
                    //點選 有電子郵件
                    if (isEmpty('Email')) { msg += '有電子郵件 請輸入電子郵件\n'; } else { emailValue = "Y";}
                }
                else {
                    //點選 無電子郵件 emailValue = "N";
                    if (trim(document.getElementById('Email').value) != '無') {
                        if (!isEmpty('Email')) { msg += '有輸入電子郵件,請選擇有電子郵件\n'; }
                    }
                }
            }
            // if (isEmpty('Email')) { msg += '請輸入電子郵件\n'; }
            if (document.forms["form1"].MIdentityID.selectedIndex == 0) { msg += '請選擇主要參訓身分別\n'; }
            //if (document.forms["form1"].IdentityID.selectedIndex==0) { msg += '請選擇參訓身分別\n'; }
            //if (document.forms["form1"].IdentityID.selectedIndex==5){ 
            //  if (document.forms["form1"].HandTypeID.selectedIndex==0) { msg += '請選擇障礙類別\n'; }	
            //  if (document.forms["form1"].HandLevelID.selectedIndex==0) { msg += '請選擇障礙等級\n'; }	
            //}
            //if (!isEmpty('PriorWorkOrg1')) { 
            //	if (isEmpty('Title1')) { msg += '請輸入職稱1\n'; }
            //}	
            //if (!isEmpty('Title1')) { 
            //	if (isEmpty('PriorWorkOrg1')) { msg += '請輸入受訓服務單位1\n'; }
            //}
            //if (!isEmpty('PriorWorkOrg2')) { 
            //	if (isEmpty('Title2')) { msg += '請輸入職稱2\n'; }
            //}	
            //if (!isEmpty('Title2')) { 
            //	if (isEmpty('PriorWorkOrg2')) { msg += '請輸入受訓服務單位2\n'; }
            //}
            //if (!isEmpty('PriorWorkPay')) {
            //    if (!isInt(getValue('PriorWorkPay'))) { msg += '受訓前薪資.請輸入數字!!!\n'; }
            //}
            //else { msg += '請輸入受訓前薪資!!!\n'; }
            //if (isEmpty('ShowDetail')) { msg += '請選擇是否提供基本資料供廠商查詢\n'; }
            /*
			if (isEmpty('AcctMode')) { 
			msg += '請選擇郵局帳號或銀行帳號\n'; }
			else {
			if (document.forms["form1"].AcctMode_0.checked) {
			if (isEmpty('PostNo_1')) { msg += '請輸入郵局-局號1\n'; }
			if (isEmpty('PostNo_2')) { msg += '請輸入郵局-局號2\n'; }    
			if (isEmpty('AcctNo1_1')) { msg += '請輸入郵局-帳號1\n'; } 
			if (isEmpty('AcctNo1_2')) { msg += '請輸入郵局-帳號2\n'; } 
			}else if(document.forms["form1"].AcctMode_1.checked){
			if (isEmpty('BankName')) { msg += '請輸入總行名稱\n'; }
			if (isEmpty('AcctheadNo')) { msg += '請輸入總行代號\n'; }
			if (isEmpty('ExBankName')) { msg += '請輸入分行名稱\n'; }
			if (isEmpty('AcctExNo')) { msg += '請輸入分行代號\n'; }     
			if (isEmpty('AcctNo2')) { msg += '請輸入銀行帳號\n'; } 
			}
			}	
			*/
            if (document.forms["form1"].MIdentityID.value != "02") { if (isEmpty('Uname')) { msg += '請輸入服務單位\n'; } }
            //if (isEmpty('Intaxno')) { msg += '請輸入服務單位統一編號\n'; }
            //if (isEmpty('ServDept')) { msg += '請輸入服務部門\n'; }
            if (document.getElementById('ServDept') && isEmpty('ServDept')) { msg += '請輸入服務部門\n'; }
            if (document.getElementById('ddlSERVDEPTID') && isEmpty('ddlSERVDEPTID')) { msg += '請選擇服務部門\n'; }
            if (document.forms["form1"].MIdentityID.value != "02") { if (isEmpty('ActName')) { msg += '請輸入投保單位名稱\n'; } }
            if (isEmpty('ActType')) { msg += '請輸入投保類別\n'; }
            var ActNo = document.getElementById('ActNo');
            //if (ActNo && document.forms["form1"].MIdentityID.value != "02") {
            //	if (isEmpty('ActNo')) {
            //		msg += '請輸入投保單位保險證號\n';
            //	} else {
            //		var flagOk1 = false;
            //		if (ActNo.value.substring(0, 1) == '0') { flagOk1 = true; }
            //		if (ActNo.value.substring(0, 1) == '1') { flagOk1 = true; }
            //		if (!flagOk1) {
            //			msg += '投保單位保險證號輸入錯誤，第1碼應為0或1\n';
            //		}
            //	}
            //}
            //if (document.forms["form1"].MIdentityID.value != "02") { if (isEmpty('ActTel')) { msg += '請輸入投保單位電話\n'; } }
            //if (document.forms["form1"].MIdentityID.value != "02") { if (isEmpty('ZipCode3')) { msg += '請選擇投保單位地址-縣市\n'; } }
            if (document.forms["form1"].MIdentityID.value != "02") {
                if (!isEmpty('ZipCode3_B3')) {
                    //checkzip23 郵遞區號
                    msg += checkzip23(true, '投保單位地址', 'ZipCode3_B3');
                    //if (!isUnsignedInt(trim(document.getElementById('ZipCode3_B3').value))) { msg += '投保單位地址郵遞區號後2碼必須為數字，且不得輸入 00\n'; }
                    //if (parseInt(trim(document.getElementById('ZipCode3_B3').value), 10) < 1) { msg += '投保單位地址郵遞區號後2碼必須為數字，得輸入 01~99 \n'; }
                    //if (parseInt(trim(document.getElementById('ZipCode3_B3').value), 10) > 99) { msg += '投保單位地址郵遞區號後2碼必須為數字，得輸入 01~99 \n'; }
                    //if (trim(document.getElementById('ZipCode3_B3').value).length != 2) { msg += '投保單位地址郵遞區號後2碼長度必須為 2 碼(例 01 或 99)\n'; }
                }
                //else { msg += '請輸入投保單位地址郵遞區號2碼!!!\n'; }
            }
            //if (document.forms["form1"].MIdentityID.value != "02") { if (isEmpty('ActAddress')) { msg += '請輸入投保單位地址\n'; } }
            //if (isEmpty('Scale')) { msg += '請輸入服務單位規模\n'; } 
            //MIdentityID.value != "02" 非自願離職者
            if (document.forms["form1"].MIdentityID.value != "02") {
                if (document.getElementById('JobTitle') && isEmpty('JobTitle')) { msg += '請輸入職務\n'; }
                if (document.getElementById('ddlJOBTITLEID') && isEmpty('ddlJOBTITLEID')) { msg += '請選擇職務\n'; }
            }
            //if (document.forms["form1"].MIdentityID.value != "02") { if (isEmpty('JobTitle')) { msg += '請輸入職稱\n'; } }
            if (isEmpty('Q1')) { msg += '請選擇參訓資料背景-是否由公司推薦參訓\n'; }
            if (isEmpty('Q2')) { msg += '請選擇參訓資料背景-參訓動機\n'; }
            //if (isEmpty('Q3')) { msg += '請選擇參訓資料背景-訓後動向\n'; }
            if (isEmpty('Q3')) { msg += '請選擇參訓資料背景-結訓後之計畫\n'; }
            if (document.forms["form1"].MIdentityID.value != "02") { if (isEmpty('Q4')) { msg += '請選擇參訓資料背景-服務單位行業別\n'; } }
            if (isEmpty('Q5')) { msg += '請選擇服務單位是否屬於中小企業\n'; }
            if (!isEmpty('Q61')) {
                if (!isInt(getValue('Q61')) && !isFloat(getValue('Q61'))) { msg += '個人工作年資.請輸入數字!!!\n'; }
            }
            else { msg += '請輸入個人工作年資!!!\n'; }
            if (!isEmpty('Q62')) {
                if (!isInt(getValue('Q62')) && !isFloat(getValue('Q62'))) { msg += '在這家公司的年資.請輸入數字!!!\n'; }
            }
            else { msg += '請輸入在這家公司的年資!!!\n'; }
            if (!isEmpty('Q63')) {
                if (!isInt(getValue('Q63')) && !isFloat(getValue('Q63'))) { msg += '在這職位的年資.請輸入數字!!!\n'; }
            }
            else { msg += '請輸入在這職位的年資!!!\n'; }

            if (!isEmpty('Q64')) {
                if (!isInt(getValue('Q64')) && !isFloat(getValue('Q64'))) { msg += '最近升遷離本職幾年.請輸入數字!!!\n'; }
            }
            else { msg += '請輸入最近升遷離本職幾年!!!\n'; }
            //if (isEmpty('IsAgree')) { msg += '請選擇是否同意個人基本資料供查詢\n'; }
            if (isEmpty('IseMail')) { msg += '請選擇是否希望收到最新課程資訊\n'; }
            else {
                var iseMailVal = getRadioValue(document.forms["form1"].IseMail);
                if (iseMailVal == "Y") {
                    if (isEmpty('Email') || document.getElementById('Email').value == '無' || emailValue=='N') {
                        msg += '您選擇「希望」收到最新課程資訊，請輸入有效的電子郵件地址\n';
                    }
                }
            }
            //if (isEmpty('IsCorrect')) { msg += '請確認以上資料為最新且正確\n'; }
            //if (isEmpty('IsAgreedata')) { msg += '請選擇是否同意個人資料用於上開所列之合理範圍內\n'; }
            var rst = true;
            if (msg == '') {
                rst = chk_Actno(); //ActNo
                if (!rst) { return false; }
                //鎖定按鈕邏輯 
                var btn = document.getElementById('<%= btnSend1.ClientID %>');
                var btnD = document.getElementById('btnSend1D');
                if (btn && btnD) { btn.style.display = 'none'; btnD.style.display = ''; }
                return true;
            } else {
                alert(msg);
                return false;
            }
        }

        function chk_Actno() {
            var ActNo = document.getElementById('ActNo');
            var flagNG1 = false;
            if (ActNo.value.substring(0, 3) == '076') { flagNG1 = true; }
            if (ActNo.value.substring(0, 3) == '075') { flagNG1 = true; }
            if (ActNo.value.substring(0, 3) == '175') { flagNG1 = true; }
            if (ActNo.value.substring(0, 2) == '09') { flagNG1 = true; }
            if (flagNG1) {
                //投保證號開頭數字為075、175（裁減續保）、076（職災續保）、09（訓）皆為不予補助對象，確定送出？
                return confirm('投保證號開頭數字為075、175、076、09為不予補助對象，是否確定送出？');
            }
            return true;
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;報名作業</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr id="trbasic1" runat="server">
                <td align="center">
                    <table id="table1" width="100%" cellspacing="1" cellpadding="1" align="center" border="0">
                        <tbody>
                            <tr>
                                <td>
                                    <table class="table_nw" id="table3" cellspacing="1" cellpadding="1" width="100%">
                                        <tr>
                                            <td align="center" class="table_title">個人報名基本資料</td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                    <table id="table8" class="table_nw" cellspacing="1" cellpadding="1" width="100%">
                        <tbody>
                            <tr>
                                <td>
                                    <table class="table_nw" id="DetailTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                        <tbody>
                                            <tr>
                                                <td class="bluecol" style="width: 20%">班級名稱</td>
                                                <td class="whitecol" colspan="3">
                                                    <asp:Label ID="ClassName" runat="server" Width="30%"></asp:Label>
                                                    <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                                    <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                                    <%--<asp:Label ID="LOCIDdate" runat="server" Width="16px" Visible="False"></asp:Label>--%>
                                                    <asp:HiddenField ID="Hid_LOCIDdate" runat="server" />
                                                </td>
                                            </tr>
                                            <tr>
                                                <td class="bluecol"></td>
                                                <td class="whitecol" colspan="3">
                                                    <asp:Label ID="GovCost" runat="server" ForeColor="Red" Font-Bold="true"></asp:Label></td>
                                            </tr>
                                            <tr>
                                                <td class="bluecol_need">中文姓名</td>
                                                <td class="whitecol" colspan="3">
                                                    <asp:TextBox ID="Name" runat="server" MaxLength="20" Columns="15" Width="20%"></asp:TextBox></td>
                                            </tr>
                                            <tr id="Tr1" runat="server">
                                                <td class="bluecol_need" style="width: 20%">身分別</td>
                                                <td class="whitecol" style="width: 30%">
                                                    <asp:RadioButtonList ID="PassPortNO" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                        <asp:ListItem Value="1">本國</asp:ListItem>
                                                        <asp:ListItem Value="2">外籍(含大陸人士)</asp:ListItem>
                                                    </asp:RadioButtonList>
                                                </td>
                                                <td class="bluecol_need" style="width: 20%">身分證號碼</td>
                                                <td class="whitecol" style="width: 30%">
                                                    <asp:TextBox ID="IDNO" runat="server" MaxLength="12" Columns="15" Enabled="False" onfocus="this.blur()" Width="50%"></asp:TextBox><font size="2"></font></td>
                                            </tr>
                                            <tr>
                                                <td class="bluecol_need">性別</td>
                                                <td class="whitecol">
                                                    <asp:RadioButtonList ID="Sex" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                        <asp:ListItem Value="M">男</asp:ListItem>
                                                        <asp:ListItem Value="F">女</asp:ListItem>
                                                    </asp:RadioButtonList>
                                                </td>
                                                <td class="bluecol_need">出生日期</td>
                                                <td class="whitecol">
                                                    <asp:TextBox ID="Birthday" runat="server" MaxLength="10" Enabled="False" onfocus="this.blur()" Width="50%"></asp:TextBox></td>
                                            </tr>
                                            <tr>
                                                <td class="bluecol_need">最高學歷</td>
                                                <td class="whitecol">
                                                    <asp:DropDownList ID="DegreeID" runat="server"></asp:DropDownList></td>
                                                <td class="bluecol_need">畢業狀況</td>
                                                <td class="whitecol">
                                                    <asp:RadioButtonList ID="GraduateStatus" runat="server" CssClass="font" RepeatDirection="Horizontal"></asp:RadioButtonList></td>
                                            </tr>
                                            <tr>
                                                <td class="bluecol_need">學校名稱</td>
                                                <td class="whitecol">
                                                    <asp:TextBox ID="School" runat="server" Columns="30" MaxLength="30" Width="90%"></asp:TextBox></td>
                                                <td class="bluecol_need">科系名稱</td>
                                                <td class="whitecol">
                                                    <asp:TextBox ID="Department" runat="server" Columns="60" MaxLength="120" Width="90%"></asp:TextBox></td>
                                            </tr>
                                            <tr>
                                                <td class="bluecol">婚姻狀況</td>
                                                <td class="whitecol" colspan="3">
                                                    <asp:DropDownList ID="MaritalStatus" runat="server">
                                                        <asp:ListItem Value="">請選擇</asp:ListItem>
                                                        <asp:ListItem Value="1">已婚</asp:ListItem>
                                                        <asp:ListItem Value="2">未婚</asp:ListItem>
                                                        <asp:ListItem Value="3">暫不提供</asp:ListItem>
                                                    </asp:DropDownList>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td class="bluecol_need">聯絡電話</td>
                                                <td class="whitecol">
                                                    <table class="font" id="table7" cellspacing="1" cellpadding="1" width="100%" border="0">
                                                        <tr>
                                                            <td style="width: 10%">(日)</td>
                                                            <td>
                                                                <asp:TextBox ID="PhoneD" runat="server" Columns="13" Width="60%" MaxLength="25"></asp:TextBox></td>
                                                        </tr>
                                                        <tr>
                                                            <td>(夜)</td>
                                                            <td>
                                                                <asp:TextBox ID="PhoneN" runat="server" Columns="13" Width="60%" MaxLength="25"></asp:TextBox></td>
                                                        </tr>
                                                    </table>
                                                </td>
                                                <td class="bluecol_need">行動電話</td>
                                                <td class="whitecol">
                                                    <asp:RadioButtonList ID="rblMobil" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                                        <asp:ListItem Value="N">無</asp:ListItem>
                                                        <asp:ListItem Value="Y">有</asp:ListItem>
                                                    </asp:RadioButtonList>
                                                    <asp:TextBox ID="CellPhone" runat="server" Columns="13" Width="60%" MaxLength="25"></asp:TextBox>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td class="bluecol_need">通訊地址</td>
                                                <td class="whitecol" colspan="3">
                                                    <input id="Bt1_city_zip" type="button" value="..." name="city_zip" runat="server" class="asp_button_Mini" />
                                                    <input id="ZipCode1" style="width: 12%;" maxlength="5" name="ZipCode1" runat="server" />－
                                                    <input id="ZipCode1_B3" style="width: 10%;" maxlength="3" name="ZipCode1_B3" runat="server" />
                                                    <input id="hidZipCode1_6W" type="hidden" runat="server" />
                                                    <asp:Literal ID="LitZip1" runat="server"></asp:Literal><br />
                                                    <%--3+3郵遞區號查詢--%>
                                                    <asp:TextBox ID="City1" runat="server" Width="25%" onfocus="this.blur()"></asp:TextBox>
                                                    <asp:TextBox ID="Address" runat="server" Width="60%" MaxLength="200"></asp:TextBox>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td class="bluecol_need">戶籍地址</td>
                                                <td class="whitecol" colspan="3">
                                                    <asp:CheckBox ID="CheckBox1" runat="server" CssClass="font" Text="同通訊地址"></asp:CheckBox><br>
                                                    <input id="Button1" type="button" value="..." name="city_zip" runat="server" class="asp_button_Mini" />
                                                    <input id="ZipCode2" style="width: 12%;" maxlength="5" name="ZipCode2" runat="server" />－
                                                    <input id="ZipCode2_B3" style="width: 10%;" maxlength="3" name="ZipCode2_B3" runat="server" />
                                                    <input id="hidZipCode2_6W" type="hidden" runat="server" />
                                                    <asp:Literal ID="LitZip2" runat="server"></asp:Literal><br />
                                                    <%--3+3郵遞區號查詢--%>
                                                    <asp:TextBox ID="City2" runat="server" Width="25%" onfocus="this.blur()"></asp:TextBox>
                                                    <asp:TextBox ID="HouseholdAddress" runat="server" Width="60%" MaxLength="200"></asp:TextBox>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td class="bluecol_need">電子郵件</td>
                                                <td class="whitecol" colspan="3">
                                                    <asp:RadioButtonList ID="rblEmail" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                                        <asp:ListItem Value="N">無</asp:ListItem>
                                                        <asp:ListItem Value="Y">有</asp:ListItem>
                                                    </asp:RadioButtonList>
                                                    <asp:TextBox ID="Email" runat="server" Width="60%" MaxLength="60"></asp:TextBox>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td class="bluecol_need">主要參訓身分別</td>
                                                <td class="whitecol" colspan="3">
                                                    <asp:DropDownList ID="MIdentityID" runat="server"></asp:DropDownList></td>
                                            </tr>
                                            <%-- <tr>
                                                <td class="bluecol">受訓前薪資</td>
                                                <td class="whitecol" colspan="3">
                                                    <asp:TextBox ID="PriorWorkPay" runat="server" Width="20%" MaxLength="10"></asp:TextBox></td>
                                            </tr>--%>
                                            <tr>
                                                <td class="bluecol">郵政/銀行帳號</td>
                                                <td onclick="ChangeAcctMode();" class="whitecol" colspan="3">
                                                    <asp:RadioButtonList ID="AcctMode" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                                        <asp:ListItem Value="0">郵局帳號</asp:ListItem>
                                                        <asp:ListItem Value="1">銀行帳號</asp:ListItem>
                                                        <asp:ListItem Value="2">訓練單位代轉現金</asp:ListItem>
                                                    </asp:RadioButtonList>
                                                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:HyperLink ID="HL_finaCodeQuery" runat="server" Target="_blank">金融機構代碼查詢</asp:HyperLink>
                                                </td>
                                            </tr>
                                            <tr id="Porttr" runat="server">
                                                <td class="bluecol">局號</td>
                                                <td class="whitecol">
                                                    <asp:TextBox ID="PostNo_1" runat="server" Width="60%" Columns="10" MaxLength="30"></asp:TextBox>
                                                </td>
                                                <td class="bluecol">帳號</td>
                                                <td class="whitecol">
                                                    <asp:TextBox ID="AcctNo1_1" runat="server" Width="60%" Columns="10" MaxLength="30"></asp:TextBox>
                                                </td>
                                            </tr>
                                            <tr id="Banktr1" runat="server">
                                                <td class="bluecol">總行名稱</td>
                                                <td class="whitecol">
                                                    <asp:TextBox ID="BankName" runat="server" Width="60%" MaxLength="30"></asp:TextBox></td>
                                                <td class="bluecol">總行代號</td>
                                                <td class="whitecol">
                                                    <asp:TextBox ID="AcctheadNo" runat="server" Width="60%" MaxLength="30"></asp:TextBox></td>
                                            </tr>
                                            <tr id="Banktr2" runat="server">
                                                <td class="bluecol">分行名稱</td>
                                                <td class="whitecol">
                                                    <asp:TextBox ID="ExBankName" runat="server" Width="60%" MaxLength="30"></asp:TextBox></td>
                                                <td class="bluecol">分行代號</td>
                                                <td class="whitecol">
                                                    <asp:TextBox ID="AcctExNo" runat="server" Width="60%" MaxLength="30"></asp:TextBox></td>
                                            </tr>
                                            <tr id="Banktr3" runat="server">
                                                <td class="bluecol">帳號</td>
                                                <td class="whitecol" colspan="3">
                                                    <asp:TextBox ID="AcctNo2" runat="server" Width="50%" MaxLength="30"></asp:TextBox></td>
                                            </tr>
                                        </tbody>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <table class="table_nw" id="BackTable" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                        <tr>
                                            <td class="table_title" colspan="4">服務單位資料</td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol_need" style="width: 20%">服務單位</td>
                                            <td class="whitecol" style="width: 30%">
                                                <asp:TextBox ID="Uname" runat="server" MaxLength="50" Width="60%"></asp:TextBox></td>
                                            <td class="bluecol" style="width: 20%">統一編號</td>
                                            <td class="whitecol" style="width: 30%">
                                                <asp:TextBox ID="Intaxno" runat="server" Width="50%" MaxLength="30"></asp:TextBox></td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol_need">服務部門</td>
                                            <td class="whitecol" colspan="3">
                                                <asp:TextBox ID="ServDept" runat="server" MaxLength="50" Width="60%"></asp:TextBox>
                                                <asp:DropDownList ID="ddlSERVDEPTID" runat="server"></asp:DropDownList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol_need">投保單位名稱</td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="ActName" runat="server" Width="60%" MaxLength="50"></asp:TextBox></td>
                                            <td class="bluecol">投保類別</td>
                                            <td class="whitecol">
                                                <asp:DropDownList ID="ActType" runat="server">
                                                    <asp:ListItem Value="1" Selected="true">勞保</asp:ListItem>
                                                    <asp:ListItem Value="2">農保</asp:ListItem>
                                                </asp:DropDownList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">投保單位保險證號</td>
                                            <td class="whitecol" colspan="3">
                                                <asp:TextBox ID="ActNo" runat="server" Width="25%" MaxLength="20"></asp:TextBox>
                                                <font color="blue">請確實填寫正確的勞工保險卡資料</font>
                                                <%-- 請依照您目前工作之勞工保險卡確實填寫，並於報名繳費時繳交勞保卡影本--%>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">投保單位電話</td>
                                            <td class="whitecol" colspan="3">
                                                <asp:TextBox ID="ActTel" runat="server" MaxLength="30" Width="20%"></asp:TextBox></td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">投保單位地址</td>
                                            <td class="whitecol" colspan="3">
                                                <input id="Button2" type="button" value="..." name="city_zip" runat="server" class="asp_button_Mini" />
                                                <input id="ZipCode3" style="width: 12%;" maxlength="3" name="ZipCode3" runat="server" />－
                                                <input id="ZipCode3_B3" style="width: 10%;" maxlength="3" name="ZipCode3_B3" runat="server" />
                                                <input id="hidZipCode3_6W" type="hidden" runat="server" />
                                                <asp:Literal ID="LitZip3" runat="server"></asp:Literal><br />
                                                <%--3+3郵遞區號查詢--%>
                                                <asp:TextBox ID="City3" runat="server" Width="25%" onfocus="this.blur()" MaxLength="100"></asp:TextBox>
                                                <asp:TextBox ID="ActAddress" runat="server" Width="60%" MaxLength="100"></asp:TextBox>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol_need">職稱</td>
                                            <td class="whitecol" colspan="3">
                                                <asp:TextBox ID="JobTitle" runat="server" MaxLength="50" Width="30%"></asp:TextBox>
                                                <asp:DropDownList ID="ddlJOBTITLEID" runat="server"></asp:DropDownList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="table_title" colspan="4">參訓背景資料</td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol_need">是否由公司推薦參訓</td>
                                            <td class="whitecol" colspan="3">
                                                <asp:RadioButtonList ID="Q1" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                                    <asp:ListItem Value="1">是</asp:ListItem>
                                                    <asp:ListItem Value="0">否</asp:ListItem>
                                                </asp:RadioButtonList>
                                                <font size="2"></font>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol_need">參訓動機</td>
                                            <td class="whitecol" colspan="3">
                                                <asp:CheckBoxList ID="Q2" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="2" CellPadding="0" CellSpacing="0">
                                                    <asp:ListItem Value="1">為補充與原專長相關之技能</asp:ListItem>
                                                    <asp:ListItem Value="2">轉換其他行職業所需技能</asp:ListItem>
                                                    <asp:ListItem Value="3">拓展工作領域及視野</asp:ListItem>
                                                    <asp:ListItem Value="4">其他</asp:ListItem>
                                                </asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol_need">結訓後之計畫</td>
                                            <td class="whitecol" onclick="check1();" colspan="3">
                                                <asp:RadioButtonList ID="Q3" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" CssClass="font">
                                                    <asp:ListItem Value="1">轉換工作</asp:ListItem>
                                                    <asp:ListItem Value="2">留任</asp:ListItem>
                                                    <asp:ListItem Value="3">其他</asp:ListItem>
                                                </asp:RadioButtonList>
                                                <asp:TextBox ID="Q3_Other" runat="server" Enabled="False" Width="25%" MaxLength="50"></asp:TextBox>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol_need">服務單位行業別</td>
                                            <td class="whitecol" colspan="3">
                                                <asp:DropDownList ID="Q4" runat="server"></asp:DropDownList></td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol_need">服務單位是否屬於中小企業</td>
                                            <td class="whitecol" colspan="3">
                                                <asp:RadioButtonList ID="Q5" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" CssClass="font">
                                                    <asp:ListItem Value="1">是</asp:ListItem>
                                                    <asp:ListItem Value="0">否</asp:ListItem>
                                                </asp:RadioButtonList>
                                                <br />
                                                <font color="red">（製造業、營造業、礦業及土石採取業常僱用員工數未滿二百人者或農林漁牧業、水電燃氣業、商業、運輸、倉儲及通信業、金融保險不動產、工商服務業、社會服務及個人服務業經常僱用員工數未滿五十人者，屬中小企業。）</font>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol_need">個人工作年資</td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="Q61" runat="server" Columns="5" Width="15%" MaxLength="5"></asp:TextBox><font color="#ff0000" size="2">年(年資開放小數點填寫，以0.5為最小單位，請自行累加計算。例：年資1~11個月，請填：0.5)</font></td>
                                            <td class="bluecol_need">在這家公司的年資</td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="Q62" runat="server" Columns="5" Width="15%" MaxLength="5"></asp:TextBox><font color="#ff0000" size="2">年(年資開放小數點填寫，以0.5為最小單位，請自行累加計算。例：年資1~11個月，請填：0.5)</font></td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol_need">在這職位的年資</td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="Q63" runat="server" Columns="5" Width="15%" MaxLength="5"></asp:TextBox><font color="#ff0000" size="2">年(年資開放小數點填寫，以0.5為最小單位，請自行累加計算。例：年資1~11個月，請填：0.5)</font></td>
                                            <td class="bluecol_need">最近升遷離本職幾年</td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="Q64" runat="server" Columns="5" Width="15%" MaxLength="5"></asp:TextBox><font color="#ff0000" size="2">年(年資開放小數點填寫，以0.5為最小單位，請自行累加計算。例：年資1~11個月，請填：0.5)</font></td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol_left" colspan="4">
                                                <font color="red">本人
                                                    <asp:RadioButtonList ID="IseMail" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" CssClass="font">
                                                        <asp:ListItem Value="Y" Selected="True">希望</asp:ListItem>
                                                        <asp:ListItem Value="N">不希望</asp:ListItem>
                                                    </asp:RadioButtonList>
                                                    &nbsp;定期收到產業人才投資方案最新課程資訊。(填選「希望」電子郵件必要有值，才能正確送出)
                                                </font>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol_left" colspan="4">
                                                <font color="red">本人
                                                    <asp:RadioButtonList ID="IsAgree" runat="server" Visible="False" RepeatDirection="Horizontal" RepeatLayout="Flow" CssClass="font">
                                                        <asp:ListItem Value="Y" Selected="true">同意</asp:ListItem>
                                                        <asp:ListItem Value="N">不同意</asp:ListItem>
                                                    </asp:RadioButtonList>
                                                    &nbsp;勞動部勞動力發展署 暨所屬機關，為本人提供職業訓練及就業服務時使用本人資料。
                                                </font>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol_left" colspan="4">
                                                <span style="color: rgb(0, 96, 128); font-family: sans-serif; font-size: 16px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(237, 247, 246); text-decoration-style: initial; text-decoration-color: initial; float: none;">*&nbsp;勞動力發展署(含所屬分署)及訓練單位，為辦理產業人才投資方案訓練課程及推動職業訓練、就業服務相關政策所需，依個人資料保護法規定，在您提供個人資料前，特告知下列事項：&nbsp;<br />
                                                    <br />
                                                </span>
                                                <ol style="box-sizing: border-box; margin-top: 0px; margin-bottom: 10px; color: rgb(51, 51, 51); font-family: sans-serif; font-size: 15px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: start; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; text-decoration-style: initial; text-decoration-color: initial;">
                                                    <li style="box-sizing: border-box; padding: 0px; margin: 0px; list-style: decimal;">個人資料蒐集目的：&nbsp;<br style="box-sizing: border-box;" />
                                                        (1) 辦理產業人才投資方案相關事宜。&nbsp;<br style="box-sizing: border-box;" />
                                                        (2) 作為政府機關辦理職業訓練或就業服務相關統計、分析。&nbsp;<br style="box-sizing: border-box;" />
                                                        (3) 寄送政府機關職業訓練或就業服務相關訊息。</li>
                                                    <li style="box-sizing: border-box; padding: 0px; margin: 0px; list-style: decimal;">個人資料類別：包含姓名、性別、出生年月日、身分證號、聯絡方式、學歷、服務單位、年資、投保狀況、身分證影本、存摺資料等(詳如學員基本資料表及補助申請書)。&nbsp;</li>
                                                    <li style="box-sizing: border-box; padding: 0px; margin: 0px; list-style: decimal;">個人資料利用之期間、地區、對象及方式：您的個人資料僅供勞動力發展署(含所屬分署)暨相關訓練單位於蒐集目的之必要範圍內，以合理方式利用至蒐集目的消失為止。&nbsp;</li>
                                                    <li style="box-sizing: border-box; padding: 0px; margin: 0px; list-style: decimal;">當事人權利：得依個人資料保護法就自身之個人資料向勞動力發展署所屬分署暨相關訓練單位行使<br style="box-sizing: border-box;" />
                                                        (1)查詢或請求閱覽。<br style="box-sizing: border-box;" />
                                                        (2)請求製給複製本。<br style="box-sizing: border-box;" />
                                                        (3)請求補充或更正。<br style="box-sizing: border-box;" />
                                                        (4)請求停止蒐集、處理或利用及。<br style="box-sizing: border-box;" />
                                                        (5)請求刪除您的個人資料之權利。<br style="box-sizing: border-box;" />
                                                        若您向勞動力發展署所屬分署申請(4)及(5)項，將終止提供您參加產業人才投資方案核定之訓練課程及相關補助訓練費用，若因此導致您的權益產生減損時，勞動力發展署所屬分署暨相關訓練單位不負相關賠償責任。&nbsp;</li>
                                                    <li style="box-sizing: border-box; padding: 0px; margin: 0px; list-style: decimal;">不提供個人資料之權益影響：若您拒絕提供個人資料為特定目的之利用，勞動力發展署所屬分署暨相關訓練單位恐無法提供您蒐集目的之相關服務。</li>
                                                </ol>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol_left" colspan="4">
                                                <font color="red">本人已充分獲知且瞭解上述事項，並&nbsp;
                                                    <asp:RadioButtonList ID="IsAgreedata" runat="server" Visible="False" RepeatDirection="Horizontal" RepeatLayout="Flow" CssClass="font">
                                                        <asp:ListItem Value="Y" Selected="true">同意</asp:ListItem>
                                                        <asp:ListItem Value="N">不同意</asp:ListItem>
                                                    </asp:RadioButtonList>
                                                    &nbsp; 於上開所列蒐集目的之合理範圍內，蒐集、處理及利用本人之個人資料。(如點選不同意將無法接受報名)</font>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol_left" colspan="4">
                                                <font color="red">本人&nbsp;
                                                    <asp:RadioButtonList ID="IsCorrect" runat="server" Visible="False" RepeatDirection="Horizontal" RepeatLayout="Flow" CssClass="font">
                                                        <asp:ListItem Value="Y" Selected="true">確認</asp:ListItem>
                                                        <asp:ListItem Value="N">未確認</asp:ListItem>
                                                    </asp:RadioButtonList>
                                                    &nbsp; 上述為個人最新及正確資料</font>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                    <table id="table4" cellspacing="1" cellpadding="1" width="100%" align="center" border="0">
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="btnSend1" runat="server" Text="送出" CssClass="asp_button_M"></asp:Button>
                                <input id="btnSend1D" runat="server" type="button" value="處理中" class="asp_button_M" style="display: none" disabled="disabled" />
                                <asp:Button ID="btnBack1" runat="server" Text="回報名作業" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <input id="hidstar3" type="hidden" name="hidstar3" runat="server" />
        <input id="hid_eSETID3" size="3" type="hidden" name="hid_eSETID3" runat="server" />
    </form>
</body>
</html>
