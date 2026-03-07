<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_01_001_add.aspx.vb" Inherits="WDAIIP.SD_01_001_add" %>


<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>報名登錄新增/修改</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function open_SD01001sch() {
            var msg = "";
            //var rqID = getUrlParameter("ID");
            var rqID = getParamValue('ID');
            var IDNO = document.getElementById("IDNO");
            var birthday = document.getElementById("birthday");
            var Name = document.getElementById("Name");
            if (IDNO.value == '') { msg += '請輸入身分證號碼!\n'; }
            if (birthday.value == '') { msg += '請輸入出生日期!\n'; }
            if (birthday.value != '' && !checkDate(birthday.value)) {
                msg += '出生日期時間格式不正確!\n';
            }
            if (msg != "") {
                alert(msg);
                return false; //結束。
            }
            var url1 = "SD_01_001_sch.aspx?ID=" + rqID + "&SPAGE=SD01001&IDNO=" + IDNO.value + "&BIRTH=" + birthday.value + "&CNAME=" + escape(Name.value);
            wopen(url1, 'MyWindow', 1040, 620, 1);
            return false;
        }

        //2009-05-19 add 依需求只允許輸入整數(排除 0/00/000)
        function CheckZIPb3_0(zipcodeb3) {
            var argsIsValid = true;
            var ZipCODEB3 = document.getElementById(zipcodeb3);
            if (!ZipCODEB3) { return argsIsValid; }
            if (isBlank(ZipCODEB3)) { return argsIsValid; }
            if (isNaN(parseInt(trim(ZipCODEB3.value), 10))) { argsIsValid = false; return argsIsValid; }
            if (!isUnsignedInt(trim(ZipCODEB3.value))) { argsIsValid = false; return argsIsValid; }
            if (parseInt(trim(ZipCODEB3.value), 10) < 1) { argsIsValid = false; return argsIsValid; }
            return argsIsValid;
        }

        //2009-05-20 add 依需求只允許輸入2碼或3碼
        function CheckZIPb3_23(zipcodeb3) {
            var argsIsValid = true;
            var ZipCODEB3 = document.getElementById(zipcodeb3);
            if (!ZipCODEB3) { return argsIsValid; }
            if (trim(ZipCODEB3.value) == "") { argsIsValid = true; return argsIsValid; }
            if (trim(ZipCODEB3.value).length == 2) { argsIsValid = true; return argsIsValid; }
            if (trim(ZipCODEB3.value).length == 3) { argsIsValid = true; return argsIsValid; }
            argsIsValid = false;
            return argsIsValid;
        }

        function Clear_Zip2() {
            $('#ZipCode2').val("");
            $('#ZipCode2_B3').val("");
            $('#ZipCode2_6W').val("");
            $('#City2').val("");
            $('#HouseholdAddress').val("");
            if ($('#CheckBox1').is(":checked")) {
                $('#ZipCode2').val($('#city_code').val());
                $('#ZipCode2_B3').val($('#ZipCODEB3').val());
                $('#ZipCode2_6W').val($('#ZipCODE6W').val());
                $('#City2').val($('#TBCity').val());
                $('#HouseholdAddress').val($('#Address').val());
            }
        }

        function chgPriorWorkType1_disabled() {
            var HidPreUseLimited17f = document.getElementById('HidPreUseLimited17f');
            if (HidPreUseLimited17f.value != 'Y') { return false; }
            var PriorWorkType1 = document.getElementById('PriorWorkType1');
            var spnStartOrg = document.getElementById('spnStartOrg');
            var spnStartDate = document.getElementById('spnStartDate');
            var spnStartNo = document.getElementById('spnStartNo');
            var PriorWorkOrg1 = document.getElementById('PriorWorkOrg1');
            var ActNo = document.getElementById('ActNo');
            var SOfficeYM1 = document.getElementById('SOfficeYM1');
            var FOfficeYM1 = document.getElementById('FOfficeYM1');
            var imgSOfficeYM1 = document.getElementById('imgSOfficeYM1');
            var imgFOfficeYM1 = document.getElementById('imgFOfficeYM1');
            var BtnCheckBli = document.getElementById('BtnCheckBli');
            var lab2017_2 = document.getElementById('lab2017_2');
            lab2017_2.innerHTML = '投保單位<br />加退保日期';
            var cst_inline = ''; // 'inline';
            var cst_none = 'none';
            if (BtnCheckBli) { BtnCheckBli.style.display = cst_none; }
            spnStartOrg.style.display = cst_inline;
            spnStartDate.style.display = cst_inline;
            spnStartNo.style.display = cst_inline;
            PriorWorkOrg1.disabled = false;
            ActNo.disabled = false;
            SOfficeYM1.disabled = false;
            FOfficeYM1.disabled = false;
            imgSOfficeYM1.disabled = false;
            imgFOfficeYM1.disabled = false;
            PriorWorkOrg1.style.display = cst_inline;
            ActNo.style.display = cst_inline;
            SOfficeYM1.style.display = cst_inline;
            FOfficeYM1.style.display = cst_inline;
            imgSOfficeYM1.style.display = cst_inline;
            imgFOfficeYM1.style.display = cst_inline;

            switch (getRadioValue(document.form1.PriorWorkType1)) {
                case '1':
                    if (BtnCheckBli) { BtnCheckBli.style.display = cst_inline; }
                    PriorWorkOrg1.disabled = true;
                    ActNo.disabled = true;
                    //PriorWorkOrg1.style.display = cst_none;
                    //ActNo.style.display = cst_none;
                    break;
                case '2':
                    //PriorWorkOrg1.disabled = true;
                    //ActNo.disabled = true;
                    //SOfficeYM1.disabled = true;
                    //FOfficeYM1.disabled = true;
                    PriorWorkOrg1.style.display = cst_none;
                    ActNo.style.display = cst_none;
                    SOfficeYM1.style.display = cst_none;
                    FOfficeYM1.style.display = cst_none;
                    imgSOfficeYM1.style.display = cst_none;
                    imgFOfficeYM1.style.display = cst_none;
                    spnStartOrg.style.display = cst_none;
                    spnStartDate.style.display = cst_none;
                    spnStartNo.style.display = cst_none;
                    break;
                case '3':
                    lab2017_2.innerHTML = '工作<br />起迄日期';
                    ActNo.style.disabled = true;
                    ActNo.style.display = cst_none;
                    //spnStartDate.style.display = cst_none;
                    spnStartNo.style.display = cst_none;
                    break;
            }
        }

        function chgPriorWorkType1() {
            var HidPreUseLimited17f = document.getElementById('HidPreUseLimited17f');
            if (HidPreUseLimited17f.value != 'Y') { return false; }
            //var PriorWorkType1 = document.getElementById('PriorWorkType1');
            var PriorWorkOrg1 = document.getElementById('PriorWorkOrg1');
            var ActNo = document.getElementById('ActNo');
            var SOfficeYM1 = document.getElementById('SOfficeYM1');
            var FOfficeYM1 = document.getElementById('FOfficeYM1');
            var hidSB4ID = document.getElementById('hidSB4ID');
            PriorWorkOrg1.value = "";
            ActNo.value = "";
            SOfficeYM1.value = "";
            FOfficeYM1.value = "";
            hidSB4ID.value = "";
            chgPriorWorkType1_disabled();
        }

        //function autoMilitary() {
        //	if (getRadioValue(document.form1.Sex) == 'F') {
        //		document.form1.MilitaryID.value == '03';
        //	}
        //}

        function table_layer(num) {
            for (var i = 1; i <= 2; i++) {
                document.getElementById('Question' + i).style.display = 'none';
                document.getElementById('menu' + i).style.backgroundColor = '#ccccff';
            }
            document.getElementById('Question' + num).style.display = 'inline';
            document.getElementById('menu' + num).style.backgroundColor = '#9999FF';
            /*
			if (document.getElementById('Question'+num).style.display=='inline'){
			   document.getElementById('Question'+num).style.display='none';
			   document.getElementById('menu'+num).style.backgroundColor='#ccccff';
			}
			else{
			   document.getElementById('Question'+num).style.display='inline';
			   document.getElementById('menu'+num).style.backgroundColor='#9999FF';
			}
            */
        }

        function choose_class(TM, OC, num) {
            var StudIDNO = $('#IDNO').val();
            var vTMID1 = $('#TMID1').val();
            if (num != 1) {
                if (vTMID1 == '') {
                    window.alert('請選擇報名班級!');
                    //window.alert('請先選擇第一志願!');
                    return;
                }
            }
            window.open('SD_01_001_ch.aspx?wish=' + num + '&StudIDNO=' + StudIDNO, '', 'width=720,height=640,location=0,status=0,menubar=0,scrollbars=1,resizable=1');
        }

        //'Button7 送出(隱藏) 檢核
        function chkdata() {
            var msg = '';
            var i;
            var j = 0;
            var ItemName = '';
            //var orgname = document.form1.orgname.value;
            //var orgname = document.getElementById('orgname').value;
            var orgname = $('#orgname').val();
            var msg1 = '';
            var msg2 = '';
            var HidPreUseLimited17f = document.getElementById('HidPreUseLimited17f');
            //if (HidPreUseLimited17f.value != 'Y') { return false; }
            //var PriorWorkOrg1 = document.getElementById('PriorWorkOrg1');
            //var ActNo = document.getElementById('ActNo');
            //var SOfficeYM1 = document.getElementById('SOfficeYM1');
            //var FOfficeYM1 = document.getElementById('FOfficeYM1');
            var rblWorkSuppIdent = document.getElementById('rblWorkSuppIdent');
            var WorkSuppIdentVal = ""; //getRBLValue('rblWorkSuppIdent'); //取得 RadioButtonList 值
            if (rblWorkSuppIdent) {
                WorkSuppIdentVal = getRBLValue('rblWorkSuppIdent'); //取得 RadioButtonList 值
            }
            var IdentityID = getCheckBoxListValue('IdentityID');
            var cblAVTCP1 = document.getElementById('cblAVTCP1');
            var MyAVTCP1 = getCheckBoxListValue('cblAVTCP1');
            //var rdo_HighEduBg = document.getElementById("rdo_HighEduBg");
            var DegreeID = document.getElementById("DegreeID");
            var GradID = document.getElementById("GradID");
            var hidstar3 = document.getElementById('hidstar3');
            var Hid_PreUseLimited18a = document.getElementById('Hid_PreUseLimited18a'); //限定2018年職前計畫
            //debugger;				
            if (document.form1.isBlack.value == 'Y') {
                msg1 = orgname + "，已列入處分名單，是否確定繼續？\n"
                msg2 = orgname + "，已列入處分名單!!\n"
                if (!confirm(msg1)) {
                    msg += msg2;
                }
            }
            if (document.form1.Name.value == '') { msg = msg + '請輸入姓名!\n'; if (ItemName == '') ItemName = 'Name'; }
            if (document.form1.birthday.value == '') { msg = msg + '請輸入出生日期!\n'; if (ItemName == '') ItemName = 'birthday'; }
            else {
                if (getDiffDay(document.form1.birthday.value, document.form1.hdatenow.value) <= 365 * 10) {
                    msg = msg + '出生日期輸入過於離譜,請重新確認!\n'; if (ItemName == '') ItemName = 'birthday';
                }
            }
            //debugger;
            if (document.form1.birthday.value != '' && !checkDate(document.form1.birthday.value)) {
                msg = msg + '出生日期時間格式不正確!\n'; if (ItemName == '') ItemName = 'birthday';
            }
            if (!isChecked(document.form1.PassPortNO)) { msg = msg + '請選擇身分別!\n'; if (ItemName == '') ItemName = 'PassPortNO'; }
            if (document.form1.IDNO.value == '') { msg = msg + '請輸入身分證號碼!\n'; if (ItemName == '') ItemName = 'IDNO'; }
            /*
			if (!checkId(document.form1.IDNO.value) && document.form1.PassPortNO[0].checked==true){
			if(!confirm('身分證號碼錯誤，是否要繼續儲存?')) msg=msg+'身分證號碼錯誤\n';
			}
			*/
            //身分證號碼錯誤 98年6月確定不再進入資料庫 by AMU
            if (!checkId(document.form1.IDNO.value) && document.form1.PassPortNO[0].checked == true) {
                msg = msg + '身分證號碼錯誤\n'; if (ItemName == '') ItemName = 'IDNO';
            }
            //if (!checkId(document.form1.IDNO.value)) msg=msg+'身分證號碼有誤\n';
            if (!isChecked(document.form1.Sex)) {
                msg = msg + '請選擇性別!\n'; if (ItemName == '') ItemName = 'Sex';
            }
            else {
                if (document.form1.PassPortNO[0].checked == true) {
                    if (document.form1.IDNO.value != '') {
                        if (!checkId(document.form1.IDNO.value)) {
                            //msg=msg+'身分證號碼有誤\n';
                        }
                        else {
                            if (document.form1.IDNO.value.charAt(1) == 1 && getRadioValue(document.form1.Sex) == 'F') {
                                msg += '性別與身分證號碼不符合\n'; if (ItemName == '') ItemName = 'Sex';
                            }
                            else if (document.form1.IDNO.value.charAt(1) == 2 && getRadioValue(document.form1.Sex) == 'M') {
                                msg += '性別與身分證號碼不符合\n'; if (ItemName == '') ItemName = 'Sex';
                            }
                        }
                    }
                    else {
                        msg += '請輸入身分證號碼!\n'; if (ItemName == '') ItemName = 'IDNO';
                    }
                }
            }
            //if (!isChecked(document.form1.MaritalStatus)) { msg = msg + '請選擇婚姻狀況!\n'; if (ItemName == '') ItemName = 'MaritalStatus'; }
            if (document.form1.DegreeID.selectedIndex == 0) { msg = msg + '請選擇最高學歷\n'; if (ItemName == '') ItemName = 'DegreeID'; }
            if (!isChecked(document.form1.GradID)) { msg = msg + '請選擇畢業狀況!\n'; if (ItemName == '') ItemName = 'GradID'; }
            if (document.form1.School.value == '') { msg += '請填寫學校名稱\n'; if (ItemName == '') ItemName = 'School'; }
            if (document.form1.Department.value == '') { msg += '請填寫科系名稱\n'; if (ItemName == '') ItemName = 'Department'; }
            /*
            //受訓前任職資料
            if (document.form1.TPlanid.value != '06') {
                if (!isChecked(document.form1.PriorWorkType1)) {
                    msg = msg + '請選擇受訓前任職狀況!\n'; if (ItemName == '') ItemName = 'PriorWorkType1';
                }
            }
            if (!isChecked(document.form1.PriorWorkType1)) {
                msg += '請選擇受訓前任職狀況!\n'; if (ItemName == '') ItemName = 'PriorWorkType1';
            }
            else {
                switch (getRadioValue(document.form1.PriorWorkType1)) {
                    case '1':
                        if (PriorWorkOrg1.value == '') { msg = msg + '請輸入 任職單位名稱!\n'; if (ItemName == '') ItemName = 'PriorWorkOrg1'; }
                        if (SOfficeYM1.value == '') { msg = msg + '請輸入 投保單位加退保日期的起日!\n'; if (ItemName == '') ItemName = 'SOfficeYM1'; }
                        if (WorkSuppIdentVal != "Y") {
                            if (FOfficeYM1.value == '' && HidPreUseLimited17f.value == 'Y') { msg = msg + '請輸入 投保單位加退保日期的迄日!\n'; if (ItemName == '') ItemName = 'FOfficeYM1'; }
                        }
                        if (ActNo.value == '') { msg = msg + '請輸入 投保單位保險證號!\n'; if (ItemName == '') ItemName = 'ActNo'; }
                        break;
                    case '2':
                        break;
                    case '3':
                        if (PriorWorkOrg1.value == '') { msg = msg + '請輸入 任職單位名稱!\n'; if (ItemName == '') ItemName = 'PriorWorkOrg1'; }
                        if (SOfficeYM1.value == '') { msg = msg + '請輸入 工作起迄日期的起日!\n'; if (ItemName == '') ItemName = 'SOfficeYM1'; }
                        if (WorkSuppIdentVal != "Y") {
                            if (FOfficeYM1.value == '' && HidPreUseLimited17f.value == 'Y') { msg = msg + '請輸入 工作起迄日期的迄日!\n'; if (ItemName == '') ItemName = 'FOfficeYM1'; }
                        }
                        break;
                    default:
                        //msg += '請選擇受訓前任職狀況!\n'; if (ItemName == '') ItemName = 'PriorWorkType1';
                }
                //if (getRadioValue(document.form1.PriorWorkType1) == '1') {}
                //if (getRadioValue(document.form1.PriorWorkType1) == '3') {}
            }
            */
            /*
			if(document.form1.PriorWorkType1[0].checked){						
			if(document.getElementById('PriorWorkOrg1').value=='') {msg=msg+'請輸入最後一次任職單位名稱!\n';if(ItemName=='') ItemName='PriorWorkOrg1';}
			if(document.getElementById('ActNo').value=='') {msg=msg+'請輸入最後投保單位保險證號!\n';if(ItemName=='') ItemName='ActNo';}
			if(document.getElementById('SOfficeYM1').value=='') {msg=msg+'請輸入最後投保單位起迄日期的起日!\n';if(ItemName=='') ItemName='SOfficeYM1';}
			if(document.getElementById('FOfficeYM1').value=='') {msg=msg+'請輸入最後投保單位起迄日期的迄日!\n';if(ItemName=='') ItemName='FOfficeYM1';}
			}
			*/
            /*
            var datetxt1 = '投保單位加退保日期';
            switch (getRadioValue(document.form1.PriorWorkType1)) {
                case '1':
                    datetxt1 = '投保單位加退保日期';
                    if (SOfficeYM1.value != '' && !checkDate(SOfficeYM1.value)) { msg += '[' + datetxt1 + ']的起日日期格式不正確\n'; if (ItemName == '') ItemName = 'SOfficeYM1'; Page = 1; }
                    if (FOfficeYM1.value != '' && !checkDate(FOfficeYM1.value)) { msg += '[' + datetxt1 + ']的迄日日期格式不正確\n'; if (ItemName == '') ItemName = 'FOfficeYM1'; Page = 1; }
                    if (msg == '' && SOfficeYM1.value != '' && FOfficeYM1.value != '' && FOfficeYM1.value < SOfficeYM1.value) {
                        msg += '[' + datetxt1 + '的迄日]必需大於[' + datetxt1 + '的起日]\n';
                        if (ItemName == '') ItemName = 'SOfficeYM1'; Page = 1;
                    }
                    break;
                case '2':
                    break;
                case '3':
                    datetxt1 = '工作起迄日期';
                    if (SOfficeYM1.value != '' && !checkDate(SOfficeYM1.value)) { msg += '[' + datetxt1 + ']的起日日期格式不正確\n'; if (ItemName == '') ItemName = 'SOfficeYM1'; Page = 1; }
                    if (FOfficeYM1.value != '' && !checkDate(FOfficeYM1.value)) { msg += '[' + datetxt1 + ']的迄日日期格式不正確\n'; if (ItemName == '') ItemName = 'FOfficeYM1'; Page = 1; }
                    if (msg == '' && SOfficeYM1.value != '' && FOfficeYM1.value != '' && FOfficeYM1.value < SOfficeYM1.value) {
                        msg += '[' + datetxt1 + '的迄日]必需大於[' + datetxt1 + '的起日]\n';
                        if (ItemName == '') ItemName = 'SOfficeYM1'; Page = 1;
                    }
                    break;
                default:
                    datetxt1 = '最後投保單位起迄日期';
                    if (SOfficeYM1.value != '' && !checkDate(SOfficeYM1.value)) { msg += '[' + datetxt1 + ']的起日日期格式不正確\n'; if (ItemName == '') ItemName = 'SOfficeYM1'; Page = 1; }
                    if (FOfficeYM1.value != '' && !checkDate(FOfficeYM1.value)) { msg += '[' + datetxt1 + ']的迄日日期格式不正確\n'; if (ItemName == '') ItemName = 'FOfficeYM1'; Page = 1; }
                    if (msg == '' && SOfficeYM1.value != '' && FOfficeYM1.value != '' && FOfficeYM1.value < SOfficeYM1.value) {
                        msg += '[' + datetxt1 + '的迄日]必需大於[' + datetxt1 + '的起日]\n';
                        if (ItemName == '') ItemName = 'SOfficeYM1'; Page = 1;
                    }
                    //msg += '請選擇受訓前任職狀況!\n'; if (ItemName == '') ItemName = 'PriorWorkType1';
            }
            */
            //if (document.form1.MilitaryID.value == 0) { msg = msg + '請選擇兵役\n'; if (ItemName == '') ItemName = 'MilitaryID'; }
            if (document.form1.EnterChannel.selectedIndex == 0) { msg = msg + '請選擇報名管道\n'; if (ItemName == '') ItemName = 'EnterChannel'; }

            if ($('#city_code').val() == '') { msg = msg + '請輸入聯絡地址郵遞區號前3碼[地區]!\n'; if (ItemName == '') ItemName = 'city_code'; }
            //checkzip23 郵遞區號
            msg += checkzip23(true, '聯絡地址', 'ZipCODEB3');
            //if ($('#ZipCODEB3').val() == '') { msg = msg + '請輸入聯絡地址郵遞區號後2碼或3碼[地區]!\n'; if (ItemName == '') ItemName = 'ZipCODEB3'; } //20090520 fix
            //if ($('#ZipCODEB3').val() != '') {
            //    if (!CheckZIPB3_0("ZipCODEB3")) { msg += '聯絡地址郵遞區號後2碼或3碼必須為數字，且不得輸入 0/00/000\n'; if (ItemName == '') ItemName = 'ZipCODEB3'; }
            //    if (!CheckZIPB3_23("ZipCODEB3")) { msg += '聯絡地址郵遞區號後2碼或3碼長度必須為 2碼或3碼\n'; if (ItemName == '') ItemName = 'ZipCODEB3'; }
            //}
            if ($('#Address').val() == '') { msg = msg + '請輸入聯絡地址!\n'; if (ItemName == '') ItemName = 'Address'; }

            if ($('#ZipCode2').val() == '') { msg = msg + '請輸入戶籍地址郵遞區號前3碼[地區]!\n'; if (ItemName == '') ItemName = 'ZipCode2'; }
            //checkzip23 郵遞區號
            msg += checkzip23(true, '戶籍地址', 'ZipCode2_B3');
            //if ($('#ZipCode2_B3').val() == '') { msg = msg + '請輸入戶籍地址郵遞區號後2碼[地區]!\n'; if (ItemName == '') ItemName = 'ZipCode2_B3'; } //20090520 fix
            //if ($('#ZipCode2_B3').val() != '') {
            //    if (!CheckZIPB3_0("ZipCode2_B3")) { msg += '戶籍地址 郵遞區號後2碼或3碼必須為數字，且不得輸入 0/00/000\n'; if (ItemName == '') ItemName = 'ZipCode2_B3'; }
            //    if (!CheckZIPB3_23("ZipCode2_B3")) { msg += '戶籍地址 郵遞區號後2碼或3碼長度必須為 2碼或3碼\n'; if (ItemName == '') ItemName = 'ZipCode2_B3'; }
            //}
            if ($('#HouseholdAddress').val() == '') { msg = msg + '請輸入戶籍地址!\n'; if (ItemName == '') ItemName = 'HouseholdAddress'; }

            //rblMobil
            if (!isChecked(document.form1.rblMobil)) {
                msg += '請選擇有無行動電話\n'; if (ItemName == '') Item = 'CellPhone';
            }
            else {
                if (getRadioValue(document.form1.rblMobil) == "Y") {
                    if (document.form1.CellPhone.value == '') { msg += '有行動電話 請輸入行動電話\n'; if (ItemName == '') ItemName = 'CellPhone'; }
                }
                else {
                    if (document.form1.Phone1.value == '') { msg += '請輸入聯絡電話(日)\n'; if (ItemName == '') ItemName = 'Phone1'; }
                    if (document.form1.CellPhone.value != '') { msg += '有輸入行動電話,請選擇有行動電話\n'; if (ItemName == '') ItemName = 'CellPhone'; }
                }
            }
            //if (document.form1.Phone1.value=='') {msg=msg+'請輸入聯絡電話(日)!\n';if(ItemName=='') ItemName='Phone1';}
            //Email.Text
            if (document.form1.Email.value != '' && document.form1.Email.value != '無' && !checkEmail(document.form1.Email.value)) { msg = msg + 'E-Mail格式不正確!\n'; if (ItemName == '') ItemName = 'Email'; }

            var vMIdentityID = document.form1.MIdentityID;
            if (vMIdentityID.selectedIndex == 0) { msg = msg + '請選擇 主要參訓身分別\n'; if (ItemName == '') ItemName = 'MIdentityID'; }
            if (parseInt(IdentityID, 10) == 0) {
                msg += '請選擇參訓身分別\n'; if (ItemName == '') ItemName = 'Identity';
            }
            else {
                for (var i = 0; i < IdentityID.length; i++) {
                    if (IdentityID.charAt(i) == '1') j++;
                }
                if (j > 5) msg += '參訓身分別最多只能選擇 五項\n'; if (ItemName == '') ItemName = 'Identity';
            }

            if (cblAVTCP1) {
                if (parseInt(MyAVTCP1, 10) == 0) {
                    msg += '請選擇 獲得職訓課程管道\n'; if (ItemName == '') ItemName = 'cblAVTCP1';
                }
            }
            /*
            if (rdo_HighEduBg) {
                if (rdo_HighEduBg.rows[0].cells[0].children[0].checked) {
                    if (DegreeID.value != "05" && DegreeID.value != "06") {
                        if (DegreeID.value == "03" || DegreeID.value == "04") {
                            if (!GradID.rows[0].cells[0].children[0].checked) {
                                msg += "專上畢業學歷失業者至少需要專科或大學以上學歷畢業。\n"; if (ItemName == '') ItemName = 'rdo_HighEduBg';
                            }
                        } else {
                            msg += "專上畢業學歷失業者至少需要專科或大學以上學歷畢業。\n"; if (ItemName == '') ItemName = 'rdo_HighEduBg';
                        }
                    }
                }
            }
            */
            //是否為在職者補助身分
            if (rblWorkSuppIdent) {
                if (isEmpty('rblWorkSuppIdent')) {
                    msg += "請選擇「是否為在職者補助身分」\n"; if (ItemName == '') ItemName = 'rblWorkSuppIdent';
                }
            }
            //服務單位、服務部門、投保單位名稱、投保類別、職稱/職務
            if ($('#Uname').val() == "") { msg = msg + '請輸入服務單位!\n'; if (ItemName == '') ItemName = 'Uname'; }
            if ($("#ddlSERVDEPTID option:selected").val() == "" || $("#ddlSERVDEPTID option:selected").text() == "") { msg = msg + '請選擇服務部門!\n'; if (ItemName == '') ItemName = 'ddlSERVDEPTID'; }
            if ($('#ActName').val() == "") { msg = msg + '請輸入投保單位名稱!\n'; if (ItemName == '') ItemName = 'ActName'; }
            if ($("#ActType option:selected").val() == "" || $("#ActType option:selected").text() == "") { msg = msg + '請選擇投保類別!\n'; if (ItemName == '') ItemName = 'ActType'; }
            if ($("#ddlJOBTITLEID option:selected").val() == "" || $("#ddlJOBTITLEID option:selected").text() == "") { msg = msg + '請選擇職稱!\n'; if (ItemName == '') ItemName = 'ddlJOBTITLEID'; }

            if (document.form1.OCID1.value == '') { msg = msg + '請選擇輸入報名班級!\n'; if (ItemName == '') ItemName = 'OCID1'; }
            if (document.form1.CellPhone.value != '' && !isCellPhone(document.form1.CellPhone.value)) { msg = msg + '行動電話號碼 長度必須是10，必定要09帶頭\n'; if (ItemName == '') ItemName = 'CellPhone'; }
            //if (document.form1.OCIDValue1.value == document.form1.OCIDValue2.value && document.form1.OCIDValue2.value != '') { msg = msg + '第一志願不能和第二志願相同\n'; if (ItemName == '') ItemName = 'OCID2'; }
            //if (document.form1.OCIDValue2.value == document.form1.OCIDValue3.value && document.form1.OCIDValue3.value != '') { msg = msg + '第二志願不能和第三志願相同\n'; if (ItemName == '') ItemName = 'OCID3'; }
            //if (document.form1.OCIDValue1.value == document.form1.OCIDValue3.value && document.form1.OCIDValue3.value != '') { msg = msg + '第一志願不能和第三志願相同\n'; if (ItemName == '') ItemName = 'OCID3'; }
            //if (!isChecked(document.form1.IsAgree)) { msg += '請選擇是否同意將個人資料提供 勞動部勞動力發展署 暨所屬機關運用\n'; if (ItemName == '') ItemName = 'IsAgree'; }
            //debugger;
            if (Hid_PreUseLimited18a.value == "") {
                if (hidstar3.value != '') {
                    if (!confirm('本次登錄之學員,仍在訓中,是否儲存,請確認!')) { msg += '學員,仍在訓中\n'; if (ItemName == '') ItemName = 'Button1'; }
                }
            }
            if (msg != '') {
                alert(msg);
                if (ItemName != '') document.getElementById(ItemName).focus();
                return false; //結束。
            }
            else {
                var rst2 = true; //正常再次檢核。
                if (rst2) rst2 = Chk_CHGIDNO(); //正常再次檢核。
                if (rst2) rst2 = Chk_IJClist(); //正常再次檢核。
                if (!rst2) {
                    return rst2; //return false; //結束。
                }
                //if (document.form1.IDNOChange.value == '1') {
                //    if (!confirm('身分證有變更過，確定要儲存?'))
                //        //return false; //結束。
                //}
                //PriorWorkOrg1.disabled = false;
                //ActNo.disabled = false;
                //SOfficeYM1.disabled = false;
                //FOfficeYM1.disabled = false;
            }
            //document.getElementById('Button7').disabled=true;
            document.getElementById('Button1').click();
        }

        function Chk_CHGIDNO() {
            var rst = true;
            var IDNOChange = document.form1.IDNOChange;
            //身分證有變更過，確定要儲存?
            if (IDNOChange.value == '1') {
                if (!confirm('身分證有變更過，確定要儲存?')) {
                    rst = false;
                }
            }
            return rst
        }

        function Chk_IJClist() {
            //因與委外實施基準條款有抵觸，請確認是否要同意此民眾的報名
            //警告訊息，但確認後可繼續儲存。
            var msg = document.getElementById('HidIJCMsg').value;
            var sName1 = document.getElementById('Name').value;
            var rst = true;
            if (msg != "") {
                msg = msg.replace(/XXX/, sName1);
                if (!confirm(msg)) {
                    rst = false;
                }
            }
            return rst;
        }

        function clear_wish(num) {
            switch (num) {
                case 1:
                    document.form1.TMID1.value = '';
                    document.form1.TMIDValue1.value = '';
                    document.form1.OCID1.value = '';
                    document.form1.OCIDValue1.value = '';
                    document.form1.ComIDNO1.value = '';
                    document.form1.SeqNO1.value = '';
                    document.form1.CCLID.value = '';
                    break;
            }
        }

        function chkadp() {
            var check = true;
            var classid = document.form1.check_class.value;
            var classary = classid.substr(1, classid.length).split(";")
            var OCIDValue1 = document.form1.OCIDValue1;
            //var OCIDValue2 = document.form1.OCIDValue2;
            //var OCIDValue3 = document.form1.OCIDValue3;
            var ocid1 = false;
            //var ocid2 = false;
            //var ocid3 = false;
            var msg = '';
            for (var i = 0; i < classary.length; i++) {
                if (OCIDValue1.value == classary[i]) {
                    document.form1.select_id.value = classary[i];
                    ocid1 = true;
                }
                //if (OCIDValue2.value == classary[i]) { ocid2 = true; }
                //if (OCIDValue3.value == classary[i]) { ocid3 = true; }
            }
            //if (ocid2) {
            //    msg += '志願二 ';
            //} else if (ocid3) {
            //    msg += '志願三 ';
            //}
            //debugger;//新增。
            if (!ocid1) { document.form1.ptype.value = "add"; }
            if ((!ocid1) && (msg != '')) {
                if (!confirm(msg + '為三合一資料，是否確定繼續?')) {
                    check = false;
                }
            }
            return check;
        }

        //檢查身分別為本國時，由身分證號第2位帶入性別
        function chkidnosex() {
            if (document.form1.PassPortNO[0].checked == true) {
                if (document.form1.IDNO.value.charAt(1) == 1) { document.form1.Sex_0.checked = true; }
                else { if (document.form1.IDNO.value.charAt(1) == 2) { document.form1.Sex_1.checked = true; } }
            }
        }

        function EnterChannelChange() {
            //職前課程邏輯
            //if (document.form1.EnterChannel.value != '4') {
            //	if (document.form1.hide_TrainMode.value != '') {
            //		if (document.form1.hide_EnterChannel.value == '4') {
            //			document.form1.EnterChannel.value = '4';
            //			alert("該學員無法從推介更改為其他報名管道。");
            //		}
            //	}
            //}
            if (document.form1.hide_EnterChannel.value == '1') {
                if (document.form1.EnterChannel.value == '2' || document.form1.EnterChannel.value == '3') {
                    document.form1.EnterChannel.value = document.form1.hide_EnterChannel.value;
                    alert("報名管道不能從網路報名更改為現場或通訊報名。");
                }
            }
        }

        //onclick()
        //function Button7_onclick() {		}
        //function Button6_onclick() {		}
        //'查詢參訓歷史 'open_StudentList
        function open_StudentList(rqID) {
            var IDNO = document.getElementById('IDNO');
            var Hid_ENCIDNO = document.getElementById('Hid_ENCIDNO');
            if (Hid_ENCIDNO && Hid_ENCIDNO.value != "") {
                window.open('../05/SD_05_010_pop.aspx?ID=' + rqID + '&SD_01_004_Type=Student&ENCIDNO=' + Hid_ENCIDNO.value, 'history', 'width=1400,height=820,scrollbars=1');
                return false;
            }
            if (IDNO) {
                window.open('../05/SD_05_010_pop.aspx?ID=' + rqID + '&SD_01_004_Type=Student&IDNO=' + IDNO.value, 'history', 'width=1400,height=820,scrollbars=1');
                return false;
            }
        }
        //function Button7_onclick() {} ->Button1

        //檢查目前選取到的『參訓身分別』,by:20180724、20180725
        function ChkSelId() {
            var str = "";
            $("input[type=checkbox]").each(function (i, el) {
                str = $(this).parent().attr("ChkValue");
                var index = str.indexOf("志願役");
                if (index > -1) {
                    if ($(this).prop('checked')) {
                        if ($(el).attr('data-hit') == "T") {
                            return false;
                        }
                        else {
                            $(el).attr('data-hit', "T");
                            alert("提醒您：志願役現役軍人應於報名截止日前將「送訓名冊」及「送訓證明正本」函送分署辦理資格審查，未符合規定者，不得參加甄試。");
                            return false;
                        }
                    }
                    else {
                        $(el).attr('data-hit', "F");
                        return false;
                    }
                }
            });
        }
    </script>
    <style type="text/css">
        .displaynone { display: none; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <input id="hdatenow" type="hidden" name="hdatenow" runat="server" />
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;報名登錄</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_nw" id="Table3" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="table_title" colspan="4">報名學員資料</td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">准考證號碼 </td>
                            <td class="whitecol" colspan="3" width="80%">
                                <asp:Label ID="ExamID" runat="server" EnableViewState="False"></asp:Label>
                                <input id="ptype" type="hidden" name="ptype" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="20%">姓名 </td>
                            <td class="whitecol" width="30%">
                                <asp:TextBox ID="Name" runat="server" Width="70%"></asp:TextBox></td>
                            <td class="bluecol_need" width="20%">出生日期 </td>
                            <td class="whitecol" width="30%">
                                <asp:TextBox ID="birthday" runat="server" MaxLength="10" Width="60%"></asp:TextBox>
                                <span id="span1" runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= birthday.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span>
                                <input id="hidBirthBtn" type="hidden" size="8" name="hidBirthBtn" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">身分別 </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="PassPortNO" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="1">本國</asp:ListItem>
                                    <asp:ListItem Value="2">外籍(含大陸人士)</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                            <td class="bluecol_need">身分證號碼<br />
                                <font color="#000000">(或護照號碼<br />
                                    或工作證號)</font> </td>
                            <td class="whitecol">
                                <asp:TextBox ID="IDNO" runat="server" Columns="12" Width="60%"></asp:TextBox>
                                <input id="IDNOChange" type="hidden" value="0" runat="server" />
                                <asp:Button ID="Button9" runat="server" Text="查詢歷史紀錄" CssClass="asp_button_M" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">性別 </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="Sex" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="M">男</asp:ListItem>
                                    <asp:ListItem Value="F">女</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                            <td class="bluecol">婚姻 </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="MaritalStatus" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="1">已婚</asp:ListItem>
                                    <asp:ListItem Value="2">未婚</asp:ListItem>
                                    <asp:ListItem Value="3" Selected="True">暫不提供</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">最高學歷 </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="DegreeID" runat="server"></asp:DropDownList></td>
                            <td class="bluecol_need">畢業狀況 </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="GradID" runat="server" CssClass="font" RepeatDirection="Horizontal"></asp:RadioButtonList></td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">學校名稱 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="School" runat="server" Width="70%" MaxLength="100"></asp:TextBox></td>
                            <td class="bluecol_need">科系 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="Department" runat="server" Width="70%" MaxLength="100"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">兵役 </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="MilitaryID" runat="server"></asp:DropDownList></td>
                            <td class="bluecol_need">報名管道 </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="EnterChannel" runat="server">
                                    <asp:ListItem Value="0">請選擇</asp:ListItem>
                                    <asp:ListItem Value="1">網路</asp:ListItem>
                                    <asp:ListItem Value="2">現場</asp:ListItem>
                                    <asp:ListItem Value="3">通訊</asp:ListItem>
                                    <%-- <asp:ListItem Value="4">推介</asp:ListItem> --%>
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">通訊地址 </td>
                            <td class="whitecol" colspan="3">
                                <input id="bt1_zipcode" runat="server" type="button" value="..." class="asp_button_Mini" />
                                <input id="city_code" onfocus="this.blur()" name="city_code" runat="server" style="width: 10%" />－
                                <input id="ZipCODEB3" maxlength="3" name="ZipCODEB3" runat="server" style="width: 10%" />
                                <input id="hidZipCODE6W" runat="server" type="hidden" />
                                <asp:Literal ID="LitZip1" runat="server"></asp:Literal><br />
                                <%--郵遞區號--%>
                                <asp:TextBox ID="TBCity" runat="server" Width="25%"></asp:TextBox>
                                <asp:TextBox ID="Address" runat="server" Width="60%" MaxLength="200"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">戶籍地址</td>
                            <td class="whitecol" colspan="3">
                                <%--<asp:CheckBox ID="CheckBox1" runat="server" CssClass="font" Text="同通訊地址"></asp:CheckBox>--%>
                                <input type="checkbox" id="CheckBox1" name="CheckBox1" runat="server" value="" title="同通訊地址" />同通訊地址<br />
                                <input id="Button6" type="button" value="..." name="city_zip" runat="server" class="asp_button_Mini" />
                                <input id="ZipCode2" style="width: 12%;" maxlength="5" name="ZipCode2" runat="server" />－
                                <input id="ZipCode2_B3" style="width: 10%;" maxlength="3" name="ZipCode2_B3" runat="server" />
                                <input id="HidZipCode2_6W" runat="server" type="hidden" />
                                <asp:Literal ID="LitZip2" runat="server"></asp:Literal><br />
                                <%--郵遞區號--%>
                                <asp:TextBox ID="City2" runat="server" Width="25%" onfocus="this.blur()"></asp:TextBox>
                                <asp:TextBox ID="HouseholdAddress" runat="server" Width="60%" MaxLength="200"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">聯絡電話(日)</td>
                            <td class="whitecol">
                                <asp:TextBox ID="Phone1" runat="server" Width="70%"></asp:TextBox></td>
                            <td class="bluecol">聯絡電話(夜)</td>
                            <td class="whitecol">
                                <asp:TextBox ID="Phone2" runat="server" Width="70%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">電子信箱 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="Email" runat="server" Width="88%"></asp:TextBox></td>
                            <td class="bluecol_need">行動電話 </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="rblMobil" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="N">無</asp:ListItem>
                                    <asp:ListItem Value="Y">有</asp:ListItem>
                                </asp:RadioButtonList>
                                <asp:TextBox ID="CellPhone" runat="server" Width="70%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">主要參訓身分別<%--<asp:Label ID="StarMIdentityID" runat="server"></asp:Label>--%></td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="MIdentityID" runat="server"></asp:DropDownList>
                                <%--<asp:Label ID="labIdentity" runat="server" ForeColor="Red"></asp:Label>--%>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">參訓身分別<br />
                                <font color="#000000">(可複選，最多五項)</font> </td>
                            <td class="whitecol" colspan="3">
                                <asp:CheckBoxList ID="IdentityID" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3" onclick="ChkSelId();"></asp:CheckBoxList>
                            </td>
                        </tr>
                        <%-- (職前課程邏輯)
                        <tr id="HGTR" runat="server">
                            <td class="bluecol">專上畢業<br />學歷失業者<br />(特別預算) </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="rdo_HighEduBg" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="Y">是</asp:ListItem>
                                    <asp:ListItem Value="N" Selected="True">否</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="DGTR" runat="server">
                            <td class="bluecol">學習券身分別 </td>
                            <td class="whitecol" colspan="3"><asp:Label ID="OBJECT_TYPE" runat="server"></asp:Label></td>
                        </tr>
                        <tr id="GovTR" runat="server">
                            <td class="bluecol">推介單個案區分 </td>
                            <td class="whitecol"><asp:Label ID="OBJECT_TYPE2" runat="server"></asp:Label></td>
                            <td class="bluecol">身分別 </td>
                            <td class="whitecol"><asp:Label ID="SPECIAL_TYPE" runat="server"></asp:Label></td>
                        </tr>
                        --%>
                        <tr id="WSITR" runat="server">
                            <td class="bluecol_need">是否為在職者補助身分 </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="rblWorkSuppIdent" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="Y">是</asp:ListItem>
                                    <asp:ListItem Value="N">否</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <%--<tr>
                            <td colspan="4">
                                <table class="table_nw" id="BackTable" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                </table>
                            </td>
                        </tr>--%>
                        <tr>
                            <td class="table_title" colspan="4">服務單位資料</td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">服務單位</td>
                            <td class="whitecol">
                                <asp:TextBox ID="Uname" runat="server" MaxLength="50" Width="70%"></asp:TextBox></td>
                            <td class="bluecol">統一編號</td>
                            <td class="whitecol">
                                <asp:TextBox ID="Intaxno" runat="server" Width="60%" MaxLength="10"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">服務部門</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="ServDept" runat="server" MaxLength="50" Width="30%" CssClass="displaynone"></asp:TextBox>
                                <asp:DropDownList ID="ddlSERVDEPTID" runat="server"></asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">投保單位名稱</td>
                            <td class="whitecol">
                                <asp:TextBox ID="ActName" runat="server" Width="70%" MaxLength="50"></asp:TextBox></td>
                            <td class="bluecol_need">投保類別</td>
                            <td class="whitecol">
                                <%--1:勞保/2:農保--%>
                                <asp:DropDownList ID="ActType" runat="server">
                                    <asp:ListItem Value="1" Selected="true">勞保</asp:ListItem>
                                    <asp:ListItem Value="2">農保</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">投保單位保險證號</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="ActNo" runat="server" Width="30%" MaxLength="20"></asp:TextBox>
                                <font color="blue">請確實填寫正確的勞工保險卡資料</font>
                                <%-- 請依照您目前工作之勞工保險卡確實填寫，並於報名繳費時繳交勞保卡影本--%>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">投保單位電話</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="ActTel" runat="server" MaxLength="30" Width="40%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">投保單位地址</td>
                            <td class="whitecol" colspan="3">
                                <input id="Button8" type="button" value="..." name="city_zip" runat="server" class="asp_button_Mini" />
                                <input id="ZipCode3" style="width: 12%;" maxlength="3" name="ZipCode3" runat="server" />－
                                <input id="ZipCode3_B3" style="width: 10%;" maxlength="3" name="ZipCode3_B3" runat="server" />
                                <input id="HidZipCode3_6W" runat="server" type="hidden" />
                                <asp:Literal ID="LitZip3" runat="server"></asp:Literal><br />
                                <%--郵遞區號--%>
                                <asp:TextBox ID="City3" runat="server" Width="25%" onfocus="this.blur()" MaxLength="100"></asp:TextBox>
                                <asp:TextBox ID="ActAddress" runat="server" Width="60%" MaxLength="110"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">職稱</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="JobTitle" runat="server" MaxLength="50" Width="30%" CssClass="displaynone"></asp:TextBox>
                                <asp:DropDownList ID="ddlJOBTITLEID" runat="server"></asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">報名日期 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="RelEnterDate" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <span id="span2" runat="server">
                                    <img id="IMG1" style="cursor: pointer" onclick="javascript:show_calendar('RelEnterDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30" /></span>
                                <input id="EnterDate" type="hidden" runat="server" />
                                <input id="check_idno" type="hidden" name="check_idno" runat="server" />
                                <input id="check_class" type="hidden" name="check_class" runat="server" />
                                <input id="select_id" type="hidden" name="select_id" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="table_title" colspan="4">報名班級資料</td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">報名班級 </td>
                            <td class="whitecol" colspan="3">
                                <table class="font" border="0" width="100%">
                                    <tr>
                                        <%--志願一：--%>
                                        <td width="12%">職類： </td>
                                        <td width="30%">
                                            <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="90%"></asp:TextBox>
                                            <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                        </td>
                                        <td width="8%">班別： </td>
                                        <td width="40%">
                                            <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="66%"></asp:TextBox>
                                            <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                            <input id="ExamNo" type="hidden" name="ExamNo" runat="server" />
                                            <input id="ComIDNO1" type="hidden" name="ComIDNO1" runat="server" />
                                            <input id="SeqNO1" type="hidden" name="SeqNO1" runat="server" />
                                            <input id="CCLID" type="hidden" runat="server" />
                                            <input id="Button5" onclick="choose_class('TMID1', 'OCID1', 1)" type="button" value="..." name="Button5" runat="server" class="asp_button_Mini" />
                                            <input id="BtnClear1" runat="server" onclick="clear_wish(1)" type="button" value="清除" class="asp_button_S" />
                                        </td>
                                    </tr>
                                    <%--4B4113DE11D52189A1254CC53E8D526B0E4CC09D6A63579572DCEAAE12EB5C4ABE5A7FFCEFC60F3CF033231EBDD90D1331FFE0EF7440C9301142033ACF71159A435C7517941BD663E030A44EF6FC31925CC07A6DF3AF2698EA3E0F6CA6DEB26B1DFA072B2861054B88FBA9D72FD6AD162DDD789425EA3D5B9513A9A3140D0D58B9B1B9026279126848CC570E69CEC38024FB5E5525BC603F34FFFA4FA5F90D9CD304255B6650C16F52540151CFEFC864C8478EE08E9F49C0CA4D715B8C9966A0CBF9D733E108BF5B8CE1E150334BCFECE20690FB75150DA5BF4F3909977D6254D9959F555997FCFB3198345C586DBFA9B9A53EE9E14E69E7A3320B5CAA60BDB15C89BD431654B43582D96E346A7957F4CC0BFC06BCBECFB3CCB6749FCC9A50D6CD2283FAE55B082739BE8AB1A233D3211FD3E63485FE1D8924C06A94A7570297810DFC9AF12622E6CF9E7D8798C747D81603E5444DA4B519382212783EC47B5DFC61F492E003A341CA85D5DD4716A1B663A7183A979D8513A725FDF00E8C7D2CA3A6609018041C14045D4E045BCC4EA3FE4F9484CB7A2A29B9292E2C1107A945BE2D2EFF6D457E8345BBB6EC99E115997BF35DB8764B58462D7448E1B7127CB766C66BB207264A51DBF7AF3414BA9732B2C33F90A520FD3B8E293A5E94F7FCAF9D2D9C76377E75EE6E02C03F207604FF99578894E728400B9B0AB4FF7E8E599B481775B588B38282AA946AC8DAAD451AD7CFC337C0452C3211F03A1021C268C5630895718A5AF0DEEC7A014A07A5ADCB8B888349857116127B9E6ADFD5A4001622E09985DA2A1DA352F22B5DE83AD87E49157A48551C108E33121EDEFBBF3FA1E39C49B5B5ABB01973046CCE7DECEC36FE164E157C12783AFA2DD5AB10AE6860658CE71432272F4ADFAE956B356C00BFBA5D2DE8A8CE36C034667C9A9056DA3772CE765C2676FCBA119B3343CB8423CE895CBF73963837AE7DE26278CF8E5D90C5542CAE514DDB99E0645435BF6A81FED3DF54ED5E603A8F0E79645BB23B9D8618DD699228660833DC1A20A3B5531DC9BC985473543F2771F450266592178480A1120AA62FF36443A456FCCA6DF2F130A7E86A99A9E9726C3C195EAA3D0647D93A4DD04A304FDC50A7321894FCCE8A9879B81A0E0A01DCA217812A56F933032568E9F8242D905031B16F6347A20295BD042985495FF513B293D6B79DAAC1EC023E3ABA666028CD1749A20F06485ABFD55D8F35CAD2A6145BC8CB6546AB220A93351E12D23CFB6F04A3DA2819A8A27D2364D76AF820D496B72A56D36BB03E80A92EA693EBB23436B35DA049E03EFC5AED86DAD30173F723A9FFB30923F845A12436EF0CA65CAC90109F649633FA40AF212BD004A8188072B41DD7FF26F8F6D7209279A9D04539999C10400520FD535CC81010466B4A562525116A28B1DF6FE7C04989C166708B7983AF8393FEAED883CEC52F1E887A327EE4DAB34A405E23448766D8EDE3D96D57DBFB5A903EF4308AA9372994DD0CDB63836A7474945AD430AD7DCDA1EB71AFA508155F15F5DE8AB88F3ACDC3308982E27F12A2C10B5E350E53D727892CD61CDC4EEA1B013F405C4B843CE3ED9D20B7E0D8C43FB5A6580EA3CC16597FFA7F037380D96B65EA4B92A2347206F753BE0A32004065EC7F8064E74AA7160D56F44AA177C7EECB9D11E6E25CEEADF08B2C265C7C59C11F29C307110925B85D8B2E6FA969C01E80E7884499E5B9F235193F582C43C5CBB014115671B1E64DEBCBCD6B52F7E8775917FD1045A7D340DEAA7D8047F9ABDC864DE51C591A53080DF31060864BDA31BB97A10E8A23CE6C247F575363DCCF16B88AF5CECBADA99908AC2F8E3C792C1BA7ECD7D1A86F7DFF837CEB3214273D9F2ED70332F984131E2F01C33FB9AC64C3480DD0877C9E97A9B3603F34E09011F09520429336E79DF43B810E04C5CB4EBA51C30C7502375844D665085A16BB3F610E6A8C5B8D41FEFA25065C5F015513164845AD4937FB1F1EFC7749B391366F2FD30D1E2DAC1A6AD9F20DA3402EA0E7D740EC046DE60D8179F39161E6E1765BC7F378F5CA6940FC9D95A4FD4B06109A1B4DDDABBC27D113B2305EF16FF2D8BC7D9B3EAB3F434770D2A9F6D04FFF19EDB0C8B7914F4F90C95FACBD3880A5C0A33A3B8F79A588ABC1AEC52DE0293693A50051CD594DE90A3EDCF0E7EBA3E0E1FBCF5E895E59AEEC57311EFB2C1F3725A239B2AEB22E7B69F54431EB48DDF1B02325ECF69FCBD95360F041DDD361B2A747B960ACC3753C062D7CE6A7D80DFE21935386F8F7786AEDAC7FECCF655BD23181E537C6FAEE7683--%>
                                </table>
                            </td>
                        </tr>
                        <tr id="trAVTCP" runat="server">
                            <td class="bluecol_need">獲得職訓課程管道 </td>
                            <td class="whitecol" colspan="3">
                                <asp:CheckBoxList ID="cblAVTCP1" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="5"></asp:CheckBoxList></td>
                        </tr>
                        <%-- 職前課程邏輯
                        <tr>
                            <td class="bluecol_need">受訓前<br />
                                任職狀況 </td>
                            <td class="whitecol" colspan="2">
                                <asp:RadioButtonList ID="PriorWorkType1" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="1">曾工作過</asp:ListItem>
                                    <asp:ListItem Value="2">未曾工作過</asp:ListItem>
                                    <asp:ListItem Value="3">先前從事為非勞保性質工作</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                            <td class="whitecol">
                                <asp:Button ID="BtnCheckBli" runat="server" Text="勞保明細檢查" ToolTip="勞保明細檢查" CssClass="asp_button_M"></asp:Button></td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                <asp:Label ID="lab2017_1" runat="server" Text="任職<br />單位名稱"></asp:Label>
                                <span id="spnStartOrg"><font color="#ff0000">*</font></span> </td>
                            <td class="whitecol">
                                <asp:TextBox ID="PriorWorkOrg1" runat="server" MaxLength="30" Width="260px"></asp:TextBox></td>
                            <td class="bluecol">
                                <asp:Label ID="lab2017_2" runat="server" Text="投保單位<br />加退保日期"></asp:Label>
                                <span id="spnStartDate"><font color="#ff0000">*</font></span> </td>
                            <td class="whitecol">
                                <asp:TextBox ID="SOfficeYM1" runat="server" Width="80px" MaxLength="10"></asp:TextBox>
                                <img id="imgSOfficeYM1" style="cursor: pointer" onclick="javascript:show_calendar('<%= SOfficeYM1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />~
							<asp:TextBox ID="FOfficeYM1" runat="server" Width="80px" MaxLength="10"></asp:TextBox>
                                <img id="imgFOfficeYM1" style="cursor: pointer" onclick="javascript:show_calendar('<%= FOfficeYM1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                <asp:Label ID="lab2017_3" runat="server" Text="投保單位<br />保險證號"></asp:Label>
                                <span id="spnStartNo"><font color="#ff0000">*</font></span> </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="ActNo" runat="server" MaxLength="9"></asp:TextBox>
                                <asp:Button ID="Button8" runat="server" Text="檢查" ToolTip="是否為協助基金補助對象" CssClass="asp_button_S"></asp:Button>
                            </td>
                        </tr>
                        --%>
                        <tr>
                            <td class="bluecol">備註 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="notes" runat="server" TextMode="MultiLine" Columns="50" Rows="8" Width="66%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="whitecol" colspan="4">
                                <center><font color="red" style="font-weight: bold;">本人同意勞動部勞動力發展署暨所屬機關，為本人提供職業訓練及就業服務時使用</font></center>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4" class="whitecol" align="center">
                                <%--Button1_Click Button1--%>
                                <input id="Button7" type="button" value="送出" name="Button7" runat="server" class="asp_button_M" />
                                <input id="Button4" type="button" value="回報名登錄" name="Button4" runat="server" class="asp_button_M" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <asp:Button ID="Button1" runat="server" Text="送出(隱藏)"></asp:Button>

        <input id="hidSB4ID" type="hidden" runat="server" />
        <asp:Literal ID="JAVASCRIPT_LITERAL" runat="server"></asp:Literal>
        <input id="STDateHidden" type="hidden" name="STDateHidden" runat="server" />
        <input id="hide_EnterChannel" type="hidden" name="hide_EnterChannel" runat="server" />
        <input id="hide_TrainMode" type="hidden" name="hide_TrainMode" runat="server" />
        <input id="R_serial" type="hidden" name="R_serial" runat="server" />
        <input id="R_EnterDate" type="hidden" name="R_EnterDate" runat="server" />
        <input id="R_SerNum" type="hidden" name="R_SerNum" runat="server" />
        <input id="hidstar3" type="hidden" name="hidstar3" runat="server" />
        <input id="isBlack" type="hidden" name="isBlack" runat="server" />
        <input id="orgname" type="hidden" name="orgname" runat="server" />
        <input id="TPlanid" type="hidden" runat="server" name="TPlanid" />
        <input id="HidIJCMsg" type="hidden" runat="server" name="HidIJCMsg" />

        <asp:HiddenField ID="Hid_ticketNo" runat="server" />
        <asp:HiddenField ID="Hid_ticketType" runat="server" />
        <asp:HiddenField ID="Hid_rqTicket" runat="server" />
        <asp:HiddenField ID="Hid_ENCIDNO" runat="server" />
        <asp:HiddenField ID="Hid_MSG1" runat="server" />
        <asp:HiddenField ID="Hid_MSG2" runat="server" />
        <asp:HiddenField ID="Hid_MSGADIDN" runat="server" />
        <asp:HiddenField ID="HidPreUseLimited17f" runat="server" />
        <asp:HiddenField ID="Hid_ACTNObli" runat="server" />
        <asp:HiddenField ID="Hid_PreUseLimited18a" runat="server" />
    </form>
</body>
</html>
