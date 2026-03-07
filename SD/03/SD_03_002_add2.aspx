<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_03_002_add2.aspx.vb" Inherits="WDAIIP.SD_03_002_add2" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD html 4.0 Transitional//EN">
<html>
<head>
    <title>學員資料維護</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1" />
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1" />
    <meta name="vs_defaultClientScript" content="JavaScript" />
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5" />
    <link rel="stylesheet" type="text/css" href="../../css/style.css" />
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>

    <script type="text/javascript">
        //勞保勾稽資料檢視查詢視窗 //職前專用
        function open_SD01001sch() {
            var msg = "";
            //var rqID = getUrlParameter("ID");
            var rqID = getParamValue('ID');
            var IDNO = document.getElementById("IDNO");
            var birthday = document.getElementById("Birthday");
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
            var url1 = "";
            url1 = "../01/SD_01_001_sch.aspx?ID=" + rqID + "&SPAGE=SD03002&IDNO=" + IDNO.value + "&BIRTH=" + birthday.value + "&CNAME=" + escape(Name.value);
            wopen(url1, 'MyWindow', 1040, 620, 1);
            return false;
        }

        //勞保勾稽資料檢視查詢視窗  //產投專用
        function open_SD01001sch2() {
            var msg = "";
            //var rqID = getUrlParameter("ID");
            var rqID = getParamValue('ID');
            var IDNO = document.getElementById("IDNO");
            var birthday = document.getElementById("Birthday");
            var Name = document.getElementById("Name");
            var Hid_OCID = document.getElementById("Hid_OCID");
            var STDateHidden = document.getElementById("STDateHidden");

            if (IDNO.value == '') { msg += '請輸入身分證號碼!\n'; }
            if (birthday.value == '') { msg += '請輸入出生日期!\n'; }
            if (birthday.value != '' && !checkDate(birthday.value)) {
                msg += '出生日期時間格式不正確!\n';
            }
            if (msg != "") {
                alert(msg);
                return false; //結束。
            }
            /*
            Dim rqOCID As String = TIMS.ClearSQM(Request("OCID"))
            Dim rqSTDATE As String = TIMS.ClearSQM(Request("STDATE"))
            */
            var url1 = "";
            url1 = "../01/SD_01_001_sch2.aspx?ID=" + rqID + "&SPAGE=SD03002&IDNO=" + IDNO.value
                + "&BIRTH=" + birthday.value + "&CNAME=" + escape(Name.value)
                + "&OCID=" + Hid_OCID.value + "&STDATE=" + escape(STDateHidden.value);
            wopen(url1, 'MyWindow', 1200, 680, 1);
            return false;
        }

        //debugger;
        var cst_inline1 = ""; //var cst_inline1 = "inline";
        var cst_inline = ''; // 'inline';
        var cst_none = 'none';

        //職前專用-受訓前任職狀況
        function chgPriorWorkType1_disabled() {
            //(使用勞保明細檢查鈕)
            var HidPreUseLimited17f = document.getElementById('HidPreUseLimited17f');
            if (HidPreUseLimited17f.value != 'Y') { return false; } //不使用此功能
            var TPlanID = document.getElementById('TPlanID');
            //28:產投 54:充電起飛計畫(補助在職勞工參訓)
            var flagTPlan_28 = false; //產投模式運作
            if (TPlanID.value == '28' || TPlanID.value == '54') flagTPlan_28 = true;
            if (flagTPlan_28) { return false; } //產投計畫 不使用此功能
            var PriorWorkType1 = document.getElementById('PriorWorkType1');
            //var spnStartOrg = document.getElementById('spnStartOrg');
            //var spnStartDate = document.getElementById('spnStartDate');
            //var spnStartNo = document.getElementById('spnStartNo');
            //var PriorWorkType1 = document.getElementById('PriorWorkType1');
            var spnPriorWorkOrg1 = document.getElementById('spnPriorWorkOrg1');
            var spnPriorWorkOrg2 = document.getElementById('spnPriorWorkOrg2');
            var spnSOfficeYM1 = document.getElementById('spnSOfficeYM1');
            var spnSOfficeYM2 = document.getElementById('spnSOfficeYM2');
            var spnTitle1 = document.getElementById('spnTitle1');
            var spnTitle2 = document.getElementById('spnTitle2');
            //var spnRealJobless = document.getElementById('spnRealJobless');
            var PriorWorkOrg1 = document.getElementById('PriorWorkOrg1');
            var PriorWorkOrg2 = document.getElementById('PriorWorkOrg2');
            var ActNo2 = document.getElementById('ActNo2'); //ActNo/ActNo2
            var SOfficeYM1 = document.getElementById('SOfficeYM1');
            var FOfficeYM1 = document.getElementById('FOfficeYM1');
            var SOfficeYM2 = document.getElementById('SOfficeYM2');
            var FOfficeYM2 = document.getElementById('FOfficeYM2');
            var imgSOfficeYM1 = document.getElementById('IMG2'); //IMG2/imgSOfficeYM1
            var imgFOfficeYM1 = document.getElementById('IMG3'); //IMG3/imgFOfficeYM1
            var imgSOfficeYM2 = document.getElementById('IMG4'); //IMG4/imgSOfficeYM2
            var imgFOfficeYM2 = document.getElementById('IMG5'); //IMG5/imgFOfficeYM3
            var PriorWorkPay = document.getElementById('PriorWorkPay');
            var Title1 = document.getElementById('Title1');
            var Title2 = document.getElementById('Title2');
            var tdPriorWorkOrg1 = document.getElementById('tdPriorWorkOrg1');
            var tdSOfficeYM1 = document.getElementById('tdSOfficeYM1');
            var tdActNo2 = document.getElementById('tdActNo2');
            var tdPriorWorkPay = document.getElementById('tdPriorWorkPay');
            var tdTitle1 = document.getElementById('tdTitle1');
            //var tdRealJobless = document.getElementById('tdRealJobless');
            var BtnCheckBli = document.getElementById('BtnCheckBli');
            var lab2017_2 = document.getElementById('lab2017_2');
            var lab2017_4 = document.getElementById('lab2017_4');
            lab2017_2.innerHTML = '投保單位<br />加退保日期';
            lab2017_4.innerHTML = '投保薪資級距';
            //lab2017_2.innerHTML = '工作<br />起迄日期';
            //lab2017_4.innerHTML = '工作薪資';
            if (BtnCheckBli) { BtnCheckBli.style.display = cst_none; }
            //spnStartOrg.style.display = cst_inline;
            //spnStartDate.style.display = cst_inline;
            //spnStartNo.style.display = cst_inline;
            spnPriorWorkOrg1.style.display = cst_inline;
            spnPriorWorkOrg2.style.display = cst_inline;
            spnSOfficeYM1.style.display = cst_inline;
            spnSOfficeYM2.style.display = cst_inline;
            spnTitle1.style.display = cst_inline;
            spnTitle2.style.display = cst_inline;
            //spnRealJobless.style.display = cst_inline;
            PriorWorkOrg1.disabled = false;
            PriorWorkOrg2.disabled = false;
            ActNo2.disabled = false;
            SOfficeYM1.disabled = false;
            FOfficeYM1.disabled = false;
            imgSOfficeYM1.disabled = false;
            imgFOfficeYM1.disabled = false;
            SOfficeYM2.disabled = false;
            FOfficeYM2.disabled = false;
            imgSOfficeYM2.disabled = false;
            imgFOfficeYM2.disabled = false;
            PriorWorkPay.disabled = false;
            PriorWorkOrg1.style.display = cst_inline;
            PriorWorkOrg2.style.display = cst_inline;
            ActNo2.style.display = cst_inline;
            SOfficeYM1.style.display = cst_inline;
            FOfficeYM1.style.display = cst_inline;
            imgSOfficeYM1.style.display = cst_inline;
            imgFOfficeYM1.style.display = cst_inline;
            SOfficeYM2.style.display = cst_inline;
            FOfficeYM2.style.display = cst_inline;
            imgSOfficeYM2.style.display = cst_inline;
            imgFOfficeYM2.style.display = cst_inline;
            PriorWorkPay.style.display = cst_inline;
            Title1.style.display = cst_inline;
            Title2.style.display = cst_inline;
            var tdPriorWorkOrg1 = document.getElementById('tdPriorWorkOrg1');
            var tdSOfficeYM1 = document.getElementById('tdSOfficeYM1');
            var tdActNo2 = document.getElementById('tdActNo2');
            var tdPriorWorkPay = document.getElementById('tdPriorWorkPay');
            var tdTitle1 = document.getElementById('tdTitle1');
            //var tdRealJobless = document.getElementById('tdRealJobless');
            tdPriorWorkOrg1.style.color = "Black"; //.setAttribute("class", "bluecol");
            tdSOfficeYM1.style.color = "Black"; //.setAttribute("class", "bluecol");
            tdActNo2.style.color = "Black"; //.setAttribute("class", "bluecol");
            tdPriorWorkPay.style.color = "Black"; //.setAttribute("class", "bluecol");
            tdTitle1.style.color = "Black"; //.setAttribute("class", "bluecol");
            //tdRealJobless.style.color = "Black"; //.setAttribute("class", "bluecol");
            switch (getRadioValue(document.form1.PriorWorkType1)) {
                case '1':
                    if (BtnCheckBli) { BtnCheckBli.style.display = cst_inline; }
                    PriorWorkOrg1.disabled = true;
                    PriorWorkOrg2.disabled = true;
                    ActNo2.disabled = true;
                    PriorWorkPay.disabled = true;
                    tdPriorWorkOrg1.style.color = "red";
                    tdSOfficeYM1.style.color = "red"; //.setAttribute("class", "bluecol_need");
                    tdActNo2.style.color = "red"; //.setAttribute("class", "bluecol_need");
                    tdPriorWorkPay.style.color = "red"; //.setAttribute("class", "bluecol_need");
                    tdTitle1.style.color = "red"; //.setAttribute("class", "bluecol_need");
                    //tdRealJobless.style.color = "red"; //.setAttribute("class", "bluecol_need");
                    //PriorWorkOrg1.style.display = cst_none;
                    //ActNo.style.display = cst_none;
                    break;
                case '2':
                    //PriorWorkOrg1.disabled = true;
                    //ActNo.disabled = true;
                    //SOfficeYM1.disabled = true;
                    //FOfficeYM1.disabled = true;
                    PriorWorkOrg1.style.display = cst_none;
                    PriorWorkOrg2.style.display = cst_none;
                    ActNo2.style.display = cst_none;
                    SOfficeYM1.style.display = cst_none;
                    FOfficeYM1.style.display = cst_none;
                    imgSOfficeYM1.style.display = cst_none;
                    imgFOfficeYM1.style.display = cst_none;
                    SOfficeYM2.style.display = cst_none;
                    FOfficeYM2.style.display = cst_none;
                    imgSOfficeYM2.style.display = cst_none;
                    imgFOfficeYM2.style.display = cst_none;
                    PriorWorkPay.style.display = cst_none;
                    Title1.style.display = cst_none;
                    Title2.style.display = cst_none;
                    spnPriorWorkOrg1.style.display = cst_none;
                    spnPriorWorkOrg2.style.display = cst_none;
                    spnSOfficeYM1.style.display = cst_none;
                    spnSOfficeYM2.style.display = cst_none;
                    spnTitle1.style.display = cst_none;
                    spnTitle2.style.display = cst_none;
                    //spnRealJobless.style.display = cst_none;
                    break;
                case '3':
                    tdPriorWorkOrg1.style.color = "red";
                    tdSOfficeYM1.style.color = "red";
                    tdPriorWorkPay.style.color = "red"; //.setAttribute("class", "bluecol_need");
                    tdTitle1.style.color = "red"; //.setAttribute("class", "bluecol_need");
                    //tdRealJobless.style.color = "red"; //.setAttribute("class", "bluecol_need");
                    lab2017_2.innerHTML = '工作<br />起迄日期';
                    lab2017_4.innerHTML = '工作薪資';
                    //PriorWorkOrg1.style.display = cst_none;
                    //PriorWorkOrg2.style.display = cst_none;
                    //spnPriorWorkOrg1.style.display = cst_none;
                    //spnPriorWorkOrg2.style.display = cst_none;
                    ActNo2.style.display = cst_none;
                    //ActNo.disabled = true;
                    break;
                case '4':
                    //PriorWorkOrg1.disabled = true;
                    //ActNo.disabled = true;
                    //SOfficeYM1.disabled = true;
                    //FOfficeYM1.disabled = true;
                    PriorWorkOrg1.style.display = cst_none;
                    PriorWorkOrg2.style.display = cst_none;
                    ActNo2.style.display = cst_none;
                    SOfficeYM1.style.display = cst_none;
                    FOfficeYM1.style.display = cst_none;
                    imgSOfficeYM1.style.display = cst_none;
                    imgFOfficeYM1.style.display = cst_none;
                    SOfficeYM2.style.display = cst_none;
                    FOfficeYM2.style.display = cst_none;
                    imgSOfficeYM2.style.display = cst_none;
                    imgFOfficeYM2.style.display = cst_none;
                    PriorWorkPay.style.display = cst_none;
                    Title1.style.display = cst_none;
                    Title2.style.display = cst_none;
                    spnPriorWorkOrg1.style.display = cst_none;
                    spnPriorWorkOrg2.style.display = cst_none;
                    spnSOfficeYM1.style.display = cst_none;
                    spnSOfficeYM2.style.display = cst_none;
                    spnTitle1.style.display = cst_none;
                    spnTitle2.style.display = cst_none;
                    //spnRealJobless.style.display = cst_none;
                    break;
            }
        }

        //職前專用-受訓前任職狀況
        function chgPriorWorkType1() {
            //(使用勞保明細檢查鈕)
            var HidPreUseLimited17f = document.getElementById('HidPreUseLimited17f');
            if (HidPreUseLimited17f.value != 'Y') { return false; } //不使用此功能
            var TPlanID = document.getElementById('TPlanID');
            //28:產投 54:充電起飛計畫(補助在職勞工參訓)
            var flagTPlan_28 = false; //產投模式運作
            if (TPlanID.value == '28' || TPlanID.value == '54') flagTPlan_28 = true;
            if (flagTPlan_28) { return false; } //產投計畫 不使用此功能
            var PriorWorkOrg1 = document.getElementById('PriorWorkOrg1');
            var PriorWorkOrg2 = document.getElementById('PriorWorkOrg2');
            var ActNo2 = document.getElementById('ActNo2'); //ActNo/ActNo2
            var SOfficeYM1 = document.getElementById('SOfficeYM1');
            var FOfficeYM1 = document.getElementById('FOfficeYM1');
            var SOfficeYM2 = document.getElementById('SOfficeYM2');
            var FOfficeYM2 = document.getElementById('FOfficeYM2');
            var PriorWorkPay = document.getElementById('PriorWorkPay');
            var hidSB4ID = document.getElementById('hidSB4ID');
            PriorWorkOrg1.value = "";
            ActNo2.value = "";
            SOfficeYM1.value = "";
            FOfficeYM1.value = "";
            SOfficeYM2.value = "";
            FOfficeYM2.value = "";
            PriorWorkPay.value = "";
            hidSB4ID.value = "";
            //職前專用-受訓前任職狀況
            //chgPriorWorkType1_disabled();
        }

        //function ChangeSubsidy() {
        //    var LabSubsidy = document.getElementById('LabSubsidy');
        //    var SubsidyID = document.getElementById('SubsidyID');
        //    var SubsidyHidden = document.getElementById('SubsidyHidden');
        //    if (!SubsidyID || !SubsidyHidden) { return false; } //未發現生活津貼 不使用此功能
        //    LabSubsidy.style.display = 'none';
        //    if (SubsidyID.selectedIndex == 3) {
        //        LabSubsidy.style.display = cst_inline1;
        //    }
        //    if (!SubsidyID || !SubsidyHidden) { return false; } //未發現生活津貼 不使用此功能
        //    if (SubsidyHidden.value == '1') {
        //        if (confirm('變更津貼類型將會將「職業訓練生活津貼申請」相關資料刪除，確定要變更?')) {
        //            SubsidyHidden.value = '0';
        //        }
        //        else {
        //            SubsidyID.selectedIndex = 3;
        //            SubsidyHidden.value = '1';
        //        }
        //    }
        //}

        function EnterChannelChange() {
            // alert("EnterChannelChange。");
            //debugger;
            var TRNDTR = document.getElementById('TRNDTR');
            var GovTR = document.getElementById('GovTR');
            var EnterChannel = document.form1.EnterChannel;
            var hide_TrainMode = document.getElementById('hide_TrainMode');
            var hide_EnterChannel = document.getElementById('hide_EnterChannel');
            var hide_EnterChannel2 = document.getElementById('hide_EnterChannel2');
            if (!EnterChannel || !TRNDTR) { return false; } //未發現報名管道 不使用此功能
            TRNDTR.style.display = 'none';
            GovTR.style.display = 'none';
            if (EnterChannel.value == '4') {
                TRNDTR.style.display = cst_inline1;
                GovTR.style.display = cst_inline1;
            }
            if (EnterChannel.value != '4') {
                //'1.網;2.現;3.通;4.推
                if (hide_EnterChannel.value == '4') {
                    TRNDTR.style.display = cst_inline1;
                    GovTR.style.display = cst_inline1;
                    alert("該學員無法從推介更改為其他報名管道。");
                    EnterChannel.value = '4';
                }
            }
            //有現場報名資料，網路可改現場報名
            //'0.沒有有現場報名'1.有現場報名
            if (hide_EnterChannel2.value == '0') {
                //'1.網;2.現;3.通;4.推
                if (hide_EnterChannel.value == '1') {
                    if (EnterChannel.value == '2' || EnterChannel.value == '3') {
                        alert("報名管道不能從網路報名更改為現場或通訊報名。");
                        EnterChannel.value = hide_EnterChannel.value;
                    }
                }
            }
        }

        function Change_MBTable(num) {
            var Page = 1;
            switch (num) {
                case 1:
                    //('在職者顯示 MenuTable / BackTable) 改變id 讓 chkdata 找到 BackTable,MenuTable
                    //if (document.getElementById('UBackTable')) {
                    //	document.getElementById('UBackTable').style.display= cst_inline1;
                    //	document.getElementById('UBackTable').id="BackTable";
                    //}
                    if (document.getElementById('UMenuTable')) {
                        document.getElementById('UMenuTable').style.display = cst_inline1;
                        document.getElementById('UMenuTable').id = "MenuTable";
                    }
                    //('在職者顯示 MenuTable / BackTable)
                    //if (document.getElementById('BackTable')) {
                    //	document.getElementById('BackTable').style.display=cst_inline1;
                    //}
                    //if (document.getElementById('ChangeMode2a')) { document.getElementById('ChangeMode2a').style.display = 'none'; }
                    //if (document.getElementById('ChangeMode2b')) { document.getElementById('ChangeMode2b').style.display = 'none'; }
                    //if (document.getElementById('ChangeMode2c')) { document.getElementById('ChangeMode2c').style.display = 'none'; }
                    if (document.getElementById('MenuTable_td_2')) { document.getElementById('MenuTable_td_2').style.display = 'none'; }

                    if (document.getElementById('MenuTable')) {
                        document.getElementById('MenuTable').style.display = cst_inline1;
                    }
                    break;
                case 2:
                    //('失業者不顯示 MenuTable / BackTable)
                    if (document.getElementById('BackTable')) {
                        document.getElementById('BackTable').style.display = 'none';
                        document.getElementById('BackTable').id = "UBackTable";
                    }
                    if (document.getElementById('MenuTable')) {
                        document.getElementById('MenuTable').style.display = 'none';
                        document.getElementById('MenuTable').id = "UMenuTable";
                    }
                    //('失業者不顯示 MenuTable / BackTable)改變id 讓 chkdata 找不到 BackTable,MenuTable
                    if (document.getElementById('UBackTable')) {
                        document.getElementById('UBackTable').style.display = 'none';
                    }
                    if (document.getElementById('UMenuTable')) {
                        document.getElementById('UMenuTable').style.display = 'none';
                    }
                    break;
            }
            ChangeMode(Page);
        }


        //改變國籍身分
        function ChangePassPort() {
            var cst_pt1 = 0;
            var cst_pt2 = 1;
            if (document.getElementsByName('ChinaOrNot').length > 2) {
                cst_pt1 = 1; //cst_pt
                cst_pt2 = 2;
            }
            var cst_pp1 = 0;
            var cst_pp2 = 1;
            if (document.getElementsByName('PPNO').length > 2) {
                cst_pp1 = 1; //cst_pp
                cst_pp2 = 2;
            }
            var cst_fs1 = 0;
            var cst_fs2 = 1;
            if (document.getElementsByName('ForeSex').length > 2) {
                cst_fs1 = 1; //cst_fs
                cst_fs2 = 2;
            }
            if (getRadioValue(document.form1.PassPortNO) == 1) {
                document.getElementById('ChinaOrNotTable').style.display = 'none';
                document.getElementById('PPNO').style.display = 'none';
                document.getElementsByName('ChinaOrNot')[cst_pt1].checked = false;
                document.getElementsByName('ChinaOrNot')[cst_pt2].checked = false;
                document.getElementById('Nationality').value = '';
                document.getElementsByName('PPNO')[cst_pp1].checked = false;
                document.getElementsByName('PPNO')[cst_pp2].checked = false;
                for (i = 1; i <= 5; i++) {
                    var ForeTrii = document.getElementById('ForeTr' + i);
                    if (ForeTrii) { ForeTrii.style.display = 'none'; }
                }
                document.getElementById('ForeName').value = '';
                document.getElementById('ForeTitle').value = '';
                document.getElementsByName('ForeSex')[cst_fs1].checked = false;
                document.getElementsByName('ForeSex')[cst_fs2].checked = false;
                document.getElementById('ForeBirth').value = '';
                document.getElementById('ForeIDNO').value = '';
                document.getElementById('City6').value = '';
                document.getElementById('ForeZip').value = '';
                document.getElementById('ForeAddr').value = '';
            }
            else {
                document.getElementById('ChinaOrNotTable').style.display = cst_inline1;
                document.getElementById('PPNO').style.display = cst_inline1;
                for (i = 1; i <= 5; i++) {
                    var ForeTrii = document.getElementById('ForeTr' + i);
                    if (ForeTrii) { ForeTrii.style.display = cst_inline1; }
                }
            }
            if (document.body) { window.scroll(0, document.body.scrollHeight); }
            if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }

        }

        //變更銀行
        function ChangeBank() {
            document.getElementById('PortTR').style.display = 'none';
            document.getElementById('BankTR1').style.display = 'none';
            document.getElementById('BankTR2').style.display = 'none';
            document.getElementById('BankTR3').style.display = 'none';
            var PostNo_1 = document.getElementById('PostNo_1');
            //var PostNo_2 = document.getElementById('PostNo_2');
            var AcctNo1_1 = document.getElementById('AcctNo1_1');
            //var AcctNo1_2 = document.getElementById('AcctNo1_2');
            var AcctheadNo = document.getElementById('AcctheadNo');
            var AcctExNo = document.getElementById('AcctExNo');
            var AcctNo2 = document.getElementById('AcctNo2');
            var HidPostNo_1 = document.getElementById('HidPostNo_1');
            //var HidPostNo_2 = document.getElementById('HidPostNo_2');
            var HidAcctNo1_1 = document.getElementById('HidAcctNo1_1');
            //var HidAcctNo1_2 = document.getElementById('HidAcctNo1_2');
            var HidAcctheadNo = document.getElementById('HidAcctheadNo');
            var HidAcctExNo = document.getElementById('HidAcctExNo');
            var HidAcctNo2 = document.getElementById('HidAcctNo2');
            if (PostNo_1.value != '') HidPostNo_1.value = PostNo_1.value;
            //if (PostNo_2.value != '') HidPostNo_2.value = PostNo_2.value;
            if (AcctNo1_1.value != '') HidAcctNo1_1.value = AcctNo1_1.value;
            //if (AcctNo1_2.value != '') HidAcctNo1_2.value = AcctNo1_2.value;
            if (AcctheadNo.value != '') HidAcctheadNo.value = AcctheadNo.value;
            if (AcctExNo.value != '') HidAcctExNo.value = AcctExNo.value;
            if (AcctNo2.value != '') HidAcctNo2.value = AcctNo2.value;
            PostNo_1.value = '';
            //PostNo_2.value = '';
            AcctNo1_1.value = '';
            //AcctNo1_2.value = '';
            AcctheadNo.value = '';
            AcctExNo.value = '';
            AcctNo2.value = '';
            //debugger;
            if (isChecked(document.getElementsByName('AcctMode'))) {
                switch (getRadioValue(document.getElementsByName('AcctMode'))) {
                    case '0':
                        document.getElementById('PortTR').style.display = cst_inline1;
                        PostNo_1.value = HidPostNo_1.value;
                        //PostNo_2.value = HidPostNo_2.value;
                        AcctNo1_1.value = HidAcctNo1_1.value
                        //AcctNo1_2.value = HidAcctNo1_2.value;
                        break;
                    case '1':
                        document.getElementById('BankTR1').style.display = cst_inline1;
                        document.getElementById('BankTR2').style.display = cst_inline1;
                        document.getElementById('BankTR3').style.display = cst_inline1;
                        AcctheadNo.value = HidAcctheadNo.value;
                        AcctExNo.value = HidAcctExNo.value;
                        AcctNo2.value = HidAcctNo2.value;
                        break;
                        //case '2':                                                                                                                                                                                                                               
                        //如果是勞工團體會有value=2的情況，不過因為不需要有輸入框，所以等於可以不用作任何判斷處理                                                                                                                                                                                                                               
                }
            }
            if (document.body) { window.scroll(0, document.body.scrollHeight); }
            if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }

        }

        //save check
        function chkdata() {
            var msg = '';
            var Item = '';
            var Page = 1;
            var TPlanID = document.getElementById('TPlanID');
            //28:產投 54:充電起飛計畫(補助在職勞工參訓)
            var flagTPlan_28 = false; //產投模式運作
            if (TPlanID.value == '28' || TPlanID.value == '54') flagTPlan_28 = true;
            //(使用勞保明細檢查鈕)
            var HidPreUseLimited17f = document.getElementById('HidPreUseLimited17f');
            var tdPriorWorkOrg1 = document.getElementById('tdPriorWorkOrg1');
            var tdSOfficeYM1 = document.getElementById('tdSOfficeYM1');
            var tdActNo2 = document.getElementById('tdActNo2');
            var tdPriorWorkPay = document.getElementById('tdPriorWorkPay');
            var tdTitle1 = document.getElementById('tdTitle1');
            //var tdRealJobless = document.getElementById('tdRealJobless');
            var PriorWorkOrg1 = document.getElementById('PriorWorkOrg1');
            var PriorWorkOrg2 = document.getElementById('PriorWorkOrg2');
            var SOfficeYM1 = document.getElementById('SOfficeYM1');
            var FOfficeYM1 = document.getElementById('FOfficeYM1');
            var ActNo2 = document.getElementById('ActNo2');
            var PriorWorkPay = document.getElementById('PriorWorkPay');
            var Title1 = document.getElementById('Title1');
            //var JoblessID = document.getElementById('JoblessID');
            //var RealJobless = document.getElementById('RealJobless');
            var lab2017_2 = document.getElementById('lab2017_2');
            var lab2017_4 = document.getElementById('lab2017_4');
            //lab2017_2.innerHTML = '投保單位<br />加退保日期';
            //lab2017_4.innerHTML = '投保薪資級距';
            var JobStateType = document.getElementById('JobStateType');
            var JobStateTypeVal = getRadioValue(document.form1.JobStateType);
            var NrblHandType = document.getElementsByName('rblHandType');
            var IrblHandType = document.getElementById('rblHandType');
            var BudID = document.getElementById("BudID");
            if (BudID) { BudID.disabled = false; }

            //企訓專用('產學訓計畫才會顯示 MenuTable / BackTable)
            var rblWorkSuppIdent = document.getElementById('rblWorkSuppIdent');
            var WorkSuppIdentVal = ""; //getRBLValue('rblWorkSuppIdent'); //取得 RadioButtonList 值
            if (rblWorkSuppIdent) {
                WorkSuppIdentVal = getRBLValue('rblWorkSuppIdent'); //取得 RadioButtonList 值
            }
            if (document.form1.LevelNo.disabled == false)
                if (document.form1.LevelNo.selectedIndex == 0) { msg += '請選擇報名階段\n'; if (Item == '') Item = 'LevelNo'; Page = 1; }
            if (document.form1.Name.value == '') { msg += '請輸入姓名\n'; if (Item == '') Item = 'Name'; Page = 1; }
            if (document.form1.StudentID.value == '') { msg += '請輸入學號\n'; if (Item == '') Item = 'StudentID'; Page = 1; }
            if (document.form1.StudentID.value != '' && !isUnsignedInt(document.form1.StudentID.value)) { msg += '學號必須為數字\n'; if (Item == '') Item = 'StudentID'; Page = 1; }
            $('#ZipCode2').val($.trim($('#ZipCode2').val()));
            $('#ZipCode2_B3').val($.trim($('#ZipCode2_B3').val()));
            $('#HouseholdAddress').val($.trim($('#HouseholdAddress').val()));
            if (document.getElementById('CheckBox3').checked == true) {
                if (document.getElementById('CheckBox1').checked == true) {
                    document.getElementById('CheckBox2').checked = true; //即同通訊地址
                    document.getElementById('CheckBox3').checked = false;
                } else {
                    if ($('#ZipCode2').val() == '') { msg += '緊急通知人地址如要同戶籍地址時，必須完整填寫戶籍地址(郵遞區號前3碼)。\n'; if (Item == '') Item = 'ZipCode2'; Page = 1; }
                    if ($('#ZipCode2_B3').val() == '') { msg += '緊急通知人地址如要同戶籍地址時，必須完整填寫戶籍地址(郵遞區號4.5碼)。\n'; if (Item == '') Item = 'ZipCode2_B3'; Page = 1; }
                    if ($('#HouseholdAddress').val() == '') { msg += '緊急通知人地址如要同戶籍地址時，必須完整填寫戶籍地址。\n'; if (Item == '') Item = 'HouseholdAddress'; Page = 1; }
                }
            }
            if (TPlanID.value == '46') {
                if (document.getElementById('BackTable')) {
                    document.getElementById('BackTable').style.display = 'none';
                    document.getElementById('BackTable').id = "UBackTable";
                }
            }
            if (TPlanID.value == '47') {
                if (document.getElementById('BackTable')) {
                    document.getElementById('BackTable').style.display = 'none';
                    document.getElementById('BackTable').id = "UBackTable";
                }
            }

            // 如果是產學訓就不擋英文姓名 緊急通知人 失業週數
            if (!flagTPlan_28) {
                //document.form1.SubsidyID.selectedvalue==''
                //(排除在職進修)
                //if (TPlanID.value != '06') {
                //    if (document.form1.SubsidyID.selectedIndex == 0) { msg += '請選擇申請津貼類別\n'; if (Item == '') Item = 'SubsidyID'; Page = 1; }
                //}
                //if (document.form1.SubsidyID.selectedIndex == 3) {
                //    if (document.form1.SubsidyIdentity.selectedIndex == 0) { msg += '津貼類別為就業促進津貼實施辦法，則津貼身分別為必填\n'; if (Item == '') Item = 'SubsidyIdentity'; Page = 1; }
                //}
                if (document.form1.LName.value == '' || document.form1.FName.value == '') {
                    msg += '請填寫英文姓名\n'; if (Item == '') Item = 'LName'; Page = 1;
                }
                else {
                    if (!isEng(document.form1.LName.value)) { msg += 'LastName必須為英文字\n'; if (Item == '') Item = 'LName'; Page = 1; }
                    if (!isEng(document.form1.FName.value)) { msg += 'FirstName必須為英文字\n'; if (Item == '') Item = 'FName'; Page = 1; }
                }

                $('#ZipCode3').val($.trim($('#ZipCode3').val()));
                $('#ZipCode3_B3').val($.trim($('#ZipCode3_B3').val()));
                $('#EmergencyAddress').val($.trim($('#EmergencyAddress').val()));
                if (document.form1.CheckBox2.checked == false && document.form1.CheckBox3.checked == false) {
                    if ($('#ZipCode3').val() != '' && $('#ZipCode3').val().length != 3) { msg += '緊急通知人地址(郵遞區號前3碼)必須為3碼\n'; if (Item == '') Item = 'ZipCode3'; Page = 1; }
                    if ($('#ZipCode3_B3').val() != '') {
                        //checkzip23 郵遞區號
                        msg += checkzip23(true, '緊急通知人地址', 'ZipCode3_B3');
                        //if (!isUnsignedInt($('#ZipCode3_B3').val())) { msg += '緊急通知人地址郵遞區號後3碼/後2碼必須為數字，且不得輸入 00\n'; if (Item == '') Item = 'ZipCode3_B3'; Page = 1; }
                        //if (parseInt($('#ZipCode3_B3').val(), 10) < 1) { msg += '緊急通知人地址郵遞區號後3碼/後2碼必須為數字，得輸入 01~999 \n'; if (Item == '') Item = 'ZipCode3_B3'; Page = 1; }
                        //if ($('#ZipCode3_B3').val().length != 2 && $('#ZipCode3_B3').val().length != 3) { msg += '緊急通知人地址郵遞區號後3碼/後2碼長度必須為 3碼或2碼(例 01 或 001)\n'; if (Item == '') Item = 'ZipCode3_B3'; Page = 1; }
                    }
                }

                //if (document.form1.ZipCode3.value=='') {msg+='請輸入緊急聯絡人通訊地址(區域)\n';if(Item=='') Item='City3';Page=1;}
                //if (document.form1.EmergencyAddress.value=='') {msg+='請輸入緊急連絡人通訊地址\n';if(Item=='') Item='EmergencyAddress';Page=1;}
                //if (RealJobless && RealJobless.value != '' && !isUnsignedInt(RealJobless.value)) { msg += '失業週數必須為數字\n'; if (Item == '') Item = 'RealJobless'; Page = 1; }
                //if (JoblessID && JoblessID.selectedIndex == 0) { msg += '請選擇受訓前失業週數\n'; if (Item == '') Item = 'JoblessID'; Page = 1; }
                /*
				for(i=0,j=0;i<document.form1.JobStateType.length;i++){
				if (!document.form1.JobStateType[i].checked) j++;
				}				
				*/
                //受訓前任職資料//(排除在職進修)
                //if (TPlanID.value != '06') {
                //    if (!isChecked(document.form1.PriorWorkType1)) { msg = msg + '請選擇受訓前任職狀況!\n'; if (Item == '') Item = 'PriorWorkType1'; Page = 1; }
                //}
                //受訓前任職資料
                if (JobStateType && JobStateTypeVal == '') { msg += '請選擇就職狀況!\n'; if (Item == '') Item = 'JobStateType'; Page = 1; }
                switch (getRadioValue(document.form1.PriorWorkType1)) {
                    case '1':
                        //if (document.getElementById('PriorWorkOrg1').value == '') { msg = msg + '請輸入受訓服務單位1\n'; if (Item == '') Item = 'PriorWorkOrg1'; Page = 1; }
                        //if (document.getElementById('SOfficeYM1').value == '') { msg = msg + '請輸入受訓前任職起迄日期的起日1!\n'; if (Item == '') Item = 'SOfficeYM1'; Page = 1; }
                        //if (document.getElementById('FOfficeYM1').value == '') { msg = msg + '請輸入受訓前任職起迄日期的迄日1!\n'; if (Item == '') Item = 'FOfficeYM1'; Page = 1; }
                        //(排除產投 與 在職進修)
                        /*
						if ((!flagTPlan_28) && (TPlanID.value != '06')) {
						if (document.getElementById('ActNo2').value == '') { msg = msg + '請輸入最後保險單位保險證號\n'; if (Item == '') Item = 'ActNo2'; Page = 1; }
						}
						*/
                        if (tdPriorWorkOrg1 && tdPriorWorkOrg1.style.color == "red") {
                            if (PriorWorkOrg1 && PriorWorkOrg1.value == '') { msg += '請輸入任職單位名稱\n'; if (Item == '') Item = 'PriorWorkOrg1'; Page = 1; }
                        }
                        if (tdSOfficeYM1 && tdSOfficeYM1.style.color == "red") {
                            if (SOfficeYM1 && SOfficeYM1.value == '') { msg += '請輸入投保單位加保日期\n'; if (Item == '') Item = 'SOfficeYM1'; Page = 1; }
                            //(使用勞保明細檢查鈕)
                            if (WorkSuppIdentVal != "Y") {
                                if (FOfficeYM1 && FOfficeYM1.value == '' && HidPreUseLimited17f.value == 'Y') { msg += '請輸入投保單位退保日期\n'; if (Item == '') Item = 'FOfficeYM1'; Page = 1; }
                            }
                        }
                        if (tdActNo2 && tdActNo2.style.color == "red") {
                            if (ActNo2 && ActNo2.value == '') { msg += '請輸入投保單位保險證號\n'; if (Item == '') Item = 'ActNo2'; Page = 1; }
                        }
                        if (tdPriorWorkPay && tdPriorWorkPay.style.color == "red") {
                            if (PriorWorkPay && PriorWorkPay.value == '') { msg += '請輸入投保薪資級距\n'; if (Item == '') Item = 'PriorWorkPay'; Page = 1; }
                            if (PriorWorkPay && PriorWorkPay.value != '' && !isUnsignedInt(PriorWorkPay.value)) { msg += '投保薪資級距必須為數字\n'; if (Item == '') Item = 'PriorWorkPay'; Page = 1; }
                        }
                        if (tdTitle1 && tdTitle1.style.color == "red") {
                            if (Title1 && Title1.value == '') { msg += '請輸入職稱\n'; if (Item == '') Item = 'Title1'; Page = 1; }
                        }
                        /*
                        if (tdRealJobless && tdRealJobless.style.color == "red") {
                            if (RealJobless && RealJobless.value == '' && JoblessID && JoblessID.selectedIndex == 0) {
                                 if (WorkSuppIdentVal != "Y") { msg += '請輸入受訓前失業週數\n'; if (Item == '') Item = 'RealJobless'; Page = 1; } 
                            }
                            else {
                                if (WorkSuppIdentVal != "Y") {
                                    if (RealJobless && RealJobless.value == '') { msg += '請輸入受訓前失業週數\n'; if (Item == '') Item = 'RealJobless'; Page = 1; }
                                }
                                if (RealJobless && RealJobless.value != '' && !isUnsignedInt(RealJobless.value)) { msg += '失業週數必須為數字\n'; if (Item == '') Item = 'RealJobless'; Page = 1; }
                                if (JoblessID && JoblessID.selectedIndex == 0) { msg += '請選擇受訓前失業週數\n'; if (Item == '') Item = 'JoblessID'; Page = 1; }
                            }
                        }
                        */
                        break;
                    case '3':
                        if (tdPriorWorkOrg1 && tdPriorWorkOrg1.style.color == "red") {
                            if (PriorWorkOrg1 && PriorWorkOrg1.value == '') { msg += '請輸入任職單位名稱\n'; if (Item == '') Item = 'Title1'; Page = 1; }
                        }
                        if (tdSOfficeYM1 && tdSOfficeYM1.style.color == "red") {
                            if (SOfficeYM1 && SOfficeYM1.value == '') { msg += '請輸入工作起始日期\n'; if (Item == '') Item = 'Title1'; Page = 1; }
                            //(使用勞保明細檢查鈕)
                            if (WorkSuppIdentVal != "Y") {
                                if (FOfficeYM1 && FOfficeYM1.value == '' && HidPreUseLimited17f.value == 'Y') { msg += '請輸入工作迄止日期\n'; if (Item == '') Item = 'Title1'; Page = 1; }
                            }
                        }
                        if (tdPriorWorkPay && tdPriorWorkPay.style.color == "red") {
                            if (PriorWorkPay && PriorWorkPay.value == '') { msg += '請輸入工作薪資\n'; if (Item == '') Item = 'Title1'; Page = 1; }
                            //if (PriorWorkPay && PriorWorkPay.value != '' && !isUnsignedInt(RealJobless.value)) { msg += '投保薪資級距必須為數字\n'; if (Item == '') Item = 'Title1'; Page = 1; }
                            if (PriorWorkPay && PriorWorkPay.value != '') { msg += '投保薪資級距必須為數字\n'; if (Item == '') Item = 'Title1'; Page = 1; }
                        }
                        if (tdTitle1 && tdTitle1.style.color == "red") {
                            if (Title1 && Title1.value == '') { msg += '請輸入職稱\n'; if (Item == '') Item = 'Title1'; Page = 1; }
                        }
                        /*
                        if (tdRealJobless && tdRealJobless.style.color == "red") {
                            if (RealJobless && RealJobless.value == '' && JoblessID && JoblessID.selectedIndex == 0) {
                                if (WorkSuppIdentVal != "Y") { msg += '請輸入受訓前失業週數\n'; if (Item == '') Item = 'RealJobless'; Page = 1; }
                            }
                            else {
                                if (WorkSuppIdentVal != "Y") {
                                    if (RealJobless && RealJobless.value == '') { msg += '請輸入受訓前失業週數\n'; if (Item == '') Item = 'RealJobless'; Page = 1; }
                                }
                                if (RealJobless && RealJobless.value != '' && !isUnsignedInt(RealJobless.value)) { msg += '失業週數必須為數字\n'; if (Item == '') Item = 'RealJobless'; Page = 1; }
                                if (JoblessID && JoblessID.selectedIndex == 0) { msg += '請選擇受訓前失業週數\n'; if (Item == '') Item = 'JoblessID'; Page = 1; }
                            }
                        }
                        */
                        break;
                }
                //if (document.form1.PriorWorkType1[0].checked) {}
                //投保單位<br />加退保日期
                //投保單位<br />加退保日期
                //if (SOfficeYM1.value != '' && FOfficeYM1.value != '' && FOfficeYM1.value < SOfficeYM1.value)
                //{ msg += '[ 加退保日期 的迄日]必需大於[ 加退保日期 的起日]\n'; if (Item == '') Item = 'FOfficeYM1'; Page = 1; }
                if (JobStateType && JobStateTypeVal == '0') {
                    //選擇失業
                    //if (RealJobless && RealJobless.value != '' && !isUnsignedInt(RealJobless.value)) { msg += '失業週數必須為數字\n'; if (Item == '') Item = 'RealJobless'; Page = 1; }
                    //if (JoblessID && JoblessID.selectedIndex == 0) { msg += '請選擇受訓前失業週數\n'; if (Item == '') Item = 'JoblessID'; Page = 1; }
                }
                /*,if (document.form1.JobStateType) {,if (!isChecked(document.form1.JobStateType)) { msg = msg + '請選擇就職狀況!\n'; if (Item == '') Item = 'JobStateType'; Page = 1; }
                ,else {,if (document.form1.JobStateType[1].checked) {,if (document.form1.JoblessID.selectedIndex == 0) { msg += '請選擇受訓前失業週數\n'; 
                if (Item == '') Item = 'JoblessID'; Page = 1; },if (document.form1.RealJobless.value != '' && !isUnsignedInt(document.form1.RealJobless.value)) 
                { msg += '失業週數必須為數字\n'; if (Item == '') Item = 'RealJobless'; Page = 1; },},},},*/

                //if (document.form1.RealJobless.value != '' && !isUnsignedInt(document.form1.RealJobless.value)) { msg += '失業週數必須為數字\n'; if (Item == '') Item = 'RealJobless'; Page = 1; }
                if (document.form1.School.value == '') { msg += '請輸入學校\n'; if (Item == '') Item = 'School'; Page = 1; }
                if (document.form1.Department.value == '') { msg += '請輸入科系\n'; if (Item == '') Item = 'Department'; Page = 1; }
                if (document.form1.GraduateStatus.selectedIndex == 0) { msg += '請選擇畢業狀況\n'; if (Item == '') Item = 'GraduateStatus'; Page = 1; }
                //if (document.form1.MilitaryID.selectedIndex == 0) { msg += '請選擇兵役狀況\n'; if (Item == '') Item = 'MilitaryID'; Page = 1; }
                /*
				if (document.form1.MIdentityID.value == '05') {
				if (document.form1.NativeID.selectedIndex == 0) { msg += '請選擇民族別\n'; if (Item == '') Item = 'NativeID'; Page = 1; }
				}
				*/
            }
            else {
                //debugger;
                //alert(' ZipCode3=['+ document.getElementById('ZipCode3').value +']\n EmergencyAddress=['+document.getElementById('EmergencyAddress').value+']\n EmergencyAddress=['+document.getElementById('EmergencyAddress').value+']');
                //20090521 fix 依需求若非必填地址項其中一欄有填寫資料，則其它相依欄位也同時改為必填(填寫完整地址資料)
                if (document.form1.CheckBox2.checked == false && document.form1.CheckBox3.checked == false) {
                    $('#ZipCode3').val($.trim($('#ZipCode3').val()));
                    $('#ZipCode3_B3').val($.trim($('#ZipCode3_B3').val()));
                    $('#EmergencyAddress').val($.trim($('#EmergencyAddress').val()));

                    if ($('#ZipCode3').val() != "" || $('#ZipCode3_B3').val() != "" || $('#EmergencyAddress').val() != "") {
                        if ($('#ZipCode3').val() == '') { msg += '緊急通知人地址項若有填值時，必須填寫完整 (郵遞區號前3碼)\n'; if (Item == '') Item = 'ZipCode3'; Page = 1; }
                        if ($('#ZipCode3_B3').val() == '') { msg += '緊急通知人地址項若有填值時，必須填寫完整 (郵遞區號3+3,3+2碼)\n'; if (Item == '') Item = 'ZipCode3_B3'; Page = 1; }
                        if ($('#EmergencyAddress').val() == '') { msg += '緊急通知人地址項若有填值時，必須填寫完整 (地址)\n'; if (Item == '') Item = 'EmergencyAddress'; Page = 1; }

                        if ($('#ZipCode3_B3').val() != '') {
                            //checkzip23 郵遞區號
                            msg += checkzip23(true, '緊急通知人地址', 'ZipCode3_B3');
                        }
                    }
                }
                else {
                    if ($('#ZipCode3').val() != '') {
                        if ($('#ZipCode3').val().length != 3) { msg += '緊急通知人地址項若有填值時，必須填寫完整 (郵遞區號前3碼)須為3碼\n'; if (Item == '') Item = 'ZipCode3'; Page = 1; }
                        if ($('#EmergencyAddress').val() == '') { msg += '緊急通知人地址項若有填值時，必須填寫完整 (地址)\n'; if (Item == '') Item = 'EmergencyAddress'; Page = 1; }
                    }
                    if ($('#ZipCode3_B3').val() != '') {
                        //checkzip23 郵遞區號
                        msg += checkzip23(true, '緊急通知人地址', 'ZipCode3_B3');
                    }
                    //if ($('#ZipCode3_B3').val() != '') {
                    //    if (!isUnsignedInt($('#ZipCode3_B3').val())) { msg += '緊急通知人地址郵遞區號後3碼/後2碼必須為數字，且不得輸入 00\n'; if (Item == '') Item = 'ZipCode3_B3'; Page = 1; }
                    //    if (parseInt($('#ZipCode3_B3').val(), 10) < 1) { msg += '緊急通知人地址郵遞區號後3碼/後2碼必須為數字，得輸入 01~999 \n'; if (Item == '') Item = 'ZipCode3_B3'; Page = 1; }
                    //    if ($('#ZipCode3_B3').val().length != 2 && $('#ZipCode3_B3').val().length != 3) { msg += '緊急通知人地址郵遞區號後3碼/後2碼長度必須為 3碼或2碼(例 01 或 001)\n'; if (Item == '') Item = 'ZipCode3_B3'; Page = 1; }
                    //}
                }
            }
            //end
            /*
            for(i=0,j=0;i<document.form1.PassPortNO.length;i++){
            if (!document.form1.PassPortNO[i].checked) j++;
            }
            //所有計畫都應該填寫下列資訊。
            */
            if (!isChecked(document.form1.PassPortNO)) { msg = msg + '請選擇身分別(國籍)!\n'; if (Item == '') Item = 'PassPortNO'; Page = 1; }
            else {
                if (document.form1.PassPortNO[1].checked) {
                    if (!isChecked(document.form1.ChinaOrNot)) { msg = msg + '請選擇是否為大陸人士!\n'; if (Item == '') Item = 'ChinaOrNot'; Page = 1; }
                    if (document.getElementById('Nationality').value == '') { msg = msg + '請輸入原屬國籍!\n'; if (Item == '') Item = 'Nationality'; Page = 1; }
                    if (!isChecked(document.form1.PPNO)) { msg = msg + '請選擇護照或居留(工作)證號!\n'; if (Item == '') Item = 'PPNO'; Page = 1; }
                }
            }
            if (document.form1.IDNO.value == '') { msg += '請輸入身分證號碼\n'; if (Item == '') Item = 'IDNO'; Page = 1; }
            else if (document.form1.PassPortNO[0].checked == true) {
                if (document.getElementById('RoleID').value != '99' || document.getElementById('Process') == 'edit') {
                    var pattern = /^[A-Z][1-2]{1}\d{8}$/;
                    if (!pattern.test(document.form1.IDNO.value)) { msg += '身分證號碼錯誤\n'; if (Item == '') Item = 'IDNO'; Page = 1; }
                }
                else {
                    if (!checkId(document.form1.IDNO.value)) { msg += '身分證號碼錯誤(如果有此身分證號碼，請聯絡系統管理者)\n'; if (Item == '') Item = 'IDNO'; Page = 1; }
                }
            }
            if (!isChecked(document.form1.Sex)) {
                msg = msg + '請選擇性別!\n'; if (Item == '') Item = 'Sex'; Page = 1;
            }
            else {
                if (document.form1.PassPortNO[0].checked == true) {
                    //if (document.form1.IDNO.value!='' && !checkId(document.form1.IDNO.value)) msg+='身分證號碼不正確\n';
                    if (document.form1.IDNO.value.charAt(1) == 1 && getRadioValue(document.form1.Sex) == 'F') { msg += '性別與身分證號碼不符合\n'; if (Item == '') Item = 'IDNO'; Page = 1; }
                    else if (document.form1.IDNO.value.charAt(1) == 2 && getRadioValue(document.form1.Sex) == 'M') { msg += '性別與身分證號碼不符合\n'; if (Item == '') Item = 'IDNO'; Page = 1; }
                }
            }
            if (document.form1.Birthday.value == '') { msg += '請輸入出生日期\n'; if (Item == '') Item = 'Birthday'; Page = 1; }
            // if (!isChecked(document.form1.MaritalStatus)) {
            //  msg = msg + '請選擇婚姻狀況!\n'; if (Item == '') Item = 'MaritalStatus'; Page = 1;
            // }
            //EnterChannel 
            if (document.form1.EnterChannel.disabled == false)
                if (document.form1.EnterChannel.selectedIndex == 0) { msg += '請選擇報名管道\n'; if (Item == '') Item = 'EnterChannel'; Page = 1; }
            /*
            for(i=0,j=0;i<document.form1.MaritalStatus.length;i++){
            if (!document.form1.MaritalStatus[i].checked) j++;
            }	
            if (j==document.form1.MaritalStatus.length) msg=msg+'請選擇婚姻狀況!\n';
            */
            if (document.form1.Birthday.value != '' && !checkDate(document.form1.Birthday.value)) { msg += '出生日期格式不正確\n'; if (Item == '') Item = 'Birthday'; Page = 1; }
            if (document.form1.EnterChannel.value == '4') {
                if (!document.form1.TRNDMode.disabled && document.form1.TRNDMode.selectedIndex == 0) {
                    msg += '請選擇推介種類\n'; if (Item == '') Item = 'TRNDMode'; Page = 1;
                }
                //else { }
                //if(document.form1.TRNDMode.value=='1' || document.form1.TRNDMode.value=='3'){
                //if (!document.form1.TRNDType.disabled && document.form1.TRNDMode.value == '1') {if (!isChecked(document.form1.TRNDType)) { msg += '請選擇券別種類\n'; if (Item == '') Item = 'TRNDMode'; Page = 1; }}
            }
            if (document.form1.OpenDate != null) {
                if (document.form1.OpenDate.value != '' && !checkDate(document.form1.OpenDate.value)) { msg += '開訓日期格式不正確\n'; if (Item == '') Item = 'OpenDate'; Page = 1; }
            }
            if (document.form1.CloseDate != null) {
                if (document.form1.CloseDate.value != '' && !checkDate(document.form1.CloseDate.value)) { msg += '結訓日期格式不正確\n'; if (Item == '') Item = 'CloseDate'; Page = 1; }
            }
            if (document.form1.EnterDate.value != '' && !checkDate(document.form1.EnterDate.value)) { msg += '報到日期格式不正確\n'; if (Item == '') Item = 'EnterDate'; Page = 1; }
            if (document.form1.DegreeID.selectedIndex == 0) { msg += '請選擇最高學歷\n'; if (Item == '') Item = 'DegreeID'; Page = 1; }
            if (document.form1.MilitaryID.selectedIndex == 5) {
                //在役中判斷同種
                if (document.form1.ServiceID.value == '') { msg += '請輸入軍種\n'; if (Item == '') Item = 'ServiceID'; Page = 1; }
                if (document.form1.MilitaryRank.value == '') { msg += '請輸入階級\n'; if (Item == '') Item = 'MilitaryRank'; Page = 1; }
                if (document.form1.ServiceOrg.value == '') { msg += '請輸入服務單位名稱\n'; if (Item == '') Item = 'ServiceOrg'; Page = 1; }
                if (document.form1.ServicePhone.value == '') { msg += '請輸入服務單位電話\n'; if (Item == '') Item = 'ServicePhone'; Page = 1; }
                if (document.form1.SServiceDate.value == '') { msg += '請輸入起始服役日期\n'; if (Item == '') Item = 'SServiceDate'; Page = 1; }
                if (document.form1.SServiceDate.value != '' && !checkDate(document.form1.SServiceDate.value)) { msg += '起始服役日期 格式不正確\n'; if (Item == '') Item = 'EnterDate'; Page = 1; }
                if (document.form1.FServiceDate.value == '') { msg += '請輸入終止服役日期\n'; if (Item == '') Item = 'FServiceDate'; Page = 1; }
                if (document.form1.FServiceDate.value != '' && !checkDate(document.form1.FServiceDate.value)) { msg += '終止服役日期 格式不正確\n'; if (Item == '') Item = 'EnterDate'; Page = 1; }
            }
            //rblMobil
            if (!isChecked(document.form1.rblMobil)) {
                msg += '請選擇有無行動電話\n'; if (Item == '') Item = 'CellPhone'; Page = 1;
            }
            else {
                if (getRadioValue(document.form1.rblMobil) == "Y") {
                    if (document.form1.CellPhone.value == '') { msg += '有行動電話 請輸入行動電話\n'; if (Item == '') Item = 'CellPhone'; Page = 1; }
                }
                else {
                    if (document.form1.PhoneD.value == '') { msg += '請輸入聯絡電話(日)\n'; if (Item == '') Item = 'PhoneD'; Page = 1; }
                    if (document.form1.CellPhone.value != '') { msg += '有輸入行動電話,請選擇有行動電話\n'; if (Item == '') Item = 'CellPhone'; Page = 1; }
                }
            }
            /*
            if (document.form1.ZipCode2.value=='') {msg+='請輸入戶籍地址(區域)\n';if(Item=='') Item='City2';Page=1;}
            if (document.form1.HouseholdAddress.value=='') {msg+='請輸入戶籍地址\n';if(Item=='') Item='HouseholdAddress';Page=1;}
            //debugger;
            if (document.form1.CheckBox2.checked==false) {
            if (document.form1.ZipCode1.value=='') {msg+='請輸入通訊地址(區域)\n';if(Item=='') Item='City1';Page=1;}
            if (document.form1.Address.value=='') {msg+='請輸入通訊地址\n';if(Item=='') Item='Address';Page=1;}
            }
            */
            //debugger;
            $('#ZipCode1').val($.trim($('#ZipCode1').val()));
            $('#ZipCode1_B3').val($.trim($('#ZipCode1_B3').val()));
            $('#Address').val($.trim($('#Address').val()));
            if ($('#ZipCode1').val() == '') { msg += '請輸入通訊地址前3碼郵遞區號\n'; if (Item == '') Item = 'ZipCode1'; Page = 1; }
            else {
                if ($('#ZipCode1').val().length != 3) { msg += '通訊地址前3碼郵遞區號必須為3碼\n'; if (Item == '') Item = 'ZipCode1'; Page = 1; }
            }

            if ($('#ZipCode1_B3').val() == '') {
                msg += '請輸入通訊地址郵遞區號後3碼\n'; if (Item == '') Item = 'ZipCode1_B3'; Page = 1;
            } else {
                //checkzip23 郵遞區號
                msg += checkzip23(true, '通訊地址', 'ZipCode1_B3');
                //if (!isUnsignedInt($('#ZipCode1_B3').val())) { msg += '通訊地址郵遞區號後3碼/後2碼必須為數字，且不得輸入 00\n'; if (Item == '') Item = 'ZipCode1_B3'; Page = 1; }
                //if (parseInt($('#ZipCode1_B3').val(), 10) < 1) { msg += '通訊地址郵遞區號後3碼/後2碼必須為數字，得輸入 01~999 \n'; if (Item == '') Item = 'ZipCode1_B3'; Page = 1; }
                //if ($('#ZipCode1_B3').val().length != 2) { msg += '通訊地址郵遞區號後3碼/後2碼長度必須為 3碼或2碼(例 01 或 001)\n'; if (Item == '') Item = 'ZipCode1_B3'; Page = 1; }
            }
            if ($('#Address').val() == '') { msg += '請輸入通訊地址\n'; if (Item == '') Item = 'Address'; Page = 1; }

            $('#ZipCode2').val($.trim($('#ZipCode2').val()));
            $('#ZipCode2_B3').val($.trim($('#ZipCode2_B3').val()));
            $('#HouseholdAddress').val($.trim($('#HouseholdAddress').val()));
            $('#Email').val($.trim($('#Email').val()));
            if (document.form1.CheckBox1.checked == false) {
                if ($('#ZipCode2').val() == '') { msg += '請輸入戶籍地址郵遞區號前3碼\n'; if (Item == '') Item = 'ZipCode2'; Page = 1; }
                else {
                    if ($('#ZipCode2').val().length != 3) { msg += '戶籍地址郵遞區號前3碼必須為3碼\n'; if (Item == '') Item = 'ZipCode2'; Page = 1; }
                }
                if ($('#ZipCode2_B3').val() == '') {
                    msg += '請輸入戶籍地址郵遞區號後3碼/後2碼\n'; if (Item == '') Item = 'ZipCode2_B3'; Page = 1;
                } else {
                    //checkzip23 郵遞區號
                    msg += checkzip23(true, '戶籍地址', 'ZipCode2_B3');
                    //if (!isUnsignedInt($('#ZipCode2_B3').val())) { msg += '戶籍地址郵遞區號後3碼/後2碼必須為數字，且不得輸入 00\n'; if (Item == '') Item = 'ZipCode2_B3'; Page = 1; }
                    //if (parseInt($('#ZipCode2_B3').val(), 10) < 1) { msg += '戶籍地址郵遞區號後3碼/後2碼必須為數字，得輸入 01~999 \n'; if (Item == '') Item = 'ZipCode2_B3'; Page = 1; }
                    //if ($('#ZipCode2_B3').val().length != 2) { msg += '戶籍地址郵遞區號後3碼/後2碼長度必須為 3碼或2碼(例 01 或 99)\n'; if (Item == '') Item = 'ZipCode2_B3'; Page = 1; }
                }
                if ($('#HouseholdAddress').val() == '') { msg += '請輸入戶籍地址\n'; if (Item == '') Item = 'HouseholdAddress'; Page = 1; }
            }
            if ($('#Email').val() == '') { msg += '請輸入Email\n'; if (Item == '') Item = 'Email'; Page = 1; }
            if ($('#Email').val() != '' && !checkEmail($('#Email').val()) && $('#Email').val() != '無') { msg += '請輸入正確的E-mail格式\n'; if (Item == '') Item = 'Email'; Page = 1; }

            var j = 0;
            var Identity = getCheckBoxListValue('IdentityID');
            //(排除在職進修)
            if (TPlanID.value != '06') {
                //非在職進修訓練 !06
                if (document.form1.MIdentityID.selectedIndex == 0) {
                    msg += '請選擇主要參訓身分別\n';
                    if (Item == '') Item = 'MIdentityID'; Page = 1;
                }
                else if (Identity.charAt(document.form1.MIdentityID.selectedIndex - 1) != '1') {
                    msg += '主要參訓身分別必須為參訓身分別之一\n';
                    if (Item == '') Item = 'MIdentityID'; Page = 1;
                }
            }
            else {
                //在職進修訓練 06
                if (document.form1.MIdentityID.selectedIndex != 0) {
                    if (Identity.charAt(document.form1.MIdentityID.selectedIndex - 1) != '1') {
                        msg += '主要參訓身分別必須為參訓身分別之一\n';
                        if (Item == '') Item = 'MIdentityID'; Page = 1;
                    }
                }
            }
            if (parseInt(Identity, 10) == 0) {
                msg += '請選擇參訓身分別\n';
            }
            else {
                for (var i = 0; i < Identity.length; i++) {
                    if (Identity.charAt(i) == '1') j++;
                }
                if (j > 5) msg += '參訓身分別最多只能選擇5項\n';
            }
            var cst_pt1 = 0;
            var cst_pt2 = 1;
            if (NrblHandType.length > 2) {
                cst_pt1 = 1; //cst_pt rblHandType
                cst_pt2 = 2;
            }
            if (IrblHandType) {
                if (NrblHandType[cst_pt1].checked) { //新制
                    if (document.form1.HandTypeID.disabled == false) {
                        var gHandTypeID2 = getCheckBoxListValue('HandTypeID2');
                        if (parseInt(gHandTypeID2, 10) == 0) { msg += '請選擇障礙類別2\n'; if (Item == '') Item = 'HandTypeID2'; Page = 1; }
                        if (document.form1.HandLevelID2.selectedIndex == 0) { msg += '請選擇障礙等級2\n'; if (Item == '') Item = 'HandLevelID2'; Page = 1; }
                    }
                }
                if (NrblHandType[cst_pt2].checked) { //舊制
                    if (document.form1.HandTypeID.disabled == false) {
                        if (document.form1.HandTypeID.selectedIndex == 0) { msg += '請選擇障礙類別\n'; if (Item == '') Item = 'HandTypeID'; Page = 1; }
                        if (document.form1.HandLevelID.selectedIndex == 0) { msg += '請選擇障礙等級\n'; if (Item == '') Item = 'HandLevelID'; Page = 1; }
                    }
                }
            }
            var RejectTDate1 = document.getElementById("RejectTDate1");
            var RejectTDate2 = document.getElementById("RejectTDate2");
            var ForeIDNO = document.getElementById("ForeIDNO");
            var SOfficeYM1 = document.getElementById("SOfficeYM1");
            var SOfficeYM2 = document.getElementById("SOfficeYM2");
            var FOfficeYM1 = document.getElementById("FOfficeYM1");
            var FOfficeYM2 = document.getElementById("FOfficeYM2");
            var PriorWorkPay = document.getElementById("PriorWorkPay");
            var ShowDetail = document.form1.ShowDetail;
            //var BudID = document.form1.BudID;
            var rdo_HighEduBg = document.getElementById("rdo_HighEduBg");

            if (RejectTDate1 && RejectTDate1.value != '' && !checkDate(RejectTDate1.value)) { msg += '離訓日期格式不正確\n'; if (Item == '') Item = 'RejectTDate1'; Page = 1; }
            if (RejectTDate2 && RejectTDate2.value != '' && !checkDate(RejectTDate2.value)) { msg += '退訓日期格式不正確\n'; if (Item == '') Item = 'RejectTDate2'; Page = 1; }
            if (ForeIDNO && ForeIDNO.value != '' && !checkId(ForeIDNO.value)) { msg += '國內聯絡人身分證號碼不正確\n'; if (Item == '') Item = 'ForeIDNO'; Page = 1; }
            if (SOfficeYM1 && SOfficeYM1.value != '' && !checkDate(SOfficeYM1.value)) { msg += '加退保日期1 起始日期格式不正確\n'; if (Item == '') Item = 'SOfficeYM1'; Page = 1; }
            if (FOfficeYM1 && FOfficeYM1.value != '' && !checkDate(FOfficeYM1.value)) { msg += '加退保日期1 迄止日期格式不正確\n'; if (Item == '') Item = 'FOfficeYM1'; Page = 1; }
            if (SOfficeYM2 && SOfficeYM2.value != '' && !checkDate(SOfficeYM2.value)) { msg += '加退保日期2 起始日期格式不正確\n'; if (Item == '') Item = 'SOfficeYM2'; Page = 1; }
            if (FOfficeYM2 && FOfficeYM2.value != '' && !checkDate(FOfficeYM2.value)) { msg += '加退保日期2 起止日期格式不正確\n'; if (Item == '') Item = 'FOfficeYM2'; Page = 1; }
            if (PriorWorkPay && PriorWorkPay.value != '' && !isUnsignedInt(PriorWorkPay.value)) { msg += '投保薪資級距 必須為數字\n'; if (Item == '') Item = 'PriorWorkPay'; Page = 1; }
            if (ShowDetail && ShowDetail.selectedIndex == 0) { msg += '請選擇是否提供基本資料供查詢\n'; if (Item == '') Item = 'ShowDetail'; Page = 1; }
            if (BudID && !isChecked(document.form1.BudID)) { msg += '請選擇預算別\n'; if (Item == '') Item = 'BudID'; Page = 1; }

            //if (document.form1.SupplyID.disabled==false)
            //if (document.form1.SupplyID.selectedIndex==0) {msg+='請選擇補助比例\n';if(Item=='') Item='SupplyID';Page=1;}
            if (rdo_HighEduBg) {
                if (rdo_HighEduBg.rows[0].cells[0].children[0].checked) {
                    if (document.getElementById("DegreeID").value != "05" && document.getElementById("DegreeID").value != "06") {
                        if (document.getElementById("DegreeID").value == "03" || document.getElementById("DegreeID").value == "04") {
                            if (document.getElementById("GraduateStatus").value != "01") {
                                msg += "專上畢業學歷失業者至少需要專科或大學以上學歷畢業。\n";
                            }
                        } else {
                            msg += "專上畢業學歷失業者至少需要專科或大學以上學歷畢業。\n";
                        }
                    }
                }
            }
            //是否為在職者補助身分
            var NrblWorkSuppIdent = document.getElementsByName('rblWorkSuppIdent');
            var IrblWorkSuppIdent = document.getElementById("rblWorkSuppIdent");
            if (IrblWorkSuppIdent) {
                var cst_pt1 = 0;
                var cst_pt2 = 1;
                if (NrblWorkSuppIdent.length > 2) {
                    cst_pt1 = 1; //cst_pt rblWorkSuppIdent
                    cst_pt2 = 2;
                }
                if (isEmpty('rblWorkSuppIdent')) {
                    msg += "請選擇「是否為在職者補助身分」\n"; if (Item == '') Item = 'rblWorkSuppIdent'; Page = 1;
                }
                else {
                    if (NrblWorkSuppIdent[cst_pt1].checked == true && getRadioValue(document.form1.JobStateType) == 0) {
                        //document.getElementsByName('rblWorkSuppIdent')[1].checked=false;
                        //document.getElementsByName('rblWorkSuppIdent')[2].checked=true;
                        msg += "就職狀況選擇為失業時，(是否為在職者補助)不可為「是」\n";
                        if (Item == '') Item = 'JobStateType'; Page = 1;
                    }
                }
            }

            if (document.getElementById('ActNo')) {
                if (document.getElementById('ActNo').value == '') { msg += '請輸入保險證號\n'; if (Item == '') Item = 'ActNo'; Page = 1; }
            }

            ///參訓身分別為身心障礙者時-必填欄位
            var GetID = document.getElementById("hide_IdentityID_06").value;
            //if (GetID == "") { return false; }
            //var NrblHandType = document.getElementsByName('rblHandType');
            //var IrblHandType = document.getElementById('rblHandType');
            if (GetID != "" && document.getElementById(GetID).checked) {
                var cst_pt1 = 0;
                var cst_pt2 = 1;
                if (NrblHandType.length > 2) {
                    cst_pt1 = 1; //cst_pt rblHandType
                    cst_pt2 = 2;
                }
                if (NrblHandType[cst_pt1].checked) { //新制
                    var gHandTypeID2 = getCheckBoxListValue('HandTypeID2');
                    if (parseInt(gHandTypeID2, 10) == 0) { msg += '參訓身分別為「身心障礙者」,請勾選障礙類別2\n'; }
                    if (document.form1.HandLevelID2.selectedIndex == 0) { msg += '參訓身分別為「身心障礙者」,障礙等級2不能為「請選擇」\n'; }
                }
                if (NrblHandType[cst_pt2].checked) { //舊制
                    if (document.form1.HandTypeID.selectedIndex == 0) { msg += '參訓身分別為「身心障礙者」,障礙類別不能為「請選擇」\n'; }
                    if (document.form1.HandTypeID.selectedIndex == 1) { msg += '參訓身分別為「身心障礙者」,障礙類別不能為「未填列」\n'; }
                    if (document.form1.HandLevelID.selectedIndex == 0) { msg += '參訓身分別為「身心障礙者」,障礙等級不能為「請選擇」\n'; }
                    if (document.form1.HandLevelID.selectedIndex == 1) { msg += '參訓身分別為「身心障礙者」,障礙等級不能為「未填列」\n'; }
                }
            }

            //企訓專用('產學訓計畫才會顯示 MenuTable / BackTable)
            //if (document.getElementById('BackTable') && document.getElementById('BackTable').style.display != 'none') { }

            /*
             if (document.getElementById('Q61').value != '' && !isUnsignedInt(document.getElementById('Q61').value)) { msg += '個人工作年資必須為數字\n'; if (Item == '') { Item = 'Q61'; Page = 2; } }
             if (document.getElementById('Q62').value != '' && !isUnsignedInt(document.getElementById('Q62').value)) { msg += '在這家公司的年資必須為數字\n'; if (Item == '') { Item = 'Q62'; Page = 2; } }
             if (document.getElementById('Q63').value != '' && !isUnsignedInt(document.getElementById('Q63').value)) { msg += '在這職位的年資必須為數字\n'; if (Item == '') { Item = 'Q63'; Page = 2; } }
             if (document.getElementById('Q64').value != '' && !isUnsignedInt(document.getElementById('Q64').value)) { msg += '最近升遷離本職幾年必須為數字\n'; if (Item == '') { Item = 'Q64'; Page = 2; } }
             */

            //else { msg += '請輸入最近升遷離本職幾年!!!\n'; }
            /*20090205 andy edit 2009年 身分為「非自願離職者」不驗証投保單位相關資料 */
            //if (! (document.getElementById('TPlanID').value=='28'  && parseInt(document.getElementById('hide_Years').value)>2008 && document.form1.MIdentityID.value=='02'))
            //if (!(flagTPlan_28 && parseInt(document.getElementById('hide_Years').value, 10) > 2008 && document.form1.MIdentityID.value == '02')) { }

            if (document.getElementById('ActName').value == '') { msg += '請輸入投保單位名稱\n'; if (Item == '') { Item = 'ActName'; Page = 2; } }
            if (document.getElementById('ActNo1').value == '') { msg += '請輸入投保單位保險證號\n'; if (Item == '') { Item = 'ActNo1'; Page = 2; } }
            //if (!isChecked(document.getElementsByName('ActType'))) { msg += '請選擇投保類別\n'; if (Item == '') { Item = 'ActType'; Page = 2; } }
            //if (document.getElementById('txt_ActPhone').value == '') { msg += '請輸入投保單位電話\n'; if (Item == '') { Item = 'txt_AxtPhone'; Page = 2; } }
            //if (document.getElementById('txt_ActZip').value == '') { msg += '請輸入投保單位地址[地區]\n'; if (Item == '') { Item = 'txt_ActCity'; Page = 2; } }
            //if(trim(document.getElementById('txt_ActZip').value).length!=3) {msg+='投保單位地址[前3碼郵遞區號]必須為3碼\n';if(Item=='') {Item='txt_ActCity';Page=2;}}
            //debugger;
            //20090520 add 投保單位地址郵遞區號後3碼 --begin

            $('#txt_ActZIPB3').val($.trim($('#txt_ActZIPB3').val()));
            if ($('#txt_ActZIPB3').val() != '') {
                //checkzip23 郵遞區號
                msg += checkzip23(true, '投保單位地址', 'txt_ActZIPB3');
                //if (!isUnsignedInt($('#txt_ActZIPB3').val())) { msg += '投保單位地址郵遞區號後3碼/後2碼必須為數字，且不得輸入 00\n'; if (Item == '') { Item = 'txt_ActZIPB3'; Page = 2; } }
                //if (parseInt($('#txt_ActZIPB3').val(), 10) < 1) { msg += '投保單位地址郵遞區號後3碼/後2碼必須為數字，得輸入 01~999 \n'; if (Item == '') Item = 'txt_ActZIPB3'; Page = 2; }
                //if ($('#txt_ActZIPB3').val().length != 2 && $('#ZipCode3_B3').val().length != 3) { msg += '投保單位地址郵遞區號後3碼/後2碼長度必須為 3碼或2碼(例 01 或 001)\n'; if (Item == '') Item = 'txt_ActZIPB3'; Page = 1; }
            }
            //20090520 add 投保單位地址郵遞區號後3碼 --end
            //if (document.getElementById('txt_ActAddress').value == '') { msg += '請輸入投保單位地址\n'; if (Item == '') { Item = 'txt_Address'; Page = 2; } }
            //if (document.getElementById('Tel').value == '') { msg += '請輸入服務單位公司電話\n'; if (Item == '') { Item = 'Tel'; Page = 2; } }
            //同投保單位地址
            //debugger;

            if (document.getElementById('Uname').value == '') { msg += '請輸入目前任職公司名稱\n'; if (Item == '') { Item = 'Uname'; Page = 2; } }
            //if (document.getElementById('JobTitle').value == '') { msg += '請輸入服務單位 職稱\n'; if (Item == '') { Item = 'JobTitle'; Page = 2; } }
            var ServDept = document.getElementById('ServDept');
            if (ServDept && isEmpty('ServDept')) { msg += '請輸入目前任職部門\n'; Page = 2; }
            if (document.getElementById('ddlSERVDEPTID') && isEmpty('ddlSERVDEPTID')) { msg += '請選擇目前任職部門\n'; Page = 2; }
            var JobTitle = document.getElementById('JobTitle');
            if (JobTitle && isEmpty('JobTitle')) { msg += '請輸入服務單位 職務\n'; Page = 2; }
            if (document.getElementById('ddlJOBTITLEID') && isEmpty('ddlJOBTITLEID')) { msg += '請選擇服務單位 職務\n'; Page = 2; }
            // if (document.getElementById('SDate').value != '' && !checkDate(document.getElementById('SDate').value)) { msg += '個人到任目前任職公司起日不是正確的日期格式\n'; if (Item == '') { Item = 'SDate'; Page = 2; } }
            // if (document.getElementById('SJDate').value != '' && !checkDate(document.getElementById('SJDate').value)) { msg += '個人到任目前職務起日不是正確的日期格式\n'; if (Item == '') { Item = 'SJDate'; Page = 2; } }
            // if (document.getElementById('SPDate').value != '' && !checkDate(document.getElementById('SPDate').value)) { msg += '最近升遷日期不是正確的日期格式\n'; if (Item == '') { Item = 'SPDate'; Page = 2; } }

            // if (document.getElementById('Q61').value != '' && !isUnsignedInt(document.getElementById('Q61').value)) { msg += '個人工作年資必須為數字\n'; if (Item == '') { Item = 'Q61'; Page = 2; } }
            // if (document.getElementById('Q62').value != '' && !isUnsignedInt(document.getElementById('Q62').value)) { msg += '在這家公司的年資必須為數字\n'; if (Item == '') { Item = 'Q62'; Page = 2; } }
            // if (document.getElementById('Q63').value != '' && !isUnsignedInt(document.getElementById('Q63').value)) { msg += '在這職位的年資必須為數字\n'; if (Item == '') { Item = 'Q63'; Page = 2; } }
            // if (document.getElementById('Q64').value != '' && !isUnsignedInt(document.getElementById('Q64').value)) { msg += '最近升遷離本職幾年必須為數字\n'; if (Item == '') { Item = 'Q64'; Page = 2; } }

            if (msg != '') {
                ChangeMode(Page);
                try {
                    if (document.getElementById(Item)) { document.getElementById(Item).focus(); }
                }
                catch (e) { }
                alert(msg);
                return false;
            }
            if (PriorWorkOrg1) {
                PriorWorkOrg1.disabled = false;
                PriorWorkOrg2.disabled = false;
                ActNo2.disabled = false;
                PriorWorkPay.disabled = false;
            }
        }

        //MIdentityID:參訓身分別
        function MIdentityChg(IdentValue) {
            //console.log('MIdentityChg,IdentValue: ' + IdentValue); //var IdentityID = document.getElementById("IdentityID");
            let MIdentityID = document.getElementById("MIdentityID");
            //var Length = MIdentityID.options.length;
            let strV = "";
            let mtContent = "";
            for (var i = 1; i < MIdentityID.options.length; i++) {
                if (i == MIdentityID.selectedIndex) {
                    strV += "1";
                    mtContent = MIdentityID[i].innerText;
                    break;
                } else { strV += "2"; }
            }
            //40:經公告之重大災害受災者 hide_MIdentityID
            if (IdentValue == "40") { $('#tr_DDL_DISASTER').show(); } else { $('#tr_DDL_DISASTER').hide(); }
            //console.log('MIdentityChg,mtContent: ' + mtContent); console.log('MIdentityChg,IdentityID: ' + IdentityID); console.log('MIdentityChg,strV: ' + strV);
            setCheckBoxList2('IdentityID', mtContent);
        }

        //只是勾選依textContent
        function setCheckBoxList2(objNM, mtContent) {
            document.querySelectorAll('#' + objNM + ' input[type="checkbox"]').forEach(checkbox => {
                // Modified line: Use strict equality (===) for textContent,,checkbox.nextElementSibling.textContent.includes(mtContent)
                if (checkbox.nextElementSibling && checkbox.nextElementSibling.textContent.trim() === mtContent.trim()) {
                    checkbox.checked = true;
                }
                // Optional: Uncheck other boxes if you want only the selected one to be checked else { checkbox.checked = false; }
            });
        }

        //MilitaryID:兵役
        function sol(nn) {
            var myTR = document.getElementById("SolTR");
            myTR.style.display = 'none';
            if (nn == '04') { myTR.style.display = cst_inline1; }

            if (document.body) { window.scroll(0, document.body.scrollHeight); }
            if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
        }

        function hard() {
            var cst_pt1 = 0;
            var cst_pt2 = 1;
            var NrblHandType = document.getElementsByName('rblHandType');
            var IrblHandType = document.getElementById('rblHandType');
            var trHandTypeID2 = document.getElementById('trHandTypeID2');
            var trHandTypeID = document.getElementById('trHandTypeID');
            if (NrblHandType.length > 2) {
                cst_pt1 = 1; //cst_pt rblHandType
                cst_pt2 = 2;
            }
            //身心障礙者
            var GetID = document.getElementById("hide_IdentityID_06").value;
            //if (GetID == "") { return false; } //非身心障礙者 不使用此功能
            var enableSelection = function () {
                $("#tdHandTypeID2").prop("disabled", false);
                $('input[id^="HandTypeID2_"]').each(function () {
                    $(this).prop("disabled", false);
                });
                $("#HandTypeID").prop("disabled", false);
                $("#HandLevelID").prop("disabled", false);
                $("#HandLevelID2").prop("disabled", false);
                //$("#tdHandTypeID2").removeAttr("disabled");
            };
            var disableSelection = function () {
                $("#tdHandTypeID2").prop("disabled", true);
                $('input[id^="HandTypeID2_"]').each(function () {
                    $(this).prop("disabled", true);
                });
                $("#HandTypeID").prop("disabled", true);
                $("#HandLevelID").prop("disabled", true);
                $("#HandLevelID2").prop("disabled", true);
                //$("#tdHandTypeID2").attr("disabled", "disabled");
            };
            if (IrblHandType) {
                trHandTypeID2.style.display = NrblHandType[cst_pt1].checked ? cst_inline1 : 'none'; //新制
                trHandTypeID.style.display = NrblHandType[cst_pt2].checked ? cst_inline1 : 'none'; //舊制
            }
            else {
                if (trHandTypeID2) trHandTypeID2.style.display = 'none'; //新制
                if (trHandTypeID) trHandTypeID.style.display = 'none'; //舊制
            }

            if (GetID != "" && document.getElementById(GetID).checked) {
                enableSelection();
                //document.form1.HandTypeID.disabled = false;
                //document.form1.HandLevelID.disabled = false;
                //document.form1.HandLevelID2.disabled = false;
            }
            else {
                disableSelection();
                //document.form1.HandTypeID.disabled = true;
                //document.form1.HandLevelID.disabled = true;
                //document.form1.HandLevelID2.disabled = true;
            }
            if (document.body) { window.scroll(0, document.body.scrollHeight); }
            if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }

        }

        function chknum(value) {
            if (value >= 48 && value <= 57) return true;
            else return false;
        }

        function ChangeMode(num) {
            //alert("ChangeMode(num:" + num + ")。"); //debugger;
            var TPlanID = document.getElementById('TPlanID');
            for (var i = 14; i <= 21; i++) {
                var starobj = document.getElementById('star' + i);
                if (starobj) { starobj.style.display = 'none'; }
            }
            var MenuTable_td_1 = $("#MenuTable_td_1");
            var MenuTable_td_2 = $("#MenuTable_td_2");
            //var MenuTable_td_3 = $("#MenuTable_td_3");
            MenuTable_td_1.removeClass();
            MenuTable_td_2.removeClass();
            //MenuTable_td_3.removeClass();
            if (document.getElementById('DetailTable') && (document.getElementById('BackTable') || document.getElementById('UBackTable'))) {
                switch (num) {
                    case 1:
                        MenuTable_td_1.addClass("active");
                        document.getElementById('DetailTable').style.display = cst_inline1;
                        if (document.getElementById('BackTable')) document.getElementById('BackTable').style.display = 'none';
                        document.getElementById('HistoryTable').style.display = 'none';
                        document.getElementById('Table4').style.display = cst_inline1;
                        break;
                    case 2:
                        MenuTable_td_2.addClass("active");
                        document.getElementById('DetailTable').style.display = 'none';
                        if (document.getElementById('BackTable')) document.getElementById('BackTable').style.display = cst_inline1;
                        document.getElementById('HistoryTable').style.display = 'none';
                        document.getElementById('Table4').style.display = cst_inline1;
                        /*20090205 andy edit 2009年 身分為「非自願離職者」不驗証投保單位相關資料 */
                        document.getElementById('hide_MIdentityID').value = getValue("MIdentityID");
                        for (var i = 14; i <= 21; i++) {
                            switch (i) {
                                case 14:
                                case 15:
                                case 18:
                                case 21:
                                    var starobj = document.getElementById('star' + i);
                                    if (starobj) { starobj.style.display = cst_inline1; }
                                    break;
                                default:
                                    break;
                            }
                        }
                        if (TPlanID.value == '28'
                            && parseInt(document.getElementById('hide_Years').value, 10) > 2008
                            && document.getElementById('hide_MIdentityID').value == '02') {
                            for (var i = 14; i <= 21; i++) {
                                switch (i) {
                                    case 14:
                                    case 15:
                                    case 18:
                                    case 21:
                                        var starobj = document.getElementById('star' + i);
                                        if (starobj) { starobj.style.display = 'none'; }
                                        break;
                                    default:
                                        break;
                                }
                            }
                        }
                        break;
                    case 3:
                        //MenuTable_td_3.addClass("active");
                        document.getElementById('DetailTable').style.display = 'none';
                        if (document.getElementById('BackTable')) document.getElementById('BackTable').style.display = 'none';
                        document.getElementById('HistoryTable').style.display = cst_inline1;
                        document.getElementById('Table4').style.display = 'none';
                        break;
                    default:
                        document.getElementById('DetailTable').style.display = 'none';
                        if (document.getElementById('BackTable')) document.getElementById('BackTable').style.display = 'none';
                        document.getElementById('HistoryTable').style.display = 'none';
                        document.getElementById('Table4').style.display = 'none';
                        break;
                }
            }

            //40:經公告之重大災害受災者 hide_MIdentityID
            if ($('#hide_MIdentityID').val() == "40") { $('#tr_DDL_DISASTER').show(); } else { $('#tr_DDL_DISASTER').hide(); }
            //chgPriorWorkType1_disabled();
            if (document.body) { window.scroll(0, document.body.scrollHeight); }
            if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }

        }

        //開啟日曆(回傳欄位,開訓日期,結訓日期,目前指定日期[可略],頁面相對位置[可略],回傳要觸發的按鈕[可略])
        //在點選「出生日期」後，開啟日期的選窗
        function callCalendar(Birthday, ButtonID) {
            openCalendar(Birthday, '1911/01/01', '2099/12/31', Date(), '', ButtonID);
        }

        //足歲45歲則選取中高齡者(45歲)。跟開訓日期比較
        function ChkBirthday() {
            var MyCheckbox = document.getElementById('IdentityID');
            var BirthDate = new Array();
            var Today = new Date();
            var TYY = Today.getYear();
            var TMM = Today.getMonth() + 1;
            var TDD = Today.getDate();
            var YY, MM, DD;
            var STDateHidden = new Array();
            var GetID = document.getElementById("hide_IdentityID_04").value;
            if (GetID == "") { return false; } //非中高齡者 不使用此功能
            if (document.getElementById('Birthday').value != '') {
                BirthDate = document.getElementById('Birthday').value.split("/");
                STDateHidden = document.getElementById('STDateHidden').value.split("/");
                YY = STDateHidden[0] - BirthDate[0];
                MM = STDateHidden[1] - BirthDate[1];
                DD = STDateHidden[2] - BirthDate[2];
            }
            //中高齡者
            if (YY >= 46 && YY <= 64) { document.getElementById(GetID).checked = true; }
            if (YY == 45) {
                if (MM > 0) { document.getElementById(GetID).checked = true; }
                if (MM == 0 && DD >= 0) { document.getElementById(GetID).checked = true; }
            }
            if (YY == 65) {
                if (MM < 0) { document.getElementById(GetID).checked = true; }
                if (MM == 0 && DD < 0) { document.getElementById(GetID).checked = true; }
            }
        }

        //檢查身分別為本國時，由身分證號第2位帶入性別
        function chkidnosex() {
            if (document.form1.PassPortNO[0].checked) {
                if (document.form1.IDNO.value.charAt(1) == 1) document.form1.Sex_0.checked = true;
                else if (document.form1.IDNO.value.charAt(1) == 2) document.form1.Sex_1.checked = true;
            }
        }

        function ChkMIdentityID() {
            var TPlanID = document.getElementById('TPlanID');
            //檢查主要參訓身分別為「身心障礙者」，參訓身分別「身心障礙者」自動勾選
            if (document.form1.MIdentityID.value == '06') {
                var GetID = document.getElementById("hide_IdentityID_06").value;
                if (GetID != "") { document.getElementById(GetID).checked = true; }
                //if (GetID == "") { return false; }
                //document.getElementById(GetID).checked = true;
            }
            //if (TPlanID.value == '28') {
            //    //產業人才投資方案 //一般身分
            //    if (document.form1.MIdentityID.value == '01') {
            //        document.form1.SupplyID.value = '1'; //補助80%
            //    }
            //}

            hard(); //身心障礙者 顯示
            document.form1.hide_MIdentityID.value = document.form1.MIdentityID.value;
        }

        //var s_tmpBudID = "none"; //初始值
        var i_tmpBudID = -1; //初始值
        function Change_BudID() {
            //alert("Change_BudID。"); //debugger;
            var rdo_HighEduBg = document.getElementById("rdo_HighEduBg");
            var BudID = document.getElementById("BudID");
            if (BudID && i_tmpBudID == -1) {
                for (i = 0; i < BudID.childNodes.length; i++) {
                    if (BudID.childNodes.item(i).type == "radio") {
                        //取得 BudID
                        if (BudID.childNodes.item(i).checked) { i_tmpBudID = i; }
                    }
                }
            }

            if (BudID && rdo_HighEduBg) {
                if (rdo_HighEduBg.rows[0].cells[0].children[0].checked) {
                    for (i = 0; i < BudID.childNodes.length; i++) {
                        if (BudID.childNodes.item(i).nodeName == "INPUT") {
                            BudID.childNodes.item(i).checked = false;
                            if (BudID.childNodes.item(i).value == "98") {
                                BudID.childNodes.item(i).style.display = cst_inline1;
                                BudID.childNodes.item(i + 1).style.display = cst_inline1;
                                BudID.childNodes.item(i).checked = true;
                            }
                        }
                    }
                    BudID.disabled = true;
                } else {
                    //for (i = 0; i < BudID.childNodes.item.length; i++) { }
                    for (i = 0; i < BudID.childNodes.length; i++) {
                        if (BudID.childNodes.item(i).nodeName == "INPUT") {
                            if (i == i_tmpBudID) {
                                BudID.childNodes.item(i).checked = true;
                            } else {
                                BudID.childNodes.item(i).checked = false;
                            }
                            if (BudID.childNodes.item(i).value == "98") {
                                BudID.childNodes.item(i).style.display = "none";
                                BudID.childNodes.item(i + 1).style.display = "none";
                                BudID.childNodes.item(i).checked = false;
                            }
                        }
                        BudID.disabled = false;
                    }
                }
            }
        }

        Date.prototype.MydateDiff = function (interval, objDate) {
            var dtEnd = new Date(objDate);
            if (isNaN(dtEnd)) return undefined;
            switch (interval) {
                case "s": return parseInt((dtEnd - this) / 1000, 10);  //秒
                case "n": return parseInt((dtEnd - this) / 60000, 10); //分 
                case "h": return parseInt((dtEnd - this) / 3600000, 10); //時
                case "d": return parseInt((dtEnd - this) / 86400000, 10); //日
                case "w": return parseInt((dtEnd - this) / (86400000 * 7), 10); //週
                case "m": return (dtEnd.getMonth() + 1) + ((dtEnd.getFullYear() - this.getFullYear()) * 12) - (this.getMonth() + 1); //月
                case "y": return dtEnd.getFullYear() - this.getFullYear(); //年
            }
        }


        //中翻英
        function sentName() {
            var msg = '';
            var Name = document.getElementById('Name');
            var strName = Name.value;
            if (isBlank(Name)) {
                msg = '請先輸入中文姓名!';
            } else if (strName.indexOf(" ") >= 0 || strName.indexOf("　") >= 0) {
                msg = '中文姓名不可輸入空格!';
            } else if (strName.length > 8) {
                msg = '中文姓名不可超過8個字!';
            }
            if (msg != '') {
                alert(msg);
            } else {
                window.open('../../common/Translation.aspx?&sn=stud&name=' + escape(strName) + '&field=LName,FName', "sch", 'width=750,height=600,top=200,left=450,location=0,status=0,menubar=0,scrollbars=1,resizable=0,scrollbars=0');
            }
            return false;
        }

        //檢查zipcode(City欄位名,Zip obj,Zip輸入內容)
        function getZipName(CityID, oZipID, ZipValue) {
            //debugger;
            var ua = window.navigator.userAgent;
            var msie = ua.indexOf("MSIE");
            var ifrmae = document.getElementById('ifmChceckZip');
            if (!isBlank(oZipID)) {
                if (isUnsignedInt(ZipValue) && ZipValue.length == 3) {
                    if (msie > 0) {
                        ifrmae.document.form1.hidCityID.value = CityID;
                        ifrmae.document.form1.hidZipID.value = oZipID.id;
                        ifrmae.document.form1.hidValue.value = ZipValue;
                        ifrmae.document.form1.submit();
                    }
                    else {
                        var ifrmaeDoc = (ifrmae.contentWindow || ifrmae.contentDocument);
                        ifrmae.contentDocument.getElementById("hidCityID").value = CityID;
                        ifrmae.contentDocument.getElementById("hidZipID").value = oZipID.id;
                        ifrmae.contentDocument.getElementById("hidValue").value = ZipValue;
                        if (ifrmaeDoc.document) ifrmaeDoc = ifrmaeDoc.document;
                        ifrmaeDoc.getElementById("form1").submit(); // ## error 
                        //ifrmae.contentWindow.formSubmit();
                    }
                } else {
                    oZipID.value = '';
                    document.getElementById(CityID).value = '';
                    oZipID.focus();
                    blockAlert('查無' + ZipValue + '郵遞區號!');
                }
            } else {
                document.getElementById(CityID).value = '';
            }
        }

        //同通訊地址-戶籍地址
        function ock_CheckBox1() {
            var CheckBox1 = document.getElementById('CheckBox1');
            if (!CheckBox1) return;
            if (CheckBox1.checked == true) {
                //if ($('#ZipCode2').val() == "") $('#ZipCode2').val($('#ZipCode1').val());
                //if ($('#ZipCode2_B3').val() == "") $('#ZipCode2_B3').val($('#ZipCode1_B3').val());
                //if ($('#hidZipCode2_6W').val() == "") $('#hidZipCode2_6W').val($('#hidZipCode1_6W').val());
                //if ($('#ZipCode2_N').val() == "") $('#ZipCode2_N').val($('#ZipCode1_N').val());
                //if ($('#City2').val() == "") $('#City2').val($('#City1').val());
                //if ($('#HouseholdAddress').val() == "") $('#HouseholdAddress').val($('#Address').val());
                $('#ZipCode2').val($('#ZipCode1').val());
                $('#ZipCode2_B3').val($('#ZipCode1_B3').val());
                $('#hidZipCode2_6W').val($('#hidZipCode1_6W').val());
                $('#ZipCode2_N').val($('#ZipCode1_N').val());
                $('#City2').val($('#City1').val());
                $('#HouseholdAddress').val($('#Address').val());
            }
        }

        //'同通訊地址-緊急通知人
        function ock_CheckBox2() {
            var CheckBox2 = document.getElementById('CheckBox2');
            var CheckBox3 = document.getElementById('CheckBox3');
            if (!CheckBox2 || !CheckBox3) { return; }
            if (CheckBox2.checked) {
                CheckBox3.checked = false;
                $('#ZipCode3').val($('#ZipCode1').val());
                $('#ZipCode3_B3').val($('#ZipCode1_B3').val());
                $('#hidZipCode3_6W').val($('#hidZipCode1_6W').val());
                $('#ZipCode3_N').val($('#ZipCode1_N').val());
                $('#City3').val($('#City1').val());
                $('#EmergencyAddress').val($('#Address').val());
            }
        }

        //'同戶籍地址-緊急通知人
        function ock_CheckBox3() {
            var CheckBox2 = document.getElementById('CheckBox2');
            var CheckBox3 = document.getElementById('CheckBox3');
            if (!CheckBox2 || !CheckBox3) { return; }
            if (CheckBox3.checked) {
                CheckBox2.checked = false;
                $('#ZipCode3').val($('#ZipCode2').val());
                $('#ZipCode3_B3').val($('#ZipCode2_B3').val());
                $('#hidZipCode3_6W').val($('#hidZipCode2_6W').val());
                $('#ZipCode3_N').val($('#ZipCode2_N').val());
                $('#City3').val($('#City2').val());
                $('#EmergencyAddress').val($('#HouseholdAddress').val());
            }
        }
    </script>
</head>
<body onload="Change_BudID();">
    <form id="form1" method="post" runat="server">
        <%--<table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="Titlelab2" runat="server">
					首頁&gt;&gt;學員動態管理&gt;&gt;學員資料管理&gt;&gt;<font color="#990000">學員資料維護</font></asp:Label>
                </td>
            </tr>
        </table>--%>
        <%--<asp:Button ID="btnTest1" runat="server" Text="btnTest1" />--%>
        <table border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <table class="font" id="MenuTable" cellspacing="0" cellpadding="0" width="50%" runat="server">
                        <tr class="newlink newlink-blue">
                            <td onclick="ChangeMode(1);" id="MenuTable_td_1">個人基本資料 </td>
                            <td onclick="ChangeMode(2);" id="MenuTable_td_2">參訓背景 </td>
                            <%--<td onclick="ChangeMode(3);" id="MenuTable_td_3">補助費用歷史頁 </td>--%>
                            <%-- <td onclick="ChangeMode(1);" background="../../images/bookmark_01.gif" width="1"></td>
                            <td onclick="ChangeMode(1);" background="../../images/bookmark_02.gif" width="100" align="center">個人基本資料 </td>
                            <td onclick="ChangeMode(1);" background="../../images/bookmark_03.gif" width="11"></td>
                            <td id="ChangeMode2a" onclick="ChangeMode(2);" background="../../images/bookmark_01.gif" width="1"></td>
                            <td id="ChangeMode2b" onclick="ChangeMode(2);" background="../../images/bookmark_02.gif" width="100" align="center"><font style="color: #009900;">參訓背景</font> </td>
                            <td id="ChangeMode2c" onclick="ChangeMode(2);" background="../../images/bookmark_03.gif" width="11"></td>
                            <td onclick="ChangeMode(3);" background="../../images/bookmark_01.gif" width="1"></td>
                            <td onclick="ChangeMode(3);" background="../../images/bookmark_02.gif" width="100" align="center"><font style="color: #0000FF;">補助費用歷史頁</font> </td>
                            <td onclick="ChangeMode(3);" background="../../images/bookmark_03.gif" width="11"></td>--%>
                        </tr>
                    </table>
                    <table id="DetailTable" class="table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr id="StdTr" runat="server">
                            <td class="bluecol_need" width="20%">學員 </td>
                            <td class="whitecol" width="30%">
                                <asp:DropDownList ID="SOCID" runat="server" CssClass="font11" AutoPostBack="true"></asp:DropDownList>
                                <asp:Label ID="labmakesocid" runat="server"></asp:Label>
                                <asp:Label ID="LabErrMsg" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                        <tr id="StdTr2" runat="server">
                            <td class="bluecol" width="20%">被遞補者 </td>
                            <td class="whitecol" width="30%">
                                <asp:DropDownList ID="RejectSOCID" runat="server" CssClass="font11"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">班別名稱 </td>
                            <td class="whitecol" width="30%">
                                <asp:Label ID="ClassName" runat="server"></asp:Label></td>
                            <td class="bluecol_need" width="20%">報名階段 </td>
                            <td class="whitecol" width="30%">
                                <asp:DropDownList ID="LevelNo" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">中文姓名 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="Name" runat="server" CssClass="font11" Columns="45" Width="80%"></asp:TextBox>
                                <input id="btnSchEng" onclick="sentName();" value="英譯" type="button" class="asp_button_M" />
                            </td>
                            <td class="bluecol_need">學號(三碼) </td>
                            <td class="whitecol">
                                <asp:TextBox ID="StudentID" runat="server" Columns="3" MaxLength="3" Width="30%"></asp:TextBox>
                                <input id="StudentIDValue" type="hidden" name="StudentIDValue" runat="server" />
                                <input id="StudentIDstring" type="hidden" name="StudentIDstring" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">英文姓名last name(姓)<asp:Label ID="star1" runat="server"></asp:Label></td>
                            <td class="whitecol">
                                <asp:TextBox ID="LName" runat="server" AutoPostBack="false" Width="80%"></asp:TextBox></td>
                            <td class="bluecol">first name(名) </td>
                            <td class="whitecol">
                                <asp:TextBox ID="FName" runat="server" AutoPostBack="false" Width="80%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">原住民羅馬拼音 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="RMPNAME" runat="server" CssClass="font11" Columns="45" Width="80%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">身分別 </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="PassPortNO" runat="server" CssClass="font" Width="100%" RepeatDirection="horizontal">
                                    <asp:ListItem Value="1">本國</asp:ListItem>
                                    <asp:ListItem Value="2">外籍(含大陸人士)</asp:ListItem>
                                </asp:RadioButtonList>
                                <table style="border-collapse: collapse" id="ChinaOrNotTable" class="font" border="1" cellspacing="0" bordercolor="darkseagreen" cellpadding="0" width="100%" runat="server">
                                    <tr>
                                        <td>
                                            <asp:RadioButtonList ID="ChinaOrNot" runat="server" CssClass="font" RepeatDirection="horizontal" RepeatLayout="flow" CellPadding="0" CellSpacing="0">
                                                <asp:ListItem Value="1">大陸人士</asp:ListItem>
                                                <asp:ListItem Value="2">非大陸人士</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:TextBox ID="Nationality" runat="server" Width="60%"></asp:TextBox></td>
                                    </tr>
                                </table>
                            </td>
                            <td class="bluecol_need">身分證號碼 </td>
                            <td class="whitecol">
                                <table id="Table5" border="0" cellspacing="1" cellpadding="1" width="100%">
                                    <tr>
                                        <td>
                                            <asp:RadioButtonList ID="PPNO" runat="server" CssClass="font" CellPadding="0" CellSpacing="0" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">護照號碼</asp:ListItem>
                                                <asp:ListItem Value="2">居留(工作)證號</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:TextBox ID="txtShowIDNO" runat="server" onfocus="this.blur()" Columns="15" Width="60%"></asp:TextBox>
                                            <asp:TextBox ID="IDNO" runat="server" Columns="15" Style="display: none"></asp:TextBox>
                                            <asp:Button ID="Button4" runat="server" Text="檢查" ToolTip="依身分證號檢查資料" CssClass="asp_button_M"></asp:Button>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">性 別 </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="Sex" runat="server" CssClass="font" RepeatDirection="horizontal">
                                    <asp:ListItem Value="M">男</asp:ListItem>
                                    <asp:ListItem Value="F">女</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                            <td class="bluecol_need">出生日期 </td>
                            <td class="whitecol">
                                <%--<asp:TextBox ID="txtShowBirthday" runat="server" onfocus="this.blur()" Width="75px"></asp:TextBox>--%>
                                <%--<asp:TextBox ID="Birthday" runat="server" Width="75px" Style="display: none" ></asp:TextBox>--%>
                                <asp:TextBox ID="Birthday" runat="server" Width="50%" onfocus="this.blur()"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" id="Img1" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30" runat="server" /></span>
                                <input id="hidBirthBtn" type="button" name="hidBirthBtn" runat="server" style="display: none" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">婚姻狀況 </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="MaritalStatus" runat="server" CssClass="font" RepeatDirection="horizontal">
                                    <asp:ListItem Value="1">已婚</asp:ListItem>
                                    <asp:ListItem Value="2">未婚</asp:ListItem>
                                    <asp:ListItem Value="3">暫不提供</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                            <td class="bluecol_need">報名管道 </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="EnterChannel" runat="server">
                                    <asp:ListItem Value="===請選擇===" Selected="true">===請選擇===</asp:ListItem>
                                    <asp:ListItem Value="1">網路</asp:ListItem>
                                    <asp:ListItem Value="2">現場</asp:ListItem>
                                    <asp:ListItem Value="3">通訊</asp:ListItem>
                                    <asp:ListItem Value="4">推介</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr id="TRNDTR" runat="server">
                            <td class="bluecol">推介種類 </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="TRNDMode" runat="server"></asp:DropDownList></td>
                            <%--
                                <td class="bluecol">券別 </td>
							    <td class="whitecol">
								    <asp:RadioButtonList ID="TRNDType" runat="server" CssClass="font" RepeatDirection="horizontal" RepeatLayout="flow">
									    <asp:ListItem Value="1">甲式</asp:ListItem>
									    <asp:ListItem Value="2">乙式</asp:ListItem>
								    </asp:RadioButtonList>
							    </td>
                            --%>
                        </tr>
                        <%--
                        <tr id="DGTR" runat="server">
							<td class="bluecol" width="100">學習券<br />身分別 </td>
							<td class="whitecol" colspan="3"><asp:Label ID="DGIdentValue" runat="server"></asp:Label></td>
						</tr>
                        --%>
                        <tr id="GovTR" runat="server">
                            <td class="bluecol">推介單<br />
                                個案區分 </td>
                            <td class="whitecol">
                                <asp:Label ID="GovSpecial_Type" runat="server"></asp:Label></td>
                            <td class="bluecol_need">推介單身分別 </td>
                            <td class="whitecol">
                                <asp:Label ID="GovObject_Type" runat="server"></asp:Label></td>
                        </tr>
                        <tr id="CDateTR" runat="server">
                            <td class="bluecol">開訓日期 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="OpenDate" runat="server" Width="40%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= OpenDate.ClientID %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30" /></span>
                            </td>
                            <td class="bluecol">結訓日期 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="CloseDate" runat="server" Width="40%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= CloseDate.ClientID %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30" /></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">報到日期 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="EnterDate" runat="server" Width="40%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= EnterDate.ClientID %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30" /></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">最高學歷 </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="DegreeID" runat="server"></asp:DropDownList></td>
                            <td class="bluecol_need">學校名稱<asp:Label ID="star9" runat="server"></asp:Label></td>
                            <td class="whitecol">
                                <asp:TextBox ID="School" runat="server" Width="70%">(未填寫)</asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">科 系<asp:Label ID="star8" runat="server"></asp:Label></td>
                            <td class="whitecol">
                                <asp:TextBox ID="Department" runat="server" Width="70%">(未填寫)</asp:TextBox></td>
                            <td class="bluecol_need">畢業狀況<br />
                                &nbsp;<font id="fontGraduateY" runat="server">畢業年份</font>
                                <asp:Label ID="star11" runat="server"></asp:Label>
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="GraduateStatus" runat="server"></asp:DropDownList>
                                &nbsp;<asp:DropDownList ID="graduatey" runat="server"></asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">兵役狀況<%--<asp:Label ID="star10" runat="server"></asp:Label>--%></td>
                            <td class="whitecol">
                                <asp:DropDownList ID="MilitaryID" runat="server"></asp:DropDownList></td>
                            <td id="jobstatetd" class="bluecol_need" runat="server">就職狀況<asp:Label ID="star7" runat="server"></asp:Label></td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="JobStateType" runat="server" CssClass="font" RepeatDirection="horizontal" RepeatLayout="flow">
                                    <asp:ListItem Value="1">在職</asp:ListItem>
                                    <asp:ListItem Value="0">失業</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="SolTR" runat="server">
                            <td colspan="4" class="whitecol" width="100%">
                                <table id="SoldierTable" class="table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                    <tr>
                                        <td class="bluecol_need" width="20%">軍種 </td>
                                        <td class="whitecol" width="30%">
                                            <asp:TextBox ID="ServiceID" runat="server" Width="60%"></asp:TextBox></td>
                                        <td class="bluecol" width="20%">職務(兵役) </td>
                                        <td class="whitecol" width="30%">
                                            <asp:TextBox ID="MilitaryAppointment" runat="server" Width="60%"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need">階級 </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="MilitaryRank" runat="server" Width="60%"></asp:TextBox></td>
                                        <td class="bluecol_need">服務單位名稱 </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="ServiceOrg" runat="server" Width="80%"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">主管階級姓名 </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="ChiefRankName" runat="server" Width="60%"></asp:TextBox></td>
                                        <td class="bluecol_need">單位電話 </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="ServicePhone" runat="server" Width="80%"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need" width="20%">服役日期 </td>
                                        <td class="whitecol" colspan="3" width="80%">
                                            <span runat="server">
                                                <asp:TextBox ID="SServiceDate" runat="server" Width="16%"></asp:TextBox>
                                                <span runat="server">
                                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= SServiceDate.ClientID %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30" /></span> ～
											    <asp:TextBox ID="FServiceDate" runat="server" Width="16%"></asp:TextBox>
                                                <span runat="server">
                                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= FServiceDate.ClientID %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30" /></span>
                                            </span>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="18%">服役單位地址 </td>
                                        <td class="whitecol" colspan="3" width="82%">
                                            <input id="ZipCode4" maxlength="3" runat="server" />－
                                            <input id="ZipCode4_B3" maxlength="3" runat="server" />
                                            <input id="hidZipCode4_6W" type="hidden" runat="server" />
                                            <asp:Literal ID="LitZipCode4" runat="server"></asp:Literal><%--郵遞區號查詢--%>
                                            <input id="hidCityName4" type="hidden" runat="server" />
                                            <input id="hidAREA4" type="hidden" runat="server" />
                                            <input id="ZipCode4_N" type="hidden" runat="server" /><br />
                                            <asp:TextBox ID="City4" runat="server" Width="20%" onfocus="this.blur()"></asp:TextBox>
                                            <input id="bt_openZip1" value="..." type="button" name="bt_openZip7" runat="server" class="asp_button_Mini" />
                                            <asp:TextBox ID="ServiceAddress" runat="server" Width="60%"></asp:TextBox>
                                            <asp:HiddenField ID="Hid_JnZipCode4" runat="server" />
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="20%">聯絡電話 </td>
                            <td class="whitecol" width="30%">
                                <table id="table7" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                                    <tr>
                                        <td width="15%">(日) </td>
                                        <td width="85%">
                                            <asp:TextBox ID="PhoneD" runat="server" Columns="13" Width="80%"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td width="15%">(夜) </td>
                                        <td width="85%">
                                            <asp:TextBox ID="PhoneN" runat="server" Columns="13" Width="80%"></asp:TextBox></td>
                                    </tr>
                                </table>
                            </td>
                            <td class="bluecol_need" width="18%">行動電話 </td>
                            <td class="whitecol" width="32%">
                                <asp:RadioButtonList ID="rblMobil" runat="server" CssClass="font" RepeatDirection="horizontal" RepeatLayout="flow">
                                    <asp:ListItem Value="N">無</asp:ListItem>
                                    <asp:ListItem Value="Y">有</asp:ListItem>
                                </asp:RadioButtonList>
                                <asp:TextBox ID="CellPhone" runat="server" Columns="13" Width="80%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="18%">通訊地址 </td>
                            <td class="whitecol" colspan="3" width="82%">
                                <input id="ZipCode1" maxlength="3" runat="server" />－
								<input id="ZipCode1_B3" maxlength="3" runat="server" />
                                <input id="hidZipCode1_6W" type="hidden" runat="server" />
                                <asp:Literal ID="LitZipCode1" runat="server"></asp:Literal><%--郵遞區號查詢--%>
                                <input id="hidCityName1" type="hidden" runat="server" />
                                <input id="hidAREA1" type="hidden" runat="server" />
                                <input id="ZipCode1_N" type="hidden" runat="server" /><br />
                                <asp:TextBox ID="City1" runat="server" Width="20%" onfocus="this.blur()"></asp:TextBox>
                                <input id="bt_openZip2" value="..." type="button" runat="server" class="asp_button_Mini" />
                                <asp:TextBox ID="Address" runat="server" Width="60%"></asp:TextBox>
                                <asp:HiddenField ID="Hid_JnZipCode1" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="18%">戶籍地址 </td>
                            <td class="whitecol" colspan="3" width="82%">
                                <asp:CheckBox ID="CheckBox1" runat="server" CssClass="font" Text="同通訊地址"></asp:CheckBox><br />
                                <input id="ZipCode2" maxlength="3" runat="server" />－
                                <input id="ZipCode2_B3" maxlength="3" runat="server" />
                                <input id="hidZipCode2_6W" type="hidden" runat="server" />
                                <asp:Literal ID="LitZipCode2" runat="server"></asp:Literal><%--郵遞區號查詢--%>
                                <input id="hidCityName2" type="hidden" runat="server" />
                                <input id="hidAREA2" type="hidden" runat="server" />
                                <input id="ZipCode2_N" type="hidden" runat="server" /><br />
                                <asp:TextBox ID="City2" runat="server" Width="20%" onfocus="this.blur()"></asp:TextBox>
                                <input id="bt_openZip3" value="..." type="button" name="bt_openZip2" runat="server" class="asp_button_Mini" />
                                <asp:TextBox ID="HouseholdAddress" runat="server" Width="60%"></asp:TextBox>
                                <asp:HiddenField ID="Hid_JnZipCode2" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="18%">電子郵件 </td>
                            <td class="whitecol" colspan="3" width="82%">
                                <asp:TextBox ID="Email" runat="server" Columns="55" Width="70%"></asp:TextBox>(如沒有請填寫"無") </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="18%">主要參訓身分別<asp:Label ID="StarMIdentityID" runat="server"></asp:Label></td>
                            <td class="whitecol" colspan="3" width="82%">
                                <asp:DropDownList ID="MIdentityID" runat="server"></asp:DropDownList>
                                <asp:Label ID="labIdentity" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                        <tr id="tr_DDL_DISASTER" runat="server" style="display: none;">
                            <td class="bluecol_need">重大災害選項</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="DDL_DISASTER" runat="server"></asp:DropDownList>
                            </td>
                        </tr>
                        <%--<tr id="trSubsidyID" runat="server"><td class="bluecol" width="18%">津貼類別 <font color="red">
                           <asp:Label ID="SubsidyLabel" runat="server">*</asp:Label></font> </td><td class="whitecol" colspan="3" width="82%">
                           <asp:DropDownList ID="SubsidyID" runat="server"></asp:DropDownList>
                           <input id="SubsidyHidden" type="hidden" name="hidden" runat="server" /></td></tr>
                           <tr id="trSubsidyIdentity" runat="server"><td class="bluecol" width="18%">津貼<br />身分別
                           <asp:Label ID="LabSubsidy" runat="server" ForeColor="red">*</asp:Label></td><td class="whitecol" colspan="3" width="82%">
                           <asp:DropDownList ID="SubsidyIdentity" runat="server"></asp:DropDownList></td></tr>--%>
                        <%--<tr id="NativeTr1" runat="server"><td class="bluecol_need" width="100">民族別 </td><td class="whitecol" colspan="3">
                           <asp:DropDownList ID="NativeID" runat="server"></asp:DropDownList></td></tr>--%>
                        <tr>
                            <td class="bluecol_need" width="18%">參訓身分別<br />
                                <font class="font">(可複選，最多五項)</font></td>
                            <td class="whitecol" colspan="3" width="82%">
                                <asp:CheckBoxList ID="IdentityID" runat="server" CssClass="font" RepeatDirection="horizontal" RepeatColumns="3"></asp:CheckBoxList>
                                <input id="hide_IdentityID_06" type="hidden" name="hide_IdentityID_06" runat="server" />
                                <input id="hide_IdentityID_04" type="hidden" name="hide_IdentityID_07" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="18%">障礙種類 </td>
                            <td class="whitecol" colspan="3" width="82%">
                                <asp:RadioButtonList ID="rblHandType" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Selected="True" Value="2">新制</asp:ListItem>
                                    <asp:ListItem Value="1">舊制</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="trHandTypeID" runat="server">
                            <td class="bluecol" width="18%">障礙類別 </td>
                            <td class="whitecol" width="32%">
                                <asp:DropDownList ID="HandTypeID" runat="server"></asp:DropDownList></td>
                            <td class="bluecol_need" width="18%">障礙等級 </td>
                            <td class="whitecol" width="32%">
                                <asp:DropDownList ID="HandLevelID" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr id="trHandTypeID2" runat="server">
                            <td class="bluecol" width="18%">障礙類別2 </td>
                            <td class="whitecol" id="tdHandTypeID2" runat="server" width="32%">
                                <asp:CheckBoxList ID="HandTypeID2" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" RepeatColumns="1"></asp:CheckBoxList></td>
                            <td class="bluecol_need" width="18%">障礙等級2 </td>
                            <td class="whitecol" width="32%">
                                <asp:DropDownList ID="HandLevelID2" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="18%">離訓日期 </td>
                            <td class="whitecol" width="32%">
                                <asp:TextBox ID="RejectTDate1" runat="server" Width="40%" onfocus="this.blur()"></asp:TextBox></td>
                            <td class="bluecol" width="18%">退訓日期 </td>
                            <td class="whitecol" width="32%">
                                <asp:TextBox ID="RejectTDate2" runat="server" Width="40%" onfocus="this.blur()"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="18%">緊急通知人<br />
                                姓名<asp:Label ID="star2" runat="server"><font color="red">*</font></asp:Label></td>
                            <td class="whitecol" width="32%">
                                <asp:TextBox ID="EmergencyContact" runat="server" Width="60%"></asp:TextBox></td>
                            <td class="bluecol" width="18%">緊急通知人<br />
                                電話<asp:Label ID="star3" runat="server"><font color="#ff0000">*</font></asp:Label></td>
                            <td class="whitecol" width="32%">
                                <asp:TextBox ID="EmergencyPhone" runat="server" MaxLength="20" Width="60%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="18%">緊急通知人<br />
                                關係<asp:Label ID="star4" runat="server"><font color="#ff0000">*</font></asp:Label></td>
                            <td class="whitecol" colspan="3" width="82%">
                                <asp:TextBox ID="EmergencyRelation" runat="server" Width="16%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="18%">緊急通知人<br />
                                地址<asp:Label ID="star5" runat="server"><font color="#ff0000">*</font></asp:Label></td>
                            <td class="whitecol" colspan="3" width="82%">
                                <asp:CheckBox ID="CheckBox2" runat="server" Text="同通訊地址"></asp:CheckBox>
                                <asp:CheckBox ID="CheckBox3" runat="server" Text="同戶籍地址"></asp:CheckBox><br />
                                <input id="ZipCode3" maxlength="3" runat="server" />－
								<input id="ZipCode3_B3" maxlength="3" runat="server" />
                                <input id="hidZipCode3_6W" type="hidden" runat="server" />
                                <asp:Literal ID="LitZipCode3" runat="server"></asp:Literal><%--郵遞區號查詢--%>
                                <input id="hidCityName3" type="hidden" runat="server" />
                                <input id="hidAREA3" type="hidden" runat="server" />
                                <input id="ZipCode3_N" type="hidden" runat="server" /><br />
                                <asp:TextBox ID="City3" runat="server" Width="20%" onfocus="this.blur()"></asp:TextBox>
                                <input id="bt_openZip4" value="..." type="button" name="bt_openZip3" runat="server" class="asp_button_Mini" />
                                <asp:TextBox ID="EmergencyAddress" runat="server" Width="60%"></asp:TextBox>
                                <asp:HiddenField ID="Hid_JnZipCode3" runat="server" />
                            </td>
                        </tr>
                        <tr id="ForeTr1" runat="server">
                            <td class="bluecol" colspan="4" align="center" width="100%">國內親屬資料 </td>
                        </tr>
                        <tr id="ForeTr2" runat="server">
                            <td class="bluecol" width="18%">姓名 </td>
                            <td class="whitecol" width="32%">
                                <asp:TextBox ID="ForeName" runat="server" Width="60%"></asp:TextBox></td>
                            <td class="bluecol" width="18%">稱謂 </td>
                            <td class="whitecol" width="32%">
                                <asp:TextBox ID="ForeTitle" runat="server" Columns="15" Width="30%"></asp:TextBox></td>
                        </tr>
                        <tr id="ForeTr3" runat="server">
                            <td class="bluecol" width="18%">性別 </td>
                            <td class="whitecol" width="32%">
                                <asp:RadioButtonList ID="ForeSex" runat="server" CssClass="font" RepeatDirection="horizontal">
                                    <asp:ListItem Value="M">男</asp:ListItem>
                                    <asp:ListItem Value="F">女</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                            <td class="bluecol" width="18%">出生日期 </td>
                            <td class="whitecol" width="32%">
                                <asp:TextBox ID="ForeBirth" runat="server" Width="36%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= ForeBirth.ClientID %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30" /></span>
                            </td>
                        </tr>
                        <tr id="ForeTr4" runat="server">
                            <td class="bluecol" width="18%">&nbsp; 身分證號碼 </td>
                            <td class="whitecol" colspan="3" width="82%">
                                <asp:TextBox ID="ForeIDNO" runat="server" Width="26%"></asp:TextBox></td>
                        </tr>
                        <tr id="ForeTr5" runat="server">
                            <td class="bluecol" width="18%">&nbsp; 戶籍地址 </td>
                            <td class="whitecol" colspan="3" width="82%">
                                <input id="ForeZip" maxlength="3" runat="server" />－
								<input id="ForeZIPB3" maxlength="3" runat="server" />
                                <input id="hidForeZIP6W" type="hidden" runat="server" />
                                <asp:Literal ID="LitForeZip" runat="server"></asp:Literal><%--郵遞區號查詢--%>
                                <input id="hidCityNameFore" type="hidden" runat="server" />
                                <input id="hidAREAFore" type="hidden" runat="server" />
                                <input id="ForeZip_N" type="hidden" runat="server" /><br />
                                <asp:TextBox ID="City6" runat="server" Width="20%" onfocus="this.blur()"></asp:TextBox>
                                <input id="bt_openZip5" value="..." type="button" name="bt_openZip4" runat="server" class="asp_button_Mini" />
                                <asp:TextBox ID="ForeAddr" runat="server" Width="60%"></asp:TextBox>
                                <asp:HiddenField ID="Hid_JnForeZip" runat="server" />
                            </td>
                        </tr>
                        <tr id="PWType_TR" runat="server">
                            <td class="bluecol_need" width="18%">受訓前任職狀況 </td>
                            <td class="whitecol" colspan="2" width="50%">
                                <asp:RadioButtonList ID="PriorWorkType1" runat="server" CssClass="font" RepeatDirection="horizontal">
                                    <asp:ListItem Value="1">曾工作過</asp:ListItem>
                                    <asp:ListItem Value="2">未曾工作過</asp:ListItem>
                                    <asp:ListItem Value="3">先前從事為非勞保性質工作</asp:ListItem>
                                    <asp:ListItem Value="4">屆退官兵</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                            <td class="whitecol" width="32%">
                                <asp:Button ID="BtnCheckBli" runat="server" Text="勞保明細檢查" ToolTip="勞保明細檢查" CssClass="asp_button_M"></asp:Button><br />
                                <asp:Label ID="labMsg2017a" Style="color: Red;" runat="server">經網路報名者，需重新勾選確認以下欄位:</asp:Label>
                            </td>
                        </tr>
                        <tr id="trPriorWorkOrg1" runat="server">
                            <td id="tdPriorWorkOrg1" class="bluecol" rowspan="2" width="18%">
                                <asp:Label ID="lab2017_1" runat="server" Text="任職單位名稱"></asp:Label></td>
                            <td class="whitecol" colspan="3" width="82%"><span id="spnPriorWorkOrg1" runat="server">1.<asp:TextBox ID="PriorWorkOrg1" runat="server" Width="30%"></asp:TextBox></span></td>
                        </tr>
                        <tr id="trPriorWorkOrg2" runat="server">
                            <td class="whitecol" colspan="3" width="82%"><span id="spnPriorWorkOrg2" runat="server">2.<asp:TextBox ID="PriorWorkOrg2" runat="server" Width="30%"></asp:TextBox></span></td>
                        </tr>
                        <tr id="trTable6" runat="server">
                            <td id="tdSOfficeYM1" class="bluecol" width="18%">
                                <asp:Label ID="lab2017_2" runat="server" Text="投保單位加退保日期"></asp:Label></td>
                            <td class="whitecol" colspan="3" width="82%">
                                <%--<td rowspan="2" width="10%">&nbsp;
                                    <font color="blue">
                                        *[受訓前失業週數]&nbsp;<br />以第1筆受訓前任職起迄年月資料的離職日至本班開訓日計算此學員失業週數共計&nbsp;
                                        <span style="text-decoration: underline">
                                        <label id="LostJobWeek">&nbsp;&nbsp;&nbsp;</label></span> 週
                                        <span style="text-decoration: underline">
                                        <label id="LostJobDay">&nbsp;&nbsp;&nbsp;</label></span>天。
                                    </font>
                                </td>--%>

                                <table id="Table6" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                                    <tr>
                                        <td colspan="4">
                                            <span id="spnSOfficeYM1" runat="server">1.<asp:TextBox ID="SOfficeYM1" runat="server" Width="20%"></asp:TextBox>
                                                <span runat="server">
                                                    <img style="cursor: pointer" id="IMG2" onclick="javascript:show_calendar('<%= SOfficeYM1.ClientID %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" runat="server" width="30" height="30" /></span>
                                                ～
                                                <asp:TextBox ID="FOfficeYM1" runat="server" Width="20%"></asp:TextBox>
                                                <span runat="server">
                                                    <img style="cursor: pointer" id="IMG3" onclick="javascript:show_calendar('<%= FOfficeYM1.ClientID %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" runat="server" width="30" height="30" /></span>
                                            </span>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td colspan="4">
                                            <span id="spnSOfficeYM2" runat="server">2.<asp:TextBox ID="SOfficeYM2" runat="server" Width="20%"></asp:TextBox>
                                                <span runat="server">
                                                    <img id="IMG4" style="cursor: pointer" onclick="javascript:show_calendar('<%= SOfficeYM2.ClientID %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30" /></span> ～
                                                <asp:TextBox ID="FOfficeYM2" runat="server" Width="20%"></asp:TextBox>
                                                <span runat="server">
                                                    <img id="IMG5" style="cursor: pointer" onclick="javascript:show_calendar('<%= FOfficeYM2.ClientID %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30" /></span>
                                            </span>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr id="ActNo2_TR" runat="server">
                            <td id="tdActNo2" class="bluecol" width="18%">
                                <asp:Label ID="lab2017_3" runat="server" Text="投保單位<br />保險證號"></asp:Label></td>
                            <td class="whitecol" width="32%">&nbsp;&nbsp;
								<asp:TextBox ID="ActNo2" runat="server" Columns="13" MaxLength="9" Width="70%"></asp:TextBox>
                                <asp:Button ID="Button8" runat="server" Text="檢查" ToolTip="是否為協助基金補助對象" CssClass="asp_button_M"></asp:Button>
                            </td>
                            <td id="tdPriorWorkPay" class="bluecol" width="18%">
                                <asp:Label ID="lab2017_4" runat="server" Text="投保薪資級距"></asp:Label></td>
                            <td class="whitecol" width="32%">
                                <asp:TextBox ID="PriorWorkPay" runat="server" Width="60%" MaxLength="10"></asp:TextBox></td>
                        </tr>
                        <tr id="trTitle1" runat="server">
                            <td id="tdTitle1" class="bluecol" rowspan="2" width="18%">職稱 </td>
                            <td class="whitecol" width="32%" colspan="3"><span id="spnTitle1">1.</span><asp:TextBox ID="Title1" runat="server" Width="70%"></asp:TextBox></td>
                            <%--<td id="tdRealJobless" class="bluecol" runat="server" rowspan="2" width="18%">受訓前失業週數</td>
                            <td class="whitecol" rowspan="2" width="32%">
                                <span id="spnRealJobless">
                                    <asp:TextBox ID="RealJobless" runat="server" Width="30%"></asp:TextBox>
                                    <asp:DropDownList ID="JoblessID" runat="server"></asp:DropDownList>
                                    <br /><asp:Label ID="RealJobless_msg" runat="server" ForeColor="red"></asp:Label>
                                </span>
                            </td>--%>
                        </tr>
                        <tr id="trTitle1b" runat="server">
                            <td class="whitecol" colspan="3"><span id="spnTitle2">2.</span><asp:TextBox ID="Title2" runat="server" Width="70%"></asp:TextBox></td>
                        </tr>
                        <tr id="trTraffic" runat="server">
                            <td class="bluecol" width="18%">交通方式 </td>
                            <td class="whitecol" colspan="3" width="82%">
                                <asp:DropDownList ID="Traffic" runat="server">
                                    <asp:ListItem Value="0">請選擇</asp:ListItem>
                                    <asp:ListItem Value="1">住宿</asp:ListItem>
                                    <asp:ListItem Value="2">通勤</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr id="trShowDetail" runat="server">
                            <td class="bluecol_need" width="18%">提供基本資料供<br />
                                求才廠商查詢 </td>
                            <td class="whitecol" colspan="3" width="82%">
                                <asp:DropDownList ID="ShowDetail" runat="server">
                                    <asp:ListItem Value="0">請選擇</asp:ListItem>
                                    <asp:ListItem Value="Y">是</asp:ListItem>
                                    <asp:ListItem Value="N">否</asp:ListItem>
                                </asp:DropDownList>
                                (姓名、出生年月日、性別、學歷、科系、電話、電子郵件帳號)
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">預算別 (經費來源別) </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="BudID" runat="server" RepeatDirection="horizontal" RepeatLayout="flow" CssClass="font"></asp:RadioButtonList>
                                <asp:Literal ID="BudIDMsg" runat="server"></asp:Literal>
                            </td>
                            <%--<td id="SupplyTD" class="bluecol_need" runat="server">補助比例 </td>
                            <td class="whitecol">&nbsp;
                                <asp:DropDownList ID="SupplyID" runat="server" AppendDataBoundItems="true">
                                    <asp:ListItem Value="0">請選擇</asp:ListItem>
                                    <asp:ListItem Value="1">一般80%</asp:ListItem>
                                    <asp:ListItem Value="2">特定100%</asp:ListItem>
                                    <asp:ListItem Value="9">0%</asp:ListItem>
                                </asp:DropDownList>
                            </td>--%>
                        </tr>
                        <tr id="HGTR" runat="server">
                            <td class="bluecol" width="18%">專上畢業學歷失<br />
                                業者(特別預算) </td>
                            <td class="whitecol" colspan="3" width="82%">
                                <asp:RadioButtonList ID="rdo_HighEduBg" runat="server" CssClass="font" RepeatDirection="horizontal">
                                    <asp:ListItem Value="Y">是</asp:ListItem>
                                    <asp:ListItem Value="N" Selected="true">否</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <%--<tr id="trPMode" runat="server">
                            <td class="bluecol" width="18%">公費/自費<br />
                                (職訓券必填) </td>
                            <td class="whitecol" colspan="3" width="82%">
                                <asp:RadioButtonList ID="PMode" runat="server" RepeatDirection="horizontal" RepeatLayout="flow" CssClass="font">
                                    <asp:ListItem Value="1">公費</asp:ListItem>
                                    <asp:ListItem Value="2">自費</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>--%>
                        <tr id="WSITR" runat="server">
                            <td class="bluecol_need" width="18%">是否為在職者<br />
                                補助身分 </td>
                            <td class="whitecol" width="32%">
                                <asp:RadioButtonList ID="rblWorkSuppIdent" runat="server" CssClass="font" RepeatDirection="horizontal">
                                    <asp:ListItem Value="Y">是</asp:ListItem>
                                    <asp:ListItem Value="N">否</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                            <td class="whitecol" colspan="2" width="50%">
                                <asp:Label ID="LabWSImsg" runat="server" ForeColor="Red"></asp:Label></td>
                        </tr>
                        <tr id="WSITR2" runat="server">
                            <td class="bluecol" width="18%">在職者已補助<br />
                                經費 </td>
                            <td class="whitecol" colspan="3" width="82%">
                                <asp:Label ID="SubsidyCost" runat="server"></asp:Label></td>
                        </tr>
                        <tr id="trHouseMatch" runat="server">
                            <td class="bluecol" width="18%">同意本署將學員個人資料提供社家署做就業媒合之用 </td>
                            <td class="whitecol" colspan="3" width="82%">
                                <asp:RadioButtonList ID="rblHouseMatch" runat="server" CssClass="font" RepeatDirection="horizontal">
                                    <asp:ListItem Value="Y" Selected="true">同意</asp:ListItem>
                                    <asp:ListItem Value="N">不同意</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" colspan="4" width="100%">本人 同意 勞動部勞動力發展署 暨所屬機關，為本人提供職業訓練及就業服務時使用 </td>
                        </tr>
                        <tr id="trTPlanid28_1" runat="server">
                            <td class="whitecol" colspan="4" width="100%"><font style="color: #009900;">*記得維護「參訓背景」。</font></td>
                        </tr>
                    </table>

                </td>
            </tr>
        </table>
        <table id="BackTable" cellspacing="1" cellpadding="1" width="100%" runat="server" class="table_nw">
            <tr>
                <td class="table_title" colspan="4" align="center">服務單位資料 </td>
            </tr>

            <tr>
                <td class="bluecol">目前任職公司名稱<asp:Label ID="star22" runat="server"><font color="#ff0000">*</font></asp:Label></td>
                <td class="whitecol">
                    <asp:TextBox ID="Uname" runat="server" Width="80%"></asp:TextBox></td>
                <td class="bluecol">統一編號 </td>
                <td class="whitecol">
                    <asp:TextBox ID="Intaxno" runat="server" Columns="19" Width="80%"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="bluecol_need">投保單位名稱<asp:Label ID="star14" runat="server"><font color="#ff0000">*</font></asp:Label></td>
                <td class="whitecol">
                    <asp:TextBox ID="ActName" runat="server" Width="80%"></asp:TextBox></td>
                <td class="bluecol_need">投保單位保險證號<%--<asp:Label ID="star15" runat="server"><font color="#ff0000">*</font></asp:Label>--%></td>
                <td class="whitecol">
                    <asp:TextBox ID="ActNo1" runat="server" Columns="13" MaxLength="9" Width="60%"></asp:TextBox>
                    <%--
                    <asp:Button ID="Button7" runat="server" Text="檢查" ToolTip="是否為協助基金補助對象" CssClass="asp_button_S"></asp:Button>
                    <asp:Button ID="BtnAutoInputActno" runat="server" Text="自動帶入" ToolTip="自動帶入投保證號" CssClass="asp_button_S"></asp:Button>
                    --%>
                    <asp:Button ID="BtnCheckBli2" runat="server" Text="選擇投保證號" ToolTip="選擇投保證號" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
            <tr>
                <td class="bluecol">投保單位電話<asp:Label ID="star16" runat="server"> </asp:Label></td>
                <td class="whitecol" colspan="3">
                    <asp:TextBox ID="txt_ActPhone" runat="server" Columns="13" Width="30%"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="bluecol">投保單位地址<asp:Label ID="star17" runat="server"> </asp:Label></td>
                <td class="whitecol" colspan="3">
                    <input id="txt_ActZip" maxlength="3" runat="server" />－
				    <input id="txt_ActZIPB3" maxlength="3" runat="server" />
                    <input id="hid_ActZIP6W" type="hidden" runat="server" />
                    <asp:Literal ID="LitActZip" runat="server"></asp:Literal><%--郵遞區號查詢--%>
                    <input id="hidCityNameAct" type="hidden" runat="server" />
                    <input id="hidAREAAct" type="hidden" runat="server" />
                    <input id="hidActZip_N" type="hidden" runat="server" /><br />
                    <asp:TextBox ID="txt_ActCity" runat="server" Width="20%" onfocus="this.blur()"></asp:TextBox>
                    <input id="bt_openZip6" value="..." type="button" name="bt_openZip5" runat="server" class="asp_button_Mini" />
                    <asp:TextBox ID="txt_ActAddress" runat="server" Width="60%"></asp:TextBox>
                    <asp:HiddenField ID="Hid_JnActZip" runat="server" />
                </td>
            </tr>
            <%--
            <tr>
			    <td id="TdActType" class="bluecol" runat="server">投保類別<asp:Label ID="star18" runat="server"><font color="#ff0000">*</font></asp:Label><br />(產業人才必填)</td>
			    <td class="whitecol" colspan="3">
				    <asp:RadioButtonList ID="ActType" runat="server" RepeatDirection="horizontal" RepeatLayout="flow" CssClass="font">
					    <asp:ListItem Value="1" Selected="true">勞</asp:ListItem>
					    <asp:ListItem Value="2">農</asp:ListItem>
				    </asp:RadioButtonList>
			    </td>
		    </tr>
            --%>

            <tr>
                <td class="bluecol">目前任職部門<font color="#ff0000">*</font> </td>
                <td class="whitecol">
                    <asp:TextBox ID="ServDept" runat="server"></asp:TextBox>
                    <asp:DropDownList ID="ddlSERVDEPTID" runat="server"></asp:DropDownList>
                </td>
                <td class="bluecol">職務<asp:Label ID="star23" runat="server"><font color="#ff0000">*</font></asp:Label></td>
                <td class="whitecol">
                    <asp:TextBox ID="JobTitle" runat="server"></asp:TextBox>
                    <asp:DropDownList ID="ddlJOBTITLEID" runat="server"></asp:DropDownList>
                </td>
            </tr>

        </table>
        <table id="HistoryTable" class="table_nw" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
            <tr>
                <td valign="top" align="center">&nbsp;&nbsp;&nbsp;
				<asp:Label ID="msg" runat="server" CssClass="font" ForeColor="red"></asp:Label><br />
                    <br />
                    <asp:Panel ID="Panel" runat="server" Width="100%" Visible="false">
                        <asp:DataGrid Style="z-index: 0" ID="DataGrid1" runat="server" CssClass="font" Width="100%" Visible="False" AutoGenerateColumns="false" CellPadding="8">
                            <AlternatingItemStyle BackColor="#EEEEEE" />
                            <HeaderStyle CssClass="head_navy" />
                            <Columns>
                                <asp:BoundColumn DataField="PlanYear" HeaderText="年度">
                                    <HeaderStyle Width="10%"></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="STDate" HeaderText="開訓日期" DataFormatString="{0:d}"></asp:BoundColumn>
                                <asp:BoundColumn DataField="FTDate" HeaderText="結訓日期" DataFormatString="{0:d}"></asp:BoundColumn>
                                <asp:BoundColumn DataField="ClassName" HeaderText="班別"></asp:BoundColumn>
                                <asp:BoundColumn DataField="BudName" HeaderText="預算別"></asp:BoundColumn>
                                <asp:TemplateColumn HeaderText="審核補助金額">
                                    <ItemStyle HorizontalAlign="right"></ItemStyle>
                                    <ItemTemplate>
                                        <asp:Label ID="SumOfMoney" runat="server" Text='<%#DataBinder.Eval(Container, "DataItem.SumOfMoney") %>'>
                                        </asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:TemplateColumn HeaderText="撥款補助金額">
                                    <ItemStyle HorizontalAlign="right"></ItemStyle>
                                    <ItemTemplate>
                                        <asp:Label ID="SumOfMoney2" runat="server" Text='<%# databinder.eval(container, "DataItem.SumOfMoney2") %>'>
                                        </asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                            </Columns>
                            <PagerStyle Visible="false"></PagerStyle>
                        </asp:DataGrid>
                    </asp:Panel>
                </td>
            </tr>
        </table>
        <table id="Table4" runat="server" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td align="center" class="whitecol">
                    <asp:Button ID="Button1" runat="server" Text="儲存回查詢頁面" CssClass="asp_button_M"></asp:Button>
                    <asp:Button ID="Button2" runat="server" Text="維護下一位學員" CssClass="asp_button_M"></asp:Button>
                    <asp:Button ID="Button3" runat="server" Text="不儲存回上一頁" CssClass="asp_button_M"></asp:Button>
                    <asp:Button ID="Button5" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
        </table>
        <input id="hidSB4ID" type="hidden" runat="server" />
        <input id="RoleID" type="hidden" runat="server" />
        <input id="Process" type="hidden" name="Process" runat="server" />
        <input id="TPlanID" type="hidden" runat="server" />
        <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
        <input id="DistValue" type="hidden" name="DistValue" runat="server" />
        <input id="Hid_OCID" type="hidden" name="Hid_OCID" runat="server" />
        <input id="STDateHidden" type="hidden" name="STDateHidden" runat="server" />
        <input id="FTDateHidden" type="hidden" name="FTDateHidden" runat="server" />
        <input id="hide_EnterChannel" type="hidden" name="hide_EnterChannel" runat="server" />
        <input id="hide_EnterChannel2" type="hidden" name="hide_EnterChannel2" runat="server" />
        <input id="hide_TrainMode" type="hidden" name="hide_TrainMode" runat="server" />
        <input id="hide_Years" type="hidden" name="hide_Years" runat="server" />
        <input id="hide_MIdentityID" type="hidden" name="hide_MIdentityID" runat="server" />
        <input id="hide_IdentityIDType" type="hidden" name="hide_IdentityIDType" runat="server" />
        <input id="hide_MakeSOCID" type="hidden" name="hide_MakeSOCID" runat="server" />
        <input id="hide_RejectSOCID" type="hidden" name="hide_RejectSOCID" runat="server" />
        <input id="hide_THours" type="hidden" name="hide_THours" runat="server" />
        <input id="hide_3Y_SupplyMoney" type="hidden" name="hide_3Y_SupplyMoney" runat="server" />
        <asp:HiddenField ID="HidMaster" runat="server" />
        <asp:HiddenField ID="HidPostNo_1" runat="server" />
        <%--<asp:HiddenField ID="HidPostNo_2" runat="server" />--%>
        <asp:HiddenField ID="HidAcctNo1_1" runat="server" />
        <%--<asp:HiddenField ID="HidAcctNo1_2" runat="server" />--%>
        <asp:HiddenField ID="HidAcctheadNo" runat="server" />
        <asp:HiddenField ID="HidAcctExNo" runat="server" />
        <asp:HiddenField ID="HidAcctNo2" runat="server" />
        <asp:HiddenField ID="HidPreUseLimited17f" runat="server" />
        <asp:HiddenField ID="Hid_MSG2" runat="server" />
        <asp:HiddenField ID="Hid_MSGADIDN" runat="server" />
        <asp:HiddenField ID="Hid_SID_C1" runat="server" />
        <asp:HiddenField ID="Hid_BIEF" runat="server" />
        <asp:HiddenField ID="Hid_out_POSITION" runat="server" />
        <asp:HiddenField ID="Hid_SETID" runat="server" />
        <asp:HiddenField ID="Hid_ETENTERDATE" runat="server" />
        <asp:HiddenField ID="Hid_SERNUM" runat="server" />

        <input id="Hid_show_actno_budid" type="hidden" runat="server" />
        <input id="Hid_nouse_SupplyID" type="hidden" runat="server" />
    </form>
    <iframe id="ifmChceckZip" height="0%" src="../../Common/CheckZip.aspx" width="0%" />
</body>
</html>
