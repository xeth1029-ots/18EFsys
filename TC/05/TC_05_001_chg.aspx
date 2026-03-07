<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_05_001_chg.aspx.vb" Inherits="WDAIIP.TC_05_001_chg" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>班級變更申請</title>
    <style type="text/css">
        DIV { scrollbar-arrow-color: #4f5f87; scrollbar-face-color: #c7cfff; scrollbar-darkshadow-color: #9fb7d7; scrollbar-highlight-color: #ffffff; scrollbar-shadow-color: #ffffff; scrollbar-track-color: #f7f7f7; scrollbar-3dlight-color: #c7cff7; }
        /*.style1 { font-size: 12px; color: Black; line-height: 22px; text-align: center; background-color: #CCD8EE; padding: 2px; height: 19px; }
        .style2 { font-size: 12px; color: Black; line-height: 22px; background-color: #FFFFFF; padding: 2px; height: 19px; }
        .auto-style1 { color: Black; text-align: right; padding: 4px 6px; background-color: #f1f9fc; border-right: 3px solid #49cbef; height: 29px; }
        .auto-style2 { height: 29px; }*/
    </style>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/OpenWin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        $(document).ready(function () {
            var isInsertNext = $("#divInsertNext").text().trim();
            if (isInsertNext == "1") {
                //'insert_next_val 0:(非)自辦計畫 /1:自辦計畫
                switch (insert_next_val) {
                    case '0':
                        insert_next('0');//(非)自辦計畫
                        break;
                    case '1':
                        insert_next('1');//自辦計畫
                        break;
                }
            }
        });

        //var cst_inline1 = "inline";
        var cst_inline1 = "";
        function insert_next(vo1) {
            var ReqID = document.getElementById("hidReqID").value;
            var ReqPlanID = document.getElementById("hidReqPlanID").value;
            var Reqcid = document.getElementById("hidReqcid").value;
            var Reqno = document.getElementById("hidReqno").value;
            var Reqcheck = document.getElementById("hidReqcheck").value;
            var sUrl = '';
            var msgX1 = "計劃變更申請成功，是否繼續新增?<br>";
            var msgX2 = "(自辦計畫 直接審核成功)<br>";
            if (vo1 == '1') { msgX1 += msgX2; }
            //if (window.confirm(msgX1)) {
            //	sUrl = 'TC_05_001_chg.aspx?ID=' + ReqID;
            //	sUrl += '&PlanID=' + ReqPlanID;
            //	sUrl += '&cid=' + Reqcid;
            //	sUrl += '&no=' + Reqno;
            //	sUrl += '&check=' + Reqcheck;
            //	location.href = sUrl;
            //}
            //else {
            //	sUrl = 'TC_05_001.aspx?ID=' + ReqID;
            //	location.href = sUrl;
            //}
            blockConfirm(msgX1, ""
                , function () {
                    sUrl = 'TC/05/TC_05_001_chg.aspx?ID=' + ReqID;
                    sUrl += '&PlanID=' + ReqPlanID;
                    sUrl += '&cid=' + Reqcid;
                    sUrl += '&no=' + Reqno;
                    sUrl += '&check=' + Reqcheck;
                    location.href = sUrl;
                }
                , function () {
                    sUrl = 'TC/05/TC_05_001.aspx?ID=' + ReqID;
                    location.href = sUrl;
                });
        }

        function ChangeState(TPlanID) {
            var vMaxChgItem = 22; //21;
            var vNotTPlanIDs = "28,54";
            vNotTPlanIDs = document.getElementById('hid_TPlanID28AppPlan').value;
            vMaxChgItem = parseInt(document.getElementById('hid_MaxChgItem').value, 10);
            //debugger;
            //alert(vMaxChgItem);
            if (document.getElementById('ChgItem').selectedIndex != 0) {
                //document.getElementById('But_Sub').style.display='inline';			
                document.getElementById('chgState').value = document.getElementById('ChgItem').value;
                document.getElementById('ReviseTable').style.display = cst_inline1;
                //document.getElementById('ReviseCont').value=document.getElementById('ReviseCont'+document.getElementById('ChgItem').value).value;
                if (document.getElementById('SearchMode').innerText == '申請') {
                    document.getElementById('changeReason').value = 0;
                    if (!(vNotTPlanIDs.search(TPlanID) > -1)) { document.getElementById('ReviseCont').value = ''; }
                }
                for (i = 1; i <= vMaxChgItem; i++) {
                    for (j = 1; j <= 3; j++) {
                        if (document.getElementById('TR' + i + '_' + j)) {
                            document.getElementById('TR' + i + '_' + j).style.display = 'none';
                            //alert('TR'+i+'_'+j);
                        }
                    }
                }
                /**  20080627 andy add 課程表	start **/
                var Tr18 = document.getElementById('Tr18');
                var x = document.getElementById('ChgItem').value;
                /*以下項目選取會一併帶出課程表予以修改*/
                //alert(x);
                switch (x) {
                    case '1':
                        /**訓練期間**/
                        Tr18.style.display = cst_inline1;
                        //if ( TPlanID  !=28 && TPlanID  !=54)  {
                        tb_New_EnterDate2.style.display = 'none';
                        tb_EnterDate2.style.display = 'none';
                        if (!(vNotTPlanIDs.search(TPlanID) > -1)) {
                            tb_New_EnterDate2.style.display = cst_inline1; /*報名起訖 (產學訓不包在內)*/
                            tb_EnterDate2.style.display = cst_inline1;
                        }
                        break;
                    case '11':      /**師資**/
                        Tr18.style.display = cst_inline1;
                        break;
                    case '20':      /**助教**/
                        Tr18.style.display = cst_inline1;
                        break;
                    case '14':      /**學(術)科場地**/
                        Tr18.style.display = cst_inline1;
                        break;
                    case '18':      /**課程表**/
                        Tr18.style.display = cst_inline1;
                        break;
                    case '22':      /**課程表 遠距教學**/
                        Tr18.style.display = cst_inline1;
                        break;
                    default:
                        Tr18.style.display = 'none';
                        break;
                }
                /**  20080627 andy add 課程表 end **/
                var ChgItem = document.getElementById('ChgItem');
                for (i = 1; i <= 3; i++) {
                    var TR_ChgItem = document.getElementById('TR' + ChgItem.value + '_' + i);
                    if (TR_ChgItem) { TR_ChgItem.style.display = cst_inline1; } //'inline';
                }
            }
            else {
                document.getElementById('chgState').value = 0;
                document.getElementById('ReviseTable').style.display = 'none';
            }
            document.getElementById('But_Sub').style.display = 'none';
            if (!(vNotTPlanIDs.search(TPlanID) > -1)) { document.getElementById('But_Sub').style.display = cst_inline1; }
            //if (!_isIE) { parent.setMainFrameHeight(); } //重新調整頁面高度
            if (parent && parent.setMainFrameHeight != undefined) { parent.setMainFrameHeight(); }
        }

        function GetPackageName54() {
            var cst_pt1 = 0;
            var cst_pt2 = 1;
            if (document.getElementsByName('PackageTypeNew').length > 2) {
                cst_pt1 = 1; //cst_pt
                cst_pt2 = 2;
            }
            if (document.getElementById('btnAddBusPackage')) {
                document.getElementById('btnAddBusPackage').style.display = cst_inline1;
            }
            if (document.getElementById('BusPackageNewTable')) {
                document.all.BusPackageNewTable.style.visibility = 'hidden';
                //document.getElementById('BusPackageNewTable').style.display='hidden';
            }
            //debugger;
            //充電起飛計畫 且 聯合企業包班 才可顯示 計畫包班事業單位
            if (document.getElementsByName("PackageTypeNew")[cst_pt2].checked == true) {
                if (document.getElementById('BusPackageNewTable')) {
                    document.all.BusPackageNewTable.style.visibility = 'visible';
                    //document.getElementById('BusPackageNewTable').style.display='none';
                }
            }
            //充電起飛計畫 且 企業包班 只顯示 計畫包班事業單位 的抬頭
            if (document.getElementsByName("PackageTypeNew")[cst_pt1].checked == true) {
                if (document.getElementById('btnAddBusPackage')) {
                    document.getElementById('btnAddBusPackage').style.display = 'none';
                }
            }
        }

        //開啟日曆(回傳欄位,開訓日期,結訓日期,目前指定日期[可略],頁面相對位置[可略],回傳要觸發的按鈕[可略])
        function openCalendar(obj, STDate, FTDate, NowDate, page, btn) {
            if (NowDate == '' || NowDate == 'undefined')
                NowDate = STDate;
            if (page == '' || page == undefined)
                page = '../../common/Calendar.aspx';
            if (btn == undefined)
                btn = '';
            if (STDate != '') {
                wopen(page + '?STDate=' + STDate + '&FTDate=' + FTDate + '&NowDate=' + NowDate + '&ValueField=' + obj + '&Button=' + btn, '', 350, 230);
            }
        }

        /**
		function showPTID(obj, ptid1, ptid2 ,ptid3){
		var objvalue=document.getElementById(obj).value;
		if (objvalue==1){
		document.getElementById(ptid1).style.display="inline";
		document.getElementById(ptid2).style.display="none";
		document.getElementById(ptid3).style.display="none";
		}
		else{
		document.getElementById(ptid1).style.display="none";
		document.getElementById(ptid2).style.display="inline";
		document.getElementById(ptid3).style.display="none";
		}			
		}	

		function LessonTeah1(opentype, fieldname, hiddenname) {
		var msg ='';
		//debugger;
		if (document.form1.RID.value != null && document.form1.RID.value != "" ) {
		wopen('../../SD/04/LessonTeah1.aspx?RID='+document.form1.RID.value+'&type='+opentype+'&hiddenname='+hiddenname+'&fieldname='+fieldname,'LessonTeah1',400,300,1);
		}
		else{
		msg+='請先選擇訓練機構\n';
		if(msg!=''){
		alert(msg);
		return false;
		}
		}
		}
		***/

        function Check_Data(TPlanID) {
            var msg = '';
            var vNotTPlanIDs = "28,54";
            vNotTPlanIDs = document.getElementById('hid_TPlanID28AppPlan').value;
            //vMaxChgItem=parseInt(document.getElementById('hid_MaxChgItem').value,10);
            if (document.getElementById('ApplyDate').value == '') msg += '請輸入申請日期\n';
            else if (!checkDate(document.getElementById('ApplyDate').value)) msg += '申請日期必須為正確的時間格式\n';
            if (document.getElementById('ChgItem').selectedIndex != 0) {
                switch (document.getElementById('ChgItem').value) {
                    case '1':
                        var SDate1 = document.getElementById('BSDate').innerHTML;
                        var EDate1 = document.getElementById('BEDate').innerHTML;
                        var SDate2 = document.getElementById('ASDate').value;
                        var EDate2 = document.getElementById('AEDate').value;
                        if (SDate2 == '') msg += '請輸入訓練起日\n';
                        else if (!checkDate(SDate2)) msg += '「訓練起日」必須為正確的時間格式\n';
                        if (EDate2 == '') msg += '請輸入訓練迄日\n';
                        else if (!checkDate(EDate2)) msg += '「訓練迄日」必須為正確的時間格式\n';
                        if (!document.getElementById('IMG1')) {
                            if (compareDate(EDate1, EDate2) == 1) msg += '「欲變更的 訓練迄日」必須比「原 訓練迄日」晚\n';
                        }
                        if ((Date.parse(SDate2)).valueOf() > (Date.parse(EDate2)).valueOf()) { msg += '「訓練起日」不得晚於「訓練迄日」\n'; }
                        if (!(vNotTPlanIDs.search(TPlanID) > -1)) {
                            //var Old_SEnterDate2 = document.getElementById('Old_SEnterDate2').value;
                            //var Old_FEnterDate2 = document.getElementById('Old_FEnterDate2').value;
                            var New_SEnterDate2 = document.getElementById('New_SEnterDate2').value;
                            var New_FEnterDate2 = document.getElementById('New_FEnterDate2').value;
                            if ((Date.parse(New_SEnterDate2)).valueOf() > (Date.parse(New_FEnterDate2)).valueOf()) { msg += '「報名起日」不得晚於「報名迄日」\n'; }
                            // if ( (Date.parse(New_FEnterDate2)).valueOf() < (new Date()).valueOf())
                            // {  msg+='「報名迄日」不得早於「今日」\n'; }    
                            if ((Date.parse(New_FEnterDate2)).valueOf() > (Date.parse(SDate2)).valueOf()) { msg += '「報名迄日」不得晚於「開訓日」\n'; }
                        }
                        else {
                            if (compareDate(SDate1, SDate2) == 0
                                && compareDate(EDate1, EDate2) == 0) msg += '「訓練起迄日」不能相同\n';
                        }
                        break;
                    /**	case '2':
                    if( document.getElementById('TimeSDate').value==document.getElementById('TimeEDate').value)
                    msg+='原始計畫變更日期不可與預定變更日期相同\n';
                    break;     
                    if(getCheckBoxListValue('TimeSClass')=='') msg+='請選擇有效的原始計畫變更日期\n';
                    else if(parseInt(getCheckBoxListValue('TimeSClass'))==0) msg+='請勾選原始計畫節次\n';
                    if(getCheckBoxListValue('TimeEClass')=='') msg+='請選擇有效的預定變更日期\n';
                    if(parseInt(getCheckBoxListValue('TimeEClass'))==0) msg+='請勾選變更節次\n';	
                    break;  
                    **/
                    case '3':
                        //if (getCheckBoxListValue('SPlace') == '') msg += '請選擇有效的原始計畫變更日期\n';
                        //else if (parseInt(getCheckBoxListValue('SPlace')) == 0) msg += '請勾選原始計畫節次\n';
                        var SPlace = document.getElementById('SPlace');
                        var STDate = document.getElementById('STDate');
                        var FDDate = document.getElementById('FDDate');
                        var PlaceDate = document.getElementById('PlaceDate');
                        var dtSTDate;
                        var dtFDDate;
                        var dtPlaceDate;
                        var strSubMsg1 = "";
                        if (STDate && !isBlank(STDate)) {
                            dtSTDate = new Date(STDate.value);
                        }
                        if (FDDate && !isBlank(FDDate)) {
                            dtFDDate = new Date(FDDate.value);
                        }
                        if (PlaceDate && !isBlank(PlaceDate)) {
                            dtPlaceDate = new Date(PlaceDate.value);
                            if (dtSTDate != undefined && dtFDDate != undefined) {
                                if (!(getDiffDay(dtSTDate, dtPlaceDate) >= 0 && getDiffDay(dtPlaceDate, dtFDDate) >= 0)) {
                                    strSubMsg1 = "請選擇有效的原始計畫變更日期\n";
                                    msg += strSubMsg1;
                                }
                            }
                        }
                        if (strSubMsg1 == "" && (SPlace && (getCheckBoxListValue('SPlace') == "" || parseInt(getCheckBoxListValue('SPlace')) == 0))) {
                            msg += '請勾選原始計畫節次\n';
                        }
                        if (document.getElementById('EPlace').value == '') { msg += '請輸入變更地點\n'; }
                        else {
                            if (document.getElementById('EPlace').value.length > 12) msg += '已輸入超過12個中文字\n';
                        }
                        break;
                    case '4':
                        if (document.getElementById('EGenSci').value == '') msg += '請輸入一般學科時數\n';
                        else if (!isUnsignedInt(document.getElementById('EGenSci').value)) msg += '一般學科時數必須為數字\n';
                        if (document.getElementById('EProSci').value == '') msg += '請輸入專業學科時數\n';
                        else if (!isUnsignedInt(document.getElementById('EProSci').value)) msg += '專業學科時數必須為數字\n';
                        if (document.getElementById('EProTech').value == '') msg += '請輸入術科時數\n';
                        else if (!isUnsignedInt(document.getElementById('EProTech').value)) msg += '術科時數必須為數字\n';
                        if (document.getElementById('EOther').value == '') msg += '請輸入一般時數\n';
                        else if (!isUnsignedInt(document.getElementById('EOther').value)) msg += '一般時數必須為數字\n';
                        break;
                    case '5':
                        //if (getCheckBoxListValue('STeacher') == '') msg += '請選擇有效的原始計畫變更日期\n';
                        //else if (parseInt(getCheckBoxListValue('STeacher')) == 0) msg += '請勾選原始計畫節次\n';
                        var STeacher = document.getElementById('STeacher');
                        var STDate = document.getElementById('STDate');
                        var FDDate = document.getElementById('FDDate');
                        var TechDate = document.getElementById('TechDate');
                        var dtSTDate;
                        var dtFDDate;
                        var dtTechDate;
                        var strSubMsg2 = "";
                        if (STDate && !isBlank(STDate)) {
                            dtSTDate = new Date(STDate.value);
                        }
                        if (FDDate && !isBlank(FDDate)) {
                            dtFDDate = new Date(FDDate.value);
                        }
                        if (TechDate && !isBlank(TechDate)) {
                            dtTechDate = new Date(TechDate.value);
                            if (dtSTDate != undefined && dtFDDate != undefined) {
                                if (!(getDiffDay(dtSTDate, dtTechDate) >= 0 && getDiffDay(dtTechDate, dtFDDate) >= 0)) {
                                    strSubMsg2 = "請選擇有效的原始計畫變更日期\n";
                                    msg += strSubMsg2;
                                }
                            }
                        }
                        if (strSubMsg2 == "" && (STeacher && (getCheckBoxListValue('STeacher') == "" || parseInt(getCheckBoxListValue('STeacher')) == 0))) {
                            msg += '請勾選原始計畫節次\n';
                        }
                        if (document.getElementById('OLessonTeah1').value == '') msg += '請輸入變更教師\n';
                        break;
                    case '6':
                        if (document.getElementById('ChangeClassCName').value == '') msg += '請輸入變更班別名稱\n';
                        break;
                    case '7':
                        if (document.getElementById('ChangeCyclType').value == '') msg += '請輸入期別\n';
                        else if (document.getElementById('ChangeCyclType').value.length != 2) msg += '期別必須為兩位數字\n';
                        else if (!isUnsignedInt(document.getElementById('ChangeCyclType').value)) msg += '期別必須為數字\n';
                        else if (parseInt(document.getElementById('ChangeCyclType').value, 10) <= 0) msg += '期別必須大於0\n';
                        else if (parseInt(document.getElementById('ChangeCyclType').value, 10) == parseInt(document.getElementById('CyclType').innerHTML, 10)) msg += '預定變更的期別不能和原計畫的相同\n';
                        break;

                    case '8':
                        $('#NewData8_3').val($.trim($("#NewData8_3").val()));
                        $('#NewData8_2').val($.trim($("#NewData8_2").val()));
                        if ($('#NewData8_1').val() == '') msg += '請輸入訓練地點城市\n';
                        if ($('#NewData8_3').val() == "") { msg += '請輸入上課地址郵遞區號後2碼或後3碼\n'; }
                        if ($('#NewData8_3').val() != '') { msg += checkzip23(true, '上課地址', 'NewData8_3'); }
                        //else {
                        //    var flag_NG = false;
                        //    if ($('#NewData8_3').val().length == 2 || $('#NewData8_3').val().length == 3) { flag_NG = true; }
                        //    if (!flag_NG && !isUnsignedInt($('#NewData8_3').val())) { flag_NG = true; }
                        //    if (!flag_NG && parseInt($('#NewData8_3').val(), 10) < 1) { flag_NG = true; }
                        //    if (flag_NG) { msg += '郵遞區號後2碼或後3碼必須為數字，且不得輸入 00\n'; }
                        //}
                        if ($('#NewData8_2').val() == '') msg += '請輸入訓練地點\n';
                        break;
                    case '9':
                        if (!isChecked(document.getElementsByName('NewData9_1'))) msg += '請勾選停辦申請\n';
                        break;
                    case '10':
                        if (document.getElementById('NewData10_1').selectedIndex == 0) msg += '請選擇上課時段\n';
                        else if (document.getElementById('OldData10_1') == document.getElementById('NewData10_1')) msg += '不能與原時段相同\n'
                        break;
                    case '11':
                        if (document.getElementById('NewData11_1').value == '') msg += '請選擇新的師資\n';
                        else if (document.getElementById('OldData11_1').value == document.getElementById('NewData11_1').value) msg += '不能與原師資相同\n';
                        break;
                    case '20':
                        if (document.getElementById('OldData20_1').value != '') {
                            if (document.getElementById('NewData20_1').value == '') msg += '請選擇新的助教\n';
                            else if (document.getElementById('OldData20_1').value == document.getElementById('NewData20_1').value) msg += '不能與原助教相同\n';
                        }
                        else {
                            if (document.getElementById('NewData20_1').value != '') msg += '原助教並無設定，不可選擇新的助教\n';
                        }
                        break;
                    case '12':
                        if (document.getElementById('NewData12_1').value == '') msg += '請輸入核定人數\n'; //'請輸入招生人數\n';
                        else if (document.getElementById('NewData12_1').value == '0') msg += '核定人數不能等於零\n';
                        else if (document.getElementById('NewData12_1').value == document.getElementById('OldData12_1').innerHTML) msg += '不能與原核定人數相同\n';
                        break;
                    case '13':
                        if (document.getElementById('NewData13_1').value == '') msg += '請選輸入增班數\n';
                        else if (document.getElementById('NewData13_1').value == document.getElementById('OldData13_1').innerHTML) msg += '不能與原班數相同\n';
                        break;
                    case '14':
                        // if(document.getElementById('NewData14_1').selectedIndex==0 && document.getElementById('NewData14_2').selectedIndex==0) msg+='學術科場地不可同時無資料\n';
                        //else if(document.getElementById('NewData14_1').selectedValue==document.form1.OldData14_1.value && document.getElementById('NewData14_2').selectedValue==document.form1.OldData14_2.value)	msg+='不能與原學(術)場地資料相同\n';				
                        break;
                    case '15':
                        var MyTable = document.getElementById('DataGrid2');
                        //if(MyTable.rows.count==0) msg+='請輸入欲變更之上課時段\n';
                        break;
                    case '16':
                        if (document.getElementById('NewData15_1').value == '') msg += '請輸入變更後其他內容\n';
                        break;
                }
                /**---檢查字數---*/
                var oTextCount = document.getElementById("ReviseCont");
                var i_max_ReviseCont_len = 250;
                var tmp = ""; // tmp = oTextCount.value;
                tmp = oTextCount.value.replace(/[^\u0000-\u00ff]/g, " a a ");
                tmp = tmp.replace(/\b([a-z0-9]+)\b/gi, " a ");
                tmp = tmp.replace(/\s/g, "");
                if (tmp.length == 0) { msg += '變更說明欄位未填寫！\n'; }
                //if (tmp.length > i_max_ReviseCont_len) { msg += '「變更說明」欄位 ' + tmp.length + ' 個字元超過欄位' + i_max_ReviseCont_len + '個字元最大限制!\n'; }
                /**---*/
                if (document.getElementById('changeReason').value == 0) msg += '請選擇變更原因\n';
            }
            else {
                msg += '請選擇要變更項目\n';
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function SetHours() {
            var num = 0;
            if (isUnsignedInt(document.getElementById('EGenSci').value)) num += parseInt(document.getElementById('EGenSci').value);
            if (isUnsignedInt(document.getElementById('EProSci').value)) num += parseInt(document.getElementById('EProSci').value);
            document.getElementById('ESumSci').innerHTML = num;
        }

        function ClearLessonTeah() {
            document.getElementById('OLessonTeah1Value').value = "";
            document.getElementById('OLessonTeah2Value').value = "";
            document.getElementById('OLessonTeah3Value').value = "";
            document.getElementById('OLessonTeah1').value = "";
            document.getElementById('OLessonTeah2').value = "";
            document.getElementById('OLessonTeah3').value = "";
            return false;
        }

        function LessonTeah3(opentype, st, fieldname, hiddenname) {
            var ExistTech1 = document.getElementById('OLessonTeah1Value');
            var ExistTech2 = document.getElementById('OLessonTeah2Value');
            var ExistTech3 = document.getElementById('OLessonTeah3Value');
            var RIDValue = document.getElementById('RIDValue');
            if (st == '1') {
                //排除助教1
                wopen('../../SD/04/LessonTeah1.aspx?RID=' + RIDValue.value + '&type=' + opentype + '&fieldname=' + fieldname + '&hiddenname=' + hiddenname + '&ExistTech=' + ExistTech2, 'LessonTeah1', 400, 300, 1);
            }
            if (st == '2') {
                //排除師資1
                if (ExistTech1.value == '') { alert('請先選擇師資(一)!'); return false; }
                wopen('../../SD/04/LessonTeah1.aspx?RID=' + RIDValue.value + '&type=' + opentype + '&fieldname=' + fieldname + '&hiddenname=' + hiddenname + '&ExistTech=' + ExistTech1, 'LessonTeah2', 400, 300, 1);
            }
            if (st == '3') {
                //排除助教1
                //if (ExistTech1.value == '') { alert('請先選擇師資(一)!'); return false; }
                if (ExistTech2.value == '') { alert('請先選擇助教(一)!'); return false; }
                wopen('../../SD/04/LessonTeah1.aspx?RID=' + RIDValue.value + '&type=' + opentype + '&fieldname=' + fieldname + '&hiddenname=' + hiddenname + '&ExistTech=' + ExistTech2, 'LessonTeah3', 400, 300, 1);
            }
        }

        /* function LessonTeah(opentype, fieldname) {,var ExistTech1 = document.getElementById('OLessonTeah1Value').value;,
var ExistTech2 = document.getElementById('OLessonTeah2Value').value;,if (opentype == 'Add') {,
wopen('../../SD/04/LessonTeah1.aspx?RID=' + document.getElementById('RID').value + '&type=' + opentype + '&fieldname=' + fieldname + '&ExistTech=' + ExistTech2,'LessonTeah1', 400, 400, 1);,
},else if (opentype == 'Add2') {,if (ExistTech1 == '') { alert('請先選擇師資(一)!'); },
else {,wopen('../../SD/04/LessonTeah1.aspx?RID=' + document.getElementById('RID').value + '&type=' + opentype + '&fieldname=' + fieldname + '&ExistTech=' + ExistTech1, 
'LessonTeah1', 400, 400, 1);,},},}, */

        //檢查上課時間
        function CheckAddTime() {
            var msg = '';
            if (document.getElementById('Weeks').selectedIndex == 0) msg += '請選擇星期\n';
            if (document.getElementById('Times').value == '') msg += '請輸入上課時間\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function chackAll(i_cells) {
            var Mytable = document.getElementById('DataGrid3');
            var jChoose1 = $('#Choose1');
            for (var i = 1; i < Mytable.rows.length; i++) {
                var mycheck = Mytable.rows[i].cells[i_cells].children[0];
                //document.form1.Choose1.checked;
                if (!mycheck.disabled) { mycheck.checked = jChoose1.prop("checked"); }
            }
        }

        function reset_Choose1() { $('#Choose1').prop('checked', false); }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_nw" id="Table2" width="100%" cellspacing="1" cellpadding="1">
                        <tr>
                            <td colspan="4" class="table_title" width="100%">班級變更資訊</td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="16%">年度 </td>
                            <td class="whitecol" width="34%">
                                <asp:Label ID="YearList" runat="server"></asp:Label><asp:Label ID="labAPPSTAGE" runat="server"></asp:Label></td>
                            <td class="bluecol" width="16%">
                                <asp:Label ID="LabTMID" runat="server">訓練職類</asp:Label></td>
                            <td class="whitecol" width="34%">
                                <asp:Label ID="TrainText" runat="server"></asp:Label><asp:Label ID="JobText" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                <asp:Label ID="Labcjob" runat="server">通俗職類</asp:Label></td>
                            <td class="whitecol">
                                <asp:Label ID="CjobName" runat="server"></asp:Label></td>
                            <td class="bluecol">申請人姓名</td>
                            <td class="whitecol">
                                <asp:Label ID="lab_REVISEACCT_Name" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練機構 </td>
                            <td class="whitecol">
                                <asp:Label ID="OrgName" runat="server"></asp:Label>
                                <input id="RIDValue" type="hidden" runat="server" />
                            </td>
                            <td class="bluecol">班別名稱 </td>
                            <td class="whitecol">
                                <asp:Label ID="ClassName" runat="server"></asp:Label>
                                <asp:Label ID="PointYN" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練期間 </td>
                            <td class="whitecol">
                                <asp:Label ID="TRange" runat="server"></asp:Label></td>
                            <td class="bluecol">是否轉班 </td>
                            <td class="whitecol">
                                <asp:Label ID="ClassFlag" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">查詢模式 </td>
                            <td class="whitecol">
                                <asp:Label ID="SearchMode" runat="server"></asp:Label></td>
                            <td class="bluecol">
                                <asp:Label ID="labTitle" runat="server"></asp:Label></td>
                            <td class="whitecol">
                                <asp:Label ID="CheckMode" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">變更項目 </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="ChgItem" runat="server">
                                    <asp:ListItem Value="0">==請選擇==</asp:ListItem>
                                    <asp:ListItem Value="1">訓練期間</asp:ListItem>
                                    <asp:ListItem Value="2">訓練時段(課程互換)</asp:ListItem>
                                    <asp:ListItem Value="3">訓練課程地點</asp:ListItem>
                                    <asp:ListItem Value="4">課程編配</asp:ListItem>
                                    <asp:ListItem Value="5">訓練師資</asp:ListItem>
                                    <asp:ListItem Value="6">班別名稱</asp:ListItem>
                                    <asp:ListItem Value="7">期別</asp:ListItem>
                                    <asp:ListItem Value="8">上課地址</asp:ListItem>
                                    <asp:ListItem Value="9">停辦</asp:ListItem>
                                    <asp:ListItem Value="10">上課時段</asp:ListItem>
                                    <asp:ListItem Value="11">師資</asp:ListItem>
                                    <asp:ListItem Value="20">助教</asp:ListItem>
                                    <asp:ListItem Value="12">核定人數</asp:ListItem>
                                    <asp:ListItem Value="13">增班</asp:ListItem>
                                    <asp:ListItem Value="14">學(術)科場地</asp:ListItem>
                                    <asp:ListItem Value="15">上課時間</asp:ListItem>
                                    <asp:ListItem Value="18">課程表</asp:ListItem>
                                    <asp:ListItem Value="17">報名日期</asp:ListItem>
                                    <asp:ListItem Value="19">包班種類</asp:ListItem>
                                    <asp:ListItem Value="21">訓練費用</asp:ListItem>
                                    <asp:ListItem Value="16">其他</asp:ListItem>
                                </asp:DropDownList>
                                <input id="chgState" type="hidden" value="0" name="chgState" runat="server" />
                            </td>
                            <td class="bluecol_need">申請變更日 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="ApplyDate" runat="server" onfocus="this.blur()" Width="40%"></asp:TextBox>
                                <img id="imgApplyDate" style="cursor: pointer" onclick="javascript:show_calendar('ApplyDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="middle" runat="server" width="30" height="30" />
                            </td>
                        </tr>
                    </table>
                    <table class="table_nw" id="ReviseTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr id="TR1_1" runat="server">
                            <td class="bluecol">注意事項</td>
                            <td class="whitecol">
                                <span style="color: #FF0000; font-weight: bolder;">*</span>當計畫變更審核完成後,將會自動修正計畫申請及開班資料中開結訓日期<br>
                                <span style="color: #FF0000; font-weight: bolder;">*</span>另開班資料中報名開始/結束日期/報到日期請自行修正<br>
                                <span style="color: #FF0000; font-weight: bolder;">*倘上課時間或星期有變動，需再變更「上課時間」</span>
                            </td>
                        </tr>
                        <tr id="TR1_2" runat="server">
                            <td class="bluecol" width="16%">原計畫內容 </td>
                            <td class="whitecol" width="84%">
                                <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td>訓練期間起訖日期： </td>
                                    </tr>
                                    <tr>
                                        <td>自<asp:Label ID="BSDate" runat="server"></asp:Label>
                                            至<asp:Label ID="BEDate" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                </table>
                                <table class="font" id="tb_EnterDate2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                    <tr>
                                        <td>報名起訖日期： </td>
                                    </tr>
                                    <tr>
                                        <td>自<asp:Label ID="Old_SEnterDate2" runat="server"></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td>至<asp:Label ID="Old_FEnterDate2" runat="server"></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td>甄試日期： </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:Label ID="Old_Examdate" runat="server"></asp:Label>
                                            <asp:Label ID="Old_ExamPeriod" runat="server"></asp:Label>
                                            <asp:HiddenField ID="HidOld_ExamPeriod" runat="server" />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>報到日期： </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:Label ID="Old_CheckInDate" runat="server"></asp:Label></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr id="TR1_3" runat="server">
                            <td class="bluecol">變更內容 </td>
                            <td class="whitecol">
                                <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td>訓練期間起訖日期：<font color="red">(請記得於本變更申請審核通過後，修正課程表。)</font> </td>
                                    </tr>
                                    <tr>
                                        <td>自<asp:TextBox ID="ASDate" runat="server" onfocus="this.blur()" Width="15%"></asp:TextBox>
                                            <img id="IMG1" style="cursor: pointer" onclick="openCalendar('ASDate','2000/1/1','2100/1/1',document.getElementById('STDate').value);" alt="" src="../../images/show-calendar.gif" align="middle" runat="server" width="30" height="30" />
                                            至<asp:TextBox ID="AEDate" runat="server" onfocus="this.blur()" Width="15%"></asp:TextBox>
                                            <img id="IMG2" style="cursor: pointer" onclick="openCalendar('AEDate','2000/1/1','2100/1/1',document.getElementById('FDDate').value);" alt="" src="../../images/show-calendar.gif" align="middle" runat="server" width="30" height="30" />
                                        </td>
                                    </tr>
                                </table>
                                <table class="font" id="tb_New_EnterDate2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                    <tr>
                                        <td>報名起訖日期： </td>
                                    </tr>
                                    <tr>
                                        <td>自<asp:TextBox ID="New_SEnterDate2" runat="server" onfocus="this.blur()" Width="15%"></asp:TextBox>
                                            <img id="Img11" style="cursor: pointer" onclick="openCalendar('New_SEnterDate2','2000/1/1','2100/1/1',document.getElementById('New_SEnterDate2').value);" alt="" src="../../images/show-calendar.gif" align="middle" runat="server" width="30" height="30" />
                                            <asp:DropDownList ID="HR3" runat="server"></asp:DropDownList>時
                                            <asp:DropDownList ID="MM3" runat="server"></asp:DropDownList>分
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>至<asp:TextBox ID="New_FEnterDate2" runat="server" onfocus="this.blur()" Width="15%"></asp:TextBox>
                                            <img id="Img12" style="cursor: pointer" onclick="openCalendar('New_FEnterDate2','2000/1/1','2100/1/1',document.getElementById('New_FEnterDate2').value);" alt="" src="../../images/show-calendar.gif" align="middle" runat="server" width="30" height="30" />
                                            <asp:DropDownList ID="HR4" runat="server"></asp:DropDownList>時
                                            <asp:DropDownList ID="MM4" runat="server"></asp:DropDownList>分
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>甄試日期：<asp:Label ID="lab_New_Examdate" runat="server" Text=""></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:TextBox ID="New_Examdate" runat="server" Width="15%"></asp:TextBox>
                                            <img id="Img13" style="cursor: pointer" onclick="openCalendar('New_Examdate','2000/1/1','2100/1/1',document.getElementById('New_Examdate').value);" alt="" src="../../images/show-calendar.gif" align="middle" runat="server" width="30" height="30" />
                                            <asp:DropDownList ID="New_ExamPeriod" runat="server"></asp:DropDownList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>報到日期： </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:TextBox ID="New_CheckInDate" runat="server" onfocus="this.blur()" Width="15%"></asp:TextBox>
                                            <img id="Img14" style="cursor: pointer" onclick="openCalendar('New_CheckInDate','2000/1/1','2100/1/1',document.getElementById('New_CheckInDate').value);" alt="" src="../../images/show-calendar.gif" align="middle" runat="server" width="30" height="30" />
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr id="TR2_1" runat="server">
                            <td class="bluecol" width="16%">原計畫內容 </td>
                            <td class="whitecol">
                                <table class="font" id="Main2_1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                    <tr>
                                        <td>日期：<asp:TextBox ID="TimeSDate" runat="server" onfocus="this.blur()" Width="15%" AutoPostBack="True"></asp:TextBox>
                                            <img id="IMG3" style="cursor: pointer" onclick="openCalendar('TimeSDate',document.getElementById('STDate').value,document.getElementById('FDDate').value,document.getElementById('ApplyDate').value,'','Button2');" alt="" src="../../images/show-calendar.gif" align="middle" runat="server" width="30" height="30" />
                                            <asp:Label ID="msg1" runat="server" ForeColor="Red"></asp:Label>
                                            <asp:Button ID="Button2" Style="display: none" runat="server" Text="顯示課程"></asp:Button>
                                            <asp:Label ID="msg5" runat="server" ForeColor="Red"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:CheckBoxList ID="TimeSClass" runat="server" CssClass="font" RepeatColumns="2" RepeatLayout="Flow"></asp:CheckBoxList></td>
                                    </tr>
                                </table>
                                <table class="font" id="Sub2_1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                    <tr>
                                        <td colspan="2">
                                            <asp:Label ID="Stime" runat="server"></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td style="width: 10%" valign="top">課程： </td>
                                        <td>
                                            <asp:Label ID="EditSClass" runat="server" Visible="False"></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td style="width: 10%" valign="top">節次： </td>
                                        <td>
                                            <asp:Label ID="EditSClassItem" runat="server" Visible="False"></asp:Label></td>
                                    </tr>
                                </table>
                                <!---變更訓時段(變更前) S  -->
                                <div id="divlist01">
                                    <table id="tb_ClassChg1" runat="server">
                                        <tr>
                                            <td>
                                                <table class="font">
                                                    <tr>
                                                        <td>可申請更換訓練時段 </td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <div id="div2_1_1" style="z-index: 102; position: relative; width: 200px; height: 200px; overflow: auto; top: 0px; left: 0px">
                                                                <asp:ListBox ID="SourceLB1" runat="server" Width="300px" Height="200px" SelectionMode="Multiple"></asp:ListBox>
                                                            </div>
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>
                                            <td>
                                                <div>
                                                    <asp:ImageButton ID="btnAdd_1" runat="server" ForeColor="Transparent" Visible="False" BackColor="Transparent" ImageUrl="../../images/right2.gif"></asp:ImageButton>
                                                </div>
                                                <div>
                                                    <asp:ImageButton ID="btnAddAll_1" runat="server" Visible="False" BackColor="Transparent" ImageUrl="../../images/right.gif"></asp:ImageButton>
                                                </div>
                                                <div>
                                                    <asp:ImageButton ID="btnRemove_1" runat="server" Visible="False" BackColor="Transparent" ImageUrl="../../images/left2.gif"></asp:ImageButton>
                                                </div>
                                                <div>
                                                    <asp:ImageButton ID="btnRemoveAll_1" runat="server" Visible="False" BackColor="Transparent" ImageUrl="../../images/left.gif"></asp:ImageButton>
                                                </div>
                                            </td>
                                            <td>
                                                <table class="font">
                                                    <tr>
                                                        <td>提出申請更換訓練時段 </td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <div id="div2_1_2" style="position: relative; width: 200px; height: 200px; overflow: auto; top: 0px">
                                                                <asp:ListBox ID="TargetLB1" runat="server" Width="300px" Height="200px" SelectionMode="Multiple"></asp:ListBox>
                                                            </div>
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>
                                        </tr>
                                    </table>
                                </div>
                            </td>
                        </tr>
                        <tr id="TR2_2" runat="server">
                            <td class="bluecol">變更內容 </td>
                            <td class="whitecol">
                                <table class="font" id="Main2_2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                    <tr>
                                        <td>日期：<asp:TextBox ID="TimeEDate" runat="server" onfocus="this.blur()" Width="15%" AutoPostBack="True"></asp:TextBox>
                                            <img id="IMG4" style="cursor: pointer" onclick="openCalendar('TimeEDate',document.getElementById('STDate').value,document.getElementById('FDDate').value,document.getElementById('ApplyDate').value,'','Button3');" alt="" src="../../images/show-calendar.gif" align="middle" runat="server" width="30" height="30" />
                                            <asp:Label ID="msg2" runat="server" ForeColor="Red"></asp:Label>
                                            <asp:Button ID="Button3" Style="display: none" runat="server" Text="顯示課程"></asp:Button>
                                            <input id="hid_chklist" type="hidden" name="hid_chklist" runat="server" />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:CheckBoxList ID="TimeEClass" runat="server" CssClass="font" RepeatColumns="2" RepeatLayout="Flow"></asp:CheckBoxList></td>
                                    </tr>
                                </table>
                                <table class="font" id="Sub2_2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                    <tr>
                                        <td colspan="2">
                                            <asp:Label ID="Etime" runat="server"></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td style="width: 10%" valign="top">課程： </td>
                                        <td>
                                            <asp:Label ID="EditEClass" runat="server" Visible="False"></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td style="width: 10%" valign="top">節次： </td>
                                        <td>
                                            <asp:Label ID="EditEClassItem" runat="server" Visible="False"></asp:Label></td>
                                    </tr>
                                </table>
                                <!---變更訓時段(變更後) S  -->
                                <div id="divlist02">
                                    <table id="tb_ClassChg2" runat="server">
                                        <tr>
                                            <td>
                                                <table class="font">
                                                    <tr>
                                                        <td>可申請更換訓練時段 </td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <div id="div2_2_1" style="z-index: 102; position: relative; width: 200px; height: 200px; overflow: auto; top: 0px; left: 0px">
                                                                <asp:ListBox ID="SourceLB2" runat="server" Width="300px" Height="200px" SelectionMode="Multiple"></asp:ListBox>
                                                            </div>
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>
                                            <td>
                                                <div>
                                                    <asp:ImageButton ID="btnAdd_2" runat="server" Visible="False" BackColor="Transparent" ImageUrl="../../images/right2.gif"></asp:ImageButton>
                                                </div>
                                                <div>
                                                    <asp:ImageButton ID="btnAddAll_2" runat="server" Visible="False" BackColor="Transparent" ImageUrl="../../images/right.gif"></asp:ImageButton>
                                                </div>
                                                <div>
                                                    <asp:ImageButton ID="btnRemove_2" runat="server" Visible="False" BackColor="Transparent" ImageUrl="../../images/left2.gif"></asp:ImageButton>
                                                </div>
                                                <div>
                                                    <asp:ImageButton ID="btnRemoveAll_2" runat="server" Visible="False" BackColor="Transparent" ImageUrl="../../images/left.gif"></asp:ImageButton>
                                                </div>
                                            </td>
                                            <td>
                                                <table class="font">
                                                    <tr>
                                                        <td>提出申請更換訓練時段 </td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <div id="div2_2_2" style="z-index: 102; position: relative; width: 200px; height: 200px; overflow: auto; top: 0px; left: 0px">
                                                                <asp:ListBox ID="TargetLB2" runat="server" Width="300px" Height="200px" SelectionMode="Multiple"></asp:ListBox>
                                                            </div>
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>
                                        </tr>
                                    </table>
                                </div>
                                <!---變更訓時段(變更後) E  -->
                            </td>
                        </tr>
                        <!--- 修改訓練時段 start--->
                        <%--<tr><td class="bluecol" style="width:20%"></td><td></td></tr>--%>
                        <!--- 修改訓練時段 end--->
                        <tr id="TR3_1" runat="server">
                            <td class="bluecol" width="16%">原計畫內容 </td>
                            <td class="whitecol">
                                <table class="font" id="Main3_1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                    <tr>
                                        <td>日期：<asp:TextBox ID="PlaceDate" runat="server" onfocus="this.blur()" Width="15%" AutoPostBack="True"></asp:TextBox>
                                            <img id="IMG5" style="cursor: pointer" onclick="openCalendar('PlaceDate',document.getElementById('STDate').value,document.getElementById('FDDate').value,document.getElementById('ApplyDate').value,'','Button4');" alt="" src="../../images/show-calendar.gif" align="middle" runat="server" width="30" height="30" />
                                            <asp:Label ID="msg3" runat="server" ForeColor="Red"></asp:Label><asp:Button ID="Button4" Style="display: none" runat="server" Text="顯示課程"></asp:Button>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:CheckBoxList ID="SPlace" runat="server" CssClass="font" RepeatColumns="2" RepeatLayout="Flow"></asp:CheckBoxList></td>
                                    </tr>
                                </table>
                                <table class="font" id="Sub3_1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                    <tr>
                                        <td>地點：<asp:Label ID="EditPlace" runat="server" Visible="False"></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td>節次：<asp:Label ID="EditPlaceItem" runat="server" Visible="False"></asp:Label></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr id="TR3_2" runat="server">
                            <td class="bluecol">變更內容 </td>
                            <td class="whitecol">地點：<asp:TextBox ID="EPlace" runat="server"></asp:TextBox><font color="red">最多只能輸入12個中文字</font></td>
                        </tr>
                        <tr id="TR4_1" runat="server">
                            <td class="bluecol" width="16%">原計畫內容 </td>
                            <td class="whitecol">
                                <table class="font" id="Table9" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td>學科：<asp:Label ID="SSumSci" runat="server"></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td>一般學科：<asp:Label ID="SGenSci" runat="server"></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td>專業學科：<asp:Label ID="SProSci" runat="server"></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td>術科：<asp:Label ID="SProTech" runat="server"></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td>其他：<asp:Label ID="SOther" runat="server"></asp:Label></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr id="TR4_2" runat="server">
                            <td class="bluecol">變更內容 </td>
                            <td class="whitecol">
                                <table class="font" id="Table10" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td>學科：<asp:Label ID="ESumSci" runat="server"></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td>一般學科：<asp:TextBox ID="EGenSci" runat="server" Columns="5" Width="20%"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td>專業學科：<asp:TextBox ID="EProSci" runat="server" Columns="5" Width="20%"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td>術科：<asp:TextBox ID="EProTech" runat="server" Columns="5" Width="20%"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td>其他：<asp:TextBox ID="EOther" runat="server" Columns="5" Width="20%"></asp:TextBox></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr id="TR5_1" runat="server">
                            <td class="bluecol" width="16%">原計畫內容 </td>
                            <td class="whitecol">
                                <table class="font" id="Main5_1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                    <tr>
                                        <td>日期：<asp:TextBox ID="TechDate" runat="server" onfocus="this.blur()" Width="15%" AutoPostBack="True"></asp:TextBox>
                                            <img id="IMG6" style="cursor: pointer" onclick="openCalendar('TechDate',document.getElementById('STDate').value,document.getElementById('FDDate').value,document.getElementById('ApplyDate').value,'','Button5');" alt="" src="../../images/show-calendar.gif" align="middle" runat="server" width="30" height="30" />
                                            <asp:Label ID="msg4" runat="server" ForeColor="Red"></asp:Label><asp:Button ID="Button5" Style="display: none" runat="server" Text="顯示課程"></asp:Button>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:CheckBoxList ID="STeacher" runat="server" CssClass="font" RepeatColumns="2" RepeatLayout="Flow"></asp:CheckBoxList></td>
                                    </tr>
                                </table>
                                <table class="font" id="Sub5_1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                    <tr>
                                        <td title="師資(一、二)： ">師資(一)：
                                            <asp:Label ID="EditTech" runat="server" Visible="False"></asp:Label>&nbsp;&nbsp;
                                            <asp:Label ID="Tech2Label" runat="server" Visible="False">&nbsp;,&nbsp;助教(1)：</asp:Label>
                                            <asp:Label ID="EditTech2" runat="server" Visible="False"></asp:Label>&nbsp;&nbsp;
                                            <asp:Label ID="Tech3Label" runat="server" Visible="False">&nbsp;,&nbsp;助教(2)：</asp:Label>
                                            <asp:Label ID="EditTech3" runat="server" Visible="False"></asp:Label>&nbsp;&nbsp;
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>節次：<asp:Label ID="EditTechItem" runat="server" Visible="False"></asp:Label></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr id="TR5_2" runat="server">
                            <td class="bluecol">變更內容</td>
                            <td class="whitecol">師資姓名：<asp:TextBox ID="OLessonTeah1" Style="cursor: pointer" runat="server" onfocus="this.blur()" Width="20%"></asp:TextBox>
                                &nbsp;,&nbsp;助教姓名(1)：<asp:TextBox ID="OLessonTeah2" Style="cursor: pointer" runat="server" onfocus="this.blur()" Width="20%"></asp:TextBox>
                                &nbsp;,&nbsp;助教姓名(2)：<asp:TextBox ID="OLessonTeah3" Style="cursor: pointer" runat="server" onfocus="this.blur()" Width="20%"></asp:TextBox>
                                &nbsp;&nbsp;&nbsp;&nbsp;<asp:Button ID="bt_clearTech" runat="server" Text="清除"></asp:Button>
                            </td>
                        </tr>
                        <tr id="TR6_1" runat="server">
                            <td colspan="2">當計畫變更審核完成後,將會自動修正計畫申請及開班資料中的班別名稱</td>
                        </tr>
                        <tr id="TR6_2" runat="server">
                            <td class="bluecol" width="16%">原計畫內容 </td>
                            <td>班別名稱：<asp:Label ID="ClassCName" runat="server"></asp:Label></td>
                        </tr>
                        <tr id="TR6_3" runat="server">
                            <td class="bluecol">變更內容</td>
                            <td class="whitecol">班別名稱：<asp:TextBox ID="ChangeClassCName" runat="server" Width="40%"></asp:TextBox></td>
                        </tr>
                        <tr id="TR7_1" runat="server">
                            <td colspan="2">當計畫變更審核完成後,將會自動修正計畫申請及開班資料中的期別</td>
                        </tr>
                        <tr id="TR7_2" runat="server">
                            <td class="bluecol" width="16%">原計畫內容 </td>
                            <td>
                                <table class="font" id="Table12" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td>班別名稱：<asp:Label ID="ClassCName2" runat="server"></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td>期別：<asp:Label ID="CyclType" runat="server"></asp:Label></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr id="TR7_3" runat="server">
                            <td class="bluecol">變更內容 </td>
                            <td class="whitecol">期別：<asp:TextBox ID="ChangeCyclType" runat="server" Width="15%" MaxLength="2"></asp:TextBox>(<font color="red">2碼</font>) </td>
                        </tr>
                        <tr id="TR8_1" runat="server">
                            <td colspan="2">當計畫變更審核完成後,將會自動修正計畫申請及開班資料中的訓練地點 </td>
                        </tr>
                        <tr id="TR8_2" runat="server">
                            <td class="bluecol" width="16%">原計畫內容 </td>
                            <td class="whitecol">
                                <asp:Label ID="TAddress" runat="server"></asp:Label>
                                <input id="OldData8_1" type="hidden" name="OldData8_1" runat="server" />
                                <input id="OldData8_2" type="hidden" name="OldData8_2" runat="server" />
                                <input id="OldData8_3" type="hidden" name="OldData8_3" runat="server" /></td>
                        </tr>
                        <tr id="TR8_3" runat="server">
                            <td class="bluecol">變更內容 </td>
                            <td class="whitecol">
                                <input id="NewData8_1" onfocus="this.blur()" maxlength="5" name="NewData8_1" runat="server" />－
                                <input id="NewData8_3" maxlength="3" name="NewData8_3" runat="server" />
                                <input id="hidNewData8_6W" type="hidden" runat="server" />
                                <asp:Literal ID="LitNewData8_1" runat="server"></asp:Literal><%--郵遞區號--%><br />
                                <asp:TextBox ID="CTName" runat="server" onfocus="this.blur()" Columns="15" Width="20%"></asp:TextBox>
                                <input id="Button1" type="button" value="..." name="Button1" runat="server" />
                                <asp:TextBox ID="NewData8_2" runat="server" Columns="40" Width="60%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr id="TR9_1" runat="server">
                            <td class="bluecol" width="16%">原計畫狀態 </td>
                            <td>
                                <asp:Label ID="OldData9_1" runat="server"></asp:Label></td>
                        </tr>
                        <tr id="TR9_2" runat="server">
                            <td class="bluecol">變更內容 </td>
                            <td>
                                <asp:CheckBox ID="NewData9_1" runat="server" Text="申請停辦"></asp:CheckBox></td>
                        </tr>
                        <tr id="TR10_1" runat="server">
                            <td class="bluecol" width="16%">原計畫內容 </td>
                            <td>
                                <asp:Label ID="TrainTime" runat="server"></asp:Label>
                                <input id="OldData10_1" type="hidden" name="OldData10_1" runat="server" /></td>
                        </tr>
                        <tr id="TR10_2" runat="server">
                            <td class="bluecol">變更內容 </td>
                            <td>
                                <asp:DropDownList ID="NewData10_1" runat="server"></asp:DropDownList></td>
                        </tr>

                        <tr id="TR11_1" runat="server">
                            <td class="bluecol" width="16%">原計畫內容 </td>
                            <td>
                                <asp:Label ID="TeacherName1" runat="server"></asp:Label><input id="OldData11_1" type="hidden" name="OldData11_1" runat="server" /></td>
                        </tr>
                        <tr id="TR11_2" runat="server">
                            <td class="bluecol">變更內容 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TeacherName1_2" runat="server" onfocus="this.blur()" Columns="30" Rows="3" TextMode="MultiLine"></asp:TextBox>
                                <input id="Button6" type="button" value="選擇" name="Button6" runat="server" class="asp_button_M" />
                                <input id="NewData11_1" type="hidden" name="NewData11_1" runat="server" />
                                <asp:HiddenField ID="Hid_NewData11_3" runat="server" />
                                <%--<br />遴選辦法說明<br /><asp:TextBox ID="TeacherDesc11" runat="server" Width="80%" Height="80px" TextMode="MultiLine"></asp:TextBox>
                                <input id="btn_TCTYPEA_11" type="button" value="..." runat="server" class="button_b_Mini" />--%>
                            </td>
                        </tr>

                        <tr id="TR20_1" runat="server">
                            <td class="bluecol" width="16%">原計畫內容 </td>
                            <td class="whitecol">
                                <asp:Label ID="TeacherName2" runat="server"></asp:Label><input id="OldData20_1" type="hidden" name="OldData20_1" runat="server" /></td>
                        </tr>
                        <tr id="TR20_2" runat="server">
                            <td class="bluecol">變更內容 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TeacherName2_2" runat="server" onfocus="this.blur()" Columns="30" Rows="3" TextMode="MultiLine"></asp:TextBox>
                                <input id="Button6_2" type="button" value="選擇" name="Button6_2" runat="server" class="asp_button_M" />
                                <input id="NewData20_1" type="hidden" name="NewData20_1" runat="server" />
                                <asp:HiddenField ID="Hid_NewData20_3" runat="server" />
                                <%--<br />遴選辦法說明<br /><asp:TextBox ID="TeacherDesc20" runat="server" Width="80%" Height="80px" TextMode="MultiLine"></asp:TextBox>
                                <input id="btn_TCTYPEB_20" type="button" value="..." runat="server" class="button_b_Mini" />--%>
                            </td>
                        </tr>

                        <tr id="TR12_1" runat="server">
                            <td class="bluecol" width="16%">原計畫內容 </td>
                            <td class="whitecol">
                                <asp:Label ID="OldData12_1" runat="server"></asp:Label>人 </td>
                        </tr>
                        <tr id="TR12_2" runat="server">
                            <td class="bluecol">變更內容 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="NewData12_1" runat="server" Columns="5" Width="10%"></asp:TextBox>人 </td>
                        </tr>
                        <tr id="TR13_1" runat="server">
                            <td class="bluecol" width="16%">原計畫內容 </td>
                            <td class="whitecol">
                                <asp:Label ID="OldData13_1" runat="server"></asp:Label>班 </td>
                        </tr>
                        <tr id="TR13_2" runat="server">
                            <td class="bluecol">變更內容 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="NewData13_1" runat="server" Columns="5" Width="10%"></asp:TextBox>班 </td>
                        </tr>
                        <%--
						<TR id="TR14_1" runat="server">
						    <td class="bluecol" width="16%">原計畫內容 </td>
							<TD>
                                &nbsp; 學科場地&nbsp;<asp:label id="SciPlaceID" runat="server" Width="420px" BorderStyle="Groove"></asp:label>
							    <INPUT id="OldData14_1" type="hidden" name="OldData14_1" runat="server" /><BR>&nbsp; 術科場地&nbsp;
							    <asp:label id="TechPlaceID" runat="server" Width="420px" BorderStyle="Ridge"></asp:label>
								<INPUT id="OldData14_2" type="hidden" name="OldData14_2" runat="server" />
							</TD>
						</TR>
						<TR id="TR14_2" runat="server" />
						    <TD class="TC_TD3">&nbsp;&nbsp;&nbsp; 變更內容</TD>
							<TD>
                                &nbsp; 學科場地&nbsp;<asp:dropdownlist id="NewData14_1" runat="server" Width="420px"></asp:dropdownlist><br>
								&nbsp; 術科場地&nbsp;<asp:dropdownlist id="NewData14_2" runat="server" Width="420px"></asp:dropdownlist>
                            </TD>
						</TR>
                        --%>
                        <tr id="TR14_1" runat="server">
                            <td class="bluecol" width="16%">原計畫內容 </td>
                            <td class="whitecol">&nbsp; 學科場地1／上課地址&nbsp;<asp:Label ID="SciPlaceIDb" runat="server"></asp:Label>
                                <input id="OldData14_1b" type="hidden" name="OldData14_1" runat="server" /><br>
                                &nbsp; 學科場地2／上課地址&nbsp;<asp:Label ID="SciPlaceID2" runat="server"></asp:Label>
                                <input id="OldData14_3" type="hidden" name="OldData14_3" runat="server" /><br>
                                &nbsp; 術科場地1／上課地址&nbsp;<asp:Label ID="TechPlaceIDb" runat="server"></asp:Label>
                                <input id="OldData14_2b" type="hidden" name="OldData14_2" runat="server" /><br>
                                &nbsp; 術科場地2／上課地址&nbsp;<asp:Label ID="TechPlaceID2" runat="server"></asp:Label>
                                <input id="OldData14_4" type="hidden" name="OldData14_4" runat="server" /><%--<br>--%>
                                <%--學科上課地址<asp:DropDownList ID="TaddressS1" runat="server" Enabled="False"></asp:DropDownList><br>
                                術科上課地址<asp:DropDownList ID="TaddressT1" runat="server" Enabled="False"></asp:DropDownList>--%>
                                <asp:HiddenField ID="Hid_OldData8_4" runat="server" />
                                <asp:HiddenField ID="Hid_OldData8_5" runat="server" />
                                <asp:HiddenField ID="Hid_OldData8_6" runat="server" />
                                <asp:HiddenField ID="Hid_OldData8_7" runat="server" />
                            </td>
                        </tr>
                        <tr id="TR14_2" runat="server">
                            <td class="bluecol">變更內容 </td>
                            <td class="whitecol">&nbsp; 學科場地1／上課地址&nbsp;<asp:DropDownList ID="NewData14_1b" runat="server"></asp:DropDownList>
                                <br>
                                &nbsp; 學科場地2／上課地址&nbsp;<asp:DropDownList ID="NewData14_3" runat="server"></asp:DropDownList>
                                <br>
                                &nbsp; 術科場地1／上課地址&nbsp;<asp:DropDownList ID="NewData14_2b" runat="server"></asp:DropDownList>
                                <br>
                                &nbsp; 術科場地2／上課地址&nbsp;<asp:DropDownList ID="NewData14_4" runat="server"></asp:DropDownList>
                                <asp:HiddenField ID="Hid_NewData8_4" runat="server" />
                                <asp:HiddenField ID="Hid_NewData8_5" runat="server" />
                                <asp:HiddenField ID="Hid_NewData8_6" runat="server" />
                                <asp:HiddenField ID="Hid_NewData8_7" runat="server" />
                                <%--<br>
                                學科上課地址<asp:DropDownList ID="TaddressS2" runat="server"></asp:DropDownList>
                                <br>
                                術科上課地址<asp:DropDownList ID="TaddressT2" runat="server"></asp:DropDownList>--%>
                            </td>
                        </tr>
                        <tr id="TR15_1" runat="server">
                            <td class="bluecol" width="16%">原計畫內容 </td>
                            <td class="whitecol" width="84%">
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:TemplateColumn HeaderText="星期">
                                            <HeaderStyle Width="20%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="OldWeeks1" runat="server"></asp:Label>
                                            </ItemTemplate>
                                            <%--<EditItemTemplate> <asp:DropDownList ID="OldWeeks2" runat="server"> </asp:DropDownList> </EditItemTemplate>--%>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="上課時段">
                                            <HeaderStyle Width="80%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="OldTimes1" runat="server"></asp:Label>
                                            </ItemTemplate>
                                            <%--<EditItemTemplate> <asp:TextBox ID="OldTimes2" runat="server"></asp:TextBox> </EditItemTemplate>--%> 
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr id="TR15_2" runat="server">
                            <td class="bluecol">變更內容 </td>
                            <td>
                                <table class="font" id="Table7" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td style="width: 20%" class="bluecol">星期 </td>
                                        <td style="width: 70%" class="bluecol">時間 </td>
                                        <td style="width: 10%" class="bluecol">功能</td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">
                                            <asp:DropDownList ID="Weeks" runat="server"></asp:DropDownList></td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="txtTimes" runat="server" Columns="70" Width="95%"></asp:TextBox>
                                            <%--<asp:TextBox ID="Times" runat="server" Columns="50"></asp:TextBox>--%></td>
                                        <td align="center" class="whitecol">
                                            <asp:Button ID="Button29" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_M"></asp:Button></td>
                                    </tr>
                                </table>
                                <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:TemplateColumn HeaderText="星期">
                                            <HeaderStyle Width="20%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="NewWeeks1" runat="server"></asp:Label>
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:DropDownList ID="NewWeeks2" runat="server">
                                                </asp:DropDownList>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="上課時段">
                                            <HeaderStyle Width="70%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="NewTimes1" runat="server"></asp:Label>
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox ID="NewTimes2" runat="server"></asp:TextBox>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:Button ID="Button7" runat="server" Text="修改" CommandName="edit" CssClass="asp_button_M"></asp:Button>
                                                <asp:Button ID="Button8" runat="server" Text="刪除" CausesValidation="False" CommandName="del" CssClass="asp_button_M"></asp:Button>
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:Button ID="Button9" runat="server" Text="確定" CommandName="save" CssClass="asp_button_M"></asp:Button>
                                                <asp:Button ID="Button10" runat="server" Text="取消" CommandName="cancel" CssClass="asp_button_M"></asp:Button>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr id="TR16_1" runat="server">
                            <td class="bluecol" width="16%">原計畫<br>
                                &nbsp;&nbsp; 其他內容 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="OldData15_1" runat="server" Columns="30" Rows="3" TextMode="MultiLine"></asp:TextBox></td>
                        </tr>
                        <tr id="TR16_2" runat="server">
                            <td class="bluecol">變更後<br>
                                其他內容 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="NewData15_1" runat="server" Columns="30" Rows="3" TextMode="MultiLine"></asp:TextBox></td>
                        </tr>
                        <!--20080825 andy add 報名起迄日 Start-->
                        <tr id="Tr17_1" runat="server">
                            <td class="bluecol" width="16%">原計畫內容 </td>
                            <td class="whitecol">
                                <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td>報名起訖日期： </td>
                                    </tr>
                                    <tr>
                                        <td>自<asp:Label ID="Old_SEnterDate" runat="server"></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td>至<asp:Label ID="Old_FEnterDate" runat="server"></asp:Label></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr id="Tr17_2" runat="server">
                            <td class="bluecol">變更內容 </td>
                            <td class="whitecol">
                                <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td>報名起訖日期： </td>
                                    </tr>
                                    <tr>
                                        <td>自<asp:TextBox ID="New_SEnterDate" runat="server" onfocus="this.blur()" Width="15%"></asp:TextBox>
                                            <img id="Img7" style="cursor: pointer" onclick="openCalendar('New_SEnterDate','2000/1/1',document.getElementById('SignUpEDate').value,document.getElementById('SEnterDate').value);&#9;" alt="" src="../../images/show-calendar.gif" align="middle" runat="server" width="30" height="30" />
                                            <asp:DropDownList ID="HR1" runat="server"></asp:DropDownList>時
										    <asp:DropDownList ID="MM1" runat="server"></asp:DropDownList>分 
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>至<asp:TextBox ID="New_FEnterDate" runat="server" onfocus="this.blur()" Width="15%"></asp:TextBox>
                                            <img id="Img8" style="cursor: pointer" onclick="openCalendar('New_FEnterDate','2000/1/1',document.getElementById('SignUpEDate').value,document.getElementById('SEnterDate').value);" alt="" src="../../images/show-calendar.gif" align="middle" runat="server" width="30" height="30" />
                                            <asp:DropDownList ID="HR2" runat="server"></asp:DropDownList>時
                                            <asp:DropDownList ID="MM2" runat="server"></asp:DropDownList>分
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr id="TR19_1" runat="server">
                            <td class="bluecol" width="16%">原計畫內容 </td>
                            <td class="whitecol">
                                <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td>&nbsp;&nbsp; 包班種類：<asp:Label ID="PackageTypeOld" runat="server"></asp:Label><input id="hidPackageTypeOld" type="hidden" name="hidPackageTypeOld" runat="server" /></td>
                                    </tr>
                                    <tr>
                                        <td>&nbsp;&nbsp; 包班事業單位：
                                            <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                                                <tr>
                                                    <td>
                                                        <asp:DataGrid ID="DG_BusPackageOld" runat="server" Width="100%" CssClass="font" AlternatingItemStyle-BackColor="WhiteSmoke" CellPadding="8">
                                                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" CssClass="head_navy"></HeaderStyle>
                                                            <PagerStyle Visible="False"></PagerStyle>
                                                        </asp:DataGrid>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr id="TR19_2" runat="server">
                            <td class="bluecol">變更內容 </td>
                            <td>
                                <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td>包班種類：
                                            <asp:RadioButtonList ID="PackageTypeNew" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="2">企業包班</asp:ListItem>
                                                <asp:ListItem Value="3">聯合企業包班</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>包班事業單位：
                                            <table class="font" id="BusPackageNewHead" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                                <tr>
                                                    <td align="center" colspan="4" class="table_title">包班事業單位資料 </td>
                                                </tr>
                                                <tr>
                                                    <td align="center" class="bluecol">企業名稱 </td>
                                                    <td align="center" class="bluecol">服務單位統一編號 </td>
                                                    <td align="center" class="bluecol">保險證號 </td>
                                                    <td align="center" class="bluecol">功能</td>
                                                </tr>
                                                <tr class="whitecol" align="center">
                                                    <td>
                                                        <asp:TextBox ID="txtUname" runat="server" Columns="33" MaxLength="50"></asp:TextBox></td>
                                                    <td>
                                                        <asp:TextBox ID="txtIntaxno" runat="server" Columns="9" MaxLength="10"></asp:TextBox></td>
                                                    <td>
                                                        <asp:TextBox ID="txtUbno" runat="server" Columns="9" MaxLength="9"></asp:TextBox></td>
                                                    <td>
                                                        <asp:Button ID="btnAddBusPackage" runat="server" Text="新增" CausesValidation="False"></asp:Button></td>
                                                </tr>
                                            </table>
                                            <table class="font" id="BusPackageNewTable" cellspacing="1" cellpadding="1" width="100%" border="0">
                                                <tr>
                                                    <td>
                                                        <asp:DataGrid ID="DG_BusPackageNew" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                            <AlternatingItemStyle BackColor="WhiteSmoke" />
                                                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" CssClass="head_navy"></HeaderStyle>
                                                            <Columns>
                                                                <asp:TemplateColumn HeaderText="企業名稱">
                                                                    <HeaderStyle Width="30%" />
                                                                    <ItemStyle Wrap="False"></ItemStyle>
                                                                    <ItemTemplate>
                                                                        <asp:Label ID="LUname" runat="server"></asp:Label>
                                                                    </ItemTemplate>
                                                                    <EditItemTemplate>
                                                                        <asp:TextBox ID="TUname" runat="server"></asp:TextBox>
                                                                    </EditItemTemplate>
                                                                </asp:TemplateColumn>
                                                                <asp:TemplateColumn HeaderText="服務單位統一編號">
                                                                    <HeaderStyle Width="30%" />
                                                                    <ItemStyle Wrap="False"></ItemStyle>
                                                                    <ItemTemplate>
                                                                        <asp:Label ID="LIntaxno" runat="server"></asp:Label>
                                                                    </ItemTemplate>
                                                                    <EditItemTemplate>
                                                                        <asp:TextBox ID="TIntaxno" runat="server"></asp:TextBox>
                                                                    </EditItemTemplate>
                                                                </asp:TemplateColumn>
                                                                <asp:TemplateColumn HeaderText="保險證號">
                                                                    <HeaderStyle Width="30%" />
                                                                    <ItemStyle Wrap="False"></ItemStyle>
                                                                    <ItemTemplate>
                                                                        <asp:Label ID="LUbno" runat="server"></asp:Label>
                                                                    </ItemTemplate>
                                                                    <EditItemTemplate>
                                                                        <asp:TextBox ID="TUbno" runat="server"></asp:TextBox>
                                                                    </EditItemTemplate>
                                                                </asp:TemplateColumn>
                                                                <asp:TemplateColumn HeaderText="功能">
                                                                    <HeaderStyle Width="10%" />
                                                                    <ItemStyle Wrap="False" HorizontalAlign="Center"></ItemStyle>
                                                                    <ItemTemplate>
                                                                        <asp:Button ID="btnxEDT" runat="server" CausesValidation="False" Text="修改" CommandName="xedit" CssClass="asp_button_M"></asp:Button>
                                                                        <asp:Button ID="btnxDEL" runat="server" CausesValidation="False" Text="刪除" CommandName="xdel" CssClass="asp_button_M"></asp:Button>
                                                                    </ItemTemplate>
                                                                    <EditItemTemplate>
                                                                        <asp:Button ID="btnxSAV" runat="server" CausesValidation="False" Text="確定" CommandName="xsave" CssClass="asp_button_M"></asp:Button>
                                                                        <asp:Button ID="btnxCLS" runat="server" CausesValidation="False" Text="取消" CommandName="xcancel" CssClass="asp_button_M"></asp:Button>
                                                                    </EditItemTemplate>
                                                                </asp:TemplateColumn>
                                                            </Columns>
                                                            <PagerStyle Visible="False"></PagerStyle>
                                                        </asp:DataGrid>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr id="TR21_1" runat="server">
                            <td class="bluecol" width="16%">原計畫內容 </td>
                            <td class="whitecol">
                                <table id="tbDataGrid21Old" runat="server" class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td class="table_title">
                                            <asp:Label ID="labcost21txt1Old" runat="server" Text="費用列表"></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:DataGrid ID="DataGrid21Old_1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                <AlternatingItemStyle BackColor="#F5F5F5" />
                                                <Columns>
                                                    <asp:TemplateColumn HeaderText="項目">
                                                        <HeaderStyle Width="60%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="CostName" runat="server">CostName</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="單價">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:HiddenField ID="HidCostID" runat="server" />
                                                            <asp:Label ID="OPrice" runat="server">OPrice</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="數量">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="Itemage" runat="server">Itemage</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="計價單位">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="ItemCost" runat="server">ItemCost</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="小計">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="subtotal" runat="server">subtotal</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                </Columns>
                                            </asp:DataGrid>
                                            <asp:DataGrid ID="DataGrid21Old_2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                <AlternatingItemStyle BackColor="#F5F5F5" />
                                                <Columns>
                                                    <asp:TemplateColumn HeaderText="單價">
                                                        <HeaderStyle Width="25%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="OPrice" runat="server">OPrice</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="人數">
                                                        <HeaderStyle Width="25%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="Itemage" runat="server">Itemage</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="時數">
                                                        <HeaderStyle Width="25%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="ItemCost" runat="server">ItemCost</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="小計">
                                                        <HeaderStyle Width="25%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="subtotal" runat="server">subtotal</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                </Columns>
                                            </asp:DataGrid>
                                            <asp:DataGrid ID="DataGrid21Old_3" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                <AlternatingItemStyle BackColor="#F5F5F5" />
                                                <Columns>
                                                    <asp:TemplateColumn HeaderText="單價">
                                                        <HeaderStyle Width="33%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="OPrice" runat="server">OPrice</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="人數">
                                                        <HeaderStyle Width="33%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="Itemage" runat="server">Itemage</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="小計">
                                                        <HeaderStyle Width="34%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="subtotal" runat="server">subtotal</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                </Columns>
                                            </asp:DataGrid>
                                            <asp:DataGrid ID="DataGrid21Old_4" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                <AlternatingItemStyle BackColor="#F5F5F5" />
                                                <Columns>
                                                    <asp:TemplateColumn HeaderText="項目">
                                                        <HeaderStyle Width="70%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="CostName" runat="server">CostName</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="單價">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="OPrice" runat="server">OPrice</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="數量">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="Itemage" runat="server">Itemage</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="小計">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="subtotal" runat="server">subtotal</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                </Columns>
                                            </asp:DataGrid>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <table cellspacing="1" cellpadding="1" width="100%" border="0">
                                                <tr id="AdmGrantTROld" runat="server">
                                                    <td width="20%" class="bluecol">行政管理費 </td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="AdmCostOld" runat="server"></asp:Label></td>
                                                </tr>
                                                <tr id="TaxGrantTROld" runat="server">
                                                    <td class="bluecol">營業稅 </td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="TaxCostOld" runat="server"></asp:Label></td>
                                                </tr>
                                                <tr id="trTotalCost1Old" runat="server">
                                                    <td class="bluecol">總計 </td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="TotalCost1Old" runat="server"></asp:Label>
                                                        <input id="HidAdmGrantOld" type="hidden" value="0" runat="server" />
                                                        <input id="HidTaxGrantOld" type="hidden" value="0" runat="server" />
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr id="TR21_2" runat="server">
                            <td class="bluecol" width="16%">變更內容</td>
                            <td>
                                <table id="tbDataGrid21New" runat="server" class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td class="table_title">
                                            <asp:Label ID="labcost21txt1New" runat="server" Text="費用列表"></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:DataGrid ID="DataGrid21New_1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                <AlternatingItemStyle BackColor="#F5F5F5" />
                                                <Columns>
                                                    <asp:TemplateColumn HeaderText="項目">
                                                        <HeaderStyle Width="50%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="CostName" runat="server">CostName</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="單價">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:HiddenField ID="HidCostID" runat="server" />
                                                            <asp:Label ID="OPrice" runat="server">OPrice</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="數量">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="Itemage" runat="server">Itemage</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="計價單位">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="ItemCost" runat="server">ItemCost</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="小計">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="subtotal" runat="server">subtotal</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="功能">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <asp:Button ID="btnDel1" runat="server" CausesValidation="False" Text="刪除" CommandName="btnDel1" CssClass="asp_button_M"></asp:Button>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                </Columns>
                                            </asp:DataGrid>
                                            <asp:DataGrid ID="DataGrid21New_2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                <AlternatingItemStyle BackColor="#F5F5F5" />
                                                <Columns>
                                                    <asp:TemplateColumn HeaderText="單價">
                                                        <HeaderStyle Width="20%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="OPrice" runat="server">OPrice</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="人數">
                                                        <HeaderStyle Width="20%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="Itemage" runat="server">Itemage</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="時數">
                                                        <HeaderStyle Width="20%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="ItemCost" runat="server">ItemCost</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="小計">
                                                        <HeaderStyle Width="20%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="subtotal" runat="server">subtotal</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="功能">
                                                        <HeaderStyle Width="20%"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <asp:Button ID="btnDel1" runat="server" CausesValidation="False" Text="刪除" CommandName="btnDel1" CssClass="asp_button_M"></asp:Button>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                </Columns>
                                            </asp:DataGrid>
                                            <asp:DataGrid ID="DataGrid21New_3" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                <AlternatingItemStyle BackColor="#F5F5F5" />
                                                <Columns>
                                                    <asp:TemplateColumn HeaderText="單價">
                                                        <HeaderStyle Width="25%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="OPrice" runat="server">OPrice</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="人數">
                                                        <HeaderStyle Width="25%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="Itemage" runat="server">Itemage</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="小計">
                                                        <HeaderStyle Width="25%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="subtotal" runat="server">subtotal</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="功能">
                                                        <HeaderStyle Width="25%"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <asp:Button ID="btnDel1" runat="server" CausesValidation="False" Text="刪除" CommandName="btnDel1" CssClass="asp_button_M"></asp:Button>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                </Columns>
                                            </asp:DataGrid>
                                            <asp:DataGrid ID="DataGrid21New_4" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                <AlternatingItemStyle BackColor="#F5F5F5" />
                                                <Columns>
                                                    <asp:TemplateColumn HeaderText="項目">
                                                        <HeaderStyle Width="60%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="CostName" runat="server">CostName</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="單價">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="OPrice" runat="server">OPrice</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="數量">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="Itemage" runat="server">Itemage</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="小計">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="subtotal" runat="server">subtotal</asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="功能">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <asp:Button ID="btnDel1" runat="server" CausesValidation="False" Text="刪除" CommandName="btnDel1" CssClass="asp_button_M"></asp:Button>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                </Columns>
                                            </asp:DataGrid>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <table cellspacing="1" cellpadding="1" width="100%" border="0">
                                                <tr id="AdmGrantTRNew" runat="server">
                                                    <td class="bluecol" width="20%">行政管理費 </td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="AdmCostNew" runat="server"></asp:Label></td>
                                                </tr>
                                                <tr id="TaxGrantTRNew" runat="server">
                                                    <td class="bluecol">營業稅 </td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="TaxCostNew" runat="server"></asp:Label></td>
                                                </tr>
                                                <tr id="trTotalCost1New" runat="server">
                                                    <td class="bluecol">總計 </td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="TotalCost1New" runat="server"></asp:Label>
                                                        <input id="HidAdmGrantNew" type="hidden" value="0" runat="server" />
                                                        <input id="HidTaxGrantNew" type="hidden" value="0" runat="server" />
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr id="TR22_1" runat="server">
                            <td class="bluecol" width="16%">原計畫內容 </td>
                            <td class="whitecol">
                                <asp:HiddenField ID="Hid_DISTANCE" runat="server" />
                                <asp:Label ID="lab_DISTANCE" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr id="TR22_2" runat="server">
                            <td class="bluecol">變更內容 </td>
                            <td class="whitecol"><%--遠距教學--%>
                                <asp:RadioButtonList ID="rbl_DISTANCE" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow"></asp:RadioButtonList>
                            </td>
                        </tr>
                        <!-- end -->
                        <tr id="Result" runat="server">
                            <td class="bluecol">審核說明 </td>
                            <td>
                                <asp:Label ID="Reason" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="16%">變更原因 </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="changeReason" runat="server">
                                    <asp:ListItem Value="0">==請選擇==</asp:ListItem>
                                    <asp:ListItem Value="1">天然災害或政策因素</asp:ListItem>
                                    <asp:ListItem Value="2">其他</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">變更說明 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="ReviseCont" runat="server" MaxLength="125" Rows="5" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td align="center" colspan="2" class="whitecol">
                                <%-- isInsertNext = divInsertNext 判斷是否要執行 insert_next (0否, 1是) --%>
                                <div id="divInsertNext" style="display: none">
                                    <asp:Literal ID="ltlInserNextFlag" runat="server"></asp:Literal>
                                </div>
                                <asp:Button ID="But_Sub" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="But_Sub28" runat="server" Text="確定" Style="z-index: 0" CssClass="asp_button_M"></asp:Button>
                                <%-- input id="back" type="button" value="回上一頁" name="back" runat="server" &nbsp;--%>
                                <asp:Button ID="btn_back" runat="server" Text="回上一頁" CssClass="asp_button_M" />
                                <asp:Button ID="But_Save28" runat="server" Text="正式送出" CssClass="asp_button_M" ToolTip="儲存後正式送出"></asp:Button>
                                <%--<asp:Button ID="But_UPLOAD28" runat="server" Text="上傳檔案" CssClass="asp_button_M" ToolTip="儲存後進入上傳檔案"></asp:Button>--%>
                            </td>
                        </tr>

                        <tr id="TR11_3" runat="server">
                            <td colspan="2">
                                <div>
                                    <table class="font" id="tbDataGrid21" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                        <tr>
                                            <td class="table_title" align="center">授課教師 </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <asp:DataGrid ID="DataGrid21" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" CellPadding="8">
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                                    <Columns>
                                                        <asp:TemplateColumn HeaderText="序號">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" Width="8%"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="seqno" runat="server"></asp:Label>
                                                                <input id="HidTechID" runat="server" type="hidden" />
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="教師姓名">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" Width="14%"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="TeachCName" runat="server"></asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="學歷">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" Width="18%"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="DegreeName" runat="server"></asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="專業領域">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" Width="20%"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="Specialty1" runat="server"></asp:Label>
                                                                <%--<asp:Label ID="ProLicense" runat="server"></asp:Label>--%>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="遴選辦法說明">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:TextBox ID="TeacherDesc" runat="server" Width="80%" Height="80px" TextMode="MultiLine"></asp:TextBox>
                                                                <input id="btn_TCTYPEA" type="button" value="..." runat="server" class="button_b_Mini" />
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                    </Columns>
                                                </asp:DataGrid>
                                            </td>
                                        </tr>
                                    </table>
                                </div>
                            </td>
                        </tr>
                        <tr id="TR20_3" runat="server">
                            <td colspan="2">
                                <div>
                                    <table class="font" id="tbDataGrid22" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                        <tr>
                                            <td align="center" class="table_title">授課助教 </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <asp:DataGrid ID="DataGrid22" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" CellPadding="8">
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                                    <Columns>
                                                        <asp:TemplateColumn HeaderText="序號">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" Width="8%"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="seqno" runat="server"></asp:Label>
                                                                <input id="HidTechID" runat="server" type="hidden" />
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="助教姓名">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" Width="14%"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="TeachCName" runat="server"></asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="學歷">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" Width="18%"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="DegreeName" runat="server"></asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="專業領域">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" Width="20%"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="Specialty1" runat="server"></asp:Label>
                                                                <%--<asp:Label ID="ProLicense" runat="server"></asp:Label>--%>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="遴選辦法說明">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:TextBox ID="TeacherDesc" runat="server" Width="80%" Height="80px" TextMode="MultiLine"></asp:TextBox>
                                                                <input id="btn_TCTYPEB" type="button" value="..." runat="server" class="button_b_Mini" />
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                    </Columns>
                                                </asp:DataGrid>
                                            </td>
                                        </tr>
                                    </table>
                                </div>
                            </td>
                        </tr>

                    </table>
                </td>
            </tr>
        </table>
        <table width="100%">
            <!-- 20080627 andy add 課程表	start-->
            <tr id="Tr18" runat="server">
                <td colspan="4">
                    <table id="Datagrid3Table" runat="server" width="100%">
                        <tr>
                            <td class="table_title" id="TD_1" align="center" runat="server">課程表申請變更後</td>
                            <td class="table_title" id="TD_2" align="center" runat="server">課程表申請變更前</td>
                        </tr>
                        <tr>
                            <td valign="top" class="whitecol">
                                <%--變更後--%>
                                <asp:DataGrid ID="DataGrid3" runat="server" Width="100%" AutoGenerateColumns="False" CellPadding="8">
                                    <HeaderStyle CssClass="head_navy" />
                                    <ItemStyle></ItemStyle>
                                    <Columns>
                                        <asp:BoundColumn Visible="False" DataField="PTDID" HeaderText="PTDID"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="日期">
                                            <ItemTemplate>
                                                <asp:TextBox ID="txt_STrainDate" runat="server" Width="120px" onfocus="this.blur()"></asp:TextBox>
                                                <img id="Img9" style="cursor: pointer" alt="" src="../../images/show-calendar.gif" width="25" runat="server" />
                                                <input id="hide_STrainDate" type="hidden" runat="server" />
                                                <input id="hide_ID1" type="hidden" runat="server" />
                                                <input id="hide_PTDRID" type="hidden" runat="server" />
                                                <input id="hide_PTDID" type="hidden" runat="server" />
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                &nbsp;<asp:Button ID="Button12" Style="display: none" runat="server"></asp:Button>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="授課時段">
                                            <ItemStyle Wrap="false"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:CheckBox ID="TPERIOD28_1t" runat="server" Text="早上" ToolTip="7:00-13:00" Enabled="false" /><br />
                                                <asp:CheckBox ID="TPERIOD28_2t" runat="server" Text="下午" ToolTip="13:00-18:00" Enabled="false" /><br />
                                                <asp:CheckBox ID="TPERIOD28_3t" runat="server" Text="晚上" ToolTip="18:00-22:00" Enabled="false" />
                                                <input id="hidTPERIOD28" type="hidden" runat="server" />
                                                <input id="hide_ID7" type="hidden" runat="server" />
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                &nbsp;
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="授課時間">
                                            <ItemTemplate>
                                                <asp:TextBox ID="txt_PName" runat="server" Width="120px" MaxLength="30"></asp:TextBox>
                                                <input id="hide_PName" type="hidden" runat="server" />
                                                <input id="hide_ID2" type="hidden" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="時數">
                                            <ItemStyle Width="6%" Wrap="false" HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:TextBox ID="txt_PHour" runat="server" Width="60px" MaxLength="5"></asp:TextBox>
                                                <input id="hide_PHour" type="hidden" runat="server" />
                                                <input id="hide_ID3" type="hidden" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="技檢訓練時數">
                                            <ItemStyle Width="6%" Wrap="false" HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:TextBox ID="txt_EHour" runat="server" Width="60px" MaxLength="5"></asp:TextBox>
                                                <input id="hide_EHour" type="hidden" runat="server" />
                                                <input id="hide_ID9" type="hidden" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="課程進度／內容">
                                            <ItemTemplate>
                                                <asp:TextBox ID="txt_PCont" runat="server" Width="150px" Height="70px" Rows="5" TextMode="MultiLine" Enabled="False"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="學／術科">
                                            <ItemTemplate>
                                                <asp:DropDownList ID="list_Classification" runat="server">
                                                    <asp:ListItem Value="1">學科</asp:ListItem>
                                                    <asp:ListItem Value="2">術科</asp:ListItem>
                                                </asp:DropDownList>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="上課地點">
                                            <ItemTemplate>
                                                <asp:DropDownList ID="list_PTID" runat="server"></asp:DropDownList>
                                                <input id="hide_PTID" type="hidden" runat="server" />
                                                <input id="hide_ID4" type="hidden" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="遠距教學">
                                            <HeaderTemplate>遠距教學<input onclick="chackAll(7);" type="checkbox" name="Choose1" id="Choose1" /></HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:CheckBox ID="bx_FARLEARN" runat="server" />
                                                <input id="hide_FARLEARN" type="hidden" runat="server" />
                                                <input id="hide_ID8" type="hidden" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="任課教師">
                                            <ItemTemplate>
                                                <asp:DropDownList ID="list_TechID" runat="server"></asp:DropDownList>
                                                <input id="hide_TechID" type="hidden" runat="server" />
                                                <input id="hide_ID5" type="hidden" runat="server" />
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                &nbsp;&nbsp;
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="助教">
                                            <ItemTemplate>
                                                <asp:DropDownList ID="list_TechID2" runat="server"></asp:DropDownList>
                                                <input id="hide_TechID2" type="hidden" runat="server" />
                                                <input id="hide_ID6" type="hidden" runat="server" />
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                &nbsp;&nbsp;
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                            <td valign="top" class="whitecol">
                                <%--變更前--%>
                                <asp:DataGrid ID="DataGrid4" runat="server" Width="100%" AutoGenerateColumns="False" CellPadding="8">
                                    <ItemStyle></ItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn DataField="PTDID" HeaderText="PTDID" Visible="False"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="日期">
                                            <ItemTemplate>
                                                <asp:Label ID="OldSTrainDateLabel" runat="server" Width="100px"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="授課時段">
                                            <ItemStyle Wrap="false"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:CheckBox ID="OldTPERIOD28_1t" runat="server" Text="早上" ToolTip="7:00-13:00" Enabled="false" /><br />
                                                <asp:CheckBox ID="OldTPERIOD28_2t" runat="server" Text="下午" ToolTip="13:00-18:00" Enabled="false" /><br />
                                                <asp:CheckBox ID="OldTPERIOD28_3t" runat="server" Text="晚上" ToolTip="18:00-22:00" Enabled="false" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="授課時間">
                                            <ItemTemplate>
                                                <asp:Label ID="OldPNameLabel" runat="server" Width="120px"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="時數">
                                            <%--<HeaderStyle Wrap="False" Width="6%"></HeaderStyle>--%>
                                            <ItemStyle Width="6%" Wrap="false" HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:Label ID="OldPHourLabel" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="技檢訓練時數">
                                            <%--<HeaderStyle Wrap="False" Width="6%"></HeaderStyle>--%>
                                            <ItemStyle Width="6%" Wrap="false" HorizontalAlign="Center" />
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:Label ID="OldEHourLabel" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="課程進度／內容">
                                            <ItemTemplate>
                                                <asp:TextBox ID="OldPContText" runat="server" Width="150px" Rows="5" TextMode="MultiLine" Height="70px" Enabled="False"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="學／術科">
                                            <ItemTemplate>
                                                <asp:DropDownList ID="OlddrpClassification1" runat="server" Enabled="False">
                                                    <asp:ListItem Value="1">學科</asp:ListItem>
                                                    <asp:ListItem Value="2">術科</asp:ListItem>
                                                </asp:DropDownList>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="上課地點">
                                            <ItemTemplate>
                                                <asp:DropDownList ID="OlddrpPTID" runat="server" Enabled="False">
                                                </asp:DropDownList>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="遠距教學">
                                            <ItemTemplate>
                                                <asp:CheckBox ID="OldFARLEARN" runat="server" Enabled="false" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="任課教師">
                                            <ItemTemplate>
                                                <input id="OldTech1Value" type="hidden" name="OldTech1Value" runat="server" />
                                                <asp:TextBox ID="OldTech1Text" runat="server" Width="100px" Enabled="False"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="助教">
                                            <ItemTemplate>
                                                <input id="OldTech2Value" type="hidden" name="OldTech2Value" runat="server" />
                                                <asp:TextBox ID="OldTech2Text" runat="server" Width="100px" Enabled="False"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <!-- 20080627 andy add 課程表	end-->
        </table>
        <div>
            <input id="OLessonTeah1Value" type="hidden" name="OLessonTeah1Value" runat="server" />
            <input id="OLessonTeah2Value" type="hidden" name="OLessonTeah2Value" runat="server" />
            <input id="OLessonTeah3Value" type="hidden" name="OLessonTeah3Value" runat="server" />
            <input id="STDate" type="hidden" name="STDate" runat="server" />
            <input id="FDDate" type="hidden" name="FTDate" runat="server" />
            <input id="ROCID" type="hidden" name="ROCID" runat="server" />
            <input id="SEnterDate" type="hidden" name="SEnterDate" runat="server" />
            <input id="FEnterDate" type="hidden" name="FEnterDate" runat="server" />
            <input id="SignUpEDate" type="hidden" name="SignUpEDate" runat="server" />
            <input id="hid_chkmsg" type="hidden" name="hid_chkmsg" runat="server" />
            <input id="hid_NoTechID2" type="hidden" name="hid_NoTechID2" runat="server" />
            <input id="hid_NoTechID3" type="hidden" name="hid_NoTechID3" runat="server" />
            <input id="hid_TPlanID28AppPlan" type="hidden" name="hid_TPlanID28AppPlan" runat="server" />
            <input id="hid_MaxChgItem" type="hidden" name="hid_MaxChgItem" runat="server" />
            <input id="hid_TMID" type="hidden" name="hid_TMID" runat="server" />
            <input id="hidReqID" type="hidden" name="hidReqID" runat="server" />
            <input id="hidReqPlanID" type="hidden" name="hidReqPlanID" runat="server" />
            <input id="hidReqcid" type="hidden" name="hidReqcid" runat="server" />
            <input id="hidReqno" type="hidden" name="hidReqno" runat="server" />
            <input id="hidReqcheck" type="hidden" name="hidReqcheck" runat="server" />
            <asp:HiddenField ID="Hid_NowDate" runat="server" />
            <asp:HiddenField ID="Hid_rCDATE" runat="server" />
            <asp:HiddenField ID="Hid_ComIDNO" runat="server" />
            <asp:HiddenField ID="Hid_RID1" runat="server" />
            <asp:HiddenField ID="Hid_COSTITEM_GUID21" runat="server" />
            <asp:HiddenField ID="Hid_PlanKind" runat="server" />
            <asp:HiddenField ID="Hid_CostMode" runat="server" />
            <asp:HiddenField ID="Hid_AdmPercent" runat="server" />
            <asp:HiddenField ID="Hid_TaxPercent" runat="server" />
            <asp:HiddenField ID="Hid_PARTREDUC_Y_CanUpdate" runat="server" />
            <%--<asp:HiddenField ID="Hid_GUID22" runat="server" />--%>
        </div>
        <asp:Literal ID="JAVASCRIPT_LITERAL" runat="server"></asp:Literal>
    </form>
</body>
</html>
