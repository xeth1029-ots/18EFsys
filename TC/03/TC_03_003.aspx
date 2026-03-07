<%@ Page Language="vb" AutoEventWireup="false" MaintainScrollPositionOnPostback="true" CodeBehind="TC_03_003.aspx.vb" Inherits="WDAIIP.TC_03_003" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html>
<head>
    <title>班級申請作業</title>
    <meta name="vs_defaultClientScript" content="JavaScript" />
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5" />
    <link rel="stylesheet" type="text/css" href="../../css/style.css" />
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        //<!--
        function showHide(type) {
            if (type == 1 && document.all.nxlayer_01.style.visibility == 'hidden') {
                document.all.nxlayer_01.style.visibility = 'visible';
                document.all.SciPlaceID.style.visibility = 'hidden';
                document.all.TechPlaceID.style.visibility = 'hidden';
                document.all.SciPlaceID2.style.visibility = 'hidden';
                document.all.TechPlaceID2.style.visibility = 'hidden';
                document.all.Taddress2.style.visibility = 'hidden';
                document.all.Taddress3.style.visibility = 'hidden';
            } else {
                document.all.nxlayer_01.style.visibility = 'hidden';
                document.all.SciPlaceID.style.visibility = 'visible';
                document.all.TechPlaceID.style.visibility = 'visible';
                document.all.SciPlaceID2.style.visibility = 'visible';
                document.all.TechPlaceID2.style.visibility = 'visible';
                document.all.Taddress2.style.visibility = 'visible';
                document.all.Taddress3.style.visibility = 'visible';
            }
            //if (!_isIE) { parent.setMainFrameHeight(); } //重新調整頁面高度，by:20180815
            //if (parent && parent.setMainFrameHeight != undefined) { parent.setMainFrameHeight(); }
        }

        function GetPackageName() {
            var PackageType = document.getElementsByName('PackageType');
            var PackageName = document.form1.PackageName;
            var cst_pt1 = 0;
            var cst_pt2 = 1;
            var cst_pt3 = 2;
            if (PackageType.length > 3) {
                cst_pt1 = 1;
                cst_pt2 = 2;
                cst_pt3 = 3;
            }
            if (PackageType[cst_pt1].checked == true) {
                PackageName.value = ''; //(非包班)
            }
            if (PackageType[cst_pt2].checked == true) {
                PackageName.value = '(企業包班)';
            }
            if (PackageType[cst_pt3].checked == true) {
                PackageName.value = '(聯合企業包班)';
            }
        }

        function GetPackageName54() {
            var btnAddBusPackage = document.getElementById('btnAddBusPackage');
            var PackageType = document.getElementsByName('PackageType');
            var PackageName = document.form1.PackageName;
            var msg = '';
            var hTPlanID54 = document.getElementById('hTPlanID54'); //確認是否為 充電起飛計畫
            var cst_pt1 = 0;
            var cst_pt2 = 1;
            var cst_pt3 = 2;
            if (PackageType.length > 3) {
                cst_pt1 = 1; //cst_pt
                cst_pt2 = 2;
                cst_pt3 = 3;
            }
            //確認是否為 充電起飛計畫
            if (hTPlanID54 != null && hTPlanID54.value == '1') {
                if (PackageType[cst_pt1].checked == true) {
                    PackageType[cst_pt1].checked = false;
                    PackageType[cst_pt1].disabled = true;
                    msg += '充電起飛計畫不可選擇非包班!!\n'; //(非包班)
                }
                if (PackageType[cst_pt2].checked == true) {
                    PackageName.value = '(企業包班)';
                }
                if (PackageType[cst_pt3].checked == true) {
                    PackageName.value = '(聯合企業包班)';
                }
                //充電起飛計畫 且 聯合企業包班 才可顯示 計畫包班事業單位				
                if (btnAddBusPackage) { btnAddBusPackage.style.display = ''; }
                if (document.getElementById('Datagrid4headTable')) {
                    document.all.Datagrid4headTable.style.visibility = 'hidden';
                }
                if (document.getElementById('Datagrid4Table')) {
                    document.all.Datagrid4Table.style.visibility = 'hidden';
                }
                //充電起飛計畫 且 聯合企業包班 才可顯示 計畫包班事業單位
                if (PackageType[cst_pt3].checked == true) {
                    if (document.getElementById('Datagrid4headTable')) {
                        document.all.Datagrid4headTable.style.visibility = 'visible';
                    }
                    if (document.getElementById('Datagrid4Table')) {
                        document.all.Datagrid4Table.style.visibility = 'visible';
                    }
                }
                //充電起飛計畫 且 企業包班 只顯示 計畫包班事業單位 的抬頭
                if (PackageType[cst_pt2].checked == true) {
                    if (document.getElementById('Datagrid4headTable')) {
                        document.all.Datagrid4headTable.style.visibility = 'visible';
                    }
                    if (btnAddBusPackage) { btnAddBusPackage.style.display = 'none'; }
                }
                if (msg != '') {
                    alert(msg);
                    return false;
                }
            }
        }

        function GetPointName() {
            var cst_pt1 = 0;
            var cst_pt2 = 1;
            var cst_pt3 = 2;
            var PointName = document.form1.PointName;
            if (PointType.length > 3) {
                cst_pt1 = 1; //cst_pt
                cst_pt2 = 2;
                cst_pt3 = 3;
            }
            if (document.getElementsByName("PointType")[cst_pt1].checked == true) {
                PointName.value = '學士學分班';
            }
            if (document.getElementsByName("PointType")[cst_pt2].checked == true) {
                PointName.value = '碩士學分班';
            }
            if (document.getElementsByName("PointType")[cst_pt3].checked == true) {
                PointName.value = '博士學分班';
            }
        }

        var ap_name = navigator.appName;
        var ap_vinfo = navigator.appVersion;
        var ap_ver = parseFloat(ap_vinfo.substring(0, ap_vinfo.indexOf('(')));
        var time_start = new Date();
        var clock_start = time_start.getTime();
        var dl_ok = false;
        var exit_second = 7; //離開時間以秒計算

        function init() {
            if (ap_name == "Netscape" && ap_ver >= 3.0)
                dl_ok = true;
            return true;
        }

        //取得現在時間與起始時間的秒差
        function get_time_spent() {
            var time_now = new Date();
            if (((time_now.getTime() - clock_start) / 1000) > exit_second) {
                //若時間差機制大於5秒鐘，表示須要重新調整時間差
                time_start = new Date();
                clock_start = time_start.getTime();
            }
            return ((time_now.getTime() - clock_start) / 1000);
        }

        // show the time user spent on the side
        function show_secs() {
            var i_total_secs = exit_second - Math.round(get_time_spent()); //5秒鐘
            var i_secs_spent = i_total_secs % 60;
            var i_mins_spent = Math.round((i_total_secs - 30) / 60);
            var s_secs_spent = "" + ((i_secs_spent > 9) ? i_secs_spent : "0" + i_secs_spent);
            var s_mins_spent = "" + ((i_mins_spent > 9) ? i_mins_spent : "0" + i_mins_spent);
            if (document.getElementById('btnAdd') != null) {
                if (document.getElementById('btnAdd').style.display == "none") {
                    document.getElementById('Labsave').style.display = "";
                }
            }
            if (document.getElementById('Button8') != null) {
                if (document.getElementById('Button8').style.display == "none") {
                    document.getElementById('Labsave').style.display = "";
                }
            }
            document.form1.time_spent.value = s_mins_spent + ":" + s_secs_spent;
            if (document.form1.time_spent.value != '00:00') {
                window.setTimeout('show_secs()', 1000);
            }
            else {
                document.getElementById('Labsave').style.display = "none";
                if (document.getElementById('btnAdd') != null) {
                    if (document.getElementById('btnAdd').style.display == "none") {
                        document.getElementById('btnAdd').style.display = "";
                    }
                }
                if (document.getElementById('Button8') != null) {
                    if (document.getElementById('Button8').style.display == "none") {
                        document.getElementById('Button8').style.display = "";
                    }
                }
                if (document.getElementById('Button24') != null) {
                    document.getElementById('Button24').style.display = "";
                }
                return false;
            }
            //if (!_isIE) { parent.setMainFrameHeight(); } //重新調整頁面高度，by:20180815
            //if (parent && parent.setMainFrameHeight != undefined) { parent.setMainFrameHeight(); }
        }

        function notshow_button() {
            if (document.getElementById('btnAdd') != null) {
                document.getElementById('btnAdd').style.display = "none";
            }
            if (document.getElementById('Button8') != null) {
                document.getElementById('Button8').style.display = "none";
            }
            if (document.getElementById('Button24') != null) {
                document.getElementById('Button24').style.display = "none";
            }
            //if (!_isIE) { parent.setMainFrameHeight(); } //重新調整頁面高度，by:20180815
            //if (parent && parent.setMainFrameHeight != undefined) { parent.setMainFrameHeight(); }
        }
        // -->
    </script>
    <script type="text/javascript">
        //若是教材費則 計價單位欄位自動幫忙帶訓練人數
        function GetItemage() {
            if (document.form1.CostID2.value == '03,人') {
                if (document.getElementById('TNum').value != '') {
                    document.getElementById('Itemage').value = document.getElementById('TNum').value;
                }
            }
        }

        function ShowItemCostName(ObjName, ObjName2, ObjName3) {
            //設定字串取得 01,XXXX
            if (document.getElementById(ObjName).value != '') {
                //debugger;
                document.getElementById(ObjName2).innerText = document.getElementById(ObjName).value.substring(3);
                if (document.getElementById(ObjName).value.substring(3) == "班") {
                    document.getElementById(ObjName3).value = "1";
                    document.getElementById(ObjName3).readOnly = true;
                }
                else {
                    document.getElementById(ObjName3).readOnly = false;
                }
            }
            else {
                document.getElementById(ObjName2).innerText = '';
            }
        }

        function Get_GovClass(fieldname) {
            var PointYN = '';
            var cst_pt1 = 0;
            var cst_pt2 = 1;
            var RadioButtonList1 = document.getElementsByName('RadioButtonList1');
            var TB_career_id = document.getElementById("TB_career_id");
            var btu_sel = document.getElementById("btu_sel");

            if (RadioButtonList1.length > 2) {
                cst_pt1 = 1; //cst_pt
                cst_pt2 = 2;
            }
            if (RadioButtonList1[cst_pt1].checked == true) { PointYN = 'Y'; }
            if (RadioButtonList1[cst_pt2].checked == true) { PointYN = 'N'; }
            if (TB_career_id.value != '') {
                btu_sel.removeAttribute("title");
                btu_sel.disabled = false;
                //不開放選擇訓練職類
                btu_sel.title = "選擇經費分類代碼，不可再選訓練職類";
                if (document.getElementById("jobValue").value != '' && PointYN == 'N') {
                    btu_sel.disabled = true;
                }
                wopen('../../common/GovClass.aspx?fieldname=' + fieldname + "&jobValue=" + document.getElementById("jobValue").value + "&PointYN=" + PointYN, 'GovClass', 930, 350, 1);
            }
            else {
                alert('請先輸入訓練業別');
                return false;
            }
        }

        function LessonTeah1(opentype, fieldname, hiddenname) {
            var RIDValue = document.getElementById("RIDValue");
            var msg = '';
            if (RIDValue.value != null && RIDValue.value != "") {
                wopen('../../SD/04/LessonTeah1.aspx?RID=' + RIDValue.value + '&type=' + opentype + '&fieldname=' + fieldname + '&hiddenname=' + hiddenname, 'LessonTeah1', 1025, 675, 1);
            }
            else {
                msg += '請先選擇訓練機構\n';
                if (msg != '') {
                    alert(msg);
                    return false;
                }
            }
        }

        function onchg_total3() {
            var msg = '';
            var Total = document.getElementById('TotalCost3').value;
            var DefGovCost = document.getElementById('DefGovCost');
            var DefStdCost = document.getElementById('DefStdCost');
            var hTPlanID54 = document.getElementById('hTPlanID54');
            if (document.form1.TotalCost3.value != '' && !isUnsignedInt(document.form1.TotalCost3.value)) msg += '總計必須為數字格式\n'
            if (msg != '') {
                alert(msg);
            }
            else {
                document.getElementById('TotalCost3').value = Total;
                DefGovCost.value = Total * 0.8;
                DefStdCost.value = Total * 0.2;
                if (hTPlanID54 != null && hTPlanID54.value == '1') {
                    DefGovCost.value = Total;
                    DefStdCost.value = 0;
                }
            }
        }

        function onchg_total2() {
            var msg = '';
            var Total = document.getElementById('TotalCost2').value;
            var DefGovCost = document.getElementById('DefGovCost');
            var DefStdCost = document.getElementById('DefStdCost');
            var hTPlanID54 = document.getElementById('hTPlanID54');
            document.getElementById('TotalCost2').value = Total;
            DefGovCost.value = Total * 0.8;
            DefStdCost.value = Total * 0.2;
            if (hTPlanID54 != null && hTPlanID54.value == '1') {
                DefGovCost.value = Total;
                DefStdCost.value = 0;
            }
        }

        function showPTID(obj, ptid1, ptid2) {
            var objvalue = document.getElementById(obj).value;
            if (objvalue == 1) {
                document.getElementById(ptid1).style.display = "";
                document.getElementById(ptid2).style.display = "none";
            }
            else {
                document.getElementById(ptid1).style.display = "none";
                document.getElementById(ptid2).style.display = "";
            }
            //if (!_isIE) { parent.setMainFrameHeight(); } //重新調整頁面高度，by:20180815
            //if (parent && parent.setMainFrameHeight != undefined) { parent.setMainFrameHeight(); }
        }

        function showCostType(obj1, table1, table2) {
            var cst_pt1 = 0;
            var cst_pt2 = 1;
            var cst_pt3 = 2;
            var PointType = document.getElementsByName('PointType');
            if (PointType.length > 3) {
                cst_pt1 = 1; //cst_pt
                cst_pt2 = 2;
                cst_pt3 = 3;
            }

            //JS DEBUG 依2/8 EMAIL問題： 產業人才投資方案經費欄位資訊錯誤，為避免顯示外網資訊錯誤，惠請協助查明並修復
            var v_obj1 = getRadioValue(document.getElementsByName(obj1));
            if (v_obj1 == 'Y') {
                //學分班
                document.getElementById('Table4').style.display = "none";
                document.getElementById(table1).style.display = "none";
                document.getElementById('Button21b').style.display = "none"; //帶入訓練費用
                document.getElementById(table2).style.display = "";
                document.getElementById('tdCredPoint').className = "bluecol_need";
                document.getElementById('PointType_TR').style.display = "";
                document.getElementById('PointName').style.display = "";
                if (document.getElementById('TotalCost3').value != '') {
                    onchg_total3();
                }
            }
            else {
                //非學分班
                document.getElementById('Table4').style.display = "";
                document.getElementById(table1).style.display = "";
                document.getElementById('Button21b').style.display = ""; //帶入訓練費用
                document.getElementById(table2).style.display = "none";
                document.getElementById('tdCredPoint').className = "bluecol";
                document.getElementById('PointType_TR').style.display = "none";
                document.getElementById('PointName').style.display = "none";
                document.getElementsByName("PointType")[cst_pt1].checked = false;
                document.getElementsByName("PointType")[cst_pt2].checked = false;
                document.getElementsByName("PointType")[cst_pt3].checked = false;
                document.form1.PointName.value = '';
                onchg_total2();
            }
            //if (!_isIE) { parent.setMainFrameHeight(); } //重新調整頁面高度，by:20180815
            //if (parent && parent.setMainFrameHeight != undefined) { parent.setMainFrameHeight(); }
        }

        function Layer_change(index) {
            //先關閉所有選項
            showHide(0);
            GetPackageName54();
            var xLayerState = document.getElementById('LayerState');
            for (i = 1; i <= 8; i++) {
                document.getElementById('TableLay' + i).style.display = "none";
            }
            for (i = 1; i <= 8; i++) {
                var mybox = document.getElementById('box' + i);
                mybox.className = "";
            }
            if (index == '') {
                if (xLayerState) { index = xLayerState.value; }
            }
            document.getElementById('TableLay' + index).style.display = "";
            var mybox2 = document.getElementById('box' + index);
            if (mybox2) { mybox2.className = "active"; }
            if (xLayerState) { xLayerState.value = index; }
            if (document.body) { window.scroll(0, document.body.scrollHeight); }
            if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
            //if (!_isIE) { parent.setMainFrameHeight(); } //重新調整頁面高度，by:20180815
            //if (parent && parent.setMainFrameHeight != undefined) { parent.setMainFrameHeight(); }
        }

        //檢查 課程大綱 (TrainDescTable)
        function CheckTrainDescTable() {
            var msg = '';
            var TPERIOD28_1 = document.getElementById("TPERIOD28_1");
            var TPERIOD28_2 = document.getElementById("TPERIOD28_2");
            var TPERIOD28_3 = document.getElementById("TPERIOD28_3");
            var objH1 = document.getElementById("ddlpnH1");
            var objH2 = document.getElementById("ddlpnH2");
            var objM1 = document.getElementById("ddlpnM1");
            var objM2 = document.getElementById("ddlpnM2");
            var STrainDate = document.getElementById('STrainDate');
            var PHour = document.getElementById('PHour');
            var PCont = document.getElementById('PCont');
            var Classification1 = document.getElementById('Classification1');
            var PTID1 = document.getElementById('PTID1');
            var PTID2 = document.getElementById('PTID2');
            var OLessonTeah1Value = document.getElementById('OLessonTeah1Value');
            var iTPERIOD28 = 0;
            if (TPERIOD28_1.checked) { iTPERIOD28 += 1 };
            if (TPERIOD28_2.checked) { iTPERIOD28 += 1 };
            if (TPERIOD28_3.checked) { iTPERIOD28 += 1 };
            if (iTPERIOD28 == 0) { msg += '授課時段:早上、下午、晚上 至少要設定其中一項\n'; }
            if (STrainDate.value == '') msg += '【日期】不可為空，請選擇\n';
            else {
                if (!checkDate(STrainDate.value)) {
                    msg += '【日期】格式有誤應為日期格式\n';
                }
            }
            var H1val = parseInt(getValue(objH1), 10);
            var H2val = parseInt(getValue(objH2), 10);
            var M1val = parseInt(getValue(objM1), 10);
            var M2val = parseInt(getValue(objM2), 10);
            if (H1val > H2val) {
                msg += '【授課時間】起始時間不得大於結束時間\n';
            }
            if (H1val == H2val) {
                if (M1val >= M2val) { msg += '【授課時間】起始時間不得大於等於結束時間\n'; }
            }
            if (PHour.value == '') msg += '【時數】不可為空，請輸入\n';
            else if (!isUnsignedInt(PHour.value)) msg += '【時數】必須為整數數字\n';
            else if (!(parseInt(PHour.value, 10) <= 4 && parseInt(PHour.value, 10) >= 1)) msg += '【時數】必須為小於4，大於0\n';
            if (PCont.value == '') msg += '【課程進度／內容】不可為空，請輸入\n';
            if (Classification1.selectedIndex == 0) msg += '【學／術科】不可為空，請選擇\n';
            if (msg == '' && Classification1.value == '1') {
                if (PTID1.selectedIndex == 0 && PTID1.value == "") msg += '【學科:上課地點】不可為空，請選擇\n';
            }
            if (msg == '' && Classification1.value == '2') {
                if (PTID2.selectedIndex == 0 && PTID2.value == "") msg += '【術科:上課地點】不可為空，請選擇\n';
            }
            if (OLessonTeah1Value.value == '') msg += '【任課教師】不可為空，請選擇\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        //檢查上課時間
        function CheckAddTime() {
            var msg = '';
            var Weeks = document.getElementById('Weeks');
            var Times = document.getElementById('Times');
            if (Weeks.selectedIndex == 0) msg += '請選擇星期\n';
            if (Times.value == '') msg += '請輸入上課時間\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function CheckAddBusPackage() {
            var msg = '';
            var txtUname = document.getElementById('txtUname');
            var txtIntaxno = document.getElementById('txtIntaxno');
            var txtUbno = document.getElementById('txtUbno');
            if (txtUname.value == '') msg += '請輸入企業名稱\n';
            //請輸入正確的統一編號
            //欄位有效時再行驗證,無效時也要進行驗證
            if (txtIntaxno.value != '') {
                if (txtIntaxno.value == "00000000") {
                    msg += '請輸入有效的服務單位統一編號，不可為00000000\n';
                }
                if ((txtIntaxno.value.length != 8) || (!isValidTWBID(txtIntaxno.value))) {
                    msg += '請輸入有效的服務單位統一編號，應為8碼\n';
                }
            }
            if (txtUbno.value != '') {
                if (txtUbno.value.length != 9) {
                    msg += '請輸入有效的保險證號，應為9碼\n';
                }
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        //經費輸入判斷
        function check_Cost2() {
            var msg = '';
            var CostID2 = document.form1.CostID2;
            var OPrice2 = document.form1.OPrice2;
            var Itemage = document.form1.Itemage;
            //97修改中
            msg += check_Cost_Detail(CostID2, OPrice2, Itemage);
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function check_Cost_Detail(CostID, OPrice, Itemage) {
            var msg = '';
            if (CostID != null) {
                if (CostID.value == '') msg += '請輸入項目\n';
            }
            if (OPrice.value == '') msg += '請輸入單價\n';
            else {
                if (!isUnsignedInt(OPrice.value)) {
                    if (!isPositiveFloat(OPrice.value)) {
                        msg += '單價必須為數字\n';
                    }
                    else {
                        if (OPrice.value.indexOf('.') < OPrice.value.length - 3) {
                            msg += '單價只能輸入到小數點第二位\n';
                        }
                    }
                }
            }
            if (Itemage.value == '') msg += '請輸入計價數量\n';
            else if (!isUnsignedInt(Itemage.value)) msg += '計價數量必須為數字\n';
            return msg;
        }

        //如果選擇其他費用，則顯示TextBox讓使用者輸入
        function ShowOther(sObjName, sObjName2) {
            var ObjName = document.getElementById(sObjName);
            var ObjName2 = document.getElementById(sObjName2);
            if (dObjName.value == '99')
                ObjName2.style.display = '';
            else
                ObjName2.style.display = 'none';
        }

        function check_PlaceID(source, args) {
            var flag = false;
            var SciPlaceID = document.form1.SciPlaceID;
            var TechPlaceID = document.form1.TechPlaceID;
            if (SciPlaceID.selectedIndex != 0 || TechPlaceID.selectedIndex != 0) {
                flag = true;
            }
            args.IsValid = flag;
        }

        function check_IsBusiness(source, args) {
            flag = true;
            args.IsValid = flag;
        }

        function check_CyclType(source, args) {
            if (!isUnsignedInt(args.Value)) args.IsValid = false;
            if (args.Value.length != 2) args.IsValid = false;
            if (parseInt(args.Value, 10) <= 0) args.IsValid = false;
        }

        //'檢查日期格式-Melody(2005/3/18)
        function check_date(source, args) {
            if (!checkDate(args.Value)) {
                args.IsValid = false;
            }
        }

        function open_hours() {
            window.open('TC_03_oper.aspx', '', 'width=1200,height=660,location=0,status=0,menubar=0,scrollbars=0,resizable=0');
        }

        function CheckDef(source, args) {
            var flag = false;
            if (document.form1.DefGovCost.value != '')
                flag = true;
            if (document.form1.DefUnitCost.value != '')
                flag = true;
            if (document.form1.DefStdCost.value != '')
                flag = true;
            args.IsValid = flag;
        }

        function CheckDef2(source, args) {
            var flag = true;
            if (!isUnsignedInt(document.getElementById('Total1').innerHTML))
                flag = false;
            if (!isUnsignedInt(document.getElementById('Total2').innerHTML))
                flag = false;
            if (!isUnsignedInt(document.getElementById('Total3').innerHTML))
                flag = false;
            args.IsValid = flag;
        }

        //判斷經費內容是否有輸入
        function CheckCost(source, args) {
            //debugger;
            var TotalCost2 = document.getElementById('TotalCost2');
            var TotalCost3 = document.getElementById('TotalCost3');
            if (parseInt(TotalCost3.value, 10) <= 0) {
                if (document.form1.GCIDValue.value == 157) {
                    args.IsValid = false;
                }
                else {
                    if (parseInt(TotalCost2.value, 10) > 0) {
                        args.IsValid = true;
                    }
                    else {
                        args.IsValid = false;
                    }
                }
            }
            else
                args.IsValid = true;
        }

        function CheckTDate(sor, args) {
            if (checkDate(document.form1.STDate.value) && checkDate(document.form1.FDDate.value)) {
                var flag = compareDate(document.form1.STDate.value, document.form1.FDDate.value);
                if (flag == 0) args.IsValid = false;
                if (flag == 1) args.IsValid = false;
            }
        }

        function CheckFactMode(sor, args) {
            if (!isChecked(document.form1.FactMode))
                args.IsValid = false;
        }

        function CheckFactModeOther(sor, args) {
            if (getRadioValue(document.form1.FactMode) == '99'
				&& document.getElementById('FactModeOther').value == '')
                args.IsValid = false;
        }

        //草稿儲存檢查
        function Check_Temp() {
            var msg = '';
            if (document.form1.center.value == '') msg += '請選擇訓練機構\n'
            if (document.form1.GenSciHours.value != '' && !isUnsignedInt(document.form1.GenSciHours.value)) msg += '一般學科必須為數字\n'
            if (document.form1.ProSciHours.value != '' && !isUnsignedInt(document.form1.ProSciHours.value)) msg += '專業術科必須為數字\n'
            if (document.form1.ProTechHours.value != '' && !isUnsignedInt(document.form1.ProTechHours.value)) msg += '必須為數字\n'
            if (document.form1.OtherHours.value != '' && !isUnsignedInt(document.form1.OtherHours.value)) msg += '其他時數必須為數字\n'
            if (document.form1.FirstSort.value != '' && !isUnsignedInt(document.form1.FirstSort.value)) msg += '優先排序必須為數字\n'
            if (document.form1.TNum.value != '' && !isUnsignedInt(document.form1.TNum.value)) msg += '訓練人數必須為數字\n'
            if (document.form1.THours.value != '' && !isUnsignedInt(document.form1.THours.value)) msg += '訓練時數必須為數字\n'
            if (document.form1.STDate.value != '' && !checkDate(document.form1.STDate.value)) msg += '訓練起日不是正確的日期格式\n'
            if (document.form1.FDDate.value != '' && !checkDate(document.form1.FDDate.value)) msg += '訓練迄日不是正確的日期格式\n'
            if (checkDate(document.form1.STDate.value) && checkDate(document.form1.FDDate.value)) {
                var flag = compareDate(document.form1.STDate.value, document.form1.FDDate.value);
                if (flag == 0) msg += '訓練起日不能和訓練迄日同一天\n';
                if (flag == 1) msg += '訓練起日不能超過訓練迄日\n';
            }
            if (document.form1.CyclType.value != '' && !isUnsignedInt(document.form1.CyclType.value)) msg += '期別必須為數字\n'
            if (document.form1.ClassCount.value != '' && !isUnsignedInt(document.form1.ClassCount.value)) msg += '班數必須為數字\n'
            if (document.form1.CredPoint.value != '' && !isUnsignedInt(document.form1.CredPoint.value)) msg += '學分數必須為數字\n'
            if (document.form1.ConNum.value != '' && !isUnsignedInt(document.form1.ConNum.value)) msg += '容納人數必須為數字\n'
            if (document.form1.DefGovCost.value != '' && !isUnsignedInt(document.form1.DefGovCost.value)) msg += '政府負擔費用必須為數字\n'
            if (document.form1.DefUnitCost.value != '' && !isUnsignedInt(document.form1.DefUnitCost.value)) msg += '企業負擔費用必須為數字\n'
            if (document.form1.DefStdCost.value != '' && !isUnsignedInt(document.form1.DefStdCost.value)) msg += '學員負擔費用必須為數字\n'
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        //計算經費來源
        function CountCostSource() {
            if (document.form1.TNum.value != '' && isUnsignedInt(document.form1.TNum.value)) {
                document.getElementById('TNum1').innerHTML = document.form1.TNum.value;
                document.getElementById('TNum2').innerHTML = document.form1.TNum.value;
                document.getElementById('TNum3').innerHTML = document.form1.TNum.value;
                if (isPositiveFloat(document.form1.DefGovCost.value) || isUnsignedInt(document.form1.DefGovCost.value)) {
                    document.getElementById('Total1').innerHTML = parseInt(document.form1.DefGovCost.value, 10) / parseInt(document.form1.TNum.value, 10);
                }
                else
                    document.getElementById('Total1').innerHTML = '0';
                if (isPositiveFloat(document.form1.DefUnitCost.value) || isUnsignedInt(document.form1.DefUnitCost.value)) {
                    document.getElementById('Total2').innerHTML = parseInt(document.form1.DefUnitCost.value, 10) / parseInt(document.form1.TNum.value, 10);
                }
                else
                    document.getElementById('Total2').innerHTML = '0';
                if (isPositiveFloat(document.form1.DefStdCost.value) || isUnsignedInt(document.form1.DefStdCost.value)) {
                    document.getElementById('Total3').innerHTML = parseInt(document.form1.DefStdCost.value, 10) / parseInt(document.form1.TNum.value, 10);
                }
                else
                    document.getElementById('Total3').innerHTML = '0';
            }
            else {
                document.getElementById('TNum1').innerHTML = '(尚未設定人數)';
                document.getElementById('TNum2').innerHTML = '(尚未設定人數)';
                document.getElementById('TNum3').innerHTML = '(尚未設定人數)';
                document.getElementById('Total1').innerHTML = '0';
                document.getElementById('Total2').innerHTML = '0';
                document.getElementById('Total3').innerHTML = '0';
            }
        }

        function check_style() {
            if (form1.STDate.disabled) {
                if (form1.date1) {
                    form1.date1.style.cursor = "";
                    form1.date1.onclick = null;
                }
            }
            if (form1.FDDate.disabled) {
                if (form1.date2) {
                    form1.date2.style.cursor = "";
                    form1.date2.onclick = null;
                }
            }
        }

        function MM_OpenWindowCenter(owcURL, owcWinName, owcWinWidth, owcWinHeight, owcFeatures) {
            var x = (screen.width - owcWinWidth) / 4;
            var y = (screen.height - owcWinHeight) / 4;
            window.open(owcURL, owcWinName, "left=" + x + ", top=" + y + ", width=" + owcWinWidth + ", height=" + owcWinHeight + owcFeatures);
        }

        //限制TextBox在MultiLine時的字數
        function checkTextLength(obj, Mlong) {
            var maxlength = new Number(Mlong); // Change number to your max length.
            if (obj.value.length > maxlength) {
                obj.value = obj.value.substring(0, maxlength);
                alert("限欄位長度不能大於" + maxlength + "個字元(含空白字元)，超出字元將自動截斷");
            }
        }

        //20090520 add 依需求只允許輸入整數(排除 00)
        function CheckZIPB3_1(source, args) {
            $('#TaddressZIPB3').val($.trim($('#TaddressZIPB3').val()));
            args.IsValid = true;
            if ($('#TaddressZIPB3').val() == "") { return; }
            if (isNaN(parseInt($('#TaddressZIPB3').val(), 10))) { args.IsValid = false; return; }
            if (!isUnsignedInt($('#TaddressZIPB3').val())) { args.IsValid = false; return; }
            if (parseInt($('#TaddressZIPB3').val(), 10) < 1) { args.IsValid = false; return; }
        }

        //20090520 add 依需求只允許輸入兩碼
        function CheckZIPB3_2(source, args) {
            $('#TaddressZIPB3').val($.trim($('#TaddressZIPB3').val()));
            args.IsValid = true;
            if ($('#TaddressZIPB3').val() == "") { return; }
            if ($('#TaddressZIPB3').val().length == 2) { return; }
            if ($('#TaddressZIPB3').val().length == 3) { return; }
            args.IsValid = false; return;
        }

        //20090521 add 依需求於郵遞區號後 2 碼進行 onchange時做格式驗證 
        function CheckZIPB3_Event() {
            var msg = '';
            $('#TaddressZIPB3').val($.trim($('#TaddressZIPB3').val()));
            if ($('#TaddressZIPB3').val() == "") { return true; }
            if (msg = '' && !isUnsignedInt($('#TaddressZIPB3').val())) { msg += '班別資料「郵遞區號後2碼」必須為數字，且不得輸入 0\n'; }
            if (msg = '' && parseInt($('#TaddressZIPB3').val(), 10) < 1) { msg += '班別資料「郵遞區號後2碼」必須為數字，且不得輸入 0\n'; }
            if (msg != '') { alert(msg); return false; }
            if ($('#TaddressZIPB3').val().length != 2 && $('#TaddressZIPB3').val().length != 3) { msg += '班別資料「郵遞區號後2碼或後3碼」長度必須為 2碼或3碼(例 01 或 001)\n'; }
            if (msg != '') { alert(msg); return false; }
            return true;
        }

        //判斷課程大剛-課程日期
        function chkTrainDate(objID) {
            var msg = '';
            var obj = document.getElementById(objID);
            var STDate = document.getElementById('STDate');
            var FDDate = document.getElementById('FDDate');
            if (isBlank(STDate)) msg += '請輸入訓練起日\n';
            else if (!checkDate(STDate.value)) msg += '訓練起日格式有誤應為日期格式\n';
            if (isBlank(FDDate)) msg += '請輸入訓練迄日\n';
            else if (!checkDate(FDDate.value)) msg += '訓練迄日格式有誤應為日期格式\n';
            if (checkDate(STDate.value) && checkDate(FDDate.value)) {
                var flag = compareDate(STDate.value, FDDate.value);
                if (flag == 0) msg += '訓練起日不能和訓練迄日同一天\n';
                else if (flag == 1) msg += '訓練起日不能超過訓練迄日\n';
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
            else {
                if (checkDate(obj.value)) openCalendar(obj.id, STDate.value, FDDate.value, obj.value);
                else openCalendar(obj.id, STDate.value, FDDate.value, '');
            }
        }
    </script>
    <style type="text/css">
        .auto-style3 { color: Black; text-align: right; padding: 4px 6px; background-color: #f1f9fc; border-right: 3px solid #49cbef; height: 49px; }
    </style>
</head>
<body onload="check_style();">
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;班級申請</asp:Label>
                </td>
            </tr>
        </table>
        <table border="0" cellspacing="0" cellpadding="0" width="100%">
            <tr>
                <td>
                    <table border="0" cellspacing="1" cellpadding="0" width="100%">
                        <tr>
                            <td class="bluecol_need" width="20%">訓練機構 </td>
                            <td class="whitecol" colspan="3" width="30%">
                                <asp:TextBox ID="center" runat="server" Width="70%" onfocus="this.blur()"></asp:TextBox>
                                <input id="Org" value="..." type="button" name="Org" runat="server" class="button_b_Mini" />
                                <asp:Button Style="display: none" ID="Button28" runat="server" Text="機構資訊(隱藏)" CausesValidation="False"></asp:Button>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="ComidValue" type="hidden" name="ComidValue" runat="server" /><br />
                                <span style="position: absolute; display: none" id="HistoryList2">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                                <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="center" Display="None" ErrorMessage="請選擇訓練機構"></asp:RequiredFieldValidator>
                            </td>
                            <td id="Table1_Email" runat="server" colspan="4" width="50%" class="whitecol">
                                <font class="font" size="2">是否要Email線上報名資料，EMail</font>
                                <asp:TextBox ID="EMail" runat="server" Columns="30" Width="70%"></asp:TextBox>
                                <asp:RegularExpressionValidator ID="check1" runat="server" ControlToValidate="EMail" Display="None" ErrorMessage="E_Mail輸入錯誤" ValidationExpression="\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"></asp:RegularExpressionValidator>
                            </td>
                        </tr>
                        <tr id="trainTR" runat="server">
                            <td class="bluecol_need">
                                <asp:Label ID="LabTMID" runat="server">訓練職類</asp:Label></td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Columns="30" Width="70%"></asp:TextBox>
                                <input id="btu_sel" onclick="openTrain(document.getElementById('trainValue').value);" value="..." type="button" runat="server" class="button_b_Mini" />
                                <asp:RequiredFieldValidator ID="fill1" runat="server" ControlToValidate="TB_career_id" Display="None" ErrorMessage="請選擇訓練業別／職類"></asp:RequiredFieldValidator>
                                <input id="trainValue" type="hidden" name="trainValue" runat="server" />
                                <input id="jobValue" type="hidden" name="jobValue" runat="server" />
                            </td>
                            <td class="bluecol_need">
                                <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label>
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="70%"></asp:TextBox>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" value="..." type="button" name="btu_sel2" runat="server" class="button_b_Mini" />
                                <asp:RequiredFieldValidator ID="fill1b" runat="server" ControlToValidate="txtCJOB_NAME" Display="None" ErrorMessage="請選擇通俗職類"></asp:RequiredFieldValidator>
                                <input id="cjobValue" type="hidden" name="cjobValue" runat="server" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <br />
                    <asp:Label ID="Label2" runat="server" CssClass="font">訓練年度:</asp:Label>
                    <asp:Label ID="Label3" runat="server" CssClass="font"></asp:Label><br />
                    <table class="font" cellspacing="0" cellpadding="0" width="90%">
                        <tr class="newlink">
                            <%--<td style="cursor: pointer" id="box1" onclick="Layer_change(1);" align="center">目標</td>
                            <td style="cursor: pointer" id="box2" onclick="Layer_change(2);" align="center">受訓資格</td>
                            <td style="cursor: pointer" id="box3" onclick="Layer_change(3);" align="center">訓練方式</td>
                            <td style="display: none; cursor: pointer" id="box4" onclick="Layer_change(4);" align="center">課程編配</td>
                            <td style="cursor: pointer" id="box5" onclick="Layer_change(5);" align="center">班別資料</td>
                            <td style="cursor: pointer" id="box6" onclick="Layer_change(6);" align="center" runat="server">訓練費用</td>
                            <td style="cursor: pointer" id="box7" onclick="Layer_change(7);" align="center" runat="server">經費來源</td>
                            <td style="cursor: pointer" id="box8" onclick="Layer_change(8);" align="center">訓練費用</td>--%>
                            <td id="box1" class="active" onclick="Layer_change(1);" runat="server">目標</td>
                            <td id="box2" onclick="Layer_change(2);" runat="server">受訓資格</td>
                            <td id="box3" onclick="Layer_change(3);" runat="server">訓練方式</td>
                            <td id="box4" onclick="Layer_change(4);" runat="server">課程編配</td>
                            <td id="box5" onclick="Layer_change(5);" runat="server">班別資料</td>
                            <td id="box6" onclick="Layer_change(6);" runat="server">訓練費用</td>
                            <td id="box7" onclick="Layer_change(7);" runat="server">經費來源</td>
                            <td id="box8" onclick="Layer_change(8);" runat="server">訓練費用</td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <div style="overflow-y: auto;">
                        <table id="TableLay1" border="0" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol_need" width="20%">緣由 </td>
                                <td class="whitecol" width="80%">
                                    <asp:TextBox ID="PlanCause" runat="server" Width="80%" Rows="4" TextMode="MultiLine" MaxLength="100"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="fill30" runat="server" ControlToValidate="PlanCause" Display="None" ErrorMessage="目標『緣由』為必填欄位"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">學科 </td>
                                <td class="whitecol" width="80%">
                                    <asp:TextBox ID="PurScience" runat="server" Width="80%" Rows="4" TextMode="MultiLine" MaxLength="100"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="fill2" runat="server" ControlToValidate="PurScience" Display="None" ErrorMessage="目標「學科」為必填欄位"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">技能 </td>
                                <td class="whitecol" width="80%">
                                    <asp:TextBox ID="PurTech" runat="server" Width="80%" Rows="4" TextMode="MultiLine" MaxLength="100"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="fill3" runat="server" ControlToValidate="PurTech" Display="None" ErrorMessage="目標「技能」為必填欄位"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">品德 </td>
                                <td class="whitecol" width="80%">
                                    <asp:TextBox ID="PurMoral" runat="server" Width="80%" Rows="4" TextMode="MultiLine" MaxLength="100"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="fill4" runat="server" ControlToValidate="PurMoral" Display="None" ErrorMessage="目標「品德」為必填欄位"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                        </table>
                        <table id="TableLay2" border="0" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol_need" width="20%">學歷 </td>
                                <td class="whitecol" width="80%">
                                    <asp:DropDownList ID="Degree" runat="server"></asp:DropDownList>(含以上)
                                    <asp:RequiredFieldValidator ID="fill5" runat="server" ControlToValidate="Degree" Display="None" ErrorMessage="請選擇受訓資格「學歷」"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">年齡 </td>
                                <td class="whitecol" width="80%">
                                    <asp:RadioButton ID="rdoAge1" runat="server" Checked="True" GroupName="GroupAge" />
                                    <asp:Label ID="l_Age" runat="server">年滿15歲以上</asp:Label>
                                    <asp:RadioButton ID="rdoAge2" runat="server" GroupName="GroupAge" />
                                    <asp:Label ID="l_Age2a" runat="server">應符合相關法規須年滿</asp:Label>
                                    <asp:TextBox ID="txtAge1" runat="server" Width="15%" MaxLength="2"></asp:TextBox>
                                    <asp:Label ID="l_Age2b" runat="server">以上</asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">其他一 </td>
                                <td class="whitecol" width="80%">
                                    <asp:TextBox onblur="checkTextLength(this,200)" ID="Other1" onkeyup="checkTextLength(this,200)" runat="server" Width="80%" Rows="2" TextMode="MultiLine" MaxLength="100" ForeColor="Black" BackColor="Silver" onChange="checkTextLength(this,200)"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">其他二 </td>
                                <td class="whitecol" width="80%">
                                    <asp:TextBox onblur="checkTextLength(this,200)" ID="Other2" onkeyup="checkTextLength(this,200)" runat="server" Width="80%" Rows="2" TextMode="MultiLine" MaxLength="100" ForeColor="Black" BackColor="Silver" onChange="checkTextLength(this,200)"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">其他三 </td>
                                <td class="whitecol" width="80%">
                                    <asp:TextBox onblur="checkTextLength(this,200)" ID="Other3" onkeyup="checkTextLength(this,200)" runat="server" Width="80%" Rows="2" TextMode="MultiLine" MaxLength="100" ForeColor="Black" BackColor="Silver" onChange="checkTextLength(this,200)"></asp:TextBox></td>
                            </tr>
                        </table>
                        <table id="TableLay3" border="0" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol_need" width="20%">訓練方式 </td>
                                <td class="whitecol" width="80%">
                                    <asp:TextBox onblur="checkTextLength(this,200)" ID="TMScience" onkeyup="checkTextLength(this,200)" runat="server" Rows="5" TextMode="MultiLine" MaxLength="200" onChange="checkTextLength(this,200)" Columns="70" Width="50%"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="fill13" runat="server" ControlToValidate="TMScience" Display="None" ErrorMessage="訓練方式為必填欄位"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                        </table>
                        <table id="TableLay4" border="0" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td rowspan="2" class="bluecol_need" width="20%">學科 </td>
                                <td rowspan="2" class="whitecol" width="30%">
                                    <asp:TextBox ID="SciHours" runat="server" Width="30%" onfocus="this.blur()"></asp:TextBox>小時 </td>
                                <td class="bluecol" width="20%">1. 一般學科 </td>
                                <td class="whitecol" width="30%">
                                    <asp:TextBox ID="GenSciHours" runat="server" Width="30%"></asp:TextBox>小時
                                    <asp:RequiredFieldValidator ID="fill15" runat="server" ControlToValidate="GenSciHours" Display="None" ErrorMessage="課程編配「一般學科」為必填欄位" Enabled="False"></asp:RequiredFieldValidator>
                                    <asp:RegularExpressionValidator ID="check4" runat="server" ControlToValidate="GenSciHours" Display="None" ErrorMessage="課程編配「一般學科」請輸入數字" ValidationExpression="[0-9]{1,4}" Enabled="False"></asp:RegularExpressionValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">2. 專業學科 </td>
                                <td class="whitecol" width="30%">
                                    <asp:TextBox ID="ProSciHours" runat="server" Width="30%"></asp:TextBox>小時
                                    <asp:RequiredFieldValidator ID="fill16" runat="server" ControlToValidate="ProSciHours" Display="None" ErrorMessage="課程編配「專業學科」為必填欄位" Enabled="False"></asp:RequiredFieldValidator>
                                    <asp:RegularExpressionValidator ID="check5" runat="server" ControlToValidate="ProSciHours" Display="None" ErrorMessage="課程編配「專業學科」請輸入數字" ValidationExpression="[0-9]{1,4}" Enabled="False"></asp:RegularExpressionValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">術科<font color="red">*</font> </td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="ProTechHours" runat="server" Width="11%"></asp:TextBox>小時
                                    <asp:RequiredFieldValidator ID="fill17" runat="server" ControlToValidate="ProTechHours" Display="None" ErrorMessage="課程編配「術科」為必填欄位" Enabled="False"></asp:RequiredFieldValidator>
                                    <asp:RegularExpressionValidator ID="check6" runat="server" ControlToValidate="ProTechHours" Display="None" ErrorMessage="課程編配「術科」請輸入數字" ValidationExpression="[0-9]{1,4}" Enabled="False"></asp:RegularExpressionValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">其他時數 </td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="OtherHours" runat="server" Width="11%"></asp:TextBox>小時
                                    <asp:RegularExpressionValidator ID="check7" runat="server" ControlToValidate="OtherHours" Display="None" ErrorMessage="課程編配「其他時數」請輸入數字" ValidationExpression="[0-9]{1,4}" Enabled="False"></asp:RegularExpressionValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">總計 </td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="TotalHours" runat="server" Width="11%" onfocus="this.blur()" MaxLength="3"></asp:TextBox>小時 </td>
                            </tr>
                        </table>
                        <%----%>
                        <table id="TableLay5" border="0" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol_need" width="20%">優先排序 </td>
                                <td colspan="3" class="whitecol" width="80%">
                                    <asp:TextBox ID="FirstSort" runat="server" Width="40%" MaxLength="3"></asp:TextBox>
                                    <asp:RegularExpressionValidator ID="FirstSort_chk1" runat="server" ControlToValidate="FirstSort" Display="None" ErrorMessage="班別資料「優先排序」請輸入數字" ValidationExpression="[0-9]{1,4}"></asp:RegularExpressionValidator>
                                    <asp:RequiredFieldValidator ID="FirstSort_chk2" runat="server" ControlToValidate="FirstSort" Display="None" ErrorMessage="班別資料「優先排序」為必填欄位"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">課程種類 </td>
                                <td class="whitecol" width="30%">
                                    <asp:RadioButtonList ID="RadioButtonList1" runat="server" Width="100%" CellSpacing="0" CellPadding="0" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="Y" Selected="True">學分班</asp:ListItem>
                                        <asp:ListItem Value="N">非學分班</asp:ListItem>
                                    </asp:RadioButtonList>
                                    <asp:CheckBox ID="IsBusiness" runat="server" Text="企業包班"></asp:CheckBox>
                                    <asp:CustomValidator ID="IsBusiness_chk1" runat="server" Display="None" ErrorMessage="班別資料『企業包班名稱』必須填寫" ClientValidationFunction="check_IsBusiness"></asp:CustomValidator>
                                </td>
                                <td class="bluecol" width="20%">企業包班名稱 </td>
                                <td class="whitecol" width="30%">
                                    <asp:TextBox ID="EnterpriseName" runat="server" MaxLength="50" Width="50%"></asp:TextBox></td>
                            </tr>
                            <tr id="PointType_TR" runat="server">
                                <td class="bluecol_need" width="20%">學分班種類 </td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:RadioButtonList ID="PointType" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">學士學分班</asp:ListItem>
                                        <asp:ListItem Value="2">碩士學分班</asp:ListItem>
                                        <asp:ListItem Value="3">博士學分班</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                            <tr id="Tr1" runat="server">
                                <td class="bluecol_need" width="20%">包班種類 </td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:RadioButtonList ID="PackageType" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非包班</asp:ListItem>
                                        <asp:ListItem Value="2">企業包班</asp:ListItem>
                                        <asp:ListItem Value="3">聯合企業包班</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">班別名稱 </td>
                                <td colspan="3" class="whitecol" width="80%">
                                    <asp:TextBox ID="ClassName" runat="server" MaxLength="35" Width="27%"></asp:TextBox>
                                    <asp:TextBox ID="PointName" runat="server" Width="27%" onfocus="this.blur()" MaxLength="10"></asp:TextBox>
                                    <asp:TextBox ID="PackageName" runat="server" Width="27%" onfocus="this.blur()" MaxLength="10"></asp:TextBox>
                                    <input id="Class_Unit" type="hidden" runat="server" />
                                    <asp:RequiredFieldValidator ID="fill18" runat="server" ControlToValidate="ClassName" Display="None" ErrorMessage="班別資料「班別名稱」為必填欄位"></asp:RequiredFieldValidator>
                                    <input onclick="open_hours()" value="時數迄日換算" type="button" class="button_b_M" />
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">訓練人數 </td>
                                <td class="whitecol" width="30%">
                                    <asp:TextBox ID="TNum" runat="server" Width="40%" MaxLength="7"></asp:TextBox>人
                                    <asp:RequiredFieldValidator ID="fill19" runat="server" ControlToValidate="TNum" Display="None" ErrorMessage="班別資料「訓練人數」為必填欄位"></asp:RequiredFieldValidator>
                                    <asp:RegularExpressionValidator ID="check8" runat="server" ControlToValidate="TNum" Display="None" ErrorMessage="班別資料「訓練人數」請輸入數字" ValidationExpression="[0-9]{1,4}"></asp:RegularExpressionValidator>
                                </td>
                                <td class="bluecol_need" width="20%">訓練時數 </td>
                                <td class="whitecol" width="30%">
                                    <asp:TextBox ID="THours" runat="server" Width="30%" MaxLength="7"></asp:TextBox>小時
                                    <asp:RequiredFieldValidator ID="fill20" runat="server" ControlToValidate="THours" Display="None" ErrorMessage="班別資料「訓練時數」為必填欄位"></asp:RequiredFieldValidator>
                                    <asp:RegularExpressionValidator ID="check9" runat="server" ControlToValidate="THours" Display="None" ErrorMessage="班別資料「訓練時數」請輸入數字" ValidationExpression="[0-9]{1,4}"></asp:RegularExpressionValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">訓練起日 </td>
                                <td class="whitecol" width="30%">
                                    <asp:TextBox ID="STDate" runat="server" Width="50%" MaxLength="10"></asp:TextBox>
                                    <span runat="server">
                                        <img style="cursor: pointer" id="date1" alt="" align="top" src="../../images/show-calendar.gif" runat="server" width="30" height="30"></span>
                                    <asp:RequiredFieldValidator ID="fill21" runat="server" ControlToValidate="STDate" Display="None" ErrorMessage="班別資料「訓練起日」請填寫"></asp:RequiredFieldValidator>
                                    <asp:CustomValidator ID="CustomValidator2" runat="server" ControlToValidate="STDate" Display="None" ErrorMessage="班別資料「訓練起日」不是正確的日期格式" ClientValidationFunction="check_date"></asp:CustomValidator>
                                </td>
                                <td class="bluecol_need" width="20%">訓練迄日 </td>
                                <td class="whitecol" width="30%">
                                    <asp:TextBox ID="FDDate" runat="server" Width="30%" MaxLength="10"></asp:TextBox>
                                    <span runat="server">
                                        <img style="cursor: pointer" id="date2" alt="" align="top" src="../../images/show-calendar.gif" runat="server" width="30" height="30"></span>
                                    <asp:RequiredFieldValidator ID="fill22" runat="server" ControlToValidate="FDDate" Display="None" ErrorMessage="班別資料「訓練迄日」請填寫"></asp:RequiredFieldValidator>
                                    <asp:CustomValidator ID="CustomValidator3" runat="server" ControlToValidate="FDDate" Display="None" ErrorMessage="班別資料「訓練迄日」不是正確的日期格式" ClientValidationFunction="check_date"></asp:CustomValidator>
                                    <asp:CustomValidator ID="CustomValidator5" runat="server" Display="None" ErrorMessage="班別資料：訓練起日不能比訓練迄日晚(或者同一天)" ClientValidationFunction="CheckTDate"></asp:CustomValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">期別(二碼) </td>
                                <td class="whitecol" width="30%">
                                    <asp:TextBox ID="CyclType" runat="server" Columns="5" MaxLength="2" Width="30%"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="fill28" runat="server" ControlToValidate="CyclType" Display="None" ErrorMessage="班別資料『期別』為必填欄位"></asp:RequiredFieldValidator>
                                    <asp:CustomValidator ID="CustomValidator4" runat="server" ControlToValidate="CyclType" Display="None" ErrorMessage="班別資料『期別』必須為大於0的兩位數字" ClientValidationFunction="check_CyclType"></asp:CustomValidator>
                                </td>
                                <td class="bluecol_need" width="20%">班數 </td>
                                <td class="whitecol" width="30%">
                                    <asp:TextBox ID="ClassCount" runat="server" onfocus="this.blur()" Columns="5" MaxLength="5" Width="20%">1</asp:TextBox>
                                    <asp:RequiredFieldValidator ID="fill29" runat="server" ControlToValidate="ClassCount" Display="None" ErrorMessage="請輸入班數"></asp:RequiredFieldValidator>
                                    <asp:RegularExpressionValidator ID="RegularExpressionValidator1" runat="server" ControlToValidate="ClassCount" Display="None" ErrorMessage="班數輸入大於0的數字" ValidationExpression="^0*[1-9](\d*$)"></asp:RegularExpressionValidator>
                                </td>
                            </tr>
                            <tr>
                                <td id="tdCredPoint" runat="server" class="bluecol_need" width="20%">學分數 </td>
                                <td class="whitecol" width="30%">
                                    <asp:TextBox ID="CredPoint" runat="server" Columns="5" MaxLength="2" Width="30%"></asp:TextBox></td>
                                <td id="RoomNameTD" runat="server" class="bluecol_need" width="20%">上課教室名稱 </td>
                                <td class="whitecol" width="30%">
                                    <asp:TextBox ID="RoomName" runat="server" Width="70%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">申請階段 </td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:DropDownList ID="AppStage" runat="server"></asp:DropDownList>
                                    <span style="cursor: pointer" title="點選可以查看說明" onclick="showHide(1)"><font color="red">設定說明</font></span>
                                    <table style="border-bottom: #9eb5cd 1px solid; position: absolute; filter: progid: DXImageTransform.Microsoft.Shadow(Color=#919899, Strength=4, Direction=135); border-left: #9eb5cd 1px solid; visibility: hidden; border-top: #9eb5cd 1px solid; border-right: #9eb5cd 1px solid" id="nxlayer_01" class="font" border="0" cellspacing="0" cellpadding="0" width="60%">
                                        <tbody>
                                            <tr>
                                                <td bgcolor="#ffffff" width="90%" align="center"><a onclick="showHide(0)" href="#"><font color="red">關閉[X]</font></a></td>
                                            </tr>
                                            <tr>
                                                <td style="height: 1px" class="dashline" height="1"><u></u></td>
                                            </tr>
                                            <tr>
                                                <td style="padding-left: 8px; padding-right: 8px" bgcolor="#f1faff" width="100%" colspan="2">
                                                    <asp:Label ID="labAppStageMsg" runat="server"></asp:Label></td>
                                            </tr>
                                        </tbody>
                                    </table>
                                </td>
                            </tr>
                            <tr id="FactModeTR" runat="server">
                                <td class="bluecol_need" width="20%">場地類型 </td>
                                <td colspan="3" class="whitecol" width="80%">
                                    <asp:RadioButtonList ID="FactMode" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                        <asp:ListItem Value="1">教室</asp:ListItem>
                                        <asp:ListItem Value="2">演講廳</asp:ListItem>
                                        <asp:ListItem Value="3">會議室</asp:ListItem>
                                        <asp:ListItem Value="99">其他(請說明)</asp:ListItem>
                                    </asp:RadioButtonList>
                                    <asp:TextBox ID="FactModeOther" runat="server" Width="30%"></asp:TextBox>
                                    <asp:CustomValidator ID="fill33" runat="server" Display="None" ErrorMessage="請選擇「場地類型」" ClientValidationFunction="CheckFactMode"></asp:CustomValidator>
                                    <asp:CustomValidator ID="fill37" runat="server" Display="None" ErrorMessage="請輸入場地類型[其他]" ClientValidationFunction="CheckFactModeOther"></asp:CustomValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">學科場地1 </td>
                                <td class="whitecol" width="30%">
                                    <asp:DropDownList ID="SciPlaceID" runat="server" AutoPostBack="True"></asp:DropDownList>
                                    <%-- <asp:requiredfieldvalidator id="fill39" runat="server" ControlToValidate="SciPlaceID" Display="None" ErrorMessage="請選擇學科場地"></asp:requiredfieldvalidator> --%>
                                </td>
                                <td class="bluecol" width="20%">術科場地1 </td>
                                <td class="whitecol" width="30%">
                                    <asp:DropDownList ID="TechPlaceID" runat="server" AutoPostBack="True"></asp:DropDownList>
                                    <%-- <asp:requiredfieldvalidator id="fill40" runat="server" ControlToValidate="TechPlaceID" Display="None" ErrorMessage="請選擇術科場地"></asp:requiredfieldvalidator> --%>
                                    <asp:CustomValidator ID="checkSciPlaceID" runat="server" Display="None" ErrorMessage="班別資料『學科場地』或『術科場地』必須填寫" ClientValidationFunction="check_PlaceID"></asp:CustomValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">學科場地2 </td>
                                <td class="whitecol" width="30%">
                                    <asp:DropDownList ID="SciPlaceID2" runat="server" AutoPostBack="True"></asp:DropDownList></td>
                                <td class="bluecol" width="20%">術科場地2 </td>
                                <td class="whitecol" width="30%">
                                    <asp:DropDownList ID="TechPlaceID2" runat="server" AutoPostBack="True"></asp:DropDownList>
                                    <asp:CustomValidator ID="checkSciPlaceID2" runat="server" Display="None" ErrorMessage="班別資料『學科場地』、『學科場地2』必項填寫 其中一項;『術科場地』、『術科場地2』必須填寫 其中一項" ClientValidationFunction="check_PlaceID"></asp:CustomValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">學科上課地址 </td>
                                <td colspan="3" class="whitecol" width="80%">
                                    <asp:DropDownList ID="Taddress2" runat="server" AutoPostBack="True"></asp:DropDownList></td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">術科上課地址 </td>
                                <td colspan="3" class="whitecol" width="80%">
                                    <asp:DropDownList ID="Taddress3" runat="server" AutoPostBack="True"></asp:DropDownList></td>
                            </tr>
                            <tr>
                                <td class="whitecol" width="20%">&nbsp; </td>
                                <td class="whitecol" colspan="3" width="80%"><font color="red">※班別資料『學科場地』、『學科場地2』、『術科場地』、『術科場地2』必須填寫 其中一項 ;<br />
                                    【學科上課地址】、【術科上課地址】至少要設定其中一項 </font></td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">容納人數 </td>
                                <td colspan="3" class="whitecol" width="80%">
                                    <asp:TextBox ID="ConNum" runat="server" Columns="5" MaxLength="5" Width="12%"></asp:TextBox>人
                                    <asp:RequiredFieldValidator ID="fill34" runat="server" ControlToValidate="ConNum" Display="None" ErrorMessage="請輸入容納人數"></asp:RequiredFieldValidator>
                                    <asp:RegularExpressionValidator ID="RegularExpressionValidator3" runat="server" ControlToValidate="ConNum" Display="None" ErrorMessage="容納人數必須為數字" ValidationExpression="^\d*"></asp:RegularExpressionValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">聯絡人 </td>
                                <td class="whitecol" width="30%">
                                    <asp:TextBox ID="ContactName" runat="server" MaxLength="50" Width="40%"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="Requiredfieldvalidator4" runat="server" ControlToValidate="ContactName" Display="None" ErrorMessage="請輸入聯絡人"></asp:RequiredFieldValidator>
                                </td>
                                <td class="bluecol_need" width="20%">電話 </td>
                                <td class="whitecol" width="30%">
                                    <asp:TextBox ID="ContactPhone" runat="server" MaxLength="50" Width="50%"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="Requiredfieldvalidator5" runat="server" ControlToValidate="ContactPhone" Display="None" ErrorMessage="請輸入電話"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">電子郵件 </td>
                                <td class="whitecol" width="30%">
                                    <asp:TextBox ID="ContactEmail" runat="server" MaxLength="64" Width="80%"></asp:TextBox>
                                    <asp:RegularExpressionValidator ID="chkContactEmail1" runat="server" ControlToValidate="ContactEmail" Display="None" ErrorMessage="電子郵件輸入錯誤" ValidationExpression="\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"></asp:RegularExpressionValidator>
                                </td>
                                <td class="bluecol" width="20%">傳真 </td>
                                <td class="whitecol" width="30%">
                                    <asp:TextBox ID="ContactFax" runat="server" MaxLength="64" Width="50%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">訓練職能<%--課程類別--%></td>
                                <td class="whitecol" width="30%">
                                    <asp:DropDownList ID="ClassCate" runat="server"></asp:DropDownList>
                                    <asp:RequiredFieldValidator ID="fill35" runat="server" ControlToValidate="ClassCate" Display="None" ErrorMessage="請選擇訓練職能"></asp:RequiredFieldValidator>
                                </td>
                                <td class="bluecol_need" width="20%">報名繳費方式<%--課程類別--%>
                                </td>
                                <td class="whitecol" width="30%">
                                    <asp:RadioButtonList ID="EnterSupplyStyle" runat="server" CssClass="font" RepeatLayout="Flow">
                                        <asp:ListItem Value="1" Selected="True">1.報名時應先繳全額訓練費用，待結訓審核通過後核撥補助款</asp:ListItem>
                                        <asp:ListItem Value="2">2.報名時應先繳50%訓練費用，待結訓審核通過後核撥補助款</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                            <tr id="ContentTR" runat="server">
                                <td class="bluecol_need" width="20%">課程大綱 </td>
                                <td colspan="3" class="whitecol" width="80%">
                                    <asp:TextBox ID="Content" runat="server" Columns="77" Rows="5" TextMode="MultiLine"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="fill36" runat="server" ControlToValidate="Content" Display="None" ErrorMessage="請輸入課程大綱"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="4" align="center" class="table_title" style="width: 100%">課程大綱 </td>
                            </tr>
                            <tr>
                                <td colspan="4">
                                    <table id="Table5A" border="0" cellspacing="1" cellpadding="1" width="100%">
                                        <tr>
                                            <td class="bluecol" style="width: 15%">日期 </td>
                                            <td class="bluecol" style="width: 10%">授課時段 </td>
                                            <td class="bluecol" style="width: 15%">授課時間 </td>
                                            <td class="bluecol" style="width: 5%">時數 </td>
                                            <td class="bluecol" style="width: 10%">課程進度／內容 </td>
                                            <td class="bluecol" style="width: 7%">學／術科 </td>
                                            <td class="bluecol" style="width: 13%">上課地點 </td>
                                            <td class="bluecol" style="width: 10%">任課教師 </td>
                                            <td class="bluecol" style="width: 10%">助教 </td>
                                            <td class="bluecol" style="width: 5%" align="center">功能 </td>
                                        </tr>
                                        <tr>
                                            <td class="whitecol" style="width: 15%"><span runat="server">
                                                <asp:TextBox ID="STrainDate" runat="server" Width="70%" MaxLength="10"></asp:TextBox><img style="cursor: pointer" id="date3" alt="" align="top" src="../../images/show-calendar.gif" runat="server" width="30" height="30" /></span></td>
                                            <td id="tdTPERIOD28" runat="server" class="whitecol" style="width: 10%" align="center">
                                                <asp:CheckBox ID="TPERIOD28_1" runat="server" Text="早上" ToolTip="7:00-13:00" /><br />
                                                <asp:CheckBox ID="TPERIOD28_2" runat="server" Text="下午" ToolTip="13:00-18:00" /><br />
                                                <asp:CheckBox ID="TPERIOD28_3" runat="server" Text="晚上" ToolTip="18:00-22:00" />
                                            </td>
                                            <td class="whitecol" style="width: 15%" align="center">
                                                <asp:DropDownList ID="ddlpnH1" runat="server"></asp:DropDownList>：<asp:DropDownList ID="ddlpnM1" runat="server"></asp:DropDownList><br />
                                                ~<br />
                                                <asp:DropDownList ID="ddlpnH2" runat="server"></asp:DropDownList>：<asp:DropDownList ID="ddlpnM2" runat="server"></asp:DropDownList><br />
                                                <asp:TextBox ID="PName" runat="server" onfocus="this.blur()" Visible="False"></asp:TextBox>
                                            </td>
                                            <td class="whitecol" style="width: 5%" align="center">
                                                <asp:TextBox ID="PHour" runat="server" Width="60%" MaxLength="3"></asp:TextBox></td>
                                            <td class="whitecol" style="width: 10%">
                                                <asp:TextBox ID="PCont" runat="server" Width="80%" Columns="50" Rows="5" TextMode="MultiLine"></asp:TextBox></td>
                                            <td class="whitecol" style="width: 7%">
                                                <asp:DropDownList ID="Classification1" runat="server" AutoPostBack="True">
                                                    <asp:ListItem Value="0">請選擇</asp:ListItem>
                                                    <asp:ListItem Value="1">學科</asp:ListItem>
                                                    <asp:ListItem Value="2">術科</asp:ListItem>
                                                </asp:DropDownList>
                                            </td>
                                            <td class="whitecol" style="width: 13%">
                                                <asp:DropDownList ID="PTID1" runat="server"></asp:DropDownList><br />
                                                <asp:DropDownList ID="PTID2" runat="server"></asp:DropDownList>
                                            </td>
                                            <td class="whitecol" style="width: 10%" align="center">
                                                <input id="OLessonTeah1Value" type="hidden" name="OLessonTeah1Value" runat="server">
                                                <asp:TextBox ID="OLessonTeah1" runat="server" onfocus="this.blur()" Columns="5" ToolTip="點選兩下可以跳出視窗選擇教師" Width="80%"></asp:TextBox>
                                            </td>
                                            <td class="whitecol" style="width: 10%" align="center">
                                                <input id="OLessonTeah2Value" type="hidden" name="OLessonTeah2Value" runat="server">
                                                <asp:TextBox ID="OLessonTeah2" runat="server" onfocus="this.blur()" Columns="5" ToolTip="點選兩下可以跳出視窗選擇助教" Width="80%"></asp:TextBox>
                                            </td>
                                            <td class="whitecol" style="width: 5%">
                                                <asp:Button ID="Button1" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_M"></asp:Button></td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="4" class="whitecol">
                                    <table id="Datagrid3Table" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                        <tr>
                                            <td>
                                                <asp:DataGrid ID="Datagrid3" Width="100%" runat="server" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                    <EditItemStyle Wrap="False"></EditItemStyle>
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                                    <Columns>
                                                        <asp:TemplateColumn HeaderText="日期">
                                                            <HeaderStyle Wrap="False" Width="10%"></HeaderStyle>
                                                            <ItemStyle Wrap="False"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="STrainDateLabel" runat="server"></asp:Label>
                                                            </ItemTemplate>
                                                            <FooterStyle Wrap="False"></FooterStyle>
                                                            <EditItemTemplate>
                                                                <asp:TextBox ID="STrainDateTxt" runat="server" Width="73px"></asp:TextBox><img id="Img2" style="cursor: pointer" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30">
                                                            </EditItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="授課時段">
                                                            <HeaderStyle Width="10%" />
                                                            <ItemStyle Wrap="false" HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:CheckBox ID="TPERIOD28_1t" runat="server" Text="早上" ToolTip="7:00-13:00" Enabled="false" />
                                                                <asp:CheckBox ID="TPERIOD28_2t" runat="server" Text="下午" ToolTip="13:00-18:00" Enabled="false" />
                                                                <asp:CheckBox ID="TPERIOD28_3t" runat="server" Text="晚上" ToolTip="18:00-22:00" Enabled="false" />
                                                            </ItemTemplate>
                                                            <EditItemTemplate>
                                                                <asp:CheckBox ID="TPERIOD28_1e" runat="server" Text="早上" ToolTip="7:00-13:00" />
                                                                <asp:CheckBox ID="TPERIOD28_2e" runat="server" Text="下午" ToolTip="13:00-18:00" />
                                                                <asp:CheckBox ID="TPERIOD28_3e" runat="server" Text="晚上" ToolTip="18:00-22:00" />
                                                            </EditItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="授課時間">
                                                            <HeaderStyle Width="10%" />
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="PNameLabel" runat="server"></asp:Label>
                                                            </ItemTemplate>
                                                            <EditItemTemplate>
                                                                <asp:DropDownList ID="Eddlh1" runat="server"></asp:DropDownList>：<asp:DropDownList ID="Eddlm1" runat="server"></asp:DropDownList>~<br>
                                                                <asp:DropDownList ID="Eddlh2" runat="server"></asp:DropDownList>：
                                                            <asp:DropDownList ID="Eddlm2" runat="server"></asp:DropDownList>
                                                                <asp:TextBox ID="PNameTxt" runat="server" onfocus="this.blur()" Visible="False"></asp:TextBox>
                                                            </EditItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="時數">
                                                            <HeaderStyle Width="7%" />
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="PHourLabel" runat="server"></asp:Label>
                                                            </ItemTemplate>
                                                            <EditItemTemplate>
                                                                <asp:TextBox ID="PHourTxt" runat="server" Width="26px"></asp:TextBox>
                                                            </EditItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="課程進度／內容">
                                                            <HeaderStyle Width="17%" />
                                                            <ItemTemplate>
                                                                <asp:TextBox ID="PContText" runat="server" onfocus="this.blur()" Width="100%" Columns="50" TextMode="MultiLine" Rows="5" Enabled="False" Height="58px"></asp:TextBox>
                                                            </ItemTemplate>
                                                            <EditItemTemplate>
                                                                <asp:TextBox ID="PContEdit" runat="server" Width="170px" Columns="50" TextMode="MultiLine" Rows="5" Height="58px"></asp:TextBox>
                                                            </EditItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="學／術科">
                                                            <HeaderStyle Width="8%" />
                                                            <ItemTemplate>
                                                                <asp:DropDownList ID="drpClassification1" runat="server" Enabled="False" AutoPostBack="True">
                                                                    <asp:ListItem Value="1">學科</asp:ListItem>
                                                                    <asp:ListItem Value="2">術科</asp:ListItem>
                                                                </asp:DropDownList>
                                                            </ItemTemplate>
                                                            <EditItemTemplate>
                                                                <asp:DropDownList ID="drpClassEdit" runat="server">
                                                                    <asp:ListItem Value="1">學科</asp:ListItem>
                                                                    <asp:ListItem Value="2">術科</asp:ListItem>
                                                                </asp:DropDownList>
                                                            </EditItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="上課地點">
                                                            <HeaderStyle Width="15%" />
                                                            <ItemTemplate>
                                                                <asp:DropDownList ID="drpPTID" runat="server" Width="160px" Enabled="False"></asp:DropDownList>
                                                            </ItemTemplate>
                                                            <EditItemTemplate>
                                                                <asp:DropDownList ID="drpPTIDEdit1" runat="server"></asp:DropDownList>
                                                                <asp:DropDownList ID="drpPTIDEdit2" runat="server"></asp:DropDownList>
                                                            </EditItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="任課教師">
                                                            <HeaderStyle Width="9%" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                            <ItemTemplate>
                                                                <input id="Tech1Value" type="hidden" name="Tech1Value" runat="server">
                                                                <asp:TextBox ID="Tech1Text" runat="server" onfocus="this.blur()" Columns="5" Enabled="False" Width="80%"></asp:TextBox>
                                                            </ItemTemplate>
                                                            <EditItemTemplate>
                                                                <input id="Tech1ValueEdit" type="hidden" runat="server">
                                                                <asp:TextBox ID="Tech1Edit" runat="server" onfocus="this.blur()" Columns="5" ToolTip="點選兩下可以跳出視窗選擇教師"></asp:TextBox>
                                                            </EditItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="助教">
                                                            <HeaderStyle Width="9%" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                            <ItemTemplate>
                                                                <input id="Tech2Value" type="hidden" name="Tech2Value" runat="server">
                                                                <asp:TextBox ID="Tech2Text" runat="server" onfocus="this.blur()" Columns="5" Enabled="False" Width="80%"></asp:TextBox>
                                                            </ItemTemplate>
                                                            <EditItemTemplate>
                                                                <input id="Tech2ValueEdit" type="hidden" runat="server">
                                                                <asp:TextBox ID="Tech2Edit" runat="server" onfocus="this.blur()" Columns="5" ToolTip="點選兩下可以跳出視窗選擇助教"></asp:TextBox>
                                                            </EditItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="功能">
                                                            <ItemStyle Width="5%"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Button ID="Button6" runat="server" Text="修改" CausesValidation="False" CommandName="edit" CssClass="asp_button_S"></asp:Button>
                                                                <asp:Button ID="Button7" runat="server" Text="刪除" CausesValidation="False" CommandName="del" CssClass="asp_button_S"></asp:Button>
                                                            </ItemTemplate>
                                                            <EditItemTemplate>
                                                                <asp:Button ID="Button10" runat="server" Text="儲存" CausesValidation="False" CommandName="save" CssClass="asp_button_S"></asp:Button>
                                                                <asp:Button ID="Button11" runat="server" Text="取消" CausesValidation="False" CommandName="cancel" CssClass="asp_button_S"></asp:Button>
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
                            <tr>
                                <td colspan="4" align="center" class="table_title" width="100%">上課時間 </td>
                            </tr>
                            <tr>
                                <td colspan="4">
                                    <table id="Table3" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                                        <tr>
                                            <td align="center" class="bluecol" width="10%">星期 </td>
                                            <td align="center" class="bluecol" width="80%">時間 </td>
                                            <td align="center" class="bluecol" width="10%">功能 </td>
                                        </tr>
                                        <tr>
                                            <td class="whitecol" align="center">
                                                <asp:DropDownList ID="Weeks" runat="server"></asp:DropDownList></td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="Times" runat="server" Columns="50" MaxLength="50" Width="60%"></asp:TextBox>
                                                <font color="red">輸入範例，半型字體例如：18:00~21:00，多筆以 ; 做分隔</font>
                                            </td>
                                            <td class="whitecol" align="center">
                                                <asp:Button ID="Button29" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_M"></asp:Button></td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="4">
                                    <table id="DataGrid1Table" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                        <tr>
                                            <td>
                                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                                    <Columns>
                                                        <asp:TemplateColumn HeaderText="星期">
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <HeaderStyle Width="20%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="Weeks1" runat="server"></asp:Label>
                                                            </ItemTemplate>
                                                            <EditItemTemplate>
                                                                <asp:DropDownList ID="Weeks2" runat="server"></asp:DropDownList>
                                                            </EditItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="上課時段">
                                                            <HeaderStyle Width="80%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="Times1" runat="server"></asp:Label>
                                                            </ItemTemplate>
                                                            <EditItemTemplate>
                                                                <asp:TextBox ID="Times2" runat="server"></asp:TextBox>
                                                            </EditItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="功能">
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <HeaderStyle Width="20%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Button ID="Button2" runat="server" CausesValidation="False" Text="修改" CommandName="edit" CssClass="asp_button_M"></asp:Button>
                                                                <asp:Button ID="Button3" runat="server" CausesValidation="False" Text="刪除" CommandName="del" CssClass="asp_button_M"></asp:Button>
                                                            </ItemTemplate>
                                                            <EditItemTemplate>
                                                                <asp:Button ID="Button4" runat="server" CausesValidation="False" Text="儲存" CommandName="save" CssClass="asp_button_M"></asp:Button>
                                                                <asp:Button ID="Button5" runat="server" CausesValidation="False" Text="取消" CommandName="cancel" CssClass="asp_button_M"></asp:Button>
                                                            </EditItemTemplate>
                                                        </asp:TemplateColumn>
                                                    </Columns>
                                                    <PagerStyle Visible="False"></PagerStyle>
                                                </asp:DataGrid>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td></td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="4">
                                    <%-- Plan_BusPackage --%>
                                    <table id="Datagrid4headTable" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                        <tr>
                                            <td colspan="4" align="center" class="table_title">包班事業單位資料 </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol" align="center" width="40%">企業名稱 </td>
                                            <td class="bluecol" align="center" width="25%">服務單位統一編號 </td>
                                            <td class="bluecol" align="center" width="25%">保險證號 </td>
                                            <td class="bluecol" align="center" width="10%">功能 </td>
                                        </tr>
                                        <tr>
                                            <td class="whitecol" align="center">
                                                <asp:TextBox ID="txtUname" runat="server" Columns="40" MaxLength="50" Width="80%"></asp:TextBox></td>
                                            <td class="whitecol" align="center">
                                                <asp:TextBox ID="txtIntaxno" runat="server" Columns="9" MaxLength="10" Width="60%"></asp:TextBox></td>
                                            <td class="whitecol" align="center">
                                                <asp:TextBox ID="txtUbno" runat="server" Columns="9" MaxLength="9" Width="60%"></asp:TextBox></td>
                                            <td class="whitecol" align="center">
                                                <asp:Button ID="btnAddBusPackage" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_M"></asp:Button></td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="4">
                                    <%-- Plan_BusPackage --%>
                                    <table id="Datagrid4Table" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                        <tr>
                                            <td>
                                                <asp:DataGrid ID="Datagrid4" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                                    <Columns>
                                                        <asp:TemplateColumn HeaderText="企業名稱">
                                                            <HeaderStyle Width="21%" />
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="slsbUname" runat="server"></asp:Label>
                                                            </ItemTemplate>
                                                            <EditItemTemplate>
                                                                <asp:TextBox ID="ttxtUname" runat="server"></asp:TextBox>
                                                            </EditItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="服務單位統一編號">
                                                            <HeaderStyle Width="21%" />
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="slabIntaxno" runat="server"></asp:Label>
                                                            </ItemTemplate>
                                                            <EditItemTemplate>
                                                                <asp:TextBox ID="ttxtIntaxno" runat="server"></asp:TextBox>
                                                            </EditItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="保險證號">
                                                            <HeaderStyle Width="21%" />
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="slabUbno" runat="server"></asp:Label>
                                                            </ItemTemplate>
                                                            <EditItemTemplate>
                                                                <asp:TextBox ID="ttxtUbno" runat="server"></asp:TextBox>
                                                            </EditItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="功能">
                                                            <HeaderStyle Width="7%"></HeaderStyle>
                                                            <ItemStyle VerticalAlign="Middle" HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Button ID="Button17" runat="server" CausesValidation="False" Text="修改" CommandName="xedit" CssClass="asp_button_M"></asp:Button>
                                                                <asp:Button ID="Button18" runat="server" CausesValidation="False" Text="刪除" CommandName="xdel" CssClass="asp_button_M"></asp:Button>
                                                            </ItemTemplate>
                                                            <EditItemTemplate>
                                                                <asp:Button ID="Button19" runat="server" CausesValidation="False" Text="儲存" CommandName="xsave" CssClass="asp_button_M"></asp:Button>
                                                                <asp:Button ID="Button20" runat="server" CausesValidation="False" Text="取消" CommandName="xcancel" CssClass="asp_button_M"></asp:Button>
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
                        <table id="TableLay6" border="0" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td colspan="4">
                                    <table id="Table6" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                        <tr>
                                            <td class="bluecol" width="20%">總計 </td>
                                            <td class="whitecol" width="80%">
                                                <asp:TextBox ID="TotalCost3" runat="server" Width="20%" MaxLength="10"></asp:TextBox>(若選擇學分班可自行輸入總計)
                                                <asp:CustomValidator ID="Customvalidator6" runat="server" Display="None" ErrorMessage="經費來源：經費必須大於0" ClientValidationFunction="CheckCost"></asp:CustomValidator>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol" width="20%">其他說明 </td>
                                            <td width="80%"></td>
                                        </tr>
                                        <tr>
                                            <td colspan="2" class="whitecol">
                                                <asp:TextBox ID="tNote2b" runat="server" Columns="77" Rows="4" TextMode="MultiLine" ToolTip="其他說明(欄位字數為1000)"></asp:TextBox>
                                                <asp:Button Style="z-index: 0" ID="btnUptNote2b" runat="server" CausesValidation="False" Text="修改" CssClass="asp_button_M"></asp:Button>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <table id="TableCost2" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                        <tr>
                                            <td width="100%">
                                                <table id="Table4" class="font" border="0" cellspacing="1" cellpadding="1" runat="server" width="100%">
                                                    <tr>
                                                        <td class="bluecol" align="center" width="40%">項目 </td>
                                                        <td class="bluecol" align="center" width="30%">單價 </td>
                                                        <td class="bluecol" align="center" width="15%">計價數量 </td>
                                                        <td class="bluecol" width="15%"></td>
                                                    </tr>
                                                    <tr>
                                                        <td class="whitecol" width="40%">
                                                            <asp:DropDownList ID="CostID2" runat="server">
                                                                <asp:ListItem Value="===請選擇===">===請選擇===</asp:ListItem>
                                                            </asp:DropDownList>
                                                        </td>
                                                        <td class="whitecol" width="30%">
                                                            <asp:TextBox ID="OPrice2" runat="server" Columns="7" MaxLength="7" Width="60%"></asp:TextBox>元／
                                                            <asp:Label ID="ItemCostName" runat="server"></asp:Label><br />
                                                            (僅准輸入整數)*<%--(可輸入小數點第二位)*--%>
                                                        </td>
                                                        <td class="whitecol" width="15%" align="center">
                                                            <input style="width: 60%" id="Itemage" class="font" value="1" maxlength="5" name="Itemage" runat="server" /></td>
                                                        <td class="whitecol" width="15%" align="center">
                                                            <asp:Button ID="Button9" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_M"></asp:Button></td>
                                                    </tr>
                                                    <tr>
                                                        <td colspan="5" class="whitecol">說明:如果新增該項目金額時，沒有計價數量時，請輸入1。<br />
                                                            計價數量:若項目是採用每小時/每次作為單價時，而計價數量則是輸入幾小時／幾次，例如:<br />
                                                            學科-鐘點費：單價15元／每小時*3小時，單價=15，計價數量=3<br />
                                                            教材費-(人)：單價15元／每人*20人，單價=15，計價數量=20<br />
                                                            場地費-(次)：單價15元／每次*3次，單價=15，計價數量=3<br />
                                                            行政管理費：單價100元／每班*1班，單價=100，計價數量=1<br />
                                                            其他費用：單價100元／每班*1班，單價=100，計價數量=1
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <table id="DataGrid2Table" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                                    <tr>
                                                        <td>
                                                            <asp:DataGrid ID="DataGrid2" runat="server" CssClass="font" AutoGenerateColumns="False" BorderColor="Gray" Width="100%" CellPadding="8">
                                                                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                                <AlternatingItemStyle BackColor="#F5F5F5" />
                                                                <Columns>
                                                                    <asp:TemplateColumn HeaderText="項目">
                                                                        <HeaderStyle Width="20%" />
                                                                        <ItemTemplate>
                                                                            <asp:DropDownList ID="drpCostID1" runat="server" Enabled="False">
                                                                                <asp:ListItem Value="===請選擇===">===請選擇===</asp:ListItem>
                                                                            </asp:DropDownList>
                                                                        </ItemTemplate>
                                                                        <EditItemTemplate>
                                                                            <asp:DropDownList ID="drpCostID2" runat="server" Enabled="False">
                                                                                <asp:ListItem Value="===請選擇===">===請選擇===</asp:ListItem>
                                                                            </asp:DropDownList>
                                                                        </EditItemTemplate>
                                                                    </asp:TemplateColumn>
                                                                    <asp:TemplateColumn HeaderText="單價">
                                                                        <HeaderStyle Width="20%" />
                                                                        <ItemStyle HorizontalAlign="Center" />
                                                                        <ItemTemplate>
                                                                            <asp:Label ID="DataGrid2Label1" runat="server"></asp:Label>
                                                                        </ItemTemplate>
                                                                        <EditItemTemplate>
                                                                            <asp:TextBox ID="DataGrid2TextBox1" runat="server" Columns="3"></asp:TextBox>
                                                                        </EditItemTemplate>
                                                                    </asp:TemplateColumn>
                                                                    <asp:TemplateColumn HeaderText="計價數量">
                                                                        <HeaderStyle Width="20%" />
                                                                        <ItemStyle HorizontalAlign="Center" />
                                                                        <ItemTemplate>
                                                                            <asp:Label ID="DataGrid2Label2" runat="server"></asp:Label>
                                                                        </ItemTemplate>
                                                                        <EditItemTemplate>
                                                                            <asp:TextBox ID="DataGrid2TextBox2" runat="server" Columns="3"></asp:TextBox>
                                                                        </EditItemTemplate>
                                                                    </asp:TemplateColumn>
                                                                    <asp:TemplateColumn HeaderText="小計">
                                                                        <HeaderStyle Width="20%" />
                                                                        <ItemStyle HorizontalAlign="Center" />
                                                                        <ItemTemplate>
                                                                            <asp:Label ID="DataGrid2Label3" runat="server"></asp:Label>
                                                                        </ItemTemplate>
                                                                        <EditItemTemplate>
                                                                            <asp:Label ID="DataGrid2Label3b" runat="server"></asp:Label>
                                                                        </EditItemTemplate>
                                                                    </asp:TemplateColumn>
                                                                    <asp:TemplateColumn HeaderText="功能">
                                                                        <HeaderStyle Width="20%" />
                                                                        <ItemStyle HorizontalAlign="Center" />
                                                                        <ItemTemplate>
                                                                            <asp:Button ID="Button13" runat="server" Text="刪除" CausesValidation="False" CommandName="del1" CssClass="asp_button_M"></asp:Button>
                                                                            <asp:Button ID="Button12" runat="server" Text="修改" CausesValidation="False" CommandName="edit1" CssClass="asp_button_M"></asp:Button>
                                                                        </ItemTemplate>
                                                                        <EditItemTemplate>
                                                                            <asp:Button ID="Button14" runat="server" Text="更新" CausesValidation="False" CommandName="update1" CssClass="asp_button_M"></asp:Button>
                                                                            <asp:Button ID="Button15" runat="server" Text="取消" CausesValidation="False" CommandName="cancel1" CssClass="asp_button_M"></asp:Button>
                                                                        </EditItemTemplate>
                                                                    </asp:TemplateColumn>
                                                                </Columns>
                                                                <PagerStyle Visible="False"></PagerStyle>
                                                            </asp:DataGrid>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td colspan="4">
                                                            <table class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                                                                <tr>
                                                                    <td class="bluecol" width="18%">&nbsp;&nbsp; 總計 </td>
                                                                    <td class="whitecol" width="82%">
                                                                        <asp:TextBox ID="TotalCost2" runat="server" Width="18%" onfocus="this.blur()" BorderStyle="None"></asp:TextBox></td>
                                                                </tr>
                                                            </table>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <table id="TableCost5" class="font" border="0" cellspacing="1" cellpadding="1" runat="server" width="100%">
                                                                <tr>
                                                                    <td class="bluecol" align="center" width="80%">材料品名 </td>
                                                                    <td class="bluecol" align="center" width="20%">&nbsp;&nbsp; </td>
                                                                </tr>
                                                                <tr>
                                                                    <td class="whitecol" width="70%" align="center">
                                                                        <asp:TextBox ID="PMcName" runat="server" Columns="44" MaxLength="50" Width="60%"></asp:TextBox></td>
                                                                    <td class="whitecol" width="30%" align="center">
                                                                        <asp:Button ID="btnAddMaterial" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_M"></asp:Button></td>
                                                                </tr>
                                                                <tr>
                                                                    <td colspan="2" width="100%">
                                                                        <asp:DataGrid ID="DataGrid5" runat="server" CssClass="font" AutoGenerateColumns="False" BorderColor="Gray" Width="100%" CellPadding="8">
                                                                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                                            <AlternatingItemStyle BackColor="#F5F5F5" />
                                                                            <Columns>
                                                                                <asp:TemplateColumn HeaderText="材料品名">
                                                                                    <HeaderStyle Width="80%" />
                                                                                    <ItemTemplate>
                                                                                        <asp:Label ID="labPMcNAME" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.cName") %>'></asp:Label>
                                                                                    </ItemTemplate>
                                                                                    <EditItemTemplate>
                                                                                        <asp:TextBox ID="txtPMcNAME" runat="server" Columns="44" MaxLength="50" Text='<%# DataBinder.Eval(Container, "DataItem.cName") %>'></asp:TextBox>
                                                                                    </EditItemTemplate>
                                                                                </asp:TemplateColumn>
                                                                                <asp:TemplateColumn HeaderText="功能">
                                                                                    <HeaderStyle Width="20%" />
                                                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                                                    <ItemTemplate>
                                                                                        <asp:Button ID="btnDEL5" runat="server" Text="刪除" CausesValidation="False" CommandName="DEL5" CssClass="asp_button_M"></asp:Button>
                                                                                        <asp:Button ID="btnEDT5" runat="server" Text="修改" CausesValidation="False" CommandName="EDT5" CssClass="asp_button_M"></asp:Button>
                                                                                    </ItemTemplate>
                                                                                    <EditItemTemplate>
                                                                                        <asp:Button ID="btnUPD5" runat="server" Text="更新" CausesValidation="False" CommandName="UPD5" CssClass="asp_button_M"></asp:Button>
                                                                                        <asp:Button ID="btnCLS5" runat="server" Text="取消" CausesValidation="False" CommandName="CLS5" CssClass="asp_button_M"></asp:Button>
                                                                                    </EditItemTemplate>
                                                                                </asp:TemplateColumn>
                                                                            </Columns>
                                                                            <PagerStyle Visible="False"></PagerStyle>
                                                                        </asp:DataGrid>
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                            <table id="TableCost6" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                                                <tr>
                                                                    <td class="table_title" colspan="8" style="width: 100%">一人份材料明細</td>
                                                                </tr>
                                                                <tr>
                                                                    <td colspan="8" class="whitecol">
                                                                        <asp:Label ID="Label7" runat="server" CssClass="font">匯入明細</asp:Label>
                                                                        <input id="File1" type="file" size="50" name="File1" runat="server" accept=".csv,.xls" />
                                                                        <asp:Button ID="BtnImport1" runat="server" CausesValidation="False" Text="匯入明細" CssClass="asp_button_M" />(必須為csv格式)
                                                                        <asp:HyperLink ID="HyperLink1" runat="server" CssClass="font" ForeColor="#8080FF" NavigateUrl="../../Doc/PersonCost_Imp.zip">下載匯入格式檔</asp:HyperLink>
                                                                    </td>
                                                                </tr>
                                                                <tr align="center">
                                                                    <td class="bluecol" style="width: 10%">項次 </td>
                                                                    <td class="bluecol" style="width: 15%">品名 </td>
                                                                    <td class="bluecol" style="width: 15%">規格 </td>
                                                                    <td class="bluecol" style="width: 10%">單位 </td>
                                                                    <td class="bluecol" style="width: 15%">單價 </td>
                                                                    <td class="bluecol" style="width: 10%">每人數量 </td>
                                                                    <td class="bluecol" style="width: 15%">用途說明 </td>
                                                                    <td class="bluecol" style="width: 10%">功能 </td>
                                                                </tr>
                                                                <tr align="center">
                                                                    <td class="whitecol" style="width: 10%">
                                                                        <asp:TextBox ID="tItemNo6" runat="server" Columns="5" MaxLength="5" Width="80%"></asp:TextBox></td>
                                                                    <td class="whitecol" style="width: 15%">
                                                                        <asp:TextBox ID="tCName6" runat="server" Columns="10" MaxLength="30" Width="90%"></asp:TextBox></td>
                                                                    <td class="whitecol" style="width: 15%">
                                                                        <asp:TextBox ID="tStandard6" runat="server" Columns="10" MaxLength="300" Width="80%"></asp:TextBox></td>
                                                                    <td class="whitecol" style="width: 10%">
                                                                        <asp:TextBox ID="tUnit6" runat="server" Columns="5" MaxLength="5"></asp:TextBox></td>
                                                                    <td class="whitecol" style="width: 15%">
                                                                        <asp:TextBox ID="tPrice6" runat="server" Columns="5" MaxLength="5"></asp:TextBox></td>
                                                                    <td class="whitecol" style="width: 10%">
                                                                        <asp:TextBox ID="tPerCount6" runat="server" Columns="5" MaxLength="5"></asp:TextBox></td>
                                                                    <td class="whitecol" style="width: 15%">
                                                                        <asp:TextBox ID="tPurpose6" runat="server" Columns="35" MaxLength="300"></asp:TextBox></td>
                                                                    <td class="whitecol" style="width: 10%">
                                                                        <asp:Button ID="btnAddCost6" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_M"></asp:Button></td>
                                                                </tr>
                                                                <tr>
                                                                    <td colspan="8" width="100%">
                                                                        <table id="DataGrid6Table" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                                                            <tr>
                                                                                <td>
                                                                                    <asp:DataGrid ID="DataGrid6" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                                                                        <Columns>
                                                                                            <asp:TemplateColumn HeaderText="項次">
                                                                                                <HeaderStyle Width="6%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lItemNo6" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eItemNo6" runat="server" Columns="5" MaxLength="5"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="品名">
                                                                                                <HeaderStyle Width="14%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lCName6" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eCName6" runat="server" Columns="10" MaxLength="30"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="規格">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lStandard6" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eStandard6" runat="server" Columns="10" MaxLength="300"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="單位">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lUnit6" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eUnit6" runat="server" Columns="3" MaxLength="5"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="單價">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lPrice6" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="ePrice6" runat="server" Columns="3" MaxLength="5"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="每人數量">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lPerCount6" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="ePerCount6" runat="server" Columns="3" MaxLength="5"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="訓練人數">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lTNum6" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eTNum6" runat="server" Columns="3"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="總數量">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lTotal6" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eTotal6" runat="server" Columns="3"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="小計">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lsubtotal6" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="esubtotal6" runat="server" Columns="3"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="用途說明">
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lPurPose6" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="ePurPose6" runat="server" Columns="15" MaxLength="300"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="功能">
                                                                                                <HeaderStyle Width="10%"></HeaderStyle>
                                                                                                <ItemStyle HorizontalAlign="Center" />
                                                                                                <ItemTemplate>
                                                                                                    <asp:Button ID="btnDel6" runat="server" Text="刪除" CausesValidation="False" CommandName="DEL6" CssClass="asp_button_M"></asp:Button>
                                                                                                    <asp:Button ID="btnEdt6" runat="server" Text="修改" CausesValidation="False" CommandName="EDT6" CssClass="asp_button_M"></asp:Button>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:Button ID="btnUpd6" runat="server" Text="更新" CausesValidation="False" CommandName="UPD6" CssClass="asp_button_M"></asp:Button>
                                                                                                    <asp:Button ID="btnCls6" runat="server" Text="取消" CausesValidation="False" CommandName="CLS6" CssClass="asp_button_M"></asp:Button>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                        </Columns>
                                                                                    </asp:DataGrid>
                                                                                </td>
                                                                            </tr>
                                                                            <tr id="trlabTotal6" runat="server">
                                                                                <td>
                                                                                    <table class="font" border="0" cellspacing="1" cellpadding="1">
                                                                                        <tr>
                                                                                            <td class="font" colspan="2">一人份材料費用合計： </td>
                                                                                            <td class="font" colspan="6">
                                                                                                <asp:Label ID="labTotal6" runat="server"></asp:Label></td>
                                                                                        </tr>
                                                                                    </table>
                                                                                </td>
                                                                            </tr>
                                                                        </table>
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                            <table id="TableCost7" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                                                <tr>
                                                                    <td class="bluecol" colspan="8">共同材料明細 </td>
                                                                </tr>
                                                                <tr>
                                                                    <td colspan="8" class="whitecol">
                                                                        <asp:Label ID="Label8" runat="server" CssClass="font">匯入明細</asp:Label>
                                                                        <input id="File2" type="file" size="50" name="File2" runat="server" accept=".csv,.xls" />
                                                                        <asp:Button ID="BtnImport2" runat="server" CausesValidation="False" Text="匯入明細" Width="88px" CssClass="asp_button_L" />(必須為csv格式)
                                                                        <asp:HyperLink ID="HyperLink2" runat="server" CssClass="font" ForeColor="#8080FF" NavigateUrl="../../Doc/CommonCost_Imp.zip">下載匯入格式檔</asp:HyperLink>
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td class="bluecol">項次 </td>
                                                                    <td class="bluecol">品名 </td>
                                                                    <td class="bluecol">規格 </td>
                                                                    <td class="bluecol">單位 </td>
                                                                    <td class="bluecol">單價 </td>
                                                                    <td class="bluecol">使用數量 </td>
                                                                    <td class="bluecol">用途說明 </td>
                                                                    <td class="bluecol">功能 </td>
                                                                </tr>
                                                                <tr>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="tItemNo7" runat="server" Columns="3" MaxLength="5"></asp:TextBox></td>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="tCName7" runat="server" Columns="10" MaxLength="30"></asp:TextBox></td>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="tStandard7" runat="server" Columns="10" MaxLength="300"></asp:TextBox></td>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="tUnit7" runat="server" Columns="3" MaxLength="5"></asp:TextBox></td>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="tPrice7" runat="server" Columns="5" MaxLength="5"></asp:TextBox></td>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="tAllCount7" runat="server" Columns="5" MaxLength="5"></asp:TextBox></td>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="tPurPose7" runat="server" Columns="35" MaxLength="300"></asp:TextBox></td>
                                                                    <td class="whitecol" align="center">
                                                                        <asp:Button ID="btnAddCost7" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_M"></asp:Button></td>
                                                                </tr>
                                                                <tr>
                                                                    <td colspan="8">
                                                                        <table id="DataGrid7Table" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                                                            <tr>
                                                                                <td class="whitecol">
                                                                                    <asp:DataGrid ID="DataGrid7" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                                                                        <Columns>
                                                                                            <asp:TemplateColumn HeaderText="項次">
                                                                                                <HeaderStyle Width="6%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lItemNo7" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eItemNo7" runat="server" Columns="5" MaxLength="5"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="品名">
                                                                                                <HeaderStyle Width="14%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lCName7" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eCName7" runat="server" Columns="10" MaxLength="30"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="規格">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lStandard7" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eStandard7" runat="server" Columns="10" MaxLength="300"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="單位">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lUnit7" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eUnit7" runat="server" Columns="3" MaxLength="5"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="單價">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lPrice7" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="ePrice7" runat="server" Columns="3" MaxLength="5"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="使用數量">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lAllCount7" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eAllCount7" runat="server" Columns="3" MaxLength="5"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="訓練人數">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lTNum7" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eTNum7" runat="server" Columns="3"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="小計">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lsubtotal7" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="esubtotal7" runat="server" Columns="3"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="每人分攤費用">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="leachCost7" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eeachCost7" runat="server" Columns="3"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="用途說明">
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lPurPose7" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="ePurPose7" runat="server" Columns="15" MaxLength="300"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="功能">
                                                                                                <HeaderStyle Width="10%"></HeaderStyle>
                                                                                                <ItemStyle HorizontalAlign="Center" />
                                                                                                <ItemTemplate>
                                                                                                    <asp:Button ID="btnDel7" runat="server" Text="刪除" CausesValidation="False" CommandName="DEL7" CssClass="asp_button_S"></asp:Button>
                                                                                                    <asp:Button ID="btnEdt7" runat="server" Text="修改" CausesValidation="False" CommandName="EDT7" CssClass="asp_button_S"></asp:Button>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:Button ID="btnUpd7" runat="server" Text="更新" CausesValidation="False" CommandName="UPD7" CssClass="asp_button_S"></asp:Button>
                                                                                                    <asp:Button ID="btnCls7" runat="server" Text="取消" CausesValidation="False" CommandName="CLS7" CssClass="asp_button_S"></asp:Button>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                        </Columns>
                                                                                    </asp:DataGrid>
                                                                                </td>
                                                                            </tr>
                                                                            <tr id="trlabTotal7" runat="server">
                                                                                <td>
                                                                                    <table class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                                                                                        <tr>
                                                                                            <td colspan="2" width="20%">共用材料費用合計： </td>
                                                                                            <td colspan="6" width="80%">
                                                                                                <asp:Label ID="labTotal7" runat="server"></asp:Label></td>
                                                                                        </tr>
                                                                                    </table>
                                                                                </td>
                                                                            </tr>
                                                                        </table>
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                            <table id="TableCostAll" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                                                <tr>
                                                                    <td class="font" colspan="8">
                                                                        <table class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                                                                            <tr>
                                                                                <td colspan="2" width="120">材料費用總計： </td>
                                                                                <td colspan="6" width="80%">
                                                                                    <asp:Label ID="labTotal67" runat="server"></asp:Label></td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td colspan="8" width="100%">
                                                                                    <asp:Label Style="z-index: 0" ID="labmemo1" runat="server" ForeColor="Red">&nbsp;註：每一項材料單品均需詳列品牌、規格(例如：型號、容量、長度...等)。</asp:Label>
                                                                                    <br />
                                                                                    <asp:Label Style="z-index: 0" ID="labmemo2" runat="server" ForeColor="Red">&nbsp;註：每人分攤費用之計算結果均四捨五入至整數位。</asp:Label>
                                                                                </td>
                                                                            </tr>
                                                                        </table>
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                            <table id="TableCost8" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                                                <tr>
                                                                    <td class="table_title" colspan="8">教材費用 </td>
                                                                </tr>
                                                                <tr>
                                                                    <td colspan="8" class="whitecol">
                                                                        <asp:Label ID="Labeld8" runat="server" CssClass="font">匯入明細</asp:Label>
                                                                        <input id="File3" type="file" size="50" name="File3" runat="server" accept=".csv,.xls" />
                                                                        <asp:Button ID="BtnImport8" runat="server" CausesValidation="False" Text="匯入明細" Width="88px" CssClass="asp_button_M" />(必須為csv格式)
                                                                        <asp:HyperLink ID="HyperLink3" runat="server" CssClass="font" ForeColor="#8080FF" NavigateUrl="../../Doc/SheetCost_Imp.zip">下載匯入格式檔</asp:HyperLink>
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td class="bluecol">項次 </td>
                                                                    <td class="bluecol">品名 </td>
                                                                    <td class="bluecol">規格 </td>
                                                                    <td class="bluecol">單位 </td>
                                                                    <td class="bluecol">單價 </td>
                                                                    <td class="bluecol">使用數量 </td>
                                                                    <td class="bluecol">用途說明 </td>
                                                                    <td class="bluecol">功能 </td>
                                                                </tr>
                                                                <tr>
                                                                    <td colspan="8" class="whitecol">
                                                                        <asp:Label ID="Labexptitle1" runat="server" Text="填寫範例" ForeColor="Red"></asp:Label></td>
                                                                </tr>
                                                                <tr>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="exptItemNo8" runat="server" Columns="5" MaxLength="5" ReadOnly="True" ForeColor="#CCCCCC">1</asp:TextBox></td>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="exptCName8" runat="server" Columns="10" MaxLength="30" ReadOnly="True" ForeColor="#CCCCCC">書籍/講義</asp:TextBox></td>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="exptStandards8" runat="server" Columns="10" MaxLength="300" ReadOnly="True" ForeColor="#CCCCCC">書名出版社/講義</asp:TextBox></td>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="exptUnit8" runat="server" Columns="5" MaxLength="5" ReadOnly="True" ForeColor="#CCCCCC">本/冊</asp:TextBox></td>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="exptPrice8" runat="server" Columns="5" MaxLength="5" ReadOnly="True" ForeColor="#CCCCCC">200</asp:TextBox></td>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="exptAllCount8" runat="server" Columns="5" MaxLength="5" ReadOnly="True" ForeColor="#CCCCCC">30</asp:TextBox></td>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="exptPurPose8" runat="server" Columns="35" MaxLength="300" ReadOnly="True" ForeColor="#CCCCCC">學科教學使用</asp:TextBox></td>
                                                                    <td class="whitecol"></td>
                                                                </tr>
                                                                <tr>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="tItemNo8" runat="server" Columns="5" MaxLength="5"></asp:TextBox></td>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="tCName8" runat="server" Columns="10" MaxLength="30"></asp:TextBox></td>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="tStandards8" runat="server" Columns="10" MaxLength="300"></asp:TextBox></td>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="tUnit8" runat="server" Columns="5" MaxLength="5"></asp:TextBox></td>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="tPrice8" runat="server" Columns="5" MaxLength="5"></asp:TextBox></td>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="tAllCount8" runat="server" Columns="5" MaxLength="5"></asp:TextBox></td>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="tPurPose8" runat="server" Columns="35" MaxLength="300"></asp:TextBox></td>
                                                                    <td class="whitecol" align="center">
                                                                        <asp:Button ID="btnAddCost8" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_M"></asp:Button></td>
                                                                </tr>
                                                                <tr>
                                                                    <td colspan="8">
                                                                        <table id="DataGrid8Table" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                                                            <tr>
                                                                                <td>
                                                                                    <asp:DataGrid ID="DataGrid8" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                                                                        <Columns>
                                                                                            <asp:TemplateColumn HeaderText="項次">
                                                                                                <HeaderStyle Width="6%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lItemNo8" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eItemNo8" runat="server" Columns="5" MaxLength="5"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="品名">
                                                                                                <HeaderStyle Width="14%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lCName8" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eCName8" runat="server" Columns="10" MaxLength="30"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="規格">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lStandards8" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eStandards8" runat="server" Columns="10" MaxLength="300"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="單位">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lUnit8" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eUnit8" runat="server" Columns="3" MaxLength="5"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="單價">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lPrice8" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="ePrice8" runat="server" Columns="3" MaxLength="5"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="使用數量">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lAllCount8" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eAllCount8" runat="server" Columns="3" MaxLength="5"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="訓練人數">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lTNum8" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eTNum8" runat="server" Columns="3"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="小計">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lsubtotal8" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="esubtotal8" runat="server" Columns="3"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="每人分攤費用">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="leachCost8" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eeachCost8" runat="server" Columns="3"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="用途說明">
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lPurPose8" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="ePurPose8" runat="server" Columns="15" MaxLength="300"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="功能">
                                                                                                <ItemStyle Width="10%" HorizontalAlign="Center"></ItemStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Button ID="btnDel8" runat="server" Text="刪除" CausesValidation="False" CommandName="DEL8" CssClass="asp_button_M"></asp:Button>
                                                                                                    <asp:Button ID="btnEdt8" runat="server" Text="修改" CausesValidation="False" CommandName="EDT8" CssClass="asp_button_M"></asp:Button>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:Button ID="btnUpd8" runat="server" Text="更新" CausesValidation="False" CommandName="UPD8" CssClass="asp_button_M"></asp:Button>
                                                                                                    <asp:Button ID="btnCls8" runat="server" Text="取消" CausesValidation="False" CommandName="CLS8" CssClass="asp_button_M"></asp:Button>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                        </Columns>
                                                                                    </asp:DataGrid>
                                                                                </td>
                                                                            </tr>
                                                                            <tr id="trlabTotal8" runat="server">
                                                                                <td>
                                                                                    <table class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                                                                                        <tr>
                                                                                            <td colspan="2" width="20%">教料費用合計： </td>
                                                                                            <td class="font" colspan="6" width="80%">
                                                                                                <asp:Label ID="labTotal8" runat="server"></asp:Label></td>
                                                                                        </tr>
                                                                                    </table>
                                                                                </td>
                                                                            </tr>
                                                                        </table>
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                            <table id="TableCost9" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                                                <tr>
                                                                    <td class="table_title" colspan="8">其他費用 </td>
                                                                </tr>
                                                                <tr>
                                                                    <td colspan="8" class="whitecol">
                                                                        <asp:Label ID="Labeld9" runat="server" CssClass="font" ForeColor="#8080FF">匯入明細</asp:Label>
                                                                        <input id="File4" type="file" size="50" name="File4" runat="server" accept=".csv,.xls" />
                                                                        <asp:Button ID="BtnImport9" runat="server" CausesValidation="False" Text="匯入明細" Width="88px" CssClass="asp_button_M" />(必須為csv格式)
                                                                        <asp:HyperLink ID="HyperLink4" runat="server" CssClass="font" ForeColor="#8080FF" NavigateUrl="../../Doc/OtherCost_Imp.zip">下載匯入格式檔</asp:HyperLink>
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td class="bluecol">項次 </td>
                                                                    <td class="bluecol">項目 </td>
                                                                    <td class="bluecol">規格 </td>
                                                                    <td class="bluecol">單位 </td>
                                                                    <td class="bluecol">單價 </td>
                                                                    <td class="bluecol">使用數量 </td>
                                                                    <td class="bluecol">用途說明 </td>
                                                                    <td class="bluecol">功能 </td>
                                                                </tr>
                                                                <tr>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="tItemNo9" runat="server" Columns="5" MaxLength="5"></asp:TextBox></td>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="tCName9" runat="server" Columns="10" MaxLength="30"></asp:TextBox></td>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="tStandards9" runat="server" Columns="10" MaxLength="300"></asp:TextBox></td>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="tUnit9" runat="server" Columns="5" MaxLength="5"></asp:TextBox></td>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="tPrice9" runat="server" Columns="5" MaxLength="5"></asp:TextBox></td>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="tAllCount9" runat="server" Columns="5" MaxLength="5"></asp:TextBox></td>
                                                                    <td class="whitecol">
                                                                        <asp:TextBox ID="tPurpose9" runat="server" Columns="35" MaxLength="300"></asp:TextBox></td>
                                                                    <td class="whitecol" align="center">
                                                                        <asp:Button ID="btnAddCost9" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_M"></asp:Button></td>
                                                                </tr>
                                                                <tr>
                                                                    <td colspan="8">
                                                                        <table id="DataGrid9Table" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                                                            <tr>
                                                                                <td>
                                                                                    <asp:DataGrid ID="DataGrid9" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                                                                        <Columns>
                                                                                            <asp:TemplateColumn HeaderText="項次">
                                                                                                <HeaderStyle Width="6%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lItemNo9" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eItemNo9" runat="server" Columns="5" MaxLength="5"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="項目">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lCName9" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eCName9" runat="server" Columns="10" MaxLength="30"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="規格">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lStandards9" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eStandards9" runat="server" Columns="10" MaxLength="300"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="單位">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lUnit9" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eUnit9" runat="server" Columns="3" MaxLength="5"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="單價">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lPrice9" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="ePrice9" runat="server" Columns="3" MaxLength="5"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="使用數量">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lAllCount9" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eAllCount9" runat="server" Columns="3" MaxLength="5"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="訓練人數">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lTNum9" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eTNum9" runat="server" Columns="3"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="小計">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lsubtotal9" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="esubtotal9" runat="server" Columns="3"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="每人分攤費用">
                                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="leachCost9" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="eeachCost9" runat="server" Columns="3"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="用途說明">
                                                                                                <ItemTemplate>
                                                                                                    <asp:Label ID="lPurPose9" runat="server"></asp:Label>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:TextBox ID="ePurPose9" runat="server" Columns="15" MaxLength="300"></asp:TextBox>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                            <asp:TemplateColumn HeaderText="功能">
                                                                                                <ItemStyle Width="10%" HorizontalAlign="Center"></ItemStyle>
                                                                                                <ItemTemplate>
                                                                                                    <asp:Button ID="btnDel9" runat="server" Text="刪除" CausesValidation="False" CommandName="DEL9" CssClass="asp_button_M"></asp:Button>
                                                                                                    <asp:Button ID="btnEdt9" runat="server" Text="修改" CausesValidation="False" CommandName="EDT9" CssClass="asp_button_M"></asp:Button>
                                                                                                </ItemTemplate>
                                                                                                <EditItemTemplate>
                                                                                                    <asp:Button ID="btnUpd9" runat="server" Text="更新" CausesValidation="False" CommandName="UPD9" CssClass="asp_button_M"></asp:Button>
                                                                                                    <asp:Button ID="btnCls9" runat="server" Text="取消" CausesValidation="False" CommandName="CLS9" CssClass="asp_button_M"></asp:Button>
                                                                                                </EditItemTemplate>
                                                                                            </asp:TemplateColumn>
                                                                                        </Columns>
                                                                                    </asp:DataGrid>
                                                                                </td>
                                                                            </tr>
                                                                            <tr id="trlabTotal9" runat="server">
                                                                                <td>
                                                                                    <table class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                                                                                        <tr>
                                                                                            <td colspan="2" width="20%">其他費用合計： </td>
                                                                                            <td class="font" colspan="6" width="80%">
                                                                                                <asp:Label ID="labTotal9" runat="server"></asp:Label></td>
                                                                                        </tr>
                                                                                    </table>
                                                                                </td>
                                                                            </tr>
                                                                        </table>
                                                                    </td>
                                                                </tr>
                                                            </table>

                                                            <table id="TableNote2" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                                                <tr>
                                                                    <td>
                                                                        <table id="tbtNote2" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                                                            <tr>
                                                                                <td class="bluecol">其他說明 </td>
                                                                                <td colspan="8" class="whitecol">
                                                                                    <asp:TextBox ID="tNote2" runat="server" Columns="88" Rows="5" TextMode="MultiLine" ToolTip="其他說明(欄位字數為1000)"></asp:TextBox>
                                                                                    <asp:Button Style="z-index: 0" ID="btnUptNote2" runat="server" CausesValidation="False" Text="修改" CssClass="asp_button_M"></asp:Button>
                                                                                </td>
                                                                            </tr>
                                                                        </table>
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                        </table>
                        <table id="TableLay7" border="0" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol_need" width="20%">經費來源 </td>
                                <td class="whitecol" width="80%">
                                    <table id="Table2" border="0" cellspacing="1" cellpadding="1" width="100%" class="whitecol">
                                        <tr>
                                            <td>政府補助金額：新臺幣(每期費用)
                                                <asp:TextBox ID="DefGovCost" runat="server" onfocus="this.blur()" Columns="7">0</asp:TextBox>元/每班人數
										        <asp:Label ID="TNum1" runat="server"></asp:Label>=
										        <asp:Label ID="Total1" runat="server"></asp:Label>元
										        <asp:RegularExpressionValidator ID="check18" runat="server" ControlToValidate="DefGovCost" Display="None" ErrorMessage="經費來源「政府負擔補助金額」請輸入數字" ValidationExpression="[0-9]{0,8}"></asp:RegularExpressionValidator>
                                            </td>
                                        </tr>
                                        <tr style="display: none">
                                            <td>企業負擔金額：新臺幣(每期費用)
										        <asp:TextBox ID="DefUnitCost" runat="server" onfocus="this.blur()" Columns="7">0</asp:TextBox>元/每班人數
										        <asp:Label ID="TNum2" runat="server"></asp:Label>=
										        <asp:Label ID="Total2" runat="server"></asp:Label>元
										        <asp:RegularExpressionValidator ID="check20" runat="server" ControlToValidate="DefUnitCost" Display="None" ErrorMessage="經費來源「民間企業度負擔」請輸入數字" ValidationExpression="[0-9]{0,8}"></asp:RegularExpressionValidator>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>學員負擔金額：新臺幣(每期費用)
										        <asp:TextBox ID="DefStdCost" runat="server" onfocus="this.blur()" Columns="7">0</asp:TextBox>元/每班人數
										        <asp:Label ID="TNum3" runat="server"></asp:Label>=
										        <asp:Label ID="Total3" runat="server"></asp:Label>元
										        <asp:RegularExpressionValidator ID="check22" runat="server" ControlToValidate="DefStdCost" Display="None" ErrorMessage="經費來源「學員負擔金額」請輸入數字" ValidationExpression="[0-9]{0,8}"></asp:RegularExpressionValidator>
                                                <asp:CustomValidator ID="CustomValidator1" runat="server" Display="None" ErrorMessage="請輸入經費來源" ClientValidationFunction="CheckDef"></asp:CustomValidator>
                                                <asp:CustomValidator ID="Customvalidator7" runat="server" Display="None" ErrorMessage="經費來源-每人的政府補助金額與學員負擔金額應為整數數字" ClientValidationFunction="CheckDef2"></asp:CustomValidator>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                        </table>
                        <table id="TableLay8" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol_need" width="20%">經費分類代碼 </td>
                                <td class="whitecol" width="80%">
                                    <asp:TextBox ID="GCIDName" runat="server" onfocus="this.blur()" Columns="70"></asp:TextBox>
                                    <input id="btn_GCID" onclick="Get_GovClass('GCIDName');" value="..." type="button" name="btn_GCID" runat="server" class="button_b_Mini" />
                                    <input style="width: 32px; height: 22px" id="GCIDValue" type="hidden" name="GCIDValue" runat="server" />
                                    <input style="width: 32px; height: 22px" id="GCID1Value" type="hidden" name="GCID1Value" runat="server" />
                                    <asp:RequiredFieldValidator ID="GCID_chk1" runat="server" ControlToValidate="GCIDName" Display="None" ErrorMessage="訓練費用編列說明：請選擇經費來源「經費分類代碼」"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">訓練費用<br />
                                    &nbsp;&nbsp; 編列說明</td>
                                <td class="whitecol" width="80%">
                                    <table>
                                        <tr>
                                            <td>
                                                <asp:TextBox ID="Note" runat="server" Columns="60" Rows="8" TextMode="MultiLine"></asp:TextBox></td>
                                            <td>
                                                <br />
                                                <asp:Button ID="Button21b" runat="server" Text="匯出EXCEL" CausesValidation="False" CssClass="asp_Export_M"></asp:Button></td>
                                        </tr>
                                    </table>
                                    <font color="red">
                                        <asp:Label ID="Labmsg3" runat="server"></asp:Label></font>
                                </td>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
            <tr>
                <td>
                    <div align="center" class="whitecol">
                        <asp:Button ID="Button8" runat="server" Text="草稿儲存" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="btnAdd" runat="server" Text="正式儲存" CssClass="asp_button_M"></asp:Button>
                        <asp:Label ID="Labsave" runat="server" CssClass="font" ForeColor="Red" ToolTip="必免使用者重複按下儲存鈕消除">儲存鈕恢復中請稍後...</asp:Label>
                        <asp:Button ID="Button24" runat="server" Text="回上一頁" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                        <asp:ValidationSummary ID="ValidationSummary1" runat="server" DisplayMode="List" ShowSummary="False" ShowMessageBox="True"></asp:ValidationSummary>
                    </div>
                </td>
            </tr>
        </table>
        <input id="LayerState" type="hidden" runat="server" />
        <input style="width: 64px; height: 22px; color: #ff0000; font-weight: bold" id="time_spent" onfocus="this.blur()" maxlength="256" size="5" type="hidden" name="time_spent" runat="server" />
        <input id="isBlack" type="hidden" name="isBlack" runat="server" />
        <input id="Blackorgname" type="hidden" name="Blackorgname" runat="server" />
        <input id="hTPlanID54" type="hidden" name="hTPlanID54" runat="server" />
        <input id="upt_PlanX" name="upt_PlanX" type="hidden" runat="server" />
        <asp:HiddenField ID="Hid_ComIDNO" runat="server" />
        <asp:HiddenField ID="Hid_PSNO28" runat="server" />
    </form>
</body>
</html>
