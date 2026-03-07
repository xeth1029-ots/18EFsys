<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_03_001.aspx.vb" Inherits="WDAIIP.TC_03_001" %>

<html>
<head>
    <title>班級申請作業</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link rel="stylesheet" type="text/css" href="../../css/style.css" />
    <script type="text/javascript" src="../../js/OpenWin/openwin.js"></script>
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript" src="../../js/autocomplete.js"></script>
    <%--專長能力標籤--%>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function Layer_change(index) {
            //先關閉所有選項
            var xLayerState = document.getElementById('LayerState');
            for (i = 1; i <= 10; i++) {
                document.getElementById('TableLay' + i).style.display = 'none';
            }
            for (i = 1; i <= 10; i++) {
                if (document.getElementById('box' + i)) {
                    document.getElementById('box' + i).className = "";
                }
            }
            //index = document.getElementById('LayerState').value;
            if (index == '') { if (xLayerState) { index = xLayerState.value; } }

            document.getElementById('SelectAll').checked = false;
            document.getElementById('TableLay' + index).style.display = '';
            document.getElementById('box' + index).className = "active";

            if (xLayerState) { xLayerState.value = index; }
            //if (!_isIE) { parent.setMainFrameHeight(); } //重新調整頁面高度，by:20180815
            //if (parent && parent.setMainFrameHeight != undefined) { parent.setMainFrameHeight(); }
            //if (document.body) { window.scroll(0, document.body.scrollHeight); }
            //if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
            show_COACHING1();
            //其他//re-binding autocomplete input event
            if (index == "8") { autocomplete_init(); }
            //if (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) { _isIE = true; }
            if (document.body) { window.scroll(0, document.body.scrollHeight); }
            if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        }

        function OpenAllLayer(Flag) {
            for (i = 1; i <= 10; i++) {
                if (Flag) {
                    if (document.getElementById('box' + i)) {
                        document.getElementById('TableLay' + i).style.display = '';
                        document.getElementById('box' + i).className = "";
                    }
                }
                else {
                    if (document.getElementById('box' + i)) {
                        document.getElementById('TableLay' + i).style.display = 'none';
                        document.getElementById('box' + i).className = "";
                    }
                }
            }
            if (!Flag) {
                var i = document.getElementById('LayerState').value;
                document.getElementById('TableLay' + i).style.display = '';
                document.getElementById('box' + i).className = "";
            }
            //其他//re-binding autocomplete input event
            autocomplete_init();

            //if (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) { _isIE = true; }
            if (document.body) { window.scroll(0, document.body.scrollHeight); }
            if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        }

        //檢查訓練內容簡介
        function CheckDescData(obj1, obj2, obj3) {
            var msg = '';
            if (document.getElementById(obj1).value == '') msg += '請輸入單元名稱\n';
            if (document.getElementById(obj2).value == '') msg += '請輸入時數\n';
            else if (!isUnsignedInt(document.getElementById(obj2).value)) msg += '時數必須為數字\n';
            if (document.getElementById(obj3).value == '') msg += '請輸入課程大綱\n';
            if (msg != '') {
                blockAlert(msg);//alert(msg);
                return false;
            }
        }

        //經費輸入判斷
        function check_Cost1() {
            var msg = '';
            if (document.form1.CostID.selectedIndex == 0) msg += '請選擇經費總類\n';
            else if (document.form1.CostID.value == '99' && document.form1.ItemOther.value == '') msg += '請輸入項目名稱\n';
            msg += check_Cost_Detail(document.form1.OPrice, document.form1.Itemage, document.form1.ItemCost);
            if (document.form1.CostID.value == '98' && document.getElementById('AdmGrantTR')) msg += '已經設定行政管理費，不能選擇此項目\n';
            if (msg != '') {
                blockAlert(msg);//alert(msg);
                return false;
            }
            else {
                //檢查是否有重複的費用
                var flag = false;
                var MyTable = document.getElementById('DataGrid1');
                if (document.form1.CostID.value != '99' && MyTable) {
                    for (var i = 1; i < MyTable.rows.length; i++) {
                        if (document.form1.CostID.value == MyTable.rows(i).cells(6).innerHTML) {
                            flag = true;
                        }
                    }
                }
                if (flag) {
                    return confirm('您已經輸入過該費用，確定還要新增?');
                }
            }
        }

        function check_Cost2() {
            var msg = '';
            msg += check_Cost_Detail(document.form1.OPrice2, document.form1.Itemage2, document.form1.ItemCost2);
            if (msg != '') {
                blockAlert(msg);//alert(msg);
                return false;
            }
        }

        function check_Cost3() {
            var msg = '';
            msg += check_Cost_Detail(document.form1.OPrice3, document.form1.Itemage3, null);
            if (msg != '') {
                blockAlert(msg);//alert(msg);
                return false;
            }
        }

        function check_Cost4() {
            var msg = '';
            if (document.form1.CostID4.selectedIndex == 0) msg += '請選擇經費總類\n';
            else if (document.form1.CostID4.value == '99' && document.form1.ItemOther4.value == '') msg += '請輸入項目名稱\n';
            msg += check_Cost_Detail(document.form1.OPrice4, document.form1.Itemage4, null);
            if (document.form1.hidden17.value != '17') {
                if (document.form1.CostID4.value == '98' && document.getElementById('AdmTR4')) msg += '已經設定行政管理費，不能選擇此項目\n';
            }
            if (msg != '') {
                blockAlert(msg);//alert(msg);
                return false;
            }
            else {
                //檢查是否有重複的費用
                var flag = false;
                var MyTable = document.getElementById('DataGrid4');
                if (document.form1.CostID4.value != '99' && MyTable) {
                    for (var i = 1; i < MyTable.rows.length; i++) {
                        if (document.form1.CostID4.value == MyTable.rows(i).cells(5).innerHTML) {
                            flag = true;
                        }
                    }
                }
                if (flag) {
                    return confirm('您已經輸入過該費用，確定還要新增?');
                }
            }
        }

        function check_Cost_Detail(OPrice, Itemage, ItemCost) {
            var msg = '';
            if (OPrice.value == '') { msg += '請輸入單價\n'; }
            else {
                if (!isUnsignedInt(OPrice.value)) {
                    if (!isPositiveFloat(OPrice.value)) { msg += '單價必須為數字\n'; }
                    else {
                        if (OPrice.value.indexOf('.') < OPrice.value.length - 3) {
                            msg += '單價只能輸入到小數點第二位\n';
                        }
                    }
                }
            }
            if (Itemage.value == '') { msg += '請輸入數量\n'; }
            else if (!isUnsignedInt(Itemage.value)) msg += '數量必須為數字\n';
            if (ItemCost != null) {
                if (ItemCost.value == '') { msg += '請輸入計價單位\n'; }
                else if (!isUnsignedInt(ItemCost.value)) msg += '計價單位必須為數字\n';
            }
            return msg;
        }

        //如果選擇其他費用，則顯示TextBox讓使用者輸入
        function ShowOther(ObjName, ObjName2) {
            var Obj1 = document.getElementById(ObjName);
            var Obj2 = document.getElementById(ObjName2);
            Obj2.style.display = 'none';
            if (Obj1.value == '99') { Obj2.style.display = ''; } //'inline';
        }

        //課程學科計算
        function set_SciHours() {
            var Sum_SciHours = 0;
            var Sum_TotalHours = 0;
            if (isUnsignedInt(form1.GenSciHours.value)) {
                Sum_SciHours += parseFloat(form1.GenSciHours.value, 10);
                Sum_TotalHours += parseFloat(form1.GenSciHours.value, 10);
            }
            if (isUnsignedInt(form1.ProSciHours.value)) {
                Sum_SciHours += parseInt(form1.ProSciHours.value, 10);
                Sum_TotalHours += parseInt(form1.ProSciHours.value, 10);
            }
            form1.SciHours.value = Sum_SciHours;
            // 計算課程總時數
            if (isUnsignedInt(form1.ProTechHours.value)) {
                Sum_TotalHours += parseInt(form1.ProTechHours.value, 10);
            }
            if (isUnsignedInt(form1.OtherHours.value)) {
                Sum_TotalHours += parseInt(form1.OtherHours.value, 10);
            }
            form1.TotalHours.value = Sum_TotalHours;
            form1.THours.value = Sum_TotalHours;
            form1.THours.title = ' 與課程編配的總時數一致';
        }

        // 計算課程編配的總時數不能大於課程內容的時數
        function Check_totalhours(source, args) {
            var flag = true;
            if (parseInt(form1.TotalHours.value, 10) > parseInt(form1.HPHour.value, 10)) {
                flag = false;
            }
            args.IsValid = flag;
        }

        //2005/01/17 Ellen判斷班別資料訓練時數是否大於總時數
        function set_THours() {
            var msg = "";
            if (form1.THours.value != '') {
                if (parseInt(form1.THours.value, 10) > parseInt(form1.TotalHours.value, 10)) {
                    form1.THours.value = form1.TotalHours.value;
                    msg = "班別資料訓練時數超過課程編配的總時數，請重新輸入";
                    blockAlert(msg);//alert("班別資料訓練時數超過課程編配的總時數，請重新輸入");
                    form1.THours.focus();
                }
            }
        }

        function check_CyclType(source, args) {
            if (!isUnsignedInt(args.Value)) args.IsValid = false;
            if (args.Value.length != 2) args.IsValid = false;
            if (parseInt(args.Value, 10) <= 0) args.IsValid = false;
        }

        //'檢查日期格式-Melody(2005/3/18)
        function check_date(source, args) {
            if (args.Value != '' && !checkDate(args.Value)) {
                args.IsValid = false;
            }
        }

        function open_hours() {
            window.open('TC_03_oper.aspx', '', 'width=1200,height=660,location=0,status=0,menubar=0,scrollbars=0,resizable=0');
        }

        function CheckTotalCost(source, args) {
            var DataGrid1Table = document.getElementById('DataGrid1Table');
            var DataGrid2Table = document.getElementById('DataGrid2Table');
            var DataGrid3Table = document.getElementById('DataGrid3Table');
            var DataGrid4Table = document.getElementById('DataGrid4Table');

            args.IsValid = false;
            if (DataGrid1Table.style.display == '' || DataGrid2Table.style.display == '' || DataGrid3Table.style.display == '' || DataGrid4Table.style.display == '') {
                args.IsValid = true;
            }
        }

        function CheckDef(source, args) {
            var flag = false;
            if (document.form1.DefGovCost.value != '') { flag = true; }
            if (document.form1.DefUnitCost.value != '') { flag = true; }
            if (document.form1.DefStdCost.value != '') { flag = true; }
            args.IsValid = flag;
        }

        function CheckTDate(sor, args) {
            args.IsValid = true;
            var STDate = document.getElementById('STDate');
            var FDDate = document.getElementById('FDDate');
            //計畫：接受企業委託訓練 可同一天'起迄日不可同一天
            var Hid_TPlanID = document.getElementById('Hid_TPlanID');
            var flag_NoSameDay1 = true;
            if (Hid_TPlanID && Hid_TPlanID.value == '07') { flag_NoSameDay1 = false; }
            if (checkDate(STDate.value) && checkDate(FDDate.value)) {
                var flag = compareDate(STDate.value, FDDate.value);
                if (flag == 0 && flag_NoSameDay1) args.IsValid = false;
                if (flag == 1) args.IsValid = false;
            }
        }

        function CheckTDate2(sor, args) {
            args.IsValid = true;
            var SEnterDate = document.getElementById('SEnterDate');
            var FEnterDate = document.getElementById('FEnterDate');
            //計畫：接受企業委託訓練 可同一天'起迄日不可同一天
            var Hid_TPlanID = document.getElementById('Hid_TPlanID');
            var flag_NoSameDay1 = true;
            if (Hid_TPlanID && Hid_TPlanID.value == '07') { flag_NoSameDay1 = false; }
            if (checkDate(SEnterDate.value) && checkDate(FEnterDate.value)) {
                var flag = compareDate(SEnterDate.value, FEnterDate.value);
                if (flag == 0 && flag_NoSameDay1) args.IsValid = false;
                if (flag == 1) args.IsValid = false;
            }
        }

        //草稿儲存檢查
        function Check_Temp() {
            var msg = '';
            var Hid_MaxTNum = document.getElementById('Hid_MaxTNum');
            var TNum = document.getElementById('TNum');
            //計畫：接受企業委託訓練 可同一天'起迄日不可同一天
            var Hid_TPlanID = document.getElementById('Hid_TPlanID');
            var flag_NoSameDay1 = true;
            if (Hid_TPlanID && Hid_TPlanID.value == '07') { flag_NoSameDay1 = false; }
            if (document.form1.center.value == '') msg += '請選擇訓練機構\n'
            if (document.form1.GenSciHours.value != '' && !isUnsignedInt(document.form1.GenSciHours.value)) msg += '一般學科必須為數字\n'
            if (document.form1.ProSciHours.value != '' && !isUnsignedInt(document.form1.ProSciHours.value)) msg += '專業學科必須為數字\n'
            if (document.form1.ProTechHours.value != '' && !isUnsignedInt(document.form1.ProTechHours.value)) msg += '術科必須為數字\n'
            if (document.form1.OtherHours.value != '' && !isUnsignedInt(document.form1.OtherHours.value)) msg += '其他時數必須為數字\n'
            if (document.form1.TNum.value != '' && !isUnsignedInt(document.form1.TNum.value)) msg += '訓練人數必須為數字\n'
            if (msg == '' && TNum.value != '' && Hid_MaxTNum.value != '') {
                if (parseInt(TNum.value, 10) > parseInt(Hid_MaxTNum.value, 10)) { msg += '訓練人數上限為' + Hid_MaxTNum.value + '人\n' }
            }
            if (document.form1.THours.value != '' && !isUnsignedInt(document.form1.THours.value)) msg += '訓練時數必須為數字\n'
            if (document.form1.STDate.value != '' && !checkDate(document.form1.STDate.value)) msg += '訓練起日不是正確的日期格式\n'
            if (document.form1.FDDate.value != '' && !checkDate(document.form1.FDDate.value)) msg += '訓練迄日不是正確的日期格式\n'
            if (checkDate(document.form1.STDate.value) && checkDate(document.form1.FDDate.value)) {
                var flag = compareDate(document.form1.STDate.value, document.form1.FDDate.value);
                if (flag == 0 && flag_NoSameDay1) msg += '訓練起日不能和訓練迄日同一天\n';
                if (flag == 1) msg += '訓練起日不能超過訓練迄日\n';
            }
            if (document.form1.CyclType.value != '' && !isUnsignedInt(document.form1.CyclType.value)) msg += '期別必須為數字\n'
            if (document.form1.ClassCount.value != '' && !isUnsignedInt(document.form1.ClassCount.value)) msg += '班數必須為數字\n'
            if (document.form1.DefGovCost.value != '' && !isUnsignedInt(document.form1.DefGovCost.value)) msg += '政府負擔費用必須為數字\n'
            if (document.form1.DefUnitCost.value != '' && !isUnsignedInt(document.form1.DefUnitCost.value)) msg += '企業負擔費用必須為數字\n'
            if (document.form1.DefStdCost.value != '' && !isUnsignedInt(document.form1.DefStdCost.value)) msg += '學員負擔費用必須為數字\n'
            if (msg != '') {
                blockAlert(msg);//alert(msg);
                return false;
            }
        }

        //計算經費來源
        function CountCostSource() {
            //旗標-看所有的值是否已經輸入
            var msg = '';
            var TNum = document.getElementById('TNum');
            var TNum1 = document.getElementById('TNum1');
            var TNum2 = document.getElementById('TNum2');
            var TNum3 = document.getElementById('TNum3');
            var Total1 = document.getElementById('Total1');
            var Total2 = document.getElementById('Total2');
            var Total3 = document.getElementById('Total3');
            var DefGovCost = document.getElementById('DefGovCost');
            var DefUnitCost = document.getElementById('DefUnitCost');
            var DefStdCost = document.getElementById('DefStdCost');

            if (TNum.value == '' || !isUnsignedInt(TNum.value)) {
                Total1.innerHTML = '0';
                Total2.innerHTML = '0';
                Total3.innerHTML = '0';

                TNum1.innerHTML = '(尚未設定人數)';
                TNum2.innerHTML = '(尚未設定人數)';
                TNum3.innerHTML = '(尚未設定人數)';
                return false;
            }
            var v_TNum = parseInt(TNum.value, 10);
            if (v_TNum <= 0) {
                msg += "請確認 訓練人數 是否有誤";
                blockAlert(msg);//alert(msg);
                return false;
            }

            TNum1.innerHTML = v_TNum;
            TNum2.innerHTML = v_TNum;
            TNum3.innerHTML = v_TNum;

            if (isUnsignedInt(DefGovCost.value)) {
                Total1.innerHTML = parseInt(DefGovCost.value, 10) * v_TNum;
                if (parseInt(DefGovCost.value, 10) > 100000) {
                    msg += '請確認政府補助金額是否有誤\n';
                }
            }
            if (isUnsignedInt(DefUnitCost.value)) {
                Total2.innerHTML = parseInt(DefUnitCost.value, 10) * v_TNum;
                if (parseInt(document.form1.DefUnitCost.value, 10) > 100000) {
                    msg += '請確認企業負擔金額是否有誤\n';
                }
            }

            if (isUnsignedInt(DefStdCost.value)) {
                Total3.innerHTML = parseInt(DefStdCost.value, 10) * v_TNum;
                if (parseInt(DefStdCost.value, 10) > 100000) {
                    msg += '請確認學員負擔金額是否有誤\n';
                }
            }

            //換算個人單價計價法
            /*
            if (document.getElementById('CostModeTR')) {
                //表示委外的情況
                if (document.form1.CostMode4.checked) {
                    //表示使用個人單價計價法
                }
            }
            */

            var Hid_cost_02_08 = document.getElementById('Hid_cost_02_08');
            var ACTHUMCOST = document.getElementById('ACTHUMCOST');
            var METCOSTPER = document.getElementById('METCOSTPER');
            var TotalHours = document.getElementById('TotalHours');
            var v_TotalHours = 0;
            if (isUnsignedInt(TotalHours.value)) { v_TotalHours = parseInt(TotalHours.value, 10) };
            var v_Hid_cost_02_08 = 0;
            if (isUnsignedInt(Hid_cost_02_08.value)) { v_Hid_cost_02_08 = parseInt(Hid_cost_02_08.value, 10) };
            var Hid_TPlanID = document.getElementById('Hid_TPlanID');
            var PerCost = document.getElementById('PerCost');
            var TotalCost4 = document.getElementById('TotalCost4');
            var hidTaxCost4 = document.getElementById('hidTaxCost4');
            var v_TotalCost4 = 0;
            if (TotalCost4) { v_TotalCost4 = parseInt(TotalCost4.innerHTML, 10); }
            var v_hidTaxCost4 = 0;
            if (hidTaxCost4) { v_hidTaxCost4 = parseInt(hidTaxCost4.value, 10); }
            PerCost.innerHTML = parseInt((v_TotalCost4 - v_hidTaxCost4) / v_TNum, 10);
            if (Hid_TPlanID && Hid_TPlanID.value == '70' && ACTHUMCOST && METCOSTPER && v_TotalHours > 0) {
                //總計/訓練時數/訓練人數
                //總計金額-TotalCost4/訓練時數-TotalHours/訓練人數-v_TNum
                ACTHUMCOST.innerHTML = toDecimal(toDecimal(toDecimal(v_TotalCost4) / v_TotalHours) / v_TNum);
                //parseInt(TotalHours.value , 10)
                //材料費小計/總計-TotalCost4
                METCOSTPER.innerHTML = toDecimal(toDecimal(v_Hid_cost_02_08) / v_TotalCost4);
            }

            if (msg != '') {
                blockAlert(msg);//alert(msg);
                return false;
            }
        }

        function toDecimal(x) {
            var f = parseFloat(x);
            if (isNaN(f)) {
                return;
            }
            f = Math.round(x * 100) / 100;
            return f;
        }

        //檢查個人單價計價中中是否有行政管理費的選項
        function checkAdm(num) {
            var msg = "";
            if (num == 1) {
                var mytable = document.getElementById('DataGrid1');
                if (mytable) {
                    for (var i = 1; i < mytable.rows.length; i++) {
                        if (mytable.rows(i).cells(6).innerHTML == '98') {
                            msg = "已經設定了行政管理費為項目!";
                            blockAlert(msg);//alert('已經設定了行政管理費為項目!');
                            return false;
                        }
                    }
                    return true;
                }
                else {
                    msg = "請先設定項目!";
                    blockAlert(msg);//alert('請先設定項目!');
                    return false;
                }
            }
            else if (num == 2) {
                var mytable = document.getElementById('DataGrid4');
                if (mytable) {
                    for (var i = 1; i < mytable.rows.length; i++) {
                        if (mytable.rows(i).cells(5).innerHTML == '98') {
                            msg = "已經設定了行政管理費為項目!";
                            blockAlert(msg);//alert('已經設定了行政管理費為項目!');
                            return false;
                        }
                    }
                    return true;
                }
                else {
                    msg = "請先設定項目!";
                    blockAlert(msg);//alert('請先設定項目!');
                    return false;
                }
            }
            else if (num == 3) {
                var mytable = document.getElementById('DataGrid1');
                if (mytable) {
                    //營業稅
                    return true;
                }
                else {
                    msg = "請先設定項目!";
                    blockAlert(msg);//alert('請先設定項目!');
                    return false;
                }
            }
            else if (num == 4) {
                var mytable = document.getElementById('DataGrid4');
                if (mytable) {
                    //營業稅
                    return true;
                }
                else {
                    msg = "請先設定項目!";
                    blockAlert(msg);//alert('請先設定項目!');
                    return false;
                }
            }
        }

        function check_style() {
            if (form1.STDate && form1.STDate.disabled && form1.date1) {
                form1.date1.style.cursor = "";
                form1.date1.onclick = null;
            }
            if (form1.FDDate && form1.FDDate.disabled && form1.date2) {
                form1.date2.style.cursor = "";
                form1.date2.onclick = null;
            }
            if (form1.SEnterDate && form1.SEnterDate.disabled && form1.imgdate1) {
                form1.imgdate1.style.cursor = "";
                form1.imgdate1.onclick = null;
            }
            if (form1.FEnterDate && form1.FEnterDate.disabled && form1.imgdate2) {
                form1.imgdate2.style.cursor = "";
                form1.imgdate2.onclick = null;
            }
            if (form1.ExamDate && form1.ExamDate.disabled && form1.imgdate3) {
                form1.imgdate3.style.cursor = "";
                form1.imgdate3.onclick = null;
            }
        }

        function CheckTrain3(sor, args) {
            if (parseInt(getCheckBoxListValue('GetTrain3'), 10) == 0)
                args.IsValid = false;
        }

        function CheckCapMilitary(sor, args) {
            if (parseInt(getCheckBoxListValue('CapMilitary'), 10) == 0)
                args.IsValid = false;
        }

        function CheckTrain3Other(sor, args) {
            if (getCheckBoxListValue('GetTrain3').charAt(getCheckBoxListValue('GetTrain3').length - 1) == '1' && document.getElementById('GetTrain3Other').value == '')
                args.IsValid = false;
        }

        function CheckTrain4(sor, args) {
            var ExamDate = document.getElementById('ExamDate');
            if (!ExamDate) { return; }
            if (ExamDate.value != '' && parseInt(getCheckBoxListValue('GetTrain4'), 10) == 0)
                args.IsValid = false;
        }

        function CheckTrain4Other(sor, args) {
            var ExamDate = document.getElementById('ExamDate');
            if (!ExamDate) { return; }
            if (ExamDate.value != '' && getCheckBoxListValue('GetTrain4').charAt(getCheckBoxListValue('GetTrain4').length - 1) == '1' && document.getElementById('GetTrain4Other').value == '')
                args.IsValid = false;
        }

        function CheckDesc(sor, args) {
            if (!document.getElementById('DataGrid5'))
                args.IsValid = false;
        }

        //限制TextBox在MultiLine時的字數
        function checkTextLength(obj, long) {
            var msg = "";
            var maxlength = new Number(long); // Change number to your max length.

            if (obj.value.length > maxlength) {
                obj.value = obj.value.substring(0, maxlength);
                msg = "限欄位長度不能大於" + maxlength + "個字元(含空白字元)，超出字元將自動截斷";
                blockAlert(msg);//alert("限欄位長度不能大於" + maxlength + "個字元(含空白字元)，超出字元將自動截斷");
            }
        }

        //twiACTNO
        function fCheckACTNO(source, args) {
            args.IsValid = true;
            if (!isBlank(args)) {
                if (args.IsValid == true) {
                    if (trim(args.Value).length < 2 || trim(args.Value).substr(0, 2) != '09') {
                        args.IsValid = false;  //訓字保保險證號為09開頭
                    }
                }
            }
        }

        //2009-05-20 add 依需求只允許輸入整數(排除 00)
        function fCheckZIPB3_1(source, args) {
            args.IsValid = true;
            if (typeof (trim(args.Value)) == "undefined") { return; }
            args.Value = trim(args.Value);
            if (isBlank(args)) { return; }
            var flag_2 = args.Value.length == 2 ? true : false;//長度2
            var flag_3 = args.Value.length == 3 ? true : false;//長度3
            if (isNaN(parseInt(args.Value, 10))) { args.IsValid = false; return; }
            if (!isUnsignedInt(args.Value)) { args.IsValid = false; return; }
            if (parseInt(args.Value, 10) < 1) { args.IsValid = false; return; }
            return;
        }

        //2009-05-20 add 依需求只允許輸入兩碼
        function fCheckZIPB3_2(source, args) {
            args.IsValid = true;
            if (typeof (trim(args.Value)) == "undefined") { return; }
            args.Value = trim(args.Value);
            if (isBlank(args)) { return; }
            var flag_2 = args.Value.length == 2 ? true : false;//長度2
            var flag_3 = args.Value.length == 3 ? true : false;//長度3
            if (!flag_2 && !flag_3) { args.IsValid = false; return; }
        }

        //20090521 add 依需求於郵遞區號後 2 碼進行 onchange時做格式驗證
        function CheckZIPB3_Event(thisArgs, txtName) {
            var msg = '';
            if (thisArgs.value != '') { thisArgs.value = trim(thisArgs.value); }
            if (isBlank(thisArgs)) { return true; }
            var flag_2 = thisArgs.value.length == 2 ? true : false;//長度2
            var flag_3 = thisArgs.value.length == 3 ? true : false;//長度3
            if (flag_2 || flag_3) {
                var flag_NG = false;
                //非數字//非整數//小於1
                if (msg == '' && isNaN(parseInt(thisArgs.value, 10))) { flag_NG = true; }
                if (msg == '' && !isUnsignedInt(parseInt(thisArgs.value, 10))) { flag_NG = true; }
                if (msg == '' && parseInt(thisArgs.value, 10) < 1) { flag_NG = true; }
                if (flag_NG) { msg += txtName + '「郵遞區號後2碼或後3碼」必須為數字，且不得輸入 0 \n'; }
            }
            else {
                msg += txtName + '「郵遞區號後2碼或後3碼」長度必須為 2碼或3碼\n';
            }
            if (msg != '') {
                blockAlert(msg);//alert(msg);
                return false;
            }
            return true;
        }

        //選擇全部，若有單選消除全部勾選
        function SelectAllcbkList(obj, hidobj) {
            var num = getCheckBoxListValue(obj).length;
            var myallcheck = document.getElementById(obj + '_' + 0);

            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
            else {
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    if ('0' == getCheckBoxListValue(obj).charAt(i)) {
                        document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(i);
                        myallcheck.checked = mycheck.checked;
                        break;
                    }
                }
            }
        }

        //COACHING / tr_COACHING1
        function show_COACHING1() {
            var RB_Y = document.getElementById("RBCOACHING_Y");
            var RB_N = document.getElementById("RBCOACHING_N");
            var myTR = document.getElementById("tr_COACHING1");
            myTR.style.display = 'none';
            if (RB_Y && RB_Y.checked) { myTR.style.display = ''; }

            //if (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) { _isIE = true; }
            if (document.body) { window.scroll(0, document.body.scrollHeight); }
            if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        }
    </script>
</head>
<body onload="check_style();">
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;班級申請作業</asp:Label>
                </td>
            </tr>
        </table>
        <table border="0" cellspacing="0" cellpadding="0" width="100%">
            <tr>
                <td>
                    <table class="table_sch" cellpadding="1" cellspacing="1">
                        <tr>
                            <td class="bluecol_need" width="20%">訓練機構 </td>
                            <td class="whitecol" colspan="3" width="30%">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="70%"></asp:TextBox>
                                <input id="Org" type="button" value="..." name="Org" runat="server" class="button_b_Mini">
                                <asp:Button ID="Button28" Style="display: none" runat="server" CausesValidation="False" Text="機構資訊(隱藏)"></asp:Button>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                <input id="ComidValue" type="hidden" name="ComidValue" runat="server"><br>
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                                <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ErrorMessage="請選擇訓練機構" Display="None" ControlToValidate="center"></asp:RequiredFieldValidator>
                            </td>
                            <td class="whitecol" colspan="4" width="50%">
                                <table id="Table1_Email" runat="server" cellpadding="1" cellspacing="1" width="100%">
                                    <tr>
                                        <td>
                                            <font class="font" size="2">是否要Email線上報名資料，EMail</font>
                                            <br />
                                            <asp:TextBox ID="EMail" runat="server" Columns="30" Width="70%"></asp:TextBox>
                                            <input id="hidden17" runat="server" type="hidden" name="hidden17" />
                                            <asp:RegularExpressionValidator ID="check1" runat="server" ErrorMessage="E_Mail輸入錯誤" Display="None" ControlToValidate="EMail" ValidationExpression="\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"></asp:RegularExpressionValidator>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">
                                <asp:Label ID="LabTMID" runat="server">訓練職類</asp:Label></td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Columns="30" Width="70%"></asp:TextBox>
                                <input id="btu_sel" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." runat="server" class="button_b_Mini">
                                <asp:RequiredFieldValidator ID="fill1" runat="server" ErrorMessage="請選擇訓練職類" Display="None" ControlToValidate="TB_career_id"></asp:RequiredFieldValidator>
                                <input id="trainValue" type="hidden" name="trainValue" runat="server" />
                            </td>
                            <td class="bluecol_need" colspan="1">
                                <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label></td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="70%"></asp:TextBox>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="button_b_Mini">
                                <asp:RequiredFieldValidator ID="fill1b" runat="server" ErrorMessage="請選擇通俗職類" Display="None" ControlToValidate="txtCJOB_NAME"></asp:RequiredFieldValidator>
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
                    <table id="tb_title" class="font" cellspacing="0" cellpadding="0" width="90%" runat="server">
                        <tr class="newlink">
                            <td id="box1" onclick="Layer_change(1);">目標</td>
                            <td id="box2" onclick="Layer_change(2);">
                                <asp:Label ID="lab_LayerC2" runat="server" Text="受訓資格"></asp:Label></td>
                            <td id="box9" onclick="Layer_change(9);" runat="server">錄訓方式</td>
                            <td id="box10" onclick="Layer_change(10);" runat="server">訓練內容簡介</td>
                            <td id="box3" onclick="Layer_change(3);">訓練方式</td>
                            <td id="box4" onclick="Layer_change(4);">課程編配</td>
                            <td id="box5" onclick="Layer_change(5);">班別資料</td>
                            <td id="box6" onclick="Layer_change(6);" runat="server">訓練費用</td>
                            <td id="box7" onclick="Layer_change(7);">經費來源</td>
                            <td id="box8" onclick="Layer_change(8);">其他</td>
                            <td align="center">
                                <input id="SelectAll" onclick="OpenAllLayer(this.checked);" type="checkbox">展開全部</td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <div style="overflow-y: auto;">
                        <table id="TableLay1" width="100%">
                            <tr>
                                <td class="bluecol_need" width="20%">緣由 </td>
                                <td class="whitecol" width="80%">(應說明辦理本訓練班次的因由及規劃屬性)<br>
                                    <asp:TextBox ID="PlanCause" runat="server" Width="77%" Rows="4" TextMode="MultiLine" MaxLength="100"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="fill30" runat="server" ErrorMessage="目標『緣由』為必填欄位" Display="None" ControlToValidate="PlanCause"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need"><%--訓練目標--%>
                                    <asp:Label ID="Title1" runat="server"></asp:Label></td>
                                <td class="whitecol">
                                    <asp:TextBox ID="PurScience" runat="server" Width="77%" Rows="4" TextMode="MultiLine" MaxLength="100"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="fill2" runat="server" ErrorMessage="目標「學科」為必填欄位" Display="None" ControlToValidate="PurScience"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need"><%--就業展望--%>
                                    <asp:Label ID="Title2" runat="server"></asp:Label></td>
                                <td class="whitecol">
                                    <asp:TextBox ID="PurTech" runat="server" Width="77%" Rows="4" TextMode="MultiLine" MaxLength="100"></asp:TextBox><asp:RequiredFieldValidator ID="fill3" runat="server" ErrorMessage="目標「技能」為必填欄位" Display="None" ControlToValidate="PurTech"></asp:RequiredFieldValidator></td>
                            </tr>
                            <tr id="TR_2005_01" runat="server">
                                <td class="bluecol_need">品德 </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="PurMoral" runat="server" Width="77%" Rows="4" TextMode="MultiLine" MaxLength="100"></asp:TextBox><asp:RequiredFieldValidator ID="fill4" runat="server" ErrorMessage="目標「品德」為必填欄位" Display="None" ControlToValidate="PurMoral"></asp:RequiredFieldValidator></td>
                            </tr>
                            <tr>
                                <td colspan="2" width="100%">
                                    <hr />
                                </td>
                            </tr>
                        </table>
                        <table id="TableLay2" width="100%">
                            <tr>
                                <td colspan="2" class="table_title" width="100%">(應說明參加本項訓練應具有之身分及相觀背景條件) </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">學歷 </td>
                                <td class="whitecol" width="80%">
                                    <asp:RadioButtonList ID="Degree" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" CellPadding="0" CellSpacing="0" RepeatColumns="6"></asp:RadioButtonList>
                                    <asp:RequiredFieldValidator ID="fill5" runat="server" ErrorMessage="請選擇受訓資格「學歷」" Display="None" ControlToValidate="Degree"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="12%">年齡 </td>
                                <td class="whitecol" width="88%">
                                    <asp:RadioButton ID="rdoAge1" runat="server" Checked="True" GroupName="GroupAge" />
                                    <asp:Label ID="l_Age" runat="server">年滿15歲以上</asp:Label>
                                    <asp:RadioButton ID="rdoAge2" runat="server" GroupName="GroupAge" />
                                    <asp:Label ID="l_Age2a" runat="server">有上限，年滿15歲~</asp:Label>
                                    <asp:TextBox ID="txtAge2" runat="server" Width="10%" MaxLength="2"></asp:TextBox>
                                    <asp:Label ID="l_Age2b" runat="server">歲</asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="12%">兵役 </td>
                                <td class="whitecol" width="88%">
                                    <asp:CheckBoxList ID="CapMilitary" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                        <asp:ListItem Value="00">不限</asp:ListItem>
                                        <asp:ListItem Value="04">在役</asp:ListItem>
                                        <asp:ListItem Value="0103">役畢(含免役)</asp:ListItem>
                                        <asp:ListItem Value="02">未役</asp:ListItem>
                                    </asp:CheckBoxList>
                                    <asp:Label ID="LabCapMilitary" runat="server" Visible="False">不限</asp:Label>
                                </td>
                            </tr>
                            <tr id="trTRNUNITNAME" runat="server">
                                <td class="bluecol_need" width="12%">委訓單位名稱 </td>
                                <td class="whitecol" width="88%">
                                    <asp:TextBox ID="TRNUNITNAME" runat="server" Width="80%" MaxLength="100"></asp:TextBox></td>
                            </tr>
                            <tr id="trTRNUNITTYPE" runat="server">
                                <td class="bluecol_need" width="12%">委訓單位類型 </td>
                                <td class="whitecol" width="88%">
                                    <%--政府機關、公民營事業機構、學校、團體、其他(請說明) --%>
                                    <asp:RadioButtonList ID="TRNUNITCHO" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" CellPadding="0" CellSpacing="0" RepeatColumns="6"></asp:RadioButtonList>
                                    <%--<asp:ListItem Value="1">政府機關</asp:ListItem>
                                        <asp:ListItem Value="2">公民營事業機構</asp:ListItem>
                                        <asp:ListItem Value="3">學校</asp:ListItem>
                                        <asp:ListItem Value="4">團體</asp:ListItem>
                                        <asp:ListItem Value="9">其他(請說明)</asp:ListItem>--%>
                                    <asp:TextBox ID="TRNUNITTYPE" runat="server" Width="26%" MaxLength="100"></asp:TextBox></td>
                            </tr>
                            <tr id="trTRNUNITEE" runat="server">
                                <td class="bluecol" width="12%">訓練對象 </td>
                                <td class="whitecol" width="88%">
                                    <asp:TextBox ID="TRNUNITEE" runat="server" Width="80%" MaxLength="100"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="12%">其他一 </td>
                                <td class="whitecol" width="88%">
                                    <asp:TextBox ID="Other1" runat="server" Width="80%" MaxLength="100"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="12%">其他二 </td>
                                <td class="whitecol" width="88%">
                                    <asp:TextBox ID="Other2" runat="server" Width="80%" MaxLength="100"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="12%">其他三 </td>
                                <td class="whitecol" width="88%">
                                    <asp:TextBox ID="Other3" runat="server" Width="80%" MaxLength="100"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td colspan="2" width="100%">
                                    <hr />
                                </td>
                            </tr>
                        </table>
                        <table id="TableLay9" width="100%">
                            <tr>
                                <td colspan="2" class="table_title" width="100%">(應說明本班次針對持推介單者採一律錄訓或優先錄訓規劃) </td>
                            </tr>
                            <tr runat="server" id="tr_GetTrain1">
                                <td class="bluecol_need" width="20%">持推介單報參訓之適用條件 </td>
                                <td class="whitecol" width="80%">
                                    <asp:RadioButtonList ID="GetTrain1" runat="server" CssClass="font" CellPadding="0" CellSpacing="0">
                                        <asp:ListItem Value="1">一律錄訓  (無須參加甄選考試，持推介單者一律錄訓)</asp:ListItem>
                                        <asp:ListItem Value="2">甄選錄訓  (詳招生簡章或其他說明)</asp:ListItem>
                                        <asp:ListItem Value="3">不適用推介機制</asp:ListItem>
                                    </asp:RadioButtonList>
                                    <asp:RequiredFieldValidator ID="fill23" runat="server" ErrorMessage="錄訓方式「持推介單報參訓之適用條件」為必填欄位" Display="None" ControlToValidate="GetTrain1"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">自行報名參訓者錄訓規定 </td>
                                <td class="whitecol" width="80%">
                                    <asp:TextBox ID="GetTrain2" runat="server" Width="77%" Rows="5" TextMode="MultiLine" MaxLength="200"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="fill24" runat="server" ErrorMessage="錄訓方式「自行報名參訓者錄訓規定」為必填欄位" Display="None" ControlToValidate="GetTrain2"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <%--td_GetTrain3-intertype--%>
                                <td runat="server" id="td_GetTrain3" class="bluecol_need" width="20%">甄試方式<br />
                                </td>
                                <td class="whitecol" width="80%">
                                    <asp:Label ID="lab_msg_GetTrain3" runat="server"></asp:Label><br />
                                    <asp:CheckBoxList ID="GetTrain3" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" AppendDataBoundItems="true"></asp:CheckBoxList>
                                    <asp:TextBox ID="GetTrain3Other" runat="server" Width="28%" MaxLength="1200"></asp:TextBox>
                                    <asp:CustomValidator ID="fill25" runat="server" ErrorMessage="錄訓方式「甄試方式」是必填勾選" Display="None" ClientValidationFunction="CheckTrain3"></asp:CustomValidator>
                                    <asp:CustomValidator ID="fill26" runat="server" ErrorMessage="錄訓方式「甄試方式」必須填寫「其他」內容" Display="None" ClientValidationFunction="CheckTrain3Other"></asp:CustomValidator>
                                    <%--<asp:Label ID="Label9" runat="server" Display="None" Visible="false"> (無辦理甄試之班級， 免填此欄位)</asp:Label>--%>
                                </td>
                            </tr>
                            <tr>
                                <td id="td_GetTrain4" runat="server" class="bluecol_need" width="20%">其他 </td>
                                <td class="whitecol" width="80%">
                                    <asp:Label ID="lab_msg_GetTrain4" runat="server"></asp:Label>
                                    <br />
                                    <asp:CheckBoxList ID="GetTrain4" runat="server" Width="100%" RepeatLayout="Flow" CssClass="font">
                                        <asp:ListItem Value="1">上述甄試方式及主要內容逕洽承訓單位。</asp:ListItem>
                                        <asp:ListItem Value="2">上述甄試方式及主要內容概述如下：(建議 50 字內)</asp:ListItem>
                                    </asp:CheckBoxList>
                                    <br />
                                    <br />
                                    <asp:TextBox ID="GetTrain4Other" runat="server" Width="77%" Rows="5" TextMode="MultiLine" MaxLength="200"></asp:TextBox>
                                    <asp:CustomValidator ID="fill27" runat="server" ErrorMessage="錄訓方式「其他」必須勾選" Display="None" ClientValidationFunction="CheckTrain4"></asp:CustomValidator><asp:CustomValidator ID="fill31" runat="server" ErrorMessage="錄訓方式「其他」必須填寫主要概述內容" Display="None" ClientValidationFunction="CheckTrain4Other"></asp:CustomValidator>
                                    <%--<asp:Label ID="Label1" runat="server" Display="None" Visible="false"> (無辦理甄試之班級， 免填此欄位)</asp:Label>--%>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2" width="100%">
                                    <hr />
                                </td>
                            </tr>
                        </table>
                        <table class="font" id="TableLay10" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <tr>
                                <td>
                                    <table class="font" id="Table7" width="100%">
                                        <tr>
                                            <td colspan="4" class="whitecol" width="100%">
                                                <asp:Label ID="Label7" runat="server" CssClass="font" ForeColor="#8080FF">匯入訓練內容簡介</asp:Label>
                                                <input id="File1" type="file" size="55" name="File1" runat="server" accept=".csv,.xls" />
                                                <asp:Button ID="Btn_TrainDescImport" runat="server" CausesValidation="False" Text="匯入簡介" CssClass="asp_button_M"></asp:Button>(必須為csv格式)
                                                <asp:HyperLink ID="HyperLink1" runat="server" CssClass="font" ForeColor="#8080FF" NavigateUrl="../../Doc/TrainDesc_Import.zip">下載整批上載格式檔</asp:HyperLink>
                                                <asp:HyperLink ID="HyperLink2" runat="server" CssClass="font" ForeColor="#8080FF" NavigateUrl="../../Doc/TrainDesc_Imp2.zip">下載整批上載格式檔</asp:HyperLink>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="table_title" width="25%">單元名稱 </td>
                                            <td class="table_title" width="15%">時數 </td>
                                            <td class="table_title" width="45%">課程大綱 </td>
                                            <td class="table_title" width="15%">&nbsp; </td>
                                        </tr>
                                        <tr>
                                            <td class="whitecol" width="25%">
                                                <asp:TextBox ID="PName" runat="server" MaxLength="50" Width="90%"></asp:TextBox></td>
                                            <td class="whitecol" width="15%">
                                                <asp:TextBox ID="PHour" runat="server" MaxLength="3" Width="50%"></asp:TextBox>
                                                <input id="HPHour" type="hidden" runat="server">
                                            </td>
                                            <td class="whitecol" width="45%">
                                                <asp:TextBox ID="PCont" runat="server" MaxLength="550" Width="90%" TextMode="MultiLine"></asp:TextBox></td>
                                            <td class="whitecol" width="15%" align="center">
                                                <asp:Button ID="Button29" runat="server" CausesValidation="False" Text="新增" CssClass="asp_button_M"></asp:Button></td>
                                        </tr>
                                        <tr>
                                            <td colspan="4" class="bluecol">應說明課程大綱(含時數)及各單元實施內容 </td>
                                        </tr>
                                    </table>
                                    <asp:CustomValidator ID="fill32" runat="server" ErrorMessage="訓練內容簡介為必填欄位" Display="None" ClientValidationFunction="CheckDesc"></asp:CustomValidator>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <table id="DataGrid5Table" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                        <tr>
                                            <td>
                                                <asp:DataGrid ID="DataGrid5" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                                    <Columns>
                                                        <asp:TemplateColumn HeaderText="課程單元">
                                                            <HeaderStyle Width="20%" />
                                                            <ItemTemplate>
                                                                <asp:Label ID="LPName" runat="server"></asp:Label>
                                                            </ItemTemplate>
                                                            <EditItemTemplate>
                                                                <asp:TextBox ID="TPName" runat="server" MaxLength="50" Width="150px"></asp:TextBox>
                                                            </EditItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="時數">
                                                            <HeaderStyle Width="10%" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                            <ItemTemplate>
                                                                <asp:Label ID="LPHour" runat="server"></asp:Label>
                                                            </ItemTemplate>
                                                            <EditItemTemplate>
                                                                <asp:TextBox ID="TPHour" runat="server" MaxLength="3" Width="50px"></asp:TextBox>
                                                            </EditItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="實施內容">
                                                            <HeaderStyle Width="60%" />
                                                            <ItemTemplate>
                                                                <asp:Label ID="LPCont" runat="server"></asp:Label>
                                                            </ItemTemplate>
                                                            <EditItemTemplate>
                                                                <asp:TextBox ID="TPCont" runat="server" MaxLength="550" Width="450px" TextMode="MultiLine"></asp:TextBox>
                                                            </EditItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="功能">
                                                            <HeaderStyle Width="10%" />
                                                            <ItemStyle HorizontalAlign="Center" />
                                                            <ItemTemplate>
                                                                <asp:Button ID="Button30" runat="server" Text="修改" CausesValidation="False" CommandName="edit" CssClass="asp_button_M"></asp:Button>
                                                                <asp:Button ID="Button31" runat="server" Text="刪除" CausesValidation="False" CommandName="del" CssClass="asp_button_M"></asp:Button>
                                                            </ItemTemplate>
                                                            <EditItemTemplate>
                                                                <asp:Button ID="Button32" runat="server" Text="儲存" CausesValidation="False" CommandName="save" CssClass="asp_button_M"></asp:Button>
                                                                <asp:Button ID="Button33" runat="server" Text="取消" CausesValidation="False" CommandName="cancel" CssClass="asp_button_M"></asp:Button>
                                                            </EditItemTemplate>
                                                        </asp:TemplateColumn>
                                                    </Columns>
                                                </asp:DataGrid>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <hr />
                                </td>
                            </tr>
                        </table>
                        <table id="TableLay3" cellspacing="1" cellpadding="1" style="width: 100%">
                            <tr>
                                <td class="bluecol_need" style="width: 20%">學科 </td>
                                <td class="whitecol" style="width: 80%">
                                    <asp:TextBox ID="TMScience" onblur="checkTextLength(this,200)" onkeyup="checkTextLength(this,200)" runat="server" Width="77%" Rows="5" TextMode="MultiLine" MaxLength="200" onChange="checkTextLength(this,200)"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="fill13" runat="server" ErrorMessage="訓練方式「學術」為必填欄位" Display="None" ControlToValidate="TMScience"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" style="width: 20%">術科 </td>
                                <td class="whitecol" style="width: 80%">
                                    <asp:TextBox ID="TMTech" onblur="checkTextLength(this,200)" onkeyup="checkTextLength(this,200)" runat="server" Width="77%" Rows="5" TextMode="MultiLine" MaxLength="200" onChange="checkTextLength(this,200)"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="fill14" runat="server" ErrorMessage="訓練方式「術科」為必填欄位" Display="None" ControlToValidate="TMTech"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2" style="width: 100%">
                                    <hr />
                                </td>
                            </tr>
                        </table>
                        <table id="TableLay4" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <tr>
                                <td rowspan="2" class="bluecol_need" width="20%">學科 </td>
                                <td rowspan="2" class="whitecol" width="30%">
                                    <asp:TextBox ID="SciHours" runat="server" onfocus="this.blur()" Width="25%"></asp:TextBox>小時 </td>
                                <td class="bluecol" width="20%">1. 一般學科 </td>
                                <td class="whitecol" width="30%">
                                    <asp:TextBox ID="GenSciHours" runat="server" Width="25%"></asp:TextBox>小時
                                    <asp:RequiredFieldValidator ID="fill15" runat="server" ErrorMessage="課程編配「一般學科」為必填欄位" Display="None" ControlToValidate="GenSciHours"></asp:RequiredFieldValidator>
                                    <asp:RegularExpressionValidator ID="check4" runat="server" ErrorMessage="課程編配「一般學科」請輸入數字" Display="None" ControlToValidate="GenSciHours" ValidationExpression="[0-9]{1,4}"></asp:RegularExpressionValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">2. 專業學科 </td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="ProSciHours" runat="server" Width="25%"></asp:TextBox>小時
                                    <asp:RequiredFieldValidator ID="fill16" runat="server" ErrorMessage="課程編配「專業學科」為必填欄位" Display="None" ControlToValidate="ProSciHours"></asp:RequiredFieldValidator><asp:RegularExpressionValidator ID="check5" runat="server" ErrorMessage="課程編配「專業學科」請輸入數字" Display="None" ControlToValidate="ProSciHours" ValidationExpression="[0-9]{1,4}"></asp:RegularExpressionValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">術科 </td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="ProTechHours" runat="server" Width="9%"></asp:TextBox>小時
                                    <asp:RequiredFieldValidator ID="fill17" runat="server" ErrorMessage="課程編配「術科」為必填欄位" Display="None" ControlToValidate="ProTechHours"></asp:RequiredFieldValidator>
                                    <asp:RegularExpressionValidator ID="check6" runat="server" ErrorMessage="課程編配「術科」請輸入數字" Display="None" ControlToValidate="ProTechHours" ValidationExpression="[0-9]{1,4}"></asp:RegularExpressionValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">其他時數 </td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="OtherHours" runat="server" Width="9%"></asp:TextBox>小時
                                    <asp:RegularExpressionValidator ID="check7" runat="server" ErrorMessage="課程編配「其他時數」請輸入數字" Display="None" ControlToValidate="OtherHours" ValidationExpression="[0-9]{1,4}"></asp:RegularExpressionValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">總計 </td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="TotalHours" runat="server" onfocus="this.blur()" Width="9%" MaxLength="3"></asp:TextBox>小時 </td>
                            </tr>
                            <tr>
                                <td colspan="4" width="100%">
                                    <hr />
                                </td>
                            </tr>
                        </table>
                        <table id="TableLay5" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <tr>
                                <td class="bluecol_need" style="width: 20%">班別名稱 </td>
                                <td colspan="3" class="whitecol" style="width: 80%">
                                    <asp:TextBox ID="ClassName" runat="server" Width="60%"></asp:TextBox>
                                    <input id="Class_Unit" type="hidden" runat="server"><asp:RequiredFieldValidator ID="fill18" runat="server" ErrorMessage="班別資料「班別名稱」為必填欄位" Display="None" ControlToValidate="ClassName"></asp:RequiredFieldValidator>
                                    <input id="Button26" type="button" value="產生班級名稱" name="Button26" runat="server" class="asp_button_M">
                                    <input id="Button34" onclick="open_hours()" type="button" value="時數迄日換算" name="Button34" runat="server" class="asp_button_M">
                                </td>
                            </tr>
                            <tr>
                                <td id="td_ClassEngName" runat="server" class="bluecol_need">班級英文名稱</td>
                                <td colspan="3" class="whitecol">
                                    <asp:TextBox ID="ClassEngName" runat="server" Columns="50" Width="70%"></asp:TextBox>
                                    <%--<asp:RequiredFieldValidator ID="Class_EName" runat="server" Display="None" ErrorMessage="請輸入班級英文名稱" ControlToValidate="ClassEngName"></asp:RequiredFieldValidator>--%>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" style="width: 20%">期別(二碼) </td>
                                <td class="whitecol" style="width: 30%">
                                    <asp:TextBox ID="CyclType" runat="server" Columns="5" MaxLength="2" Width="20%"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="fill28" runat="server" ErrorMessage="班別資料『期別』為必填欄位" Display="None" ControlToValidate="CyclType"></asp:RequiredFieldValidator>
                                    <asp:CustomValidator ID="CustomValidator4" runat="server" ErrorMessage="班別資料『期別』必須為大於0的兩位數字" Display="None" ControlToValidate="CyclType" ClientValidationFunction="check_CyclType"></asp:CustomValidator>
                                </td>
                                <td class="bluecol_need" style="width: 20%">班數 </td>
                                <td class="whitecol" style="width: 30%">
                                    <asp:TextBox ID="ClassCount" runat="server" onfocus="this.blur()" Columns="5" Text="1" Width="20%"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="fill29" runat="server" ErrorMessage="班別資料「班數」為必填欄位" Display="None" ControlToValidate="ClassCount"></asp:RequiredFieldValidator>
                                    <asp:RegularExpressionValidator ID="RegularExpressionValidator1" runat="server" ErrorMessage="班數輸入大於0的數字" Display="None" ControlToValidate="ClassCount" ValidationExpression="^0*[1-9](\d*$)"></asp:RegularExpressionValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">訓練課程類型
                                </td>
                                <td class="whitecol">
                                    <asp:RadioButtonList ID="rblADVANCE" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="01">基礎</asp:ListItem>
                                        <asp:ListItem Value="02">進階</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                                <td class="whitecol"></td>
                                <td class="whitecol"></td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" style="width: 20%">訓練人數 </td>
                                <td class="whitecol" style="width: 30%">
                                    <asp:TextBox ID="TNum" runat="server" Width="30%"></asp:TextBox>人
                                    <asp:RequiredFieldValidator ID="fill19" runat="server" ErrorMessage="班別資料「訓練人數」為必填欄位" Display="None" ControlToValidate="TNum"></asp:RequiredFieldValidator>
                                    <asp:RegularExpressionValidator ID="check8" runat="server" ErrorMessage="班別資料「訓練人數」請輸入數字" Display="None" ControlToValidate="TNum" ValidationExpression="[0-9]{1,4}"></asp:RegularExpressionValidator>
                                </td>
                                <td class="bluecol_need" style="width: 20%">訓練時數 </td>
                                <td class="whitecol" style="width: 30%">
                                    <asp:TextBox ID="THours" runat="server" Width="30%" onfocus="this.blur()"></asp:TextBox>小時
                                    <asp:RequiredFieldValidator ID="fill20" runat="server" ErrorMessage="班別資料「訓練時數」為必填欄位" Display="None" ControlToValidate="THours"></asp:RequiredFieldValidator>
                                    <asp:RegularExpressionValidator ID="check9" runat="server" ErrorMessage="班別資料「訓練時數」請輸入數字" Display="None" ControlToValidate="THours" ValidationExpression="[0-9]{1,4}"></asp:RegularExpressionValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" style="width: 20%">訓練起日 </td>
                                <td class="whitecol" style="width: 30%">
                                    <asp:TextBox ID="STDate" runat="server" Columns="12" MaxLength="11"></asp:TextBox>
                                    <span runat="server">
                                        <img id="date1" style="cursor: pointer" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30"></span>
                                    <asp:RequiredFieldValidator ID="fill21" runat="server" ErrorMessage="班別資料「訓練起日」為必填欄位" Display="None" ControlToValidate="STDate"></asp:RequiredFieldValidator>
                                    <asp:CustomValidator ID="CustomValidator2" runat="server" ErrorMessage="班別資料「訓練起日」不是正確的日期格式" Display="None" ControlToValidate="STDate" ClientValidationFunction="check_date"></asp:CustomValidator>
                                </td>
                                <td class="bluecol_need" style="width: 20%">訓練迄日 </td>
                                <td class="whitecol" style="width: 30%">
                                    <asp:TextBox ID="FDDate" runat="server" Columns="12" MaxLength="11"></asp:TextBox>
                                    <span runat="server">
                                        <img id="date2" style="cursor: pointer" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30"></span>
                                    <asp:RequiredFieldValidator ID="fill22" runat="server" ErrorMessage="班別資料「訓練迄日」為必填欄位" Display="None" ControlToValidate="FDDate"></asp:RequiredFieldValidator>
                                    <asp:CustomValidator ID="CustomValidator3" runat="server" ErrorMessage="班別資料「訓練迄日」不是正確的日期格式" Display="None" ControlToValidate="FDDate" ClientValidationFunction="check_date"></asp:CustomValidator>
                                    <asp:CustomValidator ID="CustomValidator5" runat="server" ErrorMessage="班別資料 訓練起日不能比訓練迄日晚(或者同一天)" Display="None" ClientValidationFunction="CheckTDate"></asp:CustomValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">訓練時段</td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="TPeriodList" runat="server"></asp:DropDownList>
                                    <asp:RequiredFieldValidator ID="Re_TPeriod_List" runat="server" Display="None" ErrorMessage="請選擇訓練時段" ControlToValidate="TPeriodList"></asp:RequiredFieldValidator>
                                    <span id="trTB_NOTE3" runat="server">
                                        <br />
                                        填寫方式:(每/隔) 週(一~日)00:00~23:59<br />
                                        <asp:TextBox ID="TB_NOTE3" runat="server" Columns="25" TextMode="MultiLine" Rows="5" Width="90%"></asp:TextBox>
                                    </span>
                                </td>
                                <td class="bluecol_need">訓練期限</td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="TDeadline_List" runat="server"></asp:DropDownList>
                                    <asp:RequiredFieldValidator ID="Re_TDeadline_List" runat="server" Display="None" ErrorMessage="[班別資料]請選擇訓練期限" ControlToValidate="TDeadline_List"></asp:RequiredFieldValidator>
                                </td>
                            </tr>

                            <tr>
                                <td class="bluecol_need" style="width: 20%">上課地址 </td>
                                <td colspan="3" class="whitecol" style="width: 80%">
                                    <input id="TAddressZip" name="TAddressZip" runat="server" maxlength="3" />－
                                    <input id="TAddressZIPB3" maxlength="3" name="TAddressZIPB3" runat="server" />
                                    <input id="hidTAddressZIP6W" type="hidden" runat="server" />
                                    <asp:Literal ID="LitTAddressZip" runat="server"></asp:Literal><%--郵遞區號查詢--%>
                                    <br />
                                    <asp:TextBox ID="CCTName" runat="server" onfocus="this.blur()" Style="width: 30%;"></asp:TextBox>
                                    <%--<input id="CCTName" name="CCTName" runat="server" onfocus="this.blur()" style="width: 30%;" />--%>
                                    <input id="BtnTAddressZip" type="button" value="..." name="BtnTAddressZip" runat="server" class="button_b_Mini" />
                                    <asp:TextBox ID="TAddress" runat="server" Columns="40" MaxLength="150" Width="40%"></asp:TextBox><br />
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ErrorMessage="班別資料「鄉鎮縣市」為必填欄位" Display="None" ControlToValidate="CCTName"></asp:RequiredFieldValidator>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" ErrorMessage="班別資料「訓練地點」為必填欄位" Display="None" ControlToValidate="TAddress"></asp:RequiredFieldValidator>
                                    <asp:RequiredFieldValidator ID="RequiredFieldValidator10" runat="server" ErrorMessage="班別資料「郵遞區號5碼之後2碼」為必填欄位" Display="None" ControlToValidate="TAddressZIPB3"></asp:RequiredFieldValidator>
                                    <asp:CustomValidator ID="CheckZIPB3_1" ErrorMessage="班別資料c「郵遞區號5碼之後2碼」必須為數字，且不得輸入00" Display="None" ControlToValidate="TAddressZIPB3" ClientValidationFunction="fCheckZIPB3_1" runat="server"></asp:CustomValidator>
                                    <asp:CustomValidator ID="CheckZIPB3_2" ErrorMessage="班別資料c「郵遞區號5碼之後2碼」長度必須為 2 碼(例 01 或 99)" Display="None" ControlToValidate="TAddressZIPB3" ClientValidationFunction="fCheckZIPB3_2" runat="server"></asp:CustomValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" style="width: 20%">聯絡人姓名 </td>
                                <td class="whitecol" style="width: 30%">
                                    <asp:TextBox ID="ContactName" runat="server" MaxLength="50" Width="90%"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="chkContactName1" runat="server" ControlToValidate="ContactName" Display="None" ErrorMessage="聯絡人姓名為必填欄位"></asp:RequiredFieldValidator>
                                </td>
                                <td class="bluecol_need" style="width: 20%">聯絡人電話 </td>
                                <td class="whitecol" style="width: 30%">
                                    <asp:TextBox ID="ContactPhone" runat="server" MaxLength="50" Width="60%"></asp:TextBox>
                                    <asp:RequiredFieldValidator ID="chkContactPhone1" runat="server" ControlToValidate="ContactPhone" Display="None" ErrorMessage="聯絡人電話為必填欄位"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" style="width: 20%">聯絡人<br>
                                    電子郵件 </td>
                                <td colspan="3" class="whitecol" style="width: 80%">
                                    <asp:TextBox ID="ContactEmail" runat="server" MaxLength="64" Columns="32" Width="60%"></asp:TextBox>
                                    <asp:RegularExpressionValidator ID="chkContactEmail1" runat="server" ControlToValidate="ContactEmail" Display="None" ErrorMessage="聯絡人電子郵件輸入錯誤" ValidationExpression="\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"></asp:RegularExpressionValidator>
                                    <asp:RequiredFieldValidator ID="chkContactEmail2" runat="server" ControlToValidate="ContactEmail" Display="None" ErrorMessage="請輸入聯絡人電子郵件"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" style="width: 20%">直屬主管<br>
                                    電子郵件 </td>
                                <td class="whitecol" style="width: 30%">
                                    <asp:TextBox ID="MasterEmail" runat="server" MaxLength="64" Columns="32" Width="90%"></asp:TextBox><br />
                                    <asp:RegularExpressionValidator ID="chkMasterEmail1" runat="server" ControlToValidate="MasterEmail" Display="None" ErrorMessage="直屬主管電子郵件輸入錯誤" ValidationExpression="\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"></asp:RegularExpressionValidator>
                                    <asp:RequiredFieldValidator ID="chkMasterEmail2" runat="server" ControlToValidate="MasterEmail" Display="None" ErrorMessage="請輸入直屬主管電子郵件"></asp:RequiredFieldValidator>
                                </td>
                                <td class="bluecol_need" style="width: 20%">訓字保保險證號 </td>
                                <td class="whitecol" style="width: 30%">
                                    <asp:TextBox ID="twiACTNO" runat="server" MaxLength="18" Width="60%"></asp:TextBox>
                                    <asp:CustomValidator ID="CustomValidatortwiACTNO" ErrorMessage="訓字保保險證號為09開頭" Display="None" ControlToValidate="twiACTNO" ClientValidationFunction="fCheckACTNO" runat="server"></asp:CustomValidator>
                                    <asp:RequiredFieldValidator ID="rfvtwiACTNO" runat="server" ControlToValidate="twiACTNO" Display="None" ErrorMessage="請輸入訓字保保險證號"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">報名地點 </td>
                                <td colspan="3" class="whitecol" width="80%">
                                    <input id="EAddressZip" name="EAddressZip" runat="server" maxlength="3">－
                                    <input id="EAddressZIPB3" maxlength="3" name="EAddressZIPB3" runat="server">
                                    <input id="hidEAddressZIP6W" type="hidden" runat="server" />
                                    <asp:Literal ID="LitEAddressZip" runat="server"></asp:Literal><%--郵遞區號查詢--%>
                                    <br />
                                    <asp:TextBox ID="ECTName" runat="server" onfocus="this.blur()" Style="width: 30%;"></asp:TextBox>
                                    <%--<input id="ECTName" name="ECTName" runat="server" onfocus="this.blur()" style="width: 30%;" />--%>
                                    <input id="BtnEAddressZip" type="button" value="..." runat="server" class="button_b_Mini" />
                                    <asp:TextBox ID="EAddress" runat="server" Columns="40" MaxLength="150" Width="40%"></asp:TextBox>
                                    <asp:CustomValidator ID="CheckZIPB3_19" ErrorMessage="班別資料c報名地點「郵遞區號5碼之後2碼」必須為數字，且不得輸入00" Display="None" ControlToValidate="EAddressZIPB3" ClientValidationFunction="fCheckZIPB3_1" runat="server"></asp:CustomValidator>
                                    <asp:CustomValidator ID="CheckZIPB3_110" ErrorMessage="班別資料c報名地點「郵遞區號5碼之後2碼」長度必須為 2 碼(例 01 或 99)" Display="None" ControlToValidate="EAddressZIPB3" ClientValidationFunction="fCheckZIPB3_2" runat="server"></asp:CustomValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">報名開始日期 </td>
                                <td class="whitecol" width="30%">
                                    <asp:TextBox ID="SEnterDate" runat="server" Columns="12" MaxLength="11"></asp:TextBox>
                                    <span runat="server">
                                        <img id="imgdate1" style="cursor: pointer" onclick="javascript:show_calendar('SEnterDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30"></span>
                                    <asp:RequiredFieldValidator ID="ReqFieldValidSEnterDate" runat="server" Display="None" ErrorMessage="班別資料「報名開始日期」為必填欄位" ControlToValidate="SEnterDate"></asp:RequiredFieldValidator>
                                    <asp:CustomValidator ID="CustomValidatorSEnterDate" runat="server" ErrorMessage="班別資料「報名開始日期」不是正確的日期格式" Display="None" ControlToValidate="SEnterDate" ClientValidationFunction="check_date"></asp:CustomValidator>
                                    <br />
                                    <asp:DropDownList ID="HR1" runat="server"></asp:DropDownList>時：<asp:DropDownList ID="MM1" runat="server"></asp:DropDownList>分 
                                </td>
                                <td class="bluecol_need" width="20%">報名結束日期 </td>
                                <td class="whitecol" width="30%">
                                    <asp:TextBox ID="FEnterDate" runat="server" Columns="12" MaxLength="11"></asp:TextBox>
                                    <span runat="server">
                                        <img id="imgdate2" style="cursor: pointer" onclick="javascript:show_calendar('FEnterDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30"></span>
                                    <asp:RequiredFieldValidator ID="ReqFieldValidFEnterDate" runat="server" Display="None" ErrorMessage="班別資料「報名結束日期」為必填欄位" ControlToValidate="FEnterDate"></asp:RequiredFieldValidator>
                                    <asp:CustomValidator ID="CustomValidatorFEnterDate" runat="server" ErrorMessage="班別資料「報名結束日期」不是正確的日期格式" Display="None" ControlToValidate="FEnterDate" ClientValidationFunction="check_date"></asp:CustomValidator>
                                    <asp:CustomValidator ID="CustomValidator6" runat="server" ErrorMessage="班別資料 報名開始日期不能比報名結束日期 晚(或者同一天)" Display="None" ClientValidationFunction="CheckTDate2"></asp:CustomValidator>
                                    <br />
                                    <asp:DropDownList ID="HR2" runat="server"></asp:DropDownList>時：<asp:DropDownList ID="MM2" runat="server"></asp:DropDownList>分 
                                </td>
                            </tr>
                            <tr>
                                <td id="td_ExamDate" runat="server" class="bluecol_need" width="20%">甄試日期</td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <asp:TextBox ID="ExamDate" runat="server" MaxLength="11" Columns="12"></asp:TextBox>
                                    <span runat="server">
                                        <img id="imgdate3" style="cursor: pointer" onclick="javascript:show_calendar('ExamDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>                                    <%--<asp:RequiredFieldValidator ID="ReqFieldValidatorExamDate" runat="server" Display="None" ErrorMessage="班別資料「甄試日期」為必填欄位" ControlToValidate="ExamDate"></asp:RequiredFieldValidator>--%>
                                    <asp:CustomValidator ID="CustomValidatorExamDate" runat="server" ErrorMessage="班別資料「甄試日期」不是正確的日期格式" Display="None" ControlToValidate="ExamDate" ClientValidationFunction="check_date"></asp:CustomValidator>
                                    <asp:DropDownList ID="ExamPeriod" runat="server"></asp:DropDownList><asp:Label ID="lab_msg_ExamDate" runat="server"></asp:Label>
                                    <span id="spExamDateTime" runat="server">
                                        <br />
                                        <asp:DropDownList ID="HR6" runat="server"></asp:DropDownList>時：<asp:DropDownList ID="MM6" runat="server"></asp:DropDownList>分
                                    </span>

                                    <%--<asp:Label ID="Label8" runat="server" Display="None" Visible="false"> (無辦理甄試之班級， 免填此欄位)</asp:Label>--%>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need" width="20%">報到日期 </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="CheckInDate" runat="server" MaxLength="11" Columns="12"></asp:TextBox>
                                    <span runat="server">
                                        <img id="img1" style="cursor: pointer" onclick="javascript:show_calendar('CheckInDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                    <asp:RequiredFieldValidator ID="ReqFieldValidatorCheckInDate" runat="server" Display="None" ErrorMessage="班別資料「報到日期」為必填欄位" ControlToValidate="CheckInDate"></asp:RequiredFieldValidator>
                                    <asp:CustomValidator ID="CustomValidatorCheckInDate" runat="server" ErrorMessage="班別資料「報到日期」不是正確的日期格式" Display="None" ControlToValidate="CheckInDate" ClientValidationFunction="check_date"></asp:CustomValidator>
                                </td>
                                <td class="bluecol">導師名稱</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="CTName" runat="server" Width="90%" MaxLength="40"></asp:TextBox></td>
                            </tr>
                            <tr id="trKID20" runat="server">
                                <td class="bluecol" width="20%">&nbsp; 政府政策性產業 </td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <table border="0" cellspacing="1" cellpadding="1" width="100%">
                                        <tr>
                                            <td class="bluecol" width="22%">「5+2」產業創新計畫</td>
                                            <td class="whitecol" width="78%">
                                                <asp:CheckBoxList ID="CBLKID20_1" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">台灣AI行動計畫</td>
                                            <td class="whitecol">
                                                <asp:CheckBoxList ID="CBLKID20_2" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">數位國家創新經濟<br />
                                                發展方案</td>
                                            <td class="whitecol">
                                                <asp:CheckBoxList ID="CBLKID20_3" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">國家資通安全發展方案</td>
                                            <td class="whitecol">
                                                <asp:CheckBoxList ID="CBLKID20_4" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">前瞻基礎建設計畫</td>
                                            <td class="whitecol">
                                                <asp:CheckBoxList ID="CBLKID20_5" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">新南向政策</td>
                                            <td class="whitecol" width="80%">
                                                <asp:CheckBoxList ID="CBLKID20_6" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <%--<tr><td class="bluecol_need">進階政策性產業類別</td><td class="whitecol"><asp:CheckBoxList ID="CBLKID22" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList></td></tr>--%>
                                    </table>
                                </td>
                            </tr>
                            <tr id="trKID25" runat="server">
                                <td class="bluecol">&nbsp; 政府政策性產業 </td>
                                <td class="whitecol" colspan="3" width="80%">
                                    <table border="0" cellspacing="1" cellpadding="1" width="100%">
                                        <tr>
                                            <td class="bluecol" width="22%">亞洲矽谷</td>
                                            <td class="whitecol" width="78%">
                                                <asp:CheckBoxList ID="CBLKID25_1" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">重點產業</td>
                                            <td class="whitecol">
                                                <asp:CheckBoxList ID="CBLKID25_2" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">台灣AI行動計畫</td>
                                            <td class="whitecol">
                                                <asp:CheckBoxList ID="CBLKID25_3" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">智慧國家方案</td>
                                            <td class="whitecol">
                                                <asp:CheckBoxList ID="CBLKID25_4" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">國家人才競爭力躍升方案</td>
                                            <td class="whitecol">
                                                <asp:CheckBoxList ID="CBLKID25_5" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">新南向政策</td>
                                            <td class="whitecol">
                                                <asp:CheckBoxList ID="CBLKID25_6" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <%--<tr id="trCBLKID25_7" runat="server"><td class="bluecol">AI加值應用</td><td class="whitecol"><asp:CheckBoxList ID="CBLKI
D25_7" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList></td></tr><tr id="trCBLKID25_8" 
runat="server"><td class="bluecol">職場續航</td><td class="whitecol"><asp:CheckBoxList ID="CBLKID25_8" runat="server" Repeat
Columns="3" RepeatDirection="Horizontal"></asp:CheckBoxList></td></tr><tr><td class="bluecol">進階政策性產業類別</td><td class="w
hitecol"><asp:CheckBoxList ID="CBLKID22B" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList></td></tr>--%>
                                    </table>
                                </td>
                            </tr>
                            <tr id="trCBLKID60" runat="server">
                                <td class="bluecol">產業別(管考) </td>
                                <td class="whitecol" colspan="3">
                                    <asp:CheckBoxList ID="CBLKID60" runat="server" RepeatColumns="4" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                </td>
                            </tr>
                            <%----%>
                            <tr>
                                <td colspan="4" width="100%">
                                    <hr />
                                </td>
                            </tr>
                        </table>
                        <table id="TableLay6" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <tr>
                                <td>
                                    <table id="TableCost1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                        <tr>
                                            <td class="table_title" width="32%">項目 </td>
                                            <td class="table_title" width="18%">單價 </td>
                                            <td class="table_title" width="12%">數量 </td>
                                            <td class="table_title" width="12%">計價單位 </td>
                                            <td class="table_title" width="26%"></td>
                                        </tr>
                                        <tr>
                                            <td class="whitecol" width="32%">
                                                <asp:DropDownList ID="CostID" runat="server"></asp:DropDownList>
                                                <asp:TextBox ID="ItemOther" runat="server" Width="50%"></asp:TextBox>
                                            </td>
                                            <td class="whitecol" width="18%">
                                                <asp:TextBox ID="OPrice" runat="server" Columns="3" Width="50%"></asp:TextBox>元(可輸入小數點第二位) </td>
                                            <td class="whitecol" width="12%">
                                                <asp:TextBox ID="Itemage" runat="server" Columns="2" Width="50%"></asp:TextBox>人 </td>
                                            <td class="whitecol" width="12%">
                                                <asp:TextBox ID="ItemCost" runat="server" Columns="2" Width="50%"></asp:TextBox>小時 </td>
                                            <td class="whitecol" width="26%" align="center">
                                                <asp:Button ID="Button2" runat="server" CausesValidation="False" Text="新增" CssClass="button_b_M"></asp:Button>&nbsp;
                                                <input id="Button3" type="button" value="新增行政管理費" name="Button3" runat="server" class="button_b_M">
                                                <input id="Button3b" type="button" value="新增營業稅" name="Button3b" runat="server" class="button_b_M">
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="5" class="whitecol" width="100%">說明:如果新增該項目金額時，沒有數量或計價單位時，請輸入1。<br>
                                                計價單位:若費用項目是採用人/時作為基本單價時，那計價單位就是要輸入時數，而數量則是輸入人數，例如:15*人數*時數，單價=15，數量=人數，計價單位=時數</td>
                                        </tr>
                                        <tr>
                                            <td colspan="5" width="100%">
                                                <table class="font" id="DataGrid1Table" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                                    <tr>
                                                        <td class="table_title">費用列表 </td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                                <AlternatingItemStyle BackColor="#F5F5F5" />
                                                                <Columns>
                                                                    <asp:BoundColumn HeaderText="項目" FooterText="總計"></asp:BoundColumn>
                                                                    <asp:TemplateColumn HeaderText="單價">
                                                                        <HeaderStyle Width="12%"></HeaderStyle>
                                                                        <ItemTemplate>
                                                                            <asp:Label ID="Label4" runat="server">Label</asp:Label>
                                                                        </ItemTemplate>
                                                                        <EditItemTemplate>
                                                                            <asp:TextBox ID="TextBox1" runat="server" Columns="3"></asp:TextBox>
                                                                        </EditItemTemplate>
                                                                    </asp:TemplateColumn>
                                                                    <asp:TemplateColumn HeaderText="數量">
                                                                        <HeaderStyle Width="12%"></HeaderStyle>
                                                                        <ItemTemplate>
                                                                            <asp:Label ID="Label5" runat="server">Label</asp:Label>
                                                                        </ItemTemplate>
                                                                        <EditItemTemplate>
                                                                            <asp:TextBox ID="TextBox2" runat="server" Columns="2"></asp:TextBox>
                                                                        </EditItemTemplate>
                                                                    </asp:TemplateColumn>
                                                                    <asp:TemplateColumn HeaderText="計價單位">
                                                                        <HeaderStyle Width="12%"></HeaderStyle>
                                                                        <ItemTemplate>
                                                                            <asp:Label ID="Label6" runat="server">Label</asp:Label>
                                                                        </ItemTemplate>
                                                                        <EditItemTemplate>
                                                                            <asp:TextBox ID="TextBox3" runat="server" Columns="2"></asp:TextBox>
                                                                        </EditItemTemplate>
                                                                    </asp:TemplateColumn>
                                                                    <asp:BoundColumn HeaderText="小計">
                                                                        <HeaderStyle Width="12%"></HeaderStyle>
                                                                    </asp:BoundColumn>
                                                                    <asp:TemplateColumn HeaderText="功能">
                                                                        <HeaderStyle Width="14%"></HeaderStyle>
                                                                        <ItemStyle HorizontalAlign="Center" Font-Size="Small" />
                                                                        <ItemTemplate>
                                                                            <asp:Button ID="Button4" runat="server" CausesValidation="False" Text="修改" CommandName="edit" CssClass="asp_button_M"></asp:Button>
                                                                            <asp:Button ID="Button5" runat="server" CausesValidation="False" Text="刪除" CommandName="del" CssClass="asp_button_M"></asp:Button>
                                                                        </ItemTemplate>
                                                                        <EditItemTemplate>
                                                                            <asp:Button ID="Button6" runat="server" CausesValidation="False" Text="更新" CommandName="update" CssClass="asp_button_M"></asp:Button>
                                                                            <asp:Button ID="Button7" runat="server" CausesValidation="False" Text="取消" CommandName="cancel" CssClass="asp_button_M"></asp:Button>
                                                                        </EditItemTemplate>
                                                                    </asp:TemplateColumn>
                                                                    <asp:BoundColumn DataField="CostID" HeaderText="CostID">
                                                                        <ItemStyle Width="12%" />
                                                                    </asp:BoundColumn>
                                                                </Columns>
                                                            </asp:DataGrid>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td colspan="5" width="100%">
                                                            <table id="AdmTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                                                <tr id="AdmGrantTR" runat="server">
                                                                    <td class="bluecol" width="20%">行政管理費 </td>
                                                                    <td class="whitecol" width="80%">
                                                                        <asp:Label ID="AdmCost" runat="server" Width="100%"></asp:Label></td>
                                                                </tr>
                                                                <tr id="TaxGrantTR" runat="server">
                                                                    <td class="bluecol" width="20%">營業稅 </td>
                                                                    <td class="whitecol" width="80%">
                                                                        <asp:Label ID="TaxCost" runat="server" Width="100%"></asp:Label></td>
                                                                </tr>
                                                            </table>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td colspan="5" width="100%">
                                                            <table id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                                                                <tr>
                                                                    <td class="bluecol" width="20%">總計 </td>
                                                                    <td class="whitecol" width="80%">
                                                                        <asp:Label ID="TotalCost1" runat="server"></asp:Label>
                                                                        <asp:CustomValidator ID="CheckCost1" runat="server" ErrorMessage="請輸入訓練費用" Display="None" ClientValidationFunction="CheckTotalCost"></asp:CustomValidator>
                                                                        <input id="AdmGrant" type="hidden" value="0" name="AdmGrant" runat="server">
                                                                        <input id="TaxGrant" type="hidden" value="0" name="TaxGrant" runat="server">
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>
                                        </tr>
                                    </table>
                                    <table id="TableCost2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                        <tr>
                                            <td>
                                                <table id="Table4" cellspacing="1" cellpadding="1" border="0" width="100%" class="font">
                                                    <tr>
                                                        <td class="table_title" colspan="4" width="100%">每人每時計價 </td>
                                                    </tr>
                                                    <tr>
                                                        <td class="bluecol" width="30%">單價 </td>
                                                        <td class="bluecol" width="24%">人數 </td>
                                                        <td class="bluecol" width="24%">時數 </td>
                                                        <td class="bluecol" width="22%"></td>
                                                    </tr>
                                                    <tr>
                                                        <td class="whitecol">
                                                            <asp:TextBox ID="OPrice2" runat="server" Columns="3" Width="30%"></asp:TextBox>元(可輸入小數點第二位)* </td>
                                                        <td class="whitecol">
                                                            <asp:TextBox ID="Itemage2" runat="server" Columns="3" Width="30%"></asp:TextBox>人* </td>
                                                        <td class="whitecol">
                                                            <asp:TextBox ID="ItemCost2" runat="server" Columns="3" Width="30%"></asp:TextBox>小時 </td>
                                                        <td class="whitecol" align="center">
                                                            <asp:Button ID="Button9" runat="server" CausesValidation="False" Text="新增" CssClass="asp_button_M"></asp:Button></td>
                                                    </tr>
                                                </table>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="4" width="100%">
                                                <table id="DataGrid2Table" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                                    <tr>
                                                        <td>
                                                            <asp:DataGrid ID="DataGrid2" runat="server" CssClass="font" AutoGenerateColumns="False" BorderColor="Gray" Width="100%" CellPadding="8">
                                                                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                                <AlternatingItemStyle BackColor="#F5F5F5" />
                                                                <Columns>
                                                                    <asp:TemplateColumn HeaderText="單價">
                                                                        <HeaderStyle Width="16%"></HeaderStyle>
                                                                        <ItemTemplate>
                                                                            <asp:Label ID="DataGrid2Label1" runat="server">Label</asp:Label>
                                                                        </ItemTemplate>
                                                                        <EditItemTemplate>
                                                                            <asp:TextBox ID="DataGrid2TextBox1" runat="server" Columns="3"></asp:TextBox>
                                                                        </EditItemTemplate>
                                                                    </asp:TemplateColumn>
                                                                    <asp:TemplateColumn HeaderText="人數">
                                                                        <HeaderStyle Width="16%"></HeaderStyle>
                                                                        <ItemTemplate>
                                                                            <asp:Label ID="DataGrid2Label2" runat="server">Label</asp:Label>
                                                                        </ItemTemplate>
                                                                        <EditItemTemplate>
                                                                            <asp:TextBox ID="DataGrid2TextBox2" runat="server" Columns="3"></asp:TextBox>
                                                                        </EditItemTemplate>
                                                                    </asp:TemplateColumn>
                                                                    <asp:TemplateColumn HeaderText="時數">
                                                                        <HeaderStyle Width="16%"></HeaderStyle>
                                                                        <ItemTemplate>
                                                                            <asp:Label ID="DataGrid2Label3" runat="server">Label</asp:Label>
                                                                        </ItemTemplate>
                                                                        <EditItemTemplate>
                                                                            <asp:TextBox ID="DataGrid2TextBox3" runat="server" Columns="3"></asp:TextBox>
                                                                        </EditItemTemplate>
                                                                    </asp:TemplateColumn>
                                                                    <asp:BoundColumn HeaderText="小計">
                                                                        <HeaderStyle Width="16%"></HeaderStyle>
                                                                    </asp:BoundColumn>
                                                                    <asp:TemplateColumn HeaderText="功能">
                                                                        <HeaderStyle Width="20%"></HeaderStyle>
                                                                        <ItemStyle HorizontalAlign="Center" />
                                                                        <ItemTemplate>
                                                                            <asp:Button ID="Button12" runat="server" CausesValidation="False" Text="修改" CommandName="edit" CssClass="asp_button_M"></asp:Button>
                                                                            <asp:Button ID="Button13" runat="server" CausesValidation="False" Text="刪除" CommandName="del" CssClass="asp_button_M"></asp:Button>
                                                                        </ItemTemplate>
                                                                        <EditItemTemplate>
                                                                            <asp:Button ID="Button14" runat="server" CausesValidation="False" Text="更新" CommandName="update" CssClass="asp_button_M"></asp:Button>
                                                                            <asp:Button ID="Button15" runat="server" CausesValidation="False" Text="取消" CommandName="cancel" CssClass="asp_button_M"></asp:Button>
                                                                        </EditItemTemplate>
                                                                    </asp:TemplateColumn>
                                                                    <asp:BoundColumn DataField="CostID" HeaderText="CostID">
                                                                        <ItemStyle Width="16%" />
                                                                    </asp:BoundColumn>
                                                                </Columns>
                                                            </asp:DataGrid>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <table class="font" id="Table6" cellspacing="1" cellpadding="1" width="100%" border="0">
                                                                <tr>
                                                                    <td class="bluecol" width="20%">總計 </td>
                                                                    <td class="whitecol" width="80%">
                                                                        <asp:Label ID="TotalCost2" runat="server"></asp:Label></td>
                                                                </tr>
                                                            </table>
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>
                                        </tr>
                                    </table>
                                    <table class="font" id="TableCost3" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                        <tr>
                                            <td colspan="3" width="100%">
                                                <table id="Table5" cellspacing="1" cellpadding="1" border="0" width="100%" class="font">
                                                    <tr>
                                                        <td class="table_title" colspan="3" width="100%">每人輔助計價 </td>
                                                    </tr>
                                                    <tr>
                                                        <td class="bluecol" width="40%">單價 </td>
                                                        <td class="bluecol" width="30%">人數 </td>
                                                        <td class="bluecol" width="30%"></td>
                                                    </tr>
                                                    <tr>
                                                        <td class="whitecol">
                                                            <asp:TextBox ID="OPrice3" runat="server" Columns="3" Width="18%"></asp:TextBox>元(可輸入小數點第二位)* </td>
                                                        <td class="whitecol">
                                                            <asp:TextBox ID="Itemage3" runat="server" Columns="3" Width="30%"></asp:TextBox>人 </td>
                                                        <td class="whitecol" align="center">
                                                            <asp:Button ID="Button10" runat="server" CausesValidation="False" Text="新增" CssClass="asp_button_M"></asp:Button></td>
                                                    </tr>
                                                </table>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="3">
                                                <table id="DataGrid3Table" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                                    <tr>
                                                        <td>
                                                            <asp:DataGrid ID="DataGrid3" runat="server" CssClass="font" AutoGenerateColumns="False" BorderColor="Gray" Width="100%" CellPadding="8">
                                                                <AlternatingItemStyle BackColor="#F5F5F5" />
                                                                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                                <Columns>
                                                                    <asp:TemplateColumn HeaderText="單價">
                                                                        <HeaderStyle Width="20%"></HeaderStyle>
                                                                        <ItemStyle HorizontalAlign="Center" />
                                                                        <ItemTemplate>
                                                                            <asp:Label ID="DataGrid3Label1" runat="server">Label</asp:Label>
                                                                        </ItemTemplate>
                                                                        <EditItemTemplate>
                                                                            <asp:TextBox ID="DataGrid3TextBox1" runat="server" Columns="3"></asp:TextBox>
                                                                        </EditItemTemplate>
                                                                    </asp:TemplateColumn>
                                                                    <asp:TemplateColumn HeaderText="人數">
                                                                        <HeaderStyle Width="20%"></HeaderStyle>
                                                                        <ItemStyle HorizontalAlign="Center" />
                                                                        <ItemTemplate>
                                                                            <asp:Label ID="DataGrid3Label2" runat="server">Label</asp:Label>
                                                                        </ItemTemplate>
                                                                        <EditItemTemplate>
                                                                            <asp:TextBox ID="DataGrid3TextBox2" runat="server" Columns="3"></asp:TextBox>
                                                                        </EditItemTemplate>
                                                                    </asp:TemplateColumn>
                                                                    <asp:BoundColumn HeaderText="小計">
                                                                        <HeaderStyle Width="20%"></HeaderStyle>
                                                                        <ItemStyle HorizontalAlign="Center" />
                                                                    </asp:BoundColumn>
                                                                    <asp:TemplateColumn HeaderText="功能">
                                                                        <HeaderStyle Width="20%"></HeaderStyle>
                                                                        <ItemStyle HorizontalAlign="Center" />
                                                                        <ItemTemplate>
                                                                            <asp:Button ID="Button16" runat="server" CausesValidation="False" Text="修改" CommandName="edit" CssClass="asp_button_M"></asp:Button>
                                                                            <asp:Button ID="Button17" runat="server" CausesValidation="False" Text="刪除" CommandName="del" CssClass="asp_button_M"></asp:Button>
                                                                        </ItemTemplate>
                                                                        <EditItemTemplate>
                                                                            <asp:Button ID="Button18" runat="server" CausesValidation="False" Text="更新" CommandName="update" CssClass="asp_button_M"></asp:Button>
                                                                            <asp:Button ID="Button19" runat="server" CausesValidation="False" Text="取消" CommandName="cancel" CssClass="asp_button_M"></asp:Button>
                                                                        </EditItemTemplate>
                                                                    </asp:TemplateColumn>
                                                                    <asp:BoundColumn DataField="CostID" HeaderText="CostID">
                                                                        <ItemStyle Width="20%" HorizontalAlign="Center" />
                                                                    </asp:BoundColumn>
                                                                </Columns>
                                                            </asp:DataGrid>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <table id="Table8" cellspacing="1" cellpadding="1" width="100%" border="0">
                                                                <tr>
                                                                    <td class="bluecol" width="12%">總計 </td>
                                                                    <td class="whitecol" width="88%">
                                                                        <asp:Label ID="TotalCost3" runat="server"></asp:Label></td>
                                                                </tr>
                                                            </table>
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>
                                        </tr>
                                    </table>
                                    <table id="TableCost4" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server" class="font">
                                        <tr>
                                            <td class="table_title" colspan="4">個人單價計價 </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol" width="36%">項目 </td>
                                            <td class="bluecol" width="20%">單價 </td>
                                            <td class="bluecol" width="10%">數量 </td>
                                            <td class="bluecol" width="34%"></td>
                                        </tr>
                                        <tr>
                                            <td class="whitecol">
                                                <asp:DropDownList ID="CostID4" runat="server"></asp:DropDownList>
                                                <asp:TextBox ID="ItemOther4" runat="server" Width="30%"></asp:TextBox>
                                            </td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="OPrice4" runat="server" Columns="3" Width="30%"></asp:TextBox>元(可輸入小數點第二位) </td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="Itemage4" runat="server" Columns="3" Width="60%"></asp:TextBox>人 </td>
                                            <td class="whitecol" style="text-align: center;">
                                                <asp:Button ID="Button11" runat="server" CausesValidation="False" Text="新增" CssClass="asp_button_M"></asp:Button>
                                                <input id="Button25" type="button" value="新增行政管理費" name="Button25" runat="server" class="asp_button_M" />
                                                <input id="Button25b" type="button" value="新增營業稅" name="Button25b" runat="server" class="asp_button_M" />
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="4" width="100%">
                                                <table id="DataGrid4Table" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                                    <tr>
                                                        <td>
                                                            <asp:DataGrid ID="DataGrid4" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" BorderColor="Gray" CellPadding="8">
                                                                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                                <AlternatingItemStyle BackColor="#F5F5F5" />
                                                                <Columns>
                                                                    <asp:BoundColumn HeaderText="項目">
                                                                        <ItemStyle Width="34%" />
                                                                    </asp:BoundColumn>
                                                                    <asp:TemplateColumn HeaderText="單價">
                                                                        <ItemStyle Width="12%" HorizontalAlign="Center" />
                                                                        <ItemTemplate>
                                                                            <asp:Label ID="DataGrid4Label1" runat="server">Label</asp:Label>
                                                                        </ItemTemplate>
                                                                        <EditItemTemplate>
                                                                            <asp:TextBox ID="DataGrid4TextBox1" runat="server" Columns="3"></asp:TextBox>
                                                                        </EditItemTemplate>
                                                                    </asp:TemplateColumn>
                                                                    <asp:TemplateColumn HeaderText="數量">
                                                                        <ItemStyle Width="12%" HorizontalAlign="Center" />
                                                                        <ItemTemplate>
                                                                            <asp:Label ID="DataGrid4Label2" runat="server">Label</asp:Label>
                                                                        </ItemTemplate>
                                                                        <EditItemTemplate>
                                                                            <asp:TextBox ID="DataGrid4TextBox2" runat="server" Columns="3"></asp:TextBox>
                                                                        </EditItemTemplate>
                                                                    </asp:TemplateColumn>
                                                                    <asp:BoundColumn HeaderText="小計">
                                                                        <ItemStyle Width="12%" HorizontalAlign="Center" />
                                                                    </asp:BoundColumn>
                                                                    <asp:TemplateColumn HeaderText="功能">
                                                                        <ItemStyle Width="20%" HorizontalAlign="Center" />
                                                                        <ItemTemplate>
                                                                            <asp:Button ID="Button20" runat="server" CausesValidation="False" Text="修改" CommandName="edit" CssClass="asp_button_M"></asp:Button>
                                                                            <asp:Button ID="Button21" runat="server" CausesValidation="False" Text="刪除" CommandName="del" CssClass="asp_button_M"></asp:Button>
                                                                        </ItemTemplate>
                                                                        <EditItemTemplate>
                                                                            <asp:Button ID="Button22" runat="server" CausesValidation="False" Text="更新" CommandName="update" CssClass="asp_button_M"></asp:Button>
                                                                            <asp:Button ID="Button23" runat="server" CausesValidation="False" Text="取消" CommandName="cancel" CssClass="asp_button_M"></asp:Button>
                                                                        </EditItemTemplate>
                                                                    </asp:TemplateColumn>
                                                                    <asp:BoundColumn DataField="CostID" HeaderText="CostID">
                                                                        <ItemStyle Width="10%" HorizontalAlign="Center" />
                                                                    </asp:BoundColumn>
                                                                </Columns>
                                                            </asp:DataGrid>
                                                        </td>
                                                    </tr>
                                                    <tr>
                                                        <td>
                                                            <table id="Table9" cellspacing="1" cellpadding="1" width="100%" border="0">
                                                                <tr id="AdmTR4" runat="server">
                                                                    <td width="20%" class="bluecol">行政管理費 </td>
                                                                    <td width="80%" class="whitecol">
                                                                        <asp:Label ID="AdmCost4" runat="server"></asp:Label></td>
                                                                </tr>
                                                                <tr id="TaxTR4" runat="server">
                                                                    <td class="bluecol">營業稅 </td>
                                                                    <td class="whitecol">
                                                                        <asp:Label ID="TaxCost4" runat="server"></asp:Label>
                                                                        <input id="hidTaxCost4" type="hidden" value="0" name="hidTaxCost4" runat="server">
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td class="bluecol">總計 </td>
                                                                    <td class="whitecol">
                                                                        <asp:Label ID="TotalCost4" runat="server"></asp:Label>
                                                                        <input id="AdmGrant4" type="hidden" value="0" name="AdmGrant4" runat="server">
                                                                        <input id="TaxGrant4" type="hidden" value="0" name="TaxGrant4" runat="server">
                                                                    </td>
                                                                </tr>
                                                                <tr>
                                                                    <td class="bluecol">費用/人 </td>
                                                                    <td class="whitecol">
                                                                        <asp:Label ID="PerCost" runat="server"></asp:Label></td>
                                                                </tr>
                                                                <tr id="tr_ACTHUMCOST" runat="server">
                                                                    <td class="bluecol">單一人時成本</td>
                                                                    <td class="whitecol">
                                                                        <asp:Label ID="ACTHUMCOST" runat="server"></asp:Label>&nbsp;&nbsp;
                                                                        <asp:Label ID="labmsg_ACTHUMCOST_1" runat="server" Text="(總計/訓練時數/訓練人數)"></asp:Label></td>
                                                                </tr>
                                                                <tr id="tr_METCOSTPER" runat="server">
                                                                    <td class="bluecol">材料費占比 </td>
                                                                    <td class="whitecol">
                                                                        <asp:Label ID="METCOSTPER" runat="server"></asp:Label>&nbsp;&nbsp;&nbsp;
                                                                        <asp:Label ID="labmsg_METCOSTPER" runat="server" Text="(材料費小計/總計)"></asp:Label></td>
                                                                </tr>
                                                            </table>
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>
                                        </tr>
                                    </table>
                                    <hr />
                                </td>
                            </tr>
                        </table>
                        <table id="TableLay7" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <tr>
                                <td width="20%" class="bluecol_need">經費來源 </td>
                                <td width="80%">
                                    <table id="Lay7tb2" cellspacing="1" cellpadding="1" width="100%" border="0">
                                        <tr>
                                            <td class="whitecol">政府補助金額：新臺幣(每人每期費用)
                                                <asp:TextBox ID="DefGovCost" runat="server" Columns="7" Width="15%">0</asp:TextBox>元*每班人數
                                                <asp:Label ID="TNum1" runat="server"></asp:Label>=
                                                <asp:Label ID="Total1" runat="server"></asp:Label>元
                                                <asp:RegularExpressionValidator ID="check18" runat="server" ErrorMessage="經費來源「政府負擔」請輸入數字" Display="None" ControlToValidate="DefGovCost" ValidationExpression="[0-9]{0,8}"></asp:RegularExpressionValidator>
                                            </td>
                                        </tr>
                                        <tr id="tr_DefUnitCost" runat="server">
                                            <td class="whitecol">企業負擔金額：新臺幣(每人每期費用)
                                                <asp:TextBox ID="DefUnitCost" runat="server" Columns="7" Width="15%">0</asp:TextBox>元*每班人數
                                                <asp:Label ID="TNum2" runat="server"></asp:Label>=
                                                <asp:Label ID="Total2" runat="server"></asp:Label>元
                                                <asp:RegularExpressionValidator ID="check20" runat="server" ErrorMessage="經費來源「民間企業度負擔」請輸入數字" Display="None" ControlToValidate="DefUnitCost" ValidationExpression="[0-9]{0,8}"></asp:RegularExpressionValidator>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="whitecol">學員負擔金額：新臺幣(每人每期費用)
                                                <asp:TextBox ID="DefStdCost" runat="server" Columns="7" Width="15%">0</asp:TextBox>元*每班人數
                                                <asp:Label ID="TNum3" runat="server"></asp:Label>=
                                                <asp:Label ID="Total3" runat="server"></asp:Label>元
                                                <asp:RegularExpressionValidator ID="check22" runat="server" ErrorMessage="經費來源「學員負擔共負擔」請輸入數字" Display="None" ControlToValidate="DefStdCost" ValidationExpression="[0-9]{0,8}"></asp:RegularExpressionValidator><asp:CustomValidator ID="CustomValidator1" runat="server" ErrorMessage="請輸入經費來源" Display="None" ClientValidationFunction="CheckDef"></asp:CustomValidator>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2" width="100%">
                                    <hr />
                                </td>
                            </tr>
                        </table>
                        <table id="TableLay8" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <tr>
                                <td class="bluecol_need" width="20%">是否為輔導考照班 </td>
                                <td class="whitecol" width="80%">
                                    <asp:RadioButton ID="RBCOACHING_Y" runat="server" Text="是" GroupName="RBCOACHING" /><asp:RadioButton ID="RBCOACHING_N" runat="server" Text="否" GroupName="RBCOACHING" />
                                </td>
                            </tr>
                            <tr id="tr_COACHING1">
                                <td class="bluecol" width="20%">完訓後可參加之全國技術士技能檢定職類與考試級別 </td>
                                <td class="whitecol" width="80%">
                                    <table width="100%">
                                        <tr>
                                            <td class="whitecol" width="100%">1.<asp:TextBox ID="txtGP1" runat="server" Width="30%" onfocus="this.blur()"></asp:TextBox><asp:TextBox ID="txtXM1" runat="server" Width="50%" onfocus="this.blur()"></asp:TextBox><asp:TextBox ID="txtLV1" runat="server" Width="10%" onfocus="this.blur()"></asp:TextBox>
                                                <input id="btnExamC1" onclick="openExamC('txtGP1', 'txtXM1', 'txtLV1', 'EXAM1val', 'EXLV1val', 'btnExamC2');" type="button" value="..." runat="server" class="button_b_Mini" />
                                                <input id="EXAM1val" type="hidden" name="EXAM1val" runat="server" />
                                                <input id="EXLV1val" type="hidden" name="EXLV1val" runat="server" />
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="whitecol" width="100%">2.<asp:TextBox ID="txtGP2" runat="server" Width="30%" onfocus="this.blur()"></asp:TextBox><asp:TextBox ID="txtXM2" runat="server" Width="50%" onfocus="this.blur()"></asp:TextBox><asp:TextBox ID="txtLV2" runat="server" Width="10%" onfocus="this.blur()"></asp:TextBox>
                                                <input id="btnExamC2" onclick="openExamC('txtGP2', 'txtXM2', 'txtLV2', 'EXAM2val', 'EXLV2val', 'btnExamC3');" type="button" value="..." runat="server" class="button_b_Mini" />
                                                <input id="EXAM2val" type="hidden" name="EXAM2val" runat="server" />
                                                <input id="EXLV2val" type="hidden" name="EXLV2val" runat="server" />
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="whitecol" width="100%">3.<asp:TextBox ID="txtGP3" runat="server" Width="30%" onfocus="this.blur()"></asp:TextBox><asp:TextBox ID="txtXM3" runat="server" Width="50%" onfocus="this.blur()"></asp:TextBox><asp:TextBox ID="txtLV3" runat="server" Width="10%" onfocus="this.blur()"></asp:TextBox>
                                                <input id="btnExamC3" onclick="openExamC('txtGP3', 'txtXM3', 'txtLV3', 'EXAM3val', 'EXLV3val', '');" type="button" value="..." runat="server" class="button_b_Mini" />
                                                <input id="EXAM3val" type="hidden" name="EXAM3val" runat="server" />
                                                <input id="EXLV3val" type="hidden" name="EXLV3val" runat="server" />
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">備註 </td>
                                <td class="whitecol" width="80%">(應依經費來源說明公務、就保基金、就安基金參訓名額配置數)<br>
                                    <asp:TextBox ID="Note" runat="server" Columns="50" Rows="6" TextMode="MultiLine" Width="77%"></asp:TextBox></td>
                            </tr>
                            <tr id="TR_2006_01" runat="server">
                                <td class="bluecol" width="20%">e網報名後<br>
                                    顯示訊息 </td>
                                <td class="whitecol" width="80%">
                                    <asp:TextBox ID="ESiteMsg" runat="server" Columns="50" Rows="6" TextMode="MultiLine" Width="77%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">專長能力標籤 </td>
                                <td class="whitecol">
                                    <table class="font" id="Table_9a" width="100%" cellpadding="1" cellspacing="1">
                                        <tr>
                                            <td class="bluecol_need" width="6%">1</td>
                                            <td class="whitecol" width="91%">
                                                <asp:Label ID="Label4" runat="server" Text="名稱" class="bluecol_need"></asp:Label>
                                                <asp:TextBox ID="txtABILITY1" runat="server" Width="33%" CssClass="form-control convUnsafeChar width-auto autocomplete" autocomplete_bind="Y" data-cat="APP_TEXT APP_TEXT_OTHER APP_TEXT_OTHER_SEEK" MaxLength="30" size="24" placeholder="請至少輸兩個關鍵字作搜尋"></asp:TextBox>
                                                <%--<div id="resABILITY1"></div>--%>
                                                <asp:Label ID="Label12" runat="server" Text="描述" class="bluecol"></asp:Label>
                                                <asp:TextBox ID="txtABILITY_DESC1" runat="server" Width="44%" CssClass="form-control convUnsafeChar width-auto" MaxLength="200" placeholder="字數上限為200字(名稱有值才會儲存)"></asp:TextBox>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol" width="6%">2</td>
                                            <td class="whitecol" width="91%">
                                                <asp:Label ID="Label13" runat="server" Text="名稱" class="bluecol"></asp:Label>
                                                <asp:TextBox ID="txtABILITY2" runat="server" Width="33%" CssClass="form-control convUnsafeChar width-auto autocomplete" autocomplete_bind="Y" data-cat="APP_TEXT APP_TEXT_OTHER APP_TEXT_OTHER_SEEK" MaxLength="30" size="24" placeholder="請至少輸兩個關鍵字作搜尋"></asp:TextBox>
                                                <%--<div id="resABILITY2"></div>--%>
                                                <asp:Label ID="Label14" runat="server" Text="描述" class="bluecol"></asp:Label>
                                                <asp:TextBox ID="txtABILITY_DESC2" runat="server" Width="44%" CssClass="form-control convUnsafeChar width-auto" MaxLength="200" placeholder="字數上限為200字(名稱有值才會儲存)"></asp:TextBox>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol" width="6%">3</td>
                                            <td class="whitecol" width="91%">
                                                <asp:Label ID="Label15" runat="server" Text="名稱" class="bluecol"></asp:Label>
                                                <asp:TextBox ID="txtABILITY3" runat="server" Width="33%" CssClass="form-control convUnsafeChar width-auto autocomplete" autocomplete_bind="Y" data-cat="APP_TEXT APP_TEXT_OTHER APP_TEXT_OTHER_SEEK" MaxLength="30" size="24" placeholder="請至少輸兩個關鍵字作搜尋"></asp:TextBox>
                                                <%--<div id="resABILITY3"></div>--%>
                                                <asp:Label ID="Label16" runat="server" Text="描述" class="bluecol"></asp:Label>
                                                <asp:TextBox ID="txtABILITY_DESC3" runat="server" Width="44%" CssClass="form-control convUnsafeChar width-auto" MaxLength="200" placeholder="字數上限為200字(名稱有值才會儲存)"></asp:TextBox>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol" width="6%">4</td>
                                            <td class="whitecol" width="91%">
                                                <asp:Label ID="Label17" runat="server" Text="名稱" class="bluecol"></asp:Label>
                                                <asp:TextBox ID="txtABILITY4" runat="server" Width="33%" CssClass="form-control convUnsafeChar width-auto autocomplete" autocomplete_bind="Y" data-cat="APP_TEXT APP_TEXT_OTHER APP_TEXT_OTHER_SEEK" MaxLength="30" size="24" placeholder="請至少輸兩個關鍵字作搜尋"></asp:TextBox>
                                                <%--<div id="resABILITY4"></div>--%>
                                                <asp:Label ID="Label18" runat="server" Text="描述" class="bluecol"></asp:Label>
                                                <asp:TextBox ID="txtABILITY_DESC4" runat="server" Width="44%" CssClass="form-control convUnsafeChar width-auto" MaxLength="200" placeholder="字數上限為200字(名稱有值才會儲存)"></asp:TextBox>
                                            </td>
                                        </tr>
                                        <%--<tr><td class="whitecol">&nbsp;</td></tr>--%>
                                    </table>
                                </td>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
            <tr>
                <td>
                    <div align="center" class="whitecol">
                        <asp:CustomValidator ID="cvCheckhours" runat="server" ErrorMessage="[課程編配]的[總計]小時數,不能大於[課程內容簡介]的[時數]" Display="None" ClientValidationFunction="Check_totalhours"></asp:CustomValidator>
                        <asp:Button ID="Button8" runat="server" CausesValidation="False" Text="草稿儲存" CssClass="asp_Export_M"></asp:Button>
                        <asp:Button ID="btnAdd" runat="server" Text="正式儲存" CssClass="asp_Export_M"></asp:Button>
                        <asp:Button ID="Button24" runat="server" CausesValidation="False" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                        <asp:ValidationSummary ID="ValidationSummary1" runat="server" ShowMessageBox="True" ShowSummary="False" DisplayMode="List"></asp:ValidationSummary>
                    </div>
                </td>
            </tr>
        </table>
        <input id="LayerState" type="hidden" runat="server" />
        <input id="isBlack" type="hidden" name="isBlack" runat="server" />
        <input id="Blackorgname" type="hidden" name="Blackorgname" runat="server" />
        <asp:HiddenField ID="Hid_ComIDNO" runat="server" />
        <asp:HiddenField ID="HidOrgID" runat="server" />
        <asp:HiddenField ID="Hid_RID1" runat="server" />
        <asp:HiddenField ID="HidAgeType" runat="server" />
        <asp:HiddenField ID="Hid_MaxTNum" runat="server" />
        <asp:HiddenField ID="Hid_CostItem_GUID1" runat="server" />
        <asp:HiddenField ID="Hid_TrainDesc_GUID1" runat="server" />
        <asp:HiddenField ID="Hid_TPlanID" runat="server" />
        <asp:HiddenField ID="Hid_cost_02_08" runat="server" />
        <asp:HiddenField ID="Hid_TRNUNIT" runat="server" />
    </form>
    <%--<iframe id="ifmChceckZip" height="0%" src="../../Common/CheckZip.aspx" width="0%" />--%>
</body>
</html>
