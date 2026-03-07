<%@ Page Language="vb" AutoEventWireup="false" MaintainScrollPositionOnPostback="true" CodeBehind="TC_03_006.aspx.vb" Inherits="WDAIIP.TC_03_006" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html>
<head>
    <title>班級申請作業</title>
    <meta name="vs_defaultClientScript" content="JavaScript" />
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5" />
    <link rel="stylesheet" type="text/css" href="../../css/style.css" />
    <script type="text/javascript" src="../../js/OpenWin/openwin.js"></script>
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <%--<script type="text/javascript" src="../../Scripts/sisyphus.js"></script>--%>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript" src="../../js/autocomplete.js"></script>
    <%--專長能力標籤--%>
    <script type="text/javascript">

</script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function move_scrollTop1() {
            // 檢查 URL 是否包含特定 Hash，或根據伺服器端傳遞的標記來執行 //debugger;
            var $container = $('html, body'); // 取得容器 //$('#mainFrame'); 
            var $target = $('#tdContent');    // 取得目標錨點
            var $hf = $('#hfScrollToAnchor'); // 取得值
            if ($target.length && $container.length && $container.length > 1 && $hf.length && $hf.val() == 'Y') {
                $hf.val("");
                // 方式 A：取得原生 DOM 物件後觸發 (最可靠)
                $('#lnkToContent')[0].click();
            }
        }

        $(document).ready(function () {
            // 監聽所有含有 group-check 類別的 checkbox (內部的 input)
            $(".group-check-tgov input:checkbox").click(function () {
                var $this = $(this);
                if ($this.is(":checked")) {
                    // 1. 核心邏輯：將其他的 checkbox 取消勾選
                    $(".group-check-tgov input:checkbox").not($this).prop("checked", false);
                    //debugger; //var isCY = $this.closest("span").attr("id").indexOf("TGovExamCY") != -1;
                    // 2. 進階邏輯：如果選的不是 "是(CY)"，就清空並停用文字方塊 (視需求選用)
                    var isCY = ($(this).attr("id").indexOf("TGovExamCY") != -1);
                    var isCN = ($(this).attr("id").indexOf("TGovExamCN") != -1);
                    var isCG = ($(this).attr("id").indexOf("TGovExamCG") != -1);
                    var V_TGovExam = "";
                    if (isCY) { V_TGovExam = "Y"; }
                    if (isCN) { V_TGovExam = "N"; }
                    if (isCG) { V_TGovExam = "G"; }
                    $("#Hid_TGovExam").val(V_TGovExam);
                    if (!isCY) {
                        if ($("#<%=GOVAGENAME.ClientID%>").val() != "") { $("#Hid_GOVAGENAME_og").val($("#<%=GOVAGENAME.ClientID%>").val()); }
                        if ($("#<%=TGovExamName.ClientID%>").val() != "") { $("#Hid_TGovExamName_og").val($("#<%=TGovExamName.ClientID%>").val()); }
                        $("#<%=GOVAGENAME.ClientID%>, #<%=TGovExamName.ClientID%>").val("").prop("disabled", true);
                    } else {
                        $("#<%=GOVAGENAME.ClientID%>, #<%=TGovExamName.ClientID%>").prop("disabled", false);
                        if ($("#Hid_GOVAGENAME_og").val() != "") { $("#<%=GOVAGENAME.ClientID%>").val($("#Hid_GOVAGENAME_og").val()); }
                        if ($("#Hid_TGovExamName_og").val() != "") { $("#<%=TGovExamName.ClientID%>").val($("#Hid_TGovExamName_og").val()); }
                    }
                }
            });

            // 監聽 CheckBox 的 change 事件
            $("[id$='CB_EnvZeroTrain']").on('change', function () {
                CHK_EnvZeroTrain();
            });

            /*檢查 URL 是否包含特定 Hash，或根據伺服器端傳遞的標記來執行,if (window.location.hash === '#tdContent' || $('#hfScrollToAnchor').val() === 'Y') {,$('html, body').animate({,scrollTop: $("#tdContent").offset().top,}, 900);  毫秒平滑捲動,}*/
            setTimeout(function () { move_scrollTop1(); }, 1500);
        });

        $(function () {
            return;
            if (_isIE) { return; }
            if ($('#Hid_sisyphus').val() == "N") { return; }
            //自動定時每 5 秒儲存一次輸入的資料,$('#form1').sisyphus({ timeout: 5 });,網頁載入時, 若有未送出的資料則載入它, ˊ
            var msgX1 = "是否要載入 系統自動暫存功能?(5秒暫存1次)";
            var flag_c1 = confirm(msgX1);
            if (flag_c1) {
                //alert('載入!');
                $('#form1').sisyphus({
                    timeout: 5,
                    onRestore: function () { alert('載入 未送出的資料完成!'); }
                });
            }
        });
        <%--function search0910(schInput1, result2) {
            //autocomplete_init
            var keyword = $('#' + schInput1).val();
            if (keyword.length < 2) { return; }
            //var urlxStrSegment = 'https://job.taiwanjobs.gov.tw/StrSegment.ashx?str=' + keyword;
            //var urlxStrSegment = 'http://192.168.0.56:8531/StrSegment.ashx?str=' + keyword;
            //var tfS = "2"; var tfE = "99999";
            var cat = "APP_TEXT APP_TEXT_OTHER APP_TEXT_OTHER_SEEK";
            var q = "key:(" + keyword + "* *" + keyword + ") AND cat:(" + cat + ") NOT artificial:1 AND tf:[2 TO 99999]";
            //var urlxSolrApi = 'https://job.taiwanjobs.gov.tw/SolrApi.aspx?q=' + q;
            //var urlxSolrApi = 'http://192.168.0.56:8531/SolrApi.aspx?q=' + q;
            var urlxSolrApi = "<%=ResolveUrl("~/ajax/SolrAPI.ashx")%>" + '?q=' + q;
            $.ajax({
                url: urlxSolrApi,
                type: 'GET',
                success: function (data) {
                    $('#' + result2).html(data.data);
                },
                error: function (error) {
                    console.error('發生錯誤：', error);
                    $('#' + result2).html('搜尋失敗，請稍後再試。');
                }
            });  var data = {};
            data.keyword = keyword;
            var PostURL_2 = "<%=ResolveUrl("~/ajax/SolrAPI.ashx")%>";
            var jqxhr = $.post(PostURL_2, data,
                function (r) {
                    if (r.status == true && r.message == "ok") {
                        //console.log("do ajax/SolrAPI.ashx get ok: " + r.data + ", " + JSON.stringify(data));
                        autocomplete_dropdown_refresh(target, data);
                    }
                    else {
                        console.warn("!do ajax/SolrAPI.ashx get warn: " + r.message + ": " + r.data + ", " + JSON.stringify(data));
                    }
                })
                .fail(function () {
                    console.error("!do ajax/SolrAPI.ashx fail failed: " + JSON.stringify(data));
                });
        //}--%>
        function CHK_EnvZeroTrain() {
            // 取得控制項 
            var $CB_EnvZeroTrain = $("[id$='CB_EnvZeroTrain']");
            if ($CB_EnvZeroTrain.length == 0) return;
            var tit_msg2 = '選【是】本課程是否應報請主管機關核備 與 學員是否可依個人需求參加政府機關辦理相關證照考試或技能檢定,鎖死';
            var tit_msg1 = '本課程屬環境部淨零綠領人才培育課程';
            var $Hid_REPORTE = $("[id$='Hid_REPORTE']");
            var $rbl_REPORTE_Y = $("[id$='rbl_REPORTE_Y']");
            var $rbl_REPORTE_N = $("[id$='rbl_REPORTE_N']");
            //if ($(this).is(':checked'))
            if ($CB_EnvZeroTrain.is(':checked')) {
                if ($("#<%=GOVAGENAME.ClientID%>").val() != "") { $("#Hid_GOVAGENAME_og").val($("#<%=GOVAGENAME.ClientID%>").val()); }
                if ($("#<%=TGovExamName.ClientID%>").val() != "") { $("#Hid_TGovExamName_og").val($("#<%=TGovExamName.ClientID%>").val()); }
                $("#<%=GOVAGENAME.ClientID%>, #<%=TGovExamName.ClientID%>").val("").prop("disabled", true).attr("title", tit_msg1);
                // 當勾選時，將「是」設定為選取，並觸發 change 事件（若有其他連動邏輯）
                $rbl_REPORTE_Y.prop('checked', true).change().prop("disabled", true);
                $rbl_REPORTE_N.prop('checked', false).change().prop("disabled", true);
                $("#TGovExamCY").prop("disabled", true).prop("checked", false).attr("title", tit_msg1);
                $("#TGovExamCN").prop("disabled", true).prop("checked", false).attr("title", tit_msg1);
                $("#TGovExamCG").prop("disabled", true).prop("checked", true).attr("title", tit_msg1);
                $rbl_REPORTE_Y.attr("title", tit_msg1).parent().attr("title", tit_msg1);
                $rbl_REPORTE_N.attr("title", tit_msg1);
                $CB_EnvZeroTrain.attr("title", tit_msg2);
                $Hid_REPORTE.val("Y");
                $("#Hid_TGovExam").val("G");
            }
            else {
                $("#TGovExamCY").prop("disabled", false).removeAttr("title");
                $("#TGovExamCN").prop("disabled", false).removeAttr("title");
                $("#TGovExamCG").prop("disabled", true).prop("checked", false).removeAttr("title");
                $rbl_REPORTE_Y.removeAttr("title").parent().removeAttr("title");
                $rbl_REPORTE_N.removeAttr("title");
                $CB_EnvZeroTrain.removeAttr("title");
                $rbl_REPORTE_Y.prop("disabled", false);
                $rbl_REPORTE_N.prop("disabled", false);
                $("#<%=GOVAGENAME.ClientID%>, #<%=TGovExamName.ClientID%>").prop("disabled", false).removeAttr("title");
                if ($("#Hid_GOVAGENAME_og").val() != "") { $("#<%=GOVAGENAME.ClientID%>").val($("#Hid_GOVAGENAME_og").val()); }
                if ($("#Hid_TGovExamName_og").val() != "") { $("#<%=TGovExamName.ClientID%>").val($("#Hid_TGovExamName_og").val()); }
                $Hid_REPORTE.val("");
                $("#Hid_TGovExam").val("");
            }
        }
        //CBLKID25_8.Attributes("onclick") = "chg_CBLKID25_8();" 'CapAll  Hid_D25_8_CapAll_MSG Hid_CapAll
        function chg_CBLKID25_8() {
            //$("#CBLKID25_8"). //<textarea name="CapAll" rows="6" cols="88" id="CapAll" style="width:70%;">1.對大數據分析有興趣之在職勞工。 2.從事資料分析領域工作，欲提升個人工作能力者。</textarea>
            if ($('#CBLKID25_8_0').length == 0) { return; }
            if ($('#CapAll').length == 0) { return; }
            if ($('#Hid_D25_8_CapAll_MSG').length == 0) { return; }
            if ($('#Hid_CapAll').length == 0) { return; }
            // 判斷是否勾選
            if ($('#CBLKID25_8_0').is(':checked')) {
                //console.log('Checkbox is checked'); // 執行其他動作 (例如: 送出表單, 顯示訊息)
                if ($('#Hid_CapAll').val() == '' && $('#CapAll').val() != '' && $('#CapAll').val() != $('#Hid_D25_8_CapAll_MSG').val()) {
                    $('#Hid_CapAll').val($('#CapAll').val());
                }
                $('#CapAll').val($('#Hid_D25_8_CapAll_MSG').val());
                $("#CapAll").prop("readonly", true);
            }
            else {
                //console.log('Checkbox is not checked'); // 執行其他動作
                $("#CapAll").prop("readonly", false);
                $('#CapAll').val($('#Hid_CapAll').val());
                $('#CapAll').text($('#Hid_CapAll').val());
            }
        }

        function open_ICAP1(xBlock1) {
            var ICAPNUM1 = $("#iCAPNUM").val();
            if (ICAPNUM1 == "") {
                alert('請輸入 iCAP標章證號!');
                return;
            }
            // 取得螢幕寬度
            const screenWidth = window.screen.width;
            //const screenHeight = window.screen.height;
            const screenHeight = window.innerHeight;
            // 計算視窗寬度（通常會留一些邊距）
            let winWidth = ((screenWidth * 0.7) - 20).toFixed(3); // 減去20像素作為邊距
            let winHeight = ((screenHeight * 0.7) - 20).toFixed(3); // 減去20像素作為邊距
            //,location=0,status=0,menubar=0,scrollbars=1,resizable=1
            window.open(('TC_03_ICAP.aspx?ICAPNUM1=' + ICAPNUM1), xBlock1, 'width=${winWidth},height=${winHeight}');
        }

        function Layer_change(index) {
            //先關閉所有選項
            showHide(0);
            GetPackageName54();
            chg_FIXSUMCOST();

            var xLayerState = document.getElementById('LayerState');
            var cst_max_tab_cnt = 9;
            for (i = 1; i <= cst_max_tab_cnt; i++) {
                var oTABLELAY = document.getElementById('TableLay' + i);
                if (oTABLELAY) { oTABLELAY.style.display = "none"; }
                //mybox.style.backgroundColor = '#9999ff';
                var oMYBOX = document.getElementById('box' + i);
                if (oMYBOX) { oMYBOX.className = ""; }
            }
            if (index == '') {
                if (xLayerState) { index = xLayerState.value; }
                //if ($('#LayerState')) { index = $('#LayerState').val(); } //document.getElementById('LayerState').value;
            }
            //mybox2.style.backgroundColor = '#006699';
            var BtnSAVE2 = document.getElementById('BtnSAVE2');
            var oTABLELAYidx = document.getElementById('TableLay' + index);
            if (oTABLELAYidx) { oTABLELAYidx.style.display = ""; }
            var oMYBOXidx = document.getElementById('box' + index);
            if (oMYBOXidx) { oMYBOXidx.style.display = ""; }
            oMYBOXidx.className = "active";
            if (xLayerState) { xLayerState.value = index; }
            var flag_TableLay_9 = (xLayerState && index == "9");//false;
            //if (xLayerState && index == "9") { flag_TableLay_9 = true; }
            if (BtnSAVE2) {
                BtnSAVE2.style.display = "none";
                if (flag_TableLay_9) { BtnSAVE2.style.display = ""; }
            }
            //開班計劃表資料維護//re-binding autocomplete input event
            if (index == "9") { autocomplete_init(); }

            if (document.body) { window.scroll(0, document.body.scrollHeight); }
            if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        }

        function doGetPTID(ddlplaceid, hidAddressPtid, ddlplaceid2) {
            console.log("doGetPTID center:" + $("#center").val());
            console.log("doGetPTID ComidValue:" + $("#ComidValue").val());
            console.log("doGetPTID Hid_ComIDNO:" + $("#Hid_ComIDNO").val());
            if ($("#center").val() == "" || ($("#ComidValue").val() == "" && $("#Hid_ComIDNO").val() == "")) {
                alert('請先選擇【訓練機構】');
                return;
            }
            var ddlobj = document.getElementById(ddlplaceid);
            //var selectedText = ddlobj.options[ddlobj.selectedIndex].innerHTML;
            var selectedValue = ddlobj.value;
            if (ddlplaceid2 != "") {
                var ddlobj2 = document.getElementById(ddlplaceid2);
                var selectedValue2 = ddlobj2.value;
                if (selectedValue2 != "" && selectedValue2 == selectedValue) {
                    ddlobj.value = '';
                    alert('請選擇，不同的場地資料！');
                    return;
                }
            }

            var data = {};
            data.comidno = ($("#Hid_ComIDNO").val() != "") ? $("#Hid_ComIDNO").val() : $("#ComidValue").val();
            data.placeid = selectedValue;//placeid;
            //console.log("doGetPTID ddlplaceid:" + ddlplaceid);,//console.log("doGetPTID hidAddressPtid:" + hidAddressPtid);,//console.log("doGetPTID data:" + JSON.stringify(data));
            var PostURL_1 = "<%=ResolveUrl("~/ajax/GetAddressPTID1.ashx")%>";
            var jqxhr = $.post(PostURL_1, data,
                function (r) {
                    if (r.status == true && r.message == "ok") {
                        console.log("doGetPTID get ok: " + r.data + ", " + JSON.stringify(data));
                        $("#" + hidAddressPtid).val(r.data);
                        $("#Classification1").val("");
                        $("#PTID1").val("");
                        $("#PTID2").val("");
                        console.log("doGetPTID (" + hidAddressPtid + ") ptid value: " + $("#" + hidAddressPtid).val());
                    }
                    else {
                        $("#" + hidAddressPtid).val("");
                        $("#Classification1").val("");
                        $("#PTID1").val("");
                        $("#PTID2").val("");
                        console.warn("doGetPTID get " + r.message + ": " + r.data + ", " + JSON.stringify(data));
                    }
                })
                .fail(function () {
                    $("#" + hidAddressPtid).val("");
                    $("#Classification1").val("");
                    $("#PTID1").val("");
                    $("#PTID2").val("");
                    console.error("doGetPTID get failed: " + JSON.stringify(data));
                });
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
            var PHour = document.getElementById('PHour'); //時數
            var EHour = document.getElementById('EHour'); //技檢訓練時數
            var PCont = document.getElementById('PCont');
            var PCont_PLval = PCont.attributes["placeholder"].value;
            var Classification1 = document.getElementById('Classification1');
            var PTID1 = document.getElementById('PTID1');
            var PTID2 = document.getElementById('PTID2');
            var OLessonTeah1Value = document.getElementById('OLessonTeah1Value');
            var iTPERIOD28 = 0;
            if (TPERIOD28_1.checked) { iTPERIOD28 += 1 };
            if (TPERIOD28_2.checked) { iTPERIOD28 += 1 };
            if (TPERIOD28_3.checked) { iTPERIOD28 += 1 };
            if (iTPERIOD28 == 0) { msg += '授課時段:早上、下午、晚上 至少要設定其中一項\n'; }
            if (iTPERIOD28 > 1) { msg += '授課時段:早上、下午、晚上為單選\n'; }
            if (STrainDate) {
                if (STrainDate.value == '') msg += '【日期】不可為空，請選擇\n';
                if (STrainDate.value != '' && !checkDate(STrainDate.value)) { msg += '【日期】格式有誤應為日期格式\n'; }
            }
            var H1val = parseInt(getValue(objH1), 10);/*授課時間 時數*/
            var M1val = parseInt(getValue(objM1), 10);
            var H2val = parseInt(getValue(objH2), 10);/*授課時間 時數*/
            var M2val = parseInt(getValue(objM2), 10);
            if (H1val > H2val) { msg += '【授課時間】起始時間不得大於結束時間\n'; }
            if (H1val == H2val) { if (M1val >= M2val) { msg += '【授課時間】起始時間不得大於等於結束時間\n'; } }
            /*
            if (PHour.value == '') msg += '【時數】不可為空，請輸入\n';
            else if (!isUnsignedInt(PHour.value) && !isFloat(PHour.value)) msg += '【時數】必須為數字\n';//'【時數】必須為整數數字\n';
            else if (!(parseInt(PHour.value, 10) <= 4 && parseInt(PHour.value, 10) >= 1)) msg += '【時數】必須為小於4，大於0\n';
            */
            if (PHour.value == '') msg += '【時數】不可為空，請輸入\n';
            else if (!isUnsignedInt(PHour.value)) msg += '【時數】必須為整數數字\n';
            else if (!(parseInt(PHour.value, 10) <= 4 && parseInt(PHour.value, 10) >= 1)) msg += '【時數】必須為小於4，大於0\n';

            //debugger;
            if (msg == '' && PHour.value != '' && EHour != undefined && EHour.value != '') {
                var fgOK1 = isUnsignedInt(EHour.value);//整數數字
                var fgOK2 = isFloat1(EHour.value);//允許小數點一位
                //'技檢訓練時數 產業人才投資方案(充飛不用)
                //'1.【技檢訓練時數】需<=該堂課【時數】，可允許小數點一位
                //'2.目前僅訓練業別為【[03-01]傳統及民俗復健整復課程】時需要填寫，但是當尚未儲存時應該還無法卡控。正式儲存時，檢核若為03-1才存欄位，否清空。
                //'3.權限：跟其他欄位一樣。'訓練單位可填寫，但送審後鎖住不可修改。送審後分署可修改。'3.訓練班節計畫表加上【技檢時數】欄位顯示。
                //'4.註記：'調整班級變更申請、班級變更審核、結訓證書等功能 '結訓證書上顯示： 符合申請技檢訓練時數
                if (!fgOK1 && !fgOK2) { msg += '【技檢訓練時數】必須為整數數字 或 小數點1位數字\n'; }
                else if (parseFloat(EHour.value).toFixed(1) > parseFloat(PHour.value).toFixed(1)) { msg += '【技檢訓練時數】必須小於等於該堂課【時數】\n'; }
            }

            var blnChkErrPCont = true; //若有檢核錯誤，即為false;
            if (PCont.value == '') { msg += '【課程進度／內容】不可為空，請輸入\n'; blnChkErrPCont = false; }
            if (blnChkErrPCont && PCont_PLval != '' && PCont.value == PCont_PLval) { msg += '【課程進度／內容】不可為空，請輸入\n'; }

            if (Classification1.selectedIndex == 0) { msg += '【學／術科】不可為空，請選擇\n'; }
            if (msg == '' && Classification1.value == '1') {
                if (PTID1.selectedIndex == 0 && PTID1.value == "") { msg += '【學科:上課地點】不可為空，請選擇\n'; }
            }
            if (msg == '' && Classification1.value == '2') {
                if (PTID2.selectedIndex == 0 && PTID2.value == "") { msg += '【術科:上課地點】不可為空，請選擇\n'; }
            }
            if (OLessonTeah1Value.value == '') { msg += '【任課教師】不可為空，請選擇\n'; }
            if (msg != '') {
                alert(msg);
                return false;
            }

            if (STrainDate) {
                var arySD1 = STrainDate.value.split("/");
                if (arySD1.length > 1 && arySD1[1].length == 1) { arySD1[1] = "0" + arySD1[1]; }
                if (arySD1.length > 2 && arySD1[2].length == 1) { arySD1[2] = "0" + arySD1[2]; }
                if (arySD1.length > 2) {
                    var SD_YMD = arySD1[0] + "/" + arySD1[1] + "/" + arySD1[2];
                    STrainDate.value = SD_YMD;
                }
            }
            //debugger;
            var xAddValue = "";
            var Hid_ChkTDescH1 = document.getElementById("Hid_ChkTDescH1");
            var strH1val = "" + H1val; if (H1val < 10) { strH1val = "0" + H1val; }
            xAddValue = "SD:" + SD_YMD + "/H1:" + strH1val;
            if (Hid_ChkTDescH1.value != "") {
                if (Hid_ChkTDescH1.value.indexOf(xAddValue) != -1) {
                    msg += "此上課日期輸入之授課時間(起)組合，同一天之授課時間重疊!\n";
                    alert(msg);
                    return false;
                }
            }
            if (Hid_ChkTDescH1.value != "") Hid_ChkTDescH1.value += ","
            Hid_ChkTDescH1.value += xAddValue

            var Hid_ChkTDescH2 = document.getElementById("Hid_ChkTDescH2");
            var strH2val = "" + H2val; if (H2val < 10) { strH2val = "0" + H2val; }
            xAddValue = "SD:" + SD_YMD + "/H2:" + strH2val;
            if (Hid_ChkTDescH2.value != "") {
                if (Hid_ChkTDescH2.value.indexOf(xAddValue) != -1) {
                    msg += "此上課日期輸入之授課時間(迄)組合，同一天之授課時間重疊!\n";
                    alert(msg);
                    return false;
                }
            }
            if (Hid_ChkTDescH2.value != "") Hid_ChkTDescH2.value += ","
            Hid_ChkTDescH2.value += xAddValue

            var Hid_ChkTDescH3 = document.getElementById("Hid_ChkTDescH3");
            var flag_double = Chk_doubleDAY(SD_YMD, H1val, H2val, Hid_ChkTDescH3.value);
            if (flag_double) {
                msg += "此上課日期輸入之授課時間(疊)組合，同一天之授課時間重疊!\n";
                alert(msg);
                return false;
            }

            //若沒有重複，就組合!!,//var Hid_ChkTEACHHOURS1 = document.getElementById("Hid_ChkTEACHHOURS1");,//if (Hid_ChkTEACHHOURS1.value == "Y") {,//    msg += "同一位師資授課時數已超過54小時，請確認是否為特殊情況。\n";,//    alert(msg);,//}
        }

        //檢核是否有重複狀況!!
        function Chk_doubleDAY(strSDYMD, iH1val, iH2val, strALLTDescH3) {
            if (strALLTDescH3 == "") { return false; }
            aryTDescH3 = strALLTDescH3.split(";");//取得上課日節數資訊(當以分號；分隔)
            if (aryTDescH3.length == 0) { return false; }
            if (strALLTDescH3.indexOf(strSDYMD) == -1) { return false; }
            for (i = 0; i < aryTDescH3.length; i++) {
                if (aryTDescH3[i].indexOf(strSDYMD) != -1) {
                    var strLesson = aryTDescH3[i].split(":")[1];//分離當日節數資訊(當以冒號：分隔)
                    if (strLesson == "") { return false; }//無資訊
                    aryLesson = strLesson.split(",");//取得節數資訊(當以逗號，分隔)
                    if (aryLesson.length == 0) { return false; }
                    for (j = iH1val; j < iH2val; j++) {
                        var tmpJ = "" + j;
                        if (aryLesson.indexOf(tmpJ) != -1) { return true; /*重複*/ }
                        /*,for (k = 0; k < aryLesson.length; k++) {,if (parseInt(aryLesson[k], 10)==j) { return true; },},*/
                        //text += "The number is " + i + "<br>";
                    }
                    return false;//沒有重複狀況
                }
            }
            return false;//沒有重複狀況
        }

        function chg_FIXSUMCOST() {
            //固定費用總額單一人時成本 //http://i-yow.blogspot.tw/2010/08/javascript.html //http://www.eion.com.tw/Blogger/?Pid=1173
            //toFixed(N) 取小數第N位 , parseFloat(num) 去掉小數點尾數零 //document.getElementById('tdFIXExceeDesc').className = 'newclass';
            var romsg2 = "鎖(固定費用總額已輸入)/清除總額可修改-計畫階段";
            var THours = document.getElementById('THours');
            var TNum = document.getElementById('TNum');
            var hTPlanID54 = document.getElementById('hTPlanID54');
            var Hid_TotalCost1 = document.getElementById('Hid_TotalCost1');
            var DefGovCost = document.getElementById('DefGovCost');
            var DefStdCost = document.getElementById('DefStdCost');
            var tTrainingCost = document.getElementById('tTrainingCost');//訓練費用總額：新臺幣
            var tPerPersonCost = document.getElementById('tPerPersonCost');//每人總訓練費用：新臺幣
            //'固定費用人時成本上限為140元  'ACTHUMCOST > 140  hid_MAX_iACTHUMCOST
            var vMaxiACTHUMCOST = Number(document.getElementById('hid_MAX_iACTHUMCOST').value);

            Hid_TotalCost1.value = '';
            DefGovCost.value = '';
            DefStdCost.value = '';
            var LabMsgFIX1in = document.getElementById('LabMsgFIX1');
            LabMsgFIX1in.innerHTML = '';
            var FIXSUMCOST = document.getElementById('FIXSUMCOST');//固定費用總額
            var ACTHUMCOST = document.getElementById('ACTHUMCOST');//材料費用總額
            //超出人時成本原因說明
            var hid_FIXExceeDesc = document.getElementById('hid_FIXExceeDesc');
            var FIXExceeDesc = document.getElementById('FIXExceeDesc');
            var tdFIXExceeDesc = document.getElementById('tdFIXExceeDesc');
            //超出材料費比率上限原因說明
            var hid_METExceeDesc = document.getElementById('hid_METExceeDesc');
            var METExceeDesc = document.getElementById('METExceeDesc');
            var tdMETExceeDesc = document.getElementById('tdMETExceeDesc');

            var Hid_THours = document.getElementById('Hid_THours');
            var Hid_TNum = document.getElementById('Hid_TNum');
            //var Hid_FIXExceeDesc = document.getElementById('Hid_FIXExceeDesc');
            var METSUMCOST = document.getElementById('METSUMCOST');
            var METCOSTPER = document.getElementById('METCOSTPER');
            if (tdFIXExceeDesc) { tdFIXExceeDesc.className = 'bluecol'; }
            if (tdMETExceeDesc) { tdMETExceeDesc.className = 'bluecol'; }
            if (ACTHUMCOST) { ACTHUMCOST.value = ''; }
            if (FIXSUMCOST.value == '') {
                INPUT_readOnlyU(THours);
                INPUT_readOnlyU(TNum);
                ACTHUMCOST.value = '';
                METSUMCOST.value = '';
                METCOSTPER.value = '';
                tTrainingCost.value = '';
                tPerPersonCost.value = '';
                LabMsgFIX1in.innerHTML = '(請先確認訓練時數、人數 或 固定費用總額)!';
            }
            if (METSUMCOST.value == '') { METCOSTPER.value = ''; }
            if (!(ACTHUMCOST && isInt(THours.value) && isInt(TNum.value) && isInt(FIXSUMCOST.value))) {
                if (FIXSUMCOST.value != '') {
                    LabMsgFIX1in.innerHTML = '缺少相關參數無法計算-固定費用總額(請確認訓練時數、人數)!!!';
                    FIXSUMCOST.value = '';
                    METSUMCOST.value = '';
                    METCOSTPER.value = '';
                    return false;
                }
                return "";
            }
            if (ACTHUMCOST && isInt(THours.value) && isInt(TNum.value) && isInt(FIXSUMCOST.value)) {
                ACTHUMCOST.value = (Number(FIXSUMCOST.value) / Number(TNum.value) / Number(THours.value));
                ACTHUMCOST.value = parseFloat(ACTHUMCOST.value).toFixed(2);

                if (tdFIXExceeDesc && ACTHUMCOST.value != '') {
                    tdFIXExceeDesc.className = (Number(ACTHUMCOST.value) > vMaxiACTHUMCOST) ? 'bluecol_need' : 'bluecol';
                    FIXExceeDesc.disabled = (Number(ACTHUMCOST.value) > vMaxiACTHUMCOST) ? false : true;
                    if (FIXExceeDesc.value != "") { hid_FIXExceeDesc.value = FIXExceeDesc.value; }
                    FIXExceeDesc.value = (Number(ACTHUMCOST.value) > vMaxiACTHUMCOST) ? hid_FIXExceeDesc.value : '';
                }

                INPUT_readOnly2(THours, romsg2); //.readOnly = true;
                INPUT_readOnly2(TNum, romsg2); //.readOnly = true;
            }

            if (THours.value != '') Hid_THours.value = THours.value;
            if (TNum.value != '') Hid_TNum.value = TNum.value;
            //if (FIXExceeDesc.value != '') Hid_FIXExceeDesc.value = FIXExceeDesc.value;
            var Hid_PERC100 = document.getElementById('Hid_PERC100');
            var iPERC100 = Number(Hid_PERC100.value); //超出材料費比率上限 材料費占比
            if (!(METCOSTPER && Hid_PERC100.value != '' && ACTHUMCOST.value != '' && isInt(FIXSUMCOST.value) && isInt(METSUMCOST.value))) {
                if (METSUMCOST.value != '') {
                    LabMsgFIX1in.innerHTML = '缺少相關參數無法計算 材料費用總額(請確認固定費用總額、訓練業別)!!!';
                    METSUMCOST.value = '';
                    return false;
                }
                return "";
            }
            //材料費編列比上限
            if (isNaN(Hid_PERC100.value)) {
                LabMsgFIX1in.innerHTML = '缺少相關參數無法計算 材料費用總額 材料費編列比上限(請確認訓練業別)!!!';
                return false;
            }

            //var t_msg1 = "材料費編列比上限:" + Hid_PERC100.value + "%";
            var t_msg1 = "材料費編列比上限:" + parseFloat(Hid_PERC100.value).toFixed(1) + "%";
            var lab_iPERC100 = document.getElementById('lab_iPERC100');
            METExceeDesc.title = t_msg1;
            lab_iPERC100.title = t_msg1;

            METCOSTPER.value = (Number(METSUMCOST.value) / Number(FIXSUMCOST.value) * 100);
            METCOSTPER.value = parseFloat(METCOSTPER.value).toFixed(2);
            if (tdMETExceeDesc && METCOSTPER.value != '' && !isNaN(Hid_PERC100.value)) {
                tdMETExceeDesc.className = (Number(METCOSTPER.value) > iPERC100) ? 'bluecol_need' : 'bluecol';
                METExceeDesc.disabled = (Number(METCOSTPER.value) > iPERC100) ? false : true;
                if (METExceeDesc.value != "") { hid_METExceeDesc.value = METExceeDesc.value; }
                METExceeDesc.value = (Number(METCOSTPER.value) > iPERC100) ? hid_METExceeDesc.value : '';
            }

            var iTotal = 0;
            if (FIXSUMCOST.value != '' && METSUMCOST.value != '') {
                iTotal = (Number(FIXSUMCOST.value) + Number(METSUMCOST.value));
            }
            Hid_TotalCost1.value = parseInt(iTotal, 10);
            //(1)訓練費用總額：訓練費用頁籤的固定費用總額+材料費用總額。
            tTrainingCost.value = Hid_TotalCost1.value;
            //(2)每人總訓練費用：訓練費用總額/訓練人數，計算顯示無條件捨去至小數點第1位。
            if ((tPerPersonCost && isInt(TNum.value) && isInt(tTrainingCost.value))) {
                tPerPersonCost.value = Number(tTrainingCost.value) / Number(TNum.value);
                tPerPersonCost.value = parseFloat(Math.floor(tPerPersonCost.value * 100) / 100).toFixed(2);
            }
            DefGovCost.value = parseFloat(iTotal * 0.8).toFixed();
            DefStdCost.value = parseFloat(iTotal - (iTotal * 0.8)).toFixed();
            if (hTPlanID54 != null && hTPlanID54.value == '1') {
                DefGovCost.value = iTotal;
                DefStdCost.value = 0;
            }
        }

        function openPDF20171226() {
            var Hid_PDF20171226 = document.getElementById('Hid_PDF20171226');
            if (Hid_PDF20171226 && Hid_PDF20171226.value != '') {
                window.open(Hid_PDF20171226.value, 'PDF20171226', 'location=0,status=0,menubar=0,scrollbars=0,resizable=0');
            }
            return false;
        }

        function showHide(type) {
            if ($('#nxlayer_01') == null) { return; }
            $('#nxlayer_01').hide();
            if (type == 1 && $('#nxlayer_01').is(":hidden")) {
                $('#nxlayer_01').show();
                $('#SciPlaceID').hide();
                $('#TechPlaceID').hide();
                $('#SciPlaceID2').hide();
                $('#TechPlaceID2').hide();
                //$('#Taddress2').hide(); //$('#Taddress3').hide();
            } else {
                $('#SciPlaceID').show();
                $('#TechPlaceID').show();
                $('#SciPlaceID2').show();
                $('#TechPlaceID2').show();
                //$('#Taddress2').show(); //$('#Taddress3').show();
            }
            //document.all
        }

        //包班種類 '(充飛使用)包班種類(PackageType) 1:非包班/2:企業包班/3:聯合企業包班 
        function GetPackageName() {
            var PackageName = document.getElementById('PackageName');
            if (!PackageName) { return; }
            var PackageType = document.getElementsByName('PackageType');
            var cst_pt1 = 0;
            var cst_pt2 = 1;
            var cst_pt3 = 2;
            if (PackageType.length > 3) {
                cst_pt1 = 1;
                cst_pt2 = 2;
                cst_pt3 = 3;
            }
            if (PackageType.length < 3) { return; }
            //(非包班)
            if (PackageType[cst_pt1].checked == true) { document.form1.PackageName.value = ''; }
            if (PackageType[cst_pt2].checked == true) { document.form1.PackageName.value = '(企業包班)'; }
            if (PackageType[cst_pt3].checked == true) { document.form1.PackageName.value = '(聯合企業包班)'; }
        }

        function GetPackageName54() {
            var PackageName = document.getElementById('PackageName');
            if (!PackageName) { return; }
            var msg = '';
            var btnAddBusPackage = document.getElementById('btnAddBusPackage');
            var hTPlanID54 = document.getElementById('hTPlanID54'); //確認是否為 充電起飛計畫
            if (!btnAddBusPackage) { return; }
            if (!hTPlanID54) { return; }

            var PackageType = document.getElementsByName('PackageType');
            var cst_pt1 = 0;
            var cst_pt2 = 1;
            var cst_pt3 = 2;
            if (PackageType.length > 3) {
                cst_pt1 = 1; //cst_pt
                cst_pt2 = 2;
                cst_pt3 = 3;
            }
            if (PackageType.length < 3) { return; }

            //確認是否為 充電起飛計畫
            if (hTPlanID54 != null && hTPlanID54.value == '1') {
                if (PackageType[cst_pt1].checked == true) {
                    PackageType[cst_pt1].checked = false;
                    PackageType[cst_pt1].disabled = true;
                    msg += '充電起飛計畫不可選擇非包班!!\n'; //(非包班)
                }
                if (PackageType[cst_pt2].checked == true) { document.form1.PackageName.value = '(企業包班)'; }
                if (PackageType[cst_pt3].checked == true) { document.form1.PackageName.value = '(聯合企業包班)'; }
                //充電起飛計畫 且 聯合企業包班 才可顯示 計畫包班事業單位				
                if (btnAddBusPackage) { btnAddBusPackage.style.display = ''; }

                $('#Datagrid4headTable').hide();
                $('#Datagrid4Table').hide();
                //if (document.getElementById('Datagrid4headTable')) { document.all.Datagrid4headTable.style.visibility = 'hidden';}
                //if (document.getElementById('Datagrid4Table')) { document.all.Datagrid4Table.style.visibility = 'hidden';}
                //充電起飛計畫 且 聯合企業包班 才可顯示 計畫包班事業單位
                if (PackageType[cst_pt3].checked == true) {
                    $('#Datagrid4headTable').show();
                    $('#Datagrid4Table').show();
                    //if (document.getElementById('Datagrid4headTable')) { document.all.Datagrid4headTable.style.visibility = 'visible'; }
                    //if (document.getElementById('Datagrid4Table')) { document.all.Datagrid4Table.style.visibility = 'visible';}
                }
                //充電起飛計畫 且 企業包班 只顯示 計畫包班事業單位 的抬頭
                if (PackageType[cst_pt2].checked == true) {
                    $('#Datagrid4headTable').show();
                    //if (document.getElementById('Datagrid4headTable')) { document.all.Datagrid4headTable.style.visibility = 'visible'; }
                    if (btnAddBusPackage) { btnAddBusPackage.style.display = 'none'; }
                }
                if (msg != '') {
                    alert(msg);
                    return false;
                }
            }
        }

        function GetPointName() {
            var PointName = document.getElementById('PointName');
            if (!PointName) { return; }

            var PointType = document.getElementsByName('PointType');
            if (!PointType) { return; }
            //cst_pt
            var cst_pt1 = (PointType.length > 3) ? 1 : 0;
            var cst_pt2 = (PointType.length > 3) ? 2 : 1;
            var cst_pt3 = (PointType.length > 3) ? 3 : 2;
            if (PointType.length < 3) { return; }
            if (PointType[cst_pt1].checked == true) { PointName.value = '學士學分班'; }
            if (PointType[cst_pt2].checked == true) { PointName.value = '碩士學分班'; }
            if (PointType[cst_pt3].checked == true) { PointName.value = '博士學分班'; }
            return true;
        }

        function GetAppStageMSG1() {
            var sp_AppStage_1 = document.getElementById('sp_AppStage_1');
            if (!sp_AppStage_1) { return false; }
            sp_AppStage_1.style.display = "none";
            var sp_AppStage_2 = document.getElementById('sp_AppStage_2');
            if (!sp_AppStage_2) { return false; }
            sp_AppStage_2.style.display = "none";

            var obj = document.getElementsByName("rbl_AppStage");
            for (var i = 0; i < obj.length; i++) {
                if (obj[i].checked && obj[i].value == "1") {
                    sp_AppStage_1.style.display = "";
                }
                else if (obj[i].checked && obj[i].value == "2") {
                    sp_AppStage_2.style.display = "";
                }
            }
            return true;
        }

        function Get_GovClass(fieldname) {
            var PointYN = '';
            var Radiobuttonlist1 = document.getElementsByName('Radiobuttonlist1');
            var cst_pt1 = 0;
            var cst_pt2 = 1;
            if (Radiobuttonlist1.length > 2) {
                cst_pt1 = 1; //cst_pt
                cst_pt2 = 2;
            }
            if (Radiobuttonlist1.length < 2) { return; }

            var TB_career_id = document.getElementById("TB_career_id");
            var btu_sel = document.getElementById("btu_sel");
            var jobValue = document.getElementById("jobValue");
            var trainValue = document.getElementById("trainValue");
            if (Radiobuttonlist1[cst_pt1].checked == true) { PointYN = 'Y'; }
            if (Radiobuttonlist1[cst_pt2].checked == true) { PointYN = 'N'; }
            if (TB_career_id.value == '') {
                alert('請先輸入訓練業別');
                return false;
            }
            btu_sel.removeAttribute("title");
            btu_sel.disabled = false;
            btu_sel.title = "選擇經費分類代碼，不可再選訓練職類";  //不開放選擇訓練職類
            if (jobValue.value != '' && PointYN == 'N') {
                btu_sel.disabled = true;
            }
            wopen('../../common/GovClass.aspx?fieldname=' + fieldname + "&trainValue=" + trainValue.value + "&jobValue=" + jobValue.value + "&PointYN=" + PointYN, 'GovClass3', 930, 350, 1);
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

        function showPTID(obj, ptid1, ptid2) {
            var objvalue = document.getElementById(obj).value;
            document.getElementById(ptid1).style.display = "none";
            document.getElementById(ptid2).style.display = ""; //"inline";
            if (objvalue == 1) {
                document.getElementById(ptid1).style.display = ""; //"inline";
                document.getElementById(ptid2).style.display = "none";
            }
            if (document.body) { window.scroll(0, document.body.scrollHeight); }
            if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        }

        /*function showAPPSTAGE(obj) {,var trObj = document.getElementById("tr_PolicyPreVal");,if (!trObj) { return; },var objAPPSTAGE = document.getElementById(obj);,if (!objAPPSTAGE) { return; }
         * ,var vAPPSTAGE = getValue(objAPPSTAGE);,trObj.style.display = "none";,"inline"; debugger;,if (vAPPSTAGE == "3") { trObj.style.display = ""; },if (document.body) { window.scroll(0, document.body.scrollHeight); }
         * ,if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); },}*/

        /* function noneAPPSTAGE(obj) {,var trObj = document.getElementById("tr_PolicyPreVal");,if (!trObj) { return; },trObj.style.display = "none";,}*/

        //Radiobuttonlist1
        function showCostType(obj1) {
            var PointName = document.getElementById('PointName');
            var PointType = document.getElementsByName('PointType');
            var CredPoint = document.getElementById('CredPoint');
            var Hid_CredPoint = document.getElementById('Hid_CredPoint');
            var cst_pt1 = 0;
            var cst_pt2 = 1;
            var cst_pt3 = 2;
            //cst_pt
            if (PointType.length > 3) { cst_pt1 = 1; cst_pt2 = 2; cst_pt3 = 3; }
            if (PointType.length < 3) { return; }

            //JS DEBUG 依2/8 EMAIL問題： 產業人才投資方案經費欄位資訊錯誤，為避免顯示外網資訊錯誤，惠請協助查明並修復
            var v_obj1 = getRadioValue(document.getElementsByName(obj1));
            //'Y'學分班  //非學分班
            document.getElementById('tdCredPoint').className = (v_obj1 == 'Y') ? "bluecol_need" : "bluecol";
            document.getElementById('PointType_TR').style.display = (v_obj1 == 'Y') ? "" : "none"; //學分班種類
            document.getElementById('PointName').style.display = (v_obj1 == 'Y') ? "" : "none"; //"inline";//學分班種類
            PointName.value = (v_obj1 == 'Y') ? PointName.value : '';

            if (v_obj1 != 'Y') {
                //非學分班
                PointType[cst_pt1].checked = false;
                PointType[cst_pt2].checked = false;
                PointType[cst_pt3].checked = false;
                PointName.value = '';
                if (CredPoint.value != '') { Hid_CredPoint.value = CredPoint.value; }
            }
            CredPoint.disabled = (v_obj1 == 'Y') ? false : true;
            CredPoint.value = (v_obj1 == 'Y') ? Hid_CredPoint.value : '';
            CredPoint.title = (CredPoint.disabled) ? "選擇「非學分班」，則於【學分數】欄位會顯示反灰不能填寫" : "";
            if (document.body) { window.scroll(0, document.body.scrollHeight); }
            if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        }

        //檢查上課時間
        function CheckAddTime() {
            var msg = '';
            var Hid_CheckAddTime = document.getElementById('Hid_CheckAddTime');
            var Weeks = document.getElementById('Weeks');
            var txtTimes = document.getElementById('txtTimes');
            if (Weeks.selectedIndex == 0) msg += '請選擇星期\n';
            if (txtTimes.value == '') msg += '請輸入上課時間\n';
            if (msg != '') {
                alert(msg);
                return false;
            }

            var xAddValue = "Ws:" + Weeks.selectedIndex + "/Ts:" + txtTimes.value;
            if (Hid_CheckAddTime.value != "") {
                if (Hid_CheckAddTime.value.indexOf(xAddValue) != -1) {
                    msg += "此上課時間組合，已重複新增!\n";
                    alert(msg);
                    return false;
                }
            }
            if (Hid_CheckAddTime.value != "") Hid_CheckAddTime.value += ","
            Hid_CheckAddTime.value += xAddValue
        }

        function CheckAddBusPackage() {
            var msg = '';
            var txtUname = document.getElementById('txtUname');
            var txtIntaxno = document.getElementById('txtIntaxno');
            var txtUbno = document.getElementById('txtUbno');
            if (txtUname.value == '') msg += '請輸入企業名稱\n';
            //請輸入正確的統一編號。欄位有效時再行驗證,無效時也要進行驗證
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
        /*function check_Cost2() {var msg = '';var CostID2 = document.form1.CostID2;var OPrice2 = document.form1.OPrice2;var Itemage = document.form1.Itemage;if (msg != '') {alert(msg);return false;}}*/

        function check_Cost_Detail(CostID, OPrice, Itemage) {
            var msg = '';
            if (CostID != null && CostID.value == '') msg += '請輸入項目\n';
            if (OPrice != null && OPrice.value == '') msg += '請輸入單價\n';
            if (!OPrice) return msg;
            if (OPrice.value != '' && !isUnsignedInt(OPrice.value) && !isPositiveFloat(OPrice.value)) { msg += '單價必須為數字\n'; }
            if (OPrice.value != '' && isPositiveFloat(OPrice.value) && OPrice.value.indexOf('.') < OPrice.value.length - 3) { msg += '單價只能輸入到小數點第二位\n'; }
            if (Itemage != null && Itemage.value == '') msg += '請輸入計價數量\n';
            if (!Itemage) return msg;
            if (Itemage != '' && !isUnsignedInt(Itemage.value)) msg += '計價數量必須為數字\n';
            return msg;
        }

        //如果選擇其他明細，則顯示TextBox讓使用者輸入
        function ShowOther(sObjName, sObjName2) {
            var ObjName = document.getElementById(sObjName);
            var ObjName2 = document.getElementById(sObjName2);
            ObjName2.style.display = 'none';
            if (dObjName.value == '99') { ObjName2.style.display = ''; }
            if (document.body) { window.scroll(0, document.body.scrollHeight); }
            if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
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

        //檢查日期格式-Melody(2005/3/18)
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
            if (document.form1.DefGovCost.value != '') { flag = true; }
            if (document.form1.DefStdCost.value != '') { flag = true; }
            args.IsValid = flag;
        }

        function CheckDef2(source, args) {
            var flag = true;
            if (!isUnsignedInt(document.getElementById('Total1').innerHTML)) { flag = false; }
            if (!isUnsignedInt(document.getElementById('Total3').innerHTML)) { flag = false; }
            args.IsValid = flag;
        }

        function CheckTDate(sor, args) {
            if (checkDate(document.form1.STDate.value) && checkDate(document.form1.FDDate.value)) {
                var flag = compareDate(document.form1.STDate.value, document.form1.FDDate.value);
                if (flag == 0) args.IsValid = false;
                if (flag == 1) args.IsValid = false;
            }
        }

        function CheckFactMode(sor, args) {
            if (!document.form1.FactMode) return;
            if (!isChecked(document.form1.FactMode)) { args.IsValid = false; }
        }

        function CheckFactModeOther(sor, args) {
            if (!document.form1.FactMode) return;
            if (getRadioValue(document.form1.FactMode) == '99'
                && document.getElementById('FactModeOther').value == '') { args.IsValid = false; }
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
            if (document.form1.DefStdCost.value != '' && !isUnsignedInt(document.form1.DefStdCost.value)) msg += '學員負擔費用必須為數字\n'
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        //計算經費來源
        function CountCostSource() {
            //97修改中 , 旗標-看所有的值是否已經輸入
            if (document.form1.TNum.value != '' && isUnsignedInt(document.form1.TNum.value)) {
                document.getElementById('TNum1').innerHTML = document.form1.TNum.value;
                document.getElementById('TNum3').innerHTML = document.form1.TNum.value;
                if (isPositiveFloat(document.form1.DefGovCost.value) || isUnsignedInt(document.form1.DefGovCost.value)) {
                    var vTotal1 = parseInt(document.form1.DefGovCost.value, 10) / parseInt(document.form1.TNum.value, 10);
                    if (!isInt(vTotal1)) { vTotal1 = Number(vTotal1).toFixed(3); }
                    else { vTotal1 = Number(vTotal1).toFixed(); }
                    document.getElementById('Total1').innerHTML = vTotal1;
                }
                else { document.getElementById('Total1').innerHTML = '0'; }
                if (isPositiveFloat(document.form1.DefStdCost.value) || isUnsignedInt(document.form1.DefStdCost.value)) {
                    var vTotal3 = parseInt(document.form1.DefStdCost.value, 10) / parseInt(document.form1.TNum.value, 10);
                    if (!isInt(vTotal3)) { vTotal3 = Number(vTotal3).toFixed(3); }
                    else { vTotal3 = Number(vTotal3).toFixed(); }
                    document.getElementById('Total3').innerHTML = vTotal3;
                }
                else { document.getElementById('Total3').innerHTML = '0'; }
            }
            else {
                document.getElementById('TNum1').innerHTML = '(尚未設定人數)';
                document.getElementById('TNum3').innerHTML = '(尚未設定人數)';
                document.getElementById('Total1').innerHTML = '0';
                document.getElementById('Total3').innerHTML = '0';
            }
        }

        function check_style() {
            if (form1.STDate && form1.STDate.disabled && form1.date1) {
                if (form1.date1) {
                    form1.date1.style.cursor = "";
                    form1.date1.onclick = null;
                }
            }
            if (form1.FDDate && form1.FDDate.disabled && form1.date2) {
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

        //限制TextBox在MultiLine時的字數 MaxLength
        function checkTextLength(obj, Mlong) {
            var maxlength = new Number(Mlong); // Change number to your max length.
            if (obj.value.length > maxlength) {
                obj.value = obj.value.substring(0, maxlength);
                alert("限欄位長度不能大於" + maxlength + "個字元(含空白字元)，超出字元將自動截斷");
            }
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
        .red-style1 { color: #FF0000; }
    </style>
</head>
<body onload="check_style();">
    <form id="form1" method="post" runat="server">
        <%--<asp:ScriptManager ID="ScriptManager1" runat="server" ScriptMode="Release" LoadScriptsBeforeUI="false" EnableScriptGlobalization="true" EnableScriptLocalization="true" />--%>
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
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
                    <table border="0" cellspacing="1" cellpadding="0" width="100%" class="table_sch">
                        <tr>
                            <td class="bluecol_need" width="16%">訓練機構 </td>
                            <td class="whitecol" width="34%">
                                <asp:TextBox ID="center" runat="server" Width="80%" onfocus="this.blur()"></asp:TextBox>
                                <input id="Org" value="..." type="button" name="Org" runat="server" class="button_b_Mini" />
                                <asp:Button Style="display: none" ID="Button28" runat="server" Text="機構資訊(隱藏)" CausesValidation="False"></asp:Button>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="ComidValue" type="hidden" name="ComidValue" runat="server" /><br />
                                <span style="position: absolute; display: none" id="HistoryList2">
                                    <asp:Table ID="HistoryRID" runat="server" Width="310px"></asp:Table>
                                </span>
                                <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="center" Display="None" ErrorMessage="請選擇訓練機構"></asp:RequiredFieldValidator>
                            </td>
                            <td id="Table1_Email" runat="server" width="50%" class="whitecol" colspan="2">
                                <span>是否要Email線上報名資料，EMail</span>
                                <asp:TextBox ID="EMail" runat="server" Columns="30" Width="70%"></asp:TextBox>
                                <asp:RegularExpressionValidator ID="check1" runat="server" ControlToValidate="EMail" Display="None" ErrorMessage="E_Mail輸入錯誤" ValidationExpression="\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"></asp:RegularExpressionValidator>
                            </td>
                        </tr>
                        <tr id="trainTR" runat="server">
                            <td class="bluecol_need" width="16%">
                                <asp:Label ID="LabTMID" runat="server">訓練職類</asp:Label></td>
                            <td class="whitecol" width="34%">
                                <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Width="80%"></asp:TextBox>
                                <input id="btu_sel" onclick="openTrain2(document.getElementById('trainValue').value);" value="..." type="button" runat="server" class="button_b_Mini" />
                                <asp:RequiredFieldValidator ID="fill1" runat="server" ControlToValidate="TB_career_id" Display="None" ErrorMessage="請選擇訓練業別／職類"></asp:RequiredFieldValidator>
                                <input id="trainValue" type="hidden" name="trainValue" runat="server" />
                                <input id="jobValue" type="hidden" name="jobValue" runat="server" />
                            </td>
                            <td class="bluecol_need" width="16%">
                                <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label></td>
                            <td class="whitecol" width="34%">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="80%"></asp:TextBox>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" value="..." type="button" name="btu_sel2" runat="server" class="button_b_Mini" />
                                <asp:RequiredFieldValidator ID="fill1b" runat="server" ControlToValidate="txtCJOB_NAME" Display="None" ErrorMessage="請選擇通俗職類"></asp:RequiredFieldValidator>
                                <input id="cjobValue" type="hidden" name="cjobValue" runat="server" />
                            </td>
                        </tr>
                        <tr id="trTMIDCORRECT" runat="server">
                            <%--訓練業別同意協助重新歸類--%>
                            <td class="bluecol_need">
                                <asp:Label ID="labtTMIDCORRECT" runat="server" ForeColor="Red"></asp:Label></td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="LabmsgTMIDCORRECT" runat="server" ForeColor="Red"></asp:Label><br />
                                <asp:RadioButtonList ID="rblTMIDCORRECT" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="Y">同意</asp:ListItem>
                                    <asp:ListItem Value="N">不同意</asp:ListItem>
                                </asp:RadioButtonList>

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
                    <table class="font" cellspacing="0" cellpadding="0" width="98%">
                        <tr class="newlink">
                            <td id="box1" onclick="Layer_change(1);">目標</td>
                            <td id="box2" onclick="Layer_change(2);" style="display: none;">受訓資格</td>
                            <td id="box3" onclick="Layer_change(3);" style="display: none;">訓練方式</td>
                            <td id="box4" onclick="Layer_change(4);" style="display: none;">課程編配</td>
                            <td id="box5" onclick="Layer_change(5);">班別資料</td>
                            <td id="box6" onclick="Layer_change(6);" runat="server">訓練費用</td>
                            <td id="box7" onclick="Layer_change(7);" runat="server">經費來源</td>
                            <td id="box8" onclick="Layer_change(8);">訓練費用<br />
                                編列說明</td>
                            <td id="box9" onclick="Layer_change(9);">開班計劃表<br />
                                資料維護</td>
                            <%--開班計劃表資料維護--%>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table id="TableLay1" border="0" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol_need" width="16%">單位核心能力介紹 </td>
                            <td class="whitecol" width="84%">
                                <asp:TextBox ID="PlanCause" runat="server" Width="77%" Rows="4" TextMode="MultiLine" MaxLength="100" placeholder="(了解單位與課程是否具關聯性)"></asp:TextBox>
                                <asp:RequiredFieldValidator ID="fill30" runat="server" ControlToValidate="PlanCause" Display="None" ErrorMessage="目標『單位核心能力介紹』為必填欄位"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">知識 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="PurScience" runat="server" Width="77%" Rows="4" TextMode="MultiLine" MaxLength="100" placeholder="(論述學員於課程結束後，所學習之知識內涵)"></asp:TextBox>
                                <asp:RequiredFieldValidator ID="fill2" runat="server" ControlToValidate="PurScience" Display="None" ErrorMessage="目標「知識」為必填欄位"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">技能 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="PurTech" runat="server" Width="77%" Rows="4" TextMode="MultiLine" MaxLength="100" placeholder="(論述學員於課程結束後，所學習之知識內涵)"></asp:TextBox>
                                <asp:RequiredFieldValidator ID="fill3" runat="server" ControlToValidate="PurTech" Display="None" ErrorMessage="目標「技能」為必填欄位"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">學習成效 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="PurMoral" runat="server" Width="77%" Rows="4" TextMode="MultiLine" MaxLength="100" placeholder="(綜合論述學員於學習結束後，可如何運用哪些知識、技能，產出哪些成果)"></asp:TextBox>
                                <asp:RequiredFieldValidator ID="fill4" runat="server" ControlToValidate="PurMoral" Display="None" ErrorMessage="目標「學習成效」為必填欄位"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                    </table>
                    <table id="TableLay2" border="0" cellspacing="1" cellpadding="1" width="100%">
                        <%--<tr>
                            <td class="bluecol_need" width="20%">學歷 </td>
                            <td class="whitecol" width="80%">
                                <asp:DropDownList ID="Degree" runat="server" AppendDataBoundItems="True"></asp:DropDownList>(含以上)
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
                                <asp:TextBox ID="txtAge1" runat="server" Width="40px" MaxLength="2"></asp:TextBox>
                                <asp:Label ID="l_Age2b" runat="server">以上</asp:Label>
                            </td>
                        </tr>--%>
                        <tr id="trOther1" style="display: none">
                            <td class="bluecol_need" width="16%">其他一 </td>
                            <td class="whitecol" width="84%">
                                <asp:TextBox onblur="checkTextLength(this,200)" ID="Other1" onkeyup="checkTextLength(this,200)" runat="server" Width="77%" Rows="3" TextMode="MultiLine" MaxLength="100" ForeColor="Black" BackColor="Silver" onChange="checkTextLength(this,200)"></asp:TextBox></td>
                        </tr>
                        <tr id="trOther2" style="display: none">
                            <td class="bluecol">其他二 </td>
                            <td class="whitecol">
                                <asp:TextBox onblur="checkTextLength(this,200)" ID="Other2" onkeyup="checkTextLength(this,200)" runat="server" Width="77%" Rows="3" TextMode="MultiLine" MaxLength="100" ForeColor="Black" BackColor="Silver" onChange="checkTextLength(this,200)"></asp:TextBox></td>
                        </tr>
                        <tr id="trOther3" style="display: none">
                            <td class="bluecol">其他三 </td>
                            <td class="whitecol">
                                <asp:TextBox onblur="checkTextLength(this,200)" ID="Other3" onkeyup="checkTextLength(this,200)" runat="server" Width="77%" Rows="3" TextMode="MultiLine" MaxLength="100" ForeColor="Black" BackColor="Silver" onChange="checkTextLength(this,200)"></asp:TextBox></td>
                        </tr>
                    </table>
                    <table id="TableLay3" border="0" cellspacing="1" cellpadding="1" width="100%" style="display: none">
                        <tr>
                            <td class="bluecol_need" width="16%">訓練方式 </td>
                            <td class="whitecol" width="84%">
                                <asp:TextBox onblur="checkTextLength(this,200)" ID="TMScience" onkeyup="checkTextLength(this,200)" runat="server" Rows="5" TextMode="MultiLine" MaxLength="200" onChange="checkTextLength(this,200)" Columns="70"></asp:TextBox></td>
                        </tr>
                    </table>
                    <table id="TableLay4" border="0" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td rowspan="2" class="bluecol_need" width="16%">學科 </td>
                            <td rowspan="2" class="whitecol" width="34%">
                                <asp:TextBox ID="SciHours" runat="server" Width="40%" onfocus="this.blur()"></asp:TextBox>小時 </td>
                            <td class="bluecol" width="16%">1. 一般學科 </td>
                            <td class="whitecol" width="34%">
                                <asp:TextBox ID="GenSciHours" runat="server" Width="40%"></asp:TextBox>小時
                                <asp:RequiredFieldValidator ID="fill15" runat="server" ControlToValidate="GenSciHours" Display="None" ErrorMessage="課程編配「一般學科」為必填欄位" Enabled="False"></asp:RequiredFieldValidator>
                                <asp:RegularExpressionValidator ID="check4" runat="server" ControlToValidate="GenSciHours" Display="None" ErrorMessage="課程編配「一般學科」請輸入數字" ValidationExpression="[0-9]{1,4}" Enabled="False"></asp:RegularExpressionValidator>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">2. 專業學科 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="ProSciHours" runat="server" Width="40%"></asp:TextBox>小時
                                <asp:RequiredFieldValidator ID="fill16" runat="server" ControlToValidate="ProSciHours" Display="None" ErrorMessage="課程編配「專業學科」為必填欄位" Enabled="False"></asp:RequiredFieldValidator>
                                <asp:RegularExpressionValidator ID="check5" runat="server" ControlToValidate="ProSciHours" Display="None" ErrorMessage="課程編配「專業學科」請輸入數字" ValidationExpression="[0-9]{1,4}" Enabled="False"></asp:RegularExpressionValidator>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">&nbsp;&nbsp; 術科<span style="color: #FF0000">*</span> </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="ProTechHours" runat="server" Width="18%"></asp:TextBox>小時
                                <asp:RequiredFieldValidator ID="fill17" runat="server" ControlToValidate="ProTechHours" Display="None" ErrorMessage="課程編配「術科」為必填欄位" Enabled="False"></asp:RequiredFieldValidator>
                                <asp:RegularExpressionValidator ID="check6" runat="server" ControlToValidate="ProTechHours" Display="None" ErrorMessage="課程編配「術科」請輸入數字" ValidationExpression="[0-9]{1,4}" Enabled="False"></asp:RegularExpressionValidator>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">&nbsp;&nbsp; 其他時數 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="OtherHours" runat="server" Width="18%"></asp:TextBox>小時
                                <asp:RegularExpressionValidator ID="check7" runat="server" ControlToValidate="OtherHours" Display="None" ErrorMessage="課程編配「其他時數」請輸入數字" ValidationExpression="[0-9]{1,4}" Enabled="False"></asp:RegularExpressionValidator>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">&nbsp;&nbsp; 總計 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="TotalHours" runat="server" Width="18%" onfocus="this.blur()" MaxLength="3"></asp:TextBox>小時 </td>
                        </tr>
                    </table>
                    <table id="TableLay5" border="0" cellspacing="1" cellpadding="1" width="100%">
                        <tr id="tr_AppStage_TP28" runat="server">
                            <td class="bluecol_need" width="16%">申請階段 </td>
                            <td class="whitecol" colspan="3" width="84%"><%--<asp:ListItem Value="1">1:上半年</asp:ListItem><asp:ListItem Value="2">2:下半年</asp:ListItem><asp:ListItem Value="3">3:政策性產業</asp:ListItem>--%>
                                <asp:RadioButtonList ID="rbl_AppStage" runat="server" Width="100%" CssClass="font" CellSpacing="0" CellPadding="0" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="16%">優先排序 </td>
                            <td class="whitecol" width="34%">
                                <asp:TextBox ID="FirstSort" runat="server" Width="40%" MaxLength="3"></asp:TextBox>
                                <asp:RegularExpressionValidator ID="FirstSort_chk1" runat="server" ControlToValidate="FirstSort" Display="None" ErrorMessage="班別資料「優先排序」請輸入數字" ValidationExpression="[0-9]{1,4}"></asp:RegularExpressionValidator>
                                <asp:RequiredFieldValidator ID="FirstSort_chk2" runat="server" ControlToValidate="FirstSort" Display="None" ErrorMessage="班別資料「優先排序」為必填欄位"></asp:RequiredFieldValidator>
                                <a id="lnkToContent" href="#tdContent">(移至課程大綱)</a><%--Table5A--%>
                                <a href="#Table3">(移至上課時間)</a>
                                <a href="#tbPlanDepot">(移至政府政策性產業)</a>
                            </td>
                            <td class="bluecol" width="16%">iCAP標章證號 </td>
                            <td class="whitecol" width="34%">
                                <asp:TextBox ID="iCAPNUM" runat="server" MaxLength="40" Columns="40"></asp:TextBox>
                                <input id="BtnICAPonlineC2" type="button" value="查看iCAP課程資訊" runat="server" class="button_b_M" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">課程種類 </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="Radiobuttonlist1" runat="server" Width="100%" CssClass="font" CellSpacing="0" CellPadding="0" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="Y" Selected="True">學分班</asp:ListItem>
                                    <asp:ListItem Value="N">非學分班</asp:ListItem>
                                </asp:RadioButtonList>
                                <asp:CheckBox ID="IsBusiness" runat="server" Text="企業包班"></asp:CheckBox>
                                <asp:CustomValidator ID="IsBusiness_chk1" runat="server" Display="None" ErrorMessage="班別資料『企業包班名稱』必須填寫" ClientValidationFunction="check_IsBusiness"></asp:CustomValidator>
                            </td>
                            <%--<asp:Label ID="labEnterpriseName" runat="server" Text="企業包班名稱"></asp:Label><asp:TextBox ID="EnterpriseName" runat="server" MaxLength="50" Width="70%"></asp:TextBox>--%>
                            <td class="bluecol">iCAP標章有效期限</td>
                            <td class="whitecol">
                                <asp:TextBox ID="iCAPMARKDATE" runat="server" Width="40%" MaxLength="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" id="date11" alt="" align="top" src="../../images/show-calendar.gif" runat="server" width="30" height="30" onclick="javascript:show_calendar('iCAPMARKDATE','','','CY/MM/DD');" /></span>
                                <span class="red-style1">
                                    <br />
                                    若是【iCAP標章證號】有填寫，【iCAP標章有效期限】不可為空</span></td>
                        </tr>
                        <tr id="PointType_TR" runat="server">
                            <td class="bluecol_need">學分班種類 </td>
                            <td colspan="3">
                                <asp:RadioButtonList ID="PointType" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="1">學士學分班</asp:ListItem>
                                    <asp:ListItem Value="2">碩士學分班</asp:ListItem>
                                    <asp:ListItem Value="3">博士學分班</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="trPackageType" runat="server">
                            <td class="bluecol_need">包班種類 </td>
                            <td colspan="3">
                                <asp:RadioButtonList ID="PackageType" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="1">非包班</asp:ListItem>
                                    <asp:ListItem Value="2">企業包班</asp:ListItem>
                                    <asp:ListItem Value="3">聯合企業包班</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">班別名稱 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="ClassName" runat="server" MaxLength="35" Width="35%"></asp:TextBox>
                                <asp:TextBox ID="PointName" runat="server" Width="18%" onfocus="this.blur()" MaxLength="10"></asp:TextBox>
                                <asp:TextBox ID="PackageName" runat="server" Width="18%" onfocus="this.blur()" MaxLength="10"></asp:TextBox>
                                <input id="Class_Unit" type="hidden" runat="server" />
                                <asp:RequiredFieldValidator ID="fill18" runat="server" ControlToValidate="ClassName" Display="None" ErrorMessage="班別資料「班別名稱」為必填欄位"></asp:RequiredFieldValidator>
                                <input onclick="open_hours()" value="時數迄日換算" type="button" class="button_b_M" />
                                &nbsp;<asp:CheckBox ID="CB_EnvZeroTrain" runat="server" />本課程屬環境部淨零綠領人才培育課程
                            </td>
                        </tr>
                        <tr>
                            <%--<td class="bluecol_need" width="20%">期別(二碼) </td>--%>
                            <td class="bluecol">期別(二碼) </td>
                            <td class="whitecol">
                                <asp:TextBox ID="CyclType" runat="server" Columns="5" MaxLength="2" Width="40%"></asp:TextBox>
                                <%--<asp:RequiredFieldValidator ID="fill28" runat="server" ControlToValidate="CyclType" Display="None" ErrorMessage="班別資料『期別』為必填欄位"></asp:RequiredFieldValidator>--%>
                                <asp:CustomValidator ID="CustomValidator4" runat="server" ControlToValidate="CyclType" Display="None" ErrorMessage="班別資料『期別』必須為大於0的兩位數字" ClientValidationFunction="check_CyclType"></asp:CustomValidator>
                            </td>
                            <td class="bluecol_need">班數 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="ClassCount" runat="server" onfocus="this.blur()" Columns="5" MaxLength="5" Width="40%">1</asp:TextBox>
                                <asp:RequiredFieldValidator ID="fill29" runat="server" ControlToValidate="ClassCount" Display="None" ErrorMessage="請輸入班數"></asp:RequiredFieldValidator>
                                <asp:RegularExpressionValidator ID="RegularExpressionValidator1" runat="server" ControlToValidate="ClassCount" Display="None" ErrorMessage="班數輸入大於0的數字" ValidationExpression="^0*[1-9](\d*$)"></asp:RegularExpressionValidator>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">訓練人數 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TNum" runat="server" Width="40%" MaxLength="7"></asp:TextBox>人
                                <asp:RequiredFieldValidator ID="fill19" runat="server" ControlToValidate="TNum" Display="None" ErrorMessage="班別資料「訓練人數」為必填欄位"></asp:RequiredFieldValidator>
                                <asp:RegularExpressionValidator ID="check8" runat="server" ControlToValidate="TNum" Display="None" ErrorMessage="班別資料「訓練人數」請輸入數字" ValidationExpression="[0-9]{1,4}"></asp:RegularExpressionValidator>
                            </td>
                            <td class="bluecol_need">訓練時數 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="THours" runat="server" Width="40%" MaxLength="7"></asp:TextBox>小時
                                <asp:RequiredFieldValidator ID="fill20" runat="server" ControlToValidate="THours" Display="None" ErrorMessage="班別資料「訓練時數」為必填欄位"></asp:RequiredFieldValidator>
                                <asp:RegularExpressionValidator ID="check9" runat="server" ControlToValidate="THours" Display="None" ErrorMessage="班別資料「訓練時數」請輸入數字" ValidationExpression="[0-9]{1,4}"></asp:RegularExpressionValidator>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol"></td>
                            <td class="whitecol" colspan="3">
                                <span style="color: #FF0000">
                                    <asp:Label ID="LabMsg8" runat="server"></asp:Label></span>
                                <%--<font color="red"></font>--%>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">訓練起日 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="STDate" runat="server" Width="40%" MaxLength="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" id="date1" alt="" align="top" src="../../images/show-calendar.gif" runat="server" width="30" height="30" /></span><br />
                                <asp:RequiredFieldValidator ID="fill21" runat="server" ControlToValidate="STDate" Display="None" ErrorMessage="班別資料「訓練起日」請填寫"></asp:RequiredFieldValidator>
                                <asp:CustomValidator ID="CustomValidator2" runat="server" ControlToValidate="STDate" Display="None" ErrorMessage="班別資料「訓練起日」不是正確的日期格式" ClientValidationFunction="check_date"></asp:CustomValidator>
                            </td>
                            <td class="bluecol_need">訓練迄日 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="FDDate" runat="server" Width="40%" MaxLength="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" id="date2" alt="" align="top" src="../../images/show-calendar.gif" runat="server" width="30" height="30" /></span><br />
                                <asp:RequiredFieldValidator ID="fill22" runat="server" ControlToValidate="FDDate" Display="None" ErrorMessage="班別資料「訓練迄日」請填寫"></asp:RequiredFieldValidator>
                                <asp:CustomValidator ID="CustomValidator3" runat="server" ControlToValidate="FDDate" Display="None" ErrorMessage="班別資料「訓練迄日」不是正確的日期格式" ClientValidationFunction="check_date"></asp:CustomValidator>
                                <asp:CustomValidator ID="CustomValidator5" runat="server" Display="None" ErrorMessage="班別資料：訓練起日不能比訓練迄日晚(或者同一天)" ClientValidationFunction="CheckTDate"></asp:CustomValidator>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol"></td>
                            <td colspan="3" class="whitecol">
                                <span id="sp_AppStage_1" class="red-style1" runat="server">※訓練期間將會影響審查計分表3-2及4-2計算範圍<br />
                                    3-2學員滿意度調查表：為前一年度7月1日開訓至當年度7月11日前結訓之班級<br />
                                    4-2訓後動態調查表：為前一年度7月1日開訓至當年度4月2日前結訓之班級</span>
                                <span id="sp_AppStage_2" class="red-style1" runat="server">※訓練期間將會影響審查計分表3-2及4-2計算範圍<br />
                                    3-2學員滿意度調查表：為前一年度1月1日開訓至當年度1月11日前結訓之班級<br />
                                    4-2訓後動態調查表：為前一年度1月1日開訓至當年度10月3日前結訓之班級</span>
                            </td>
                        </tr>
                        <tr id="trCredPoint" runat="server">
                            <td id="tdCredPoint" runat="server" class="bluecol_need">學分數 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="CredPoint" runat="server" Columns="5" MaxLength="2" Width="30%"></asp:TextBox>
                                <asp:HiddenField ID="Hid_CredPoint" runat="server" />
                            </td>
                        </tr>
                        <tr id="trRoomName" runat="server">
                            <td class="bluecol_need">上課教室名稱</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="RoomName" runat="server" Width="80%"></asp:TextBox></td>
                        </tr>
                        <tr id="tr_rbl_DISTANCE" runat="server">
                            <td class="bluecol_need">辦理方式 </td>
                            <td class="whitecol" colspan="3"><%--辦理方式/遠距教學--%>
                                <asp:RadioButtonList ID="rbl_DISTANCE" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow"></asp:RadioButtonList>
                                <asp:Label ID="lab_DISTANCE" runat="server"></asp:Label>
                                <%--DISTANCE--%><asp:HiddenField ID="Hid_DISTANCE" runat="server" />
                            </td>
                        </tr>
                        <%--政策性產業課程可辦理班數--%>
                        <%--<tr id="tr_PolicyPreVal" runat="server">
                            <td class="bluecol" width="20%">政策性產業課程可辦理班數</td>
                            <td colspan="3" class="whitecol" width="80%">
                                <asp:DataGrid ID="DataGrid33" Width="33%" runat="server" CssClass="font" AutoGenerateColumns="False">
                                    <EditItemStyle Wrap="False"></EditItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <Columns>
                                        <asp:TemplateColumn HeaderText="年度">
                                            <ItemStyle Wrap="False" Width="10%" HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:HiddenField ID="Hid_PREID" runat="server" />
                                                <asp:HiddenField ID="Hid_YEARS_V1" runat="server" />
                                                <asp:Label ID="lab_YEARS" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="可辦理班數">
                                            <ItemStyle Wrap="False" Width="10%" HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:TextBox ID="txt_PRECLASSCNT" runat="server" MaxLength="1"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>--%>
                        <tr id="FactModeTR" runat="server">
                            <td class="bluecol_need">場地類型 </td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="FactMode" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="1">教室</asp:ListItem>
                                    <asp:ListItem Value="2">演講廳</asp:ListItem>
                                    <asp:ListItem Value="3">會議室</asp:ListItem>
                                    <asp:ListItem Value="99">其他(請說明)</asp:ListItem>
                                </asp:RadioButtonList>
                                <asp:TextBox ID="FactModeOther" runat="server" Width="40%"></asp:TextBox>
                                <%--<asp:CustomValidator ID="fill33" runat="server" Display="None" ErrorMessage="請選擇「場地類型」" ClientValidationFunction="CheckFactMode"></asp:CustomValidator>
                                <asp:CustomValidator ID="fill37" runat="server" Display="None" ErrorMessage="請輸入場地類型[其他]" ClientValidationFunction="CheckFactModeOther"></asp:CustomValidator>--%>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">學科場地1／上課地址</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="SciPlaceID" runat="server" Width="80%"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol">學科場地2／上課地址</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="SciPlaceID2" runat="server" Width="80%"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol">術科場地1／上課地址</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="TechPlaceID" runat="server" Width="80%"></asp:DropDownList><br />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">術科場地2／上課地址</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="TechPlaceID2" runat="server" Width="80%"></asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">遠距課程環境1</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="ddl_REMOTEID1" runat="server" Width="80%"></asp:DropDownList>
                                <%--<asp:HiddenField ID="Hid_RMTID1" runat="server" />--%>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">遠距課程環境2</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="ddl_REMOTEID2" runat="server" Width="80%"></asp:DropDownList>
                                <%--<asp:HiddenField ID="Hid_RMTID2" runat="server" />--%>
                            </td>
                        </tr>
                        <%-- <tr id="tr_Taddress2" runat="server" style="display: none"><td class="bluecol" style="display: none">學科上課地址/術科上課地址</td>
                        <td colspan="3" style="display: none"><asp:DropDownList ID="Taddress2" runat="server" AutoPostBack="true"></asp:DropDownList>
                        <asp:DropDownList ID="Taddress3" runat="server" AutoPostBack="true"></asp:DropDownList></td></tr>--%>
                        <tr>
                            <td class="whitecol">&nbsp; </td>
                            <td class="whitecol" colspan="3">
                                <span style="color: #FF0000">※【班別資料】『學科場地』、『學科場地2』、『術科場地』、『術科場地2』必須填寫 其中一項;<br />
                                    【學科上課地址】、【術科上課地址】至少要設定其中一項<br />
                                    ※【辦理方式】選擇『混成課程』，選基本儲存、正式儲存時，【遠距課程環境】必須設定<br />
                                </span>
                                <%--<font color="red"> </font>--%>
                                <asp:CustomValidator ID="checkSciPlaceID" runat="server" Display="None" ErrorMessage="班別資料『學科場地1』或『術科場地1』必須填寫" ClientValidationFunction="check_PlaceID"></asp:CustomValidator>
                                <asp:CustomValidator ID="checkSciPlaceID2" runat="server" Display="None" ErrorMessage="班別資料『學科場地1』、『學科場地2』必項填寫 其中一項;『術科場地1』、『術科場地2』必須填寫 其中一項" ClientValidationFunction="check_PlaceID"></asp:CustomValidator>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">課程內容有室外教學</td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="rbl_OUTDOOR" runat="server" Width="100%" CellSpacing="0" CellPadding="0" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="Y">是</asp:ListItem>
                                    <asp:ListItem Value="N">否</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">容納人數 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="ConNum" runat="server" Columns="5" MaxLength="5" Width="18%"></asp:TextBox>人
                                <asp:RequiredFieldValidator ID="fill34" runat="server" ControlToValidate="ConNum" Display="None" ErrorMessage="請輸入容納人數"></asp:RequiredFieldValidator>
                                <asp:RegularExpressionValidator ID="RegularExpressionValidator3" runat="server" ControlToValidate="ConNum" Display="None" ErrorMessage="容納人數必須為數字" ValidationExpression="^\d*"></asp:RegularExpressionValidator>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">聯絡人 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="ContactName" runat="server" MaxLength="50" Width="30%"></asp:TextBox>
                                <asp:RequiredFieldValidator ID="Requiredfieldvalidator4" runat="server" ControlToValidate="ContactName" Display="None" ErrorMessage="請輸入聯絡人"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr id="trContactPhone_2023_O1" runat="server">
                            <td class="bluecol_need">電話</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="ContactPhone" runat="server" MaxLength="40" Width="44%"></asp:TextBox></td>
                        </tr>
                        <tr id="trContactPhone_2024_N1" runat="server">
                            <td class="bluecol_need">辦公室電話</td>
                            <td class="whitecol">
                                <asp:TextBox ID="ContactPhone_1" runat="server" MaxLength="10" Width="18%" ToolTip="區碼(2~4碼)" placeholder="區碼(0開頭)"></asp:TextBox>-
                                <asp:TextBox ID="ContactPhone_2" runat="server" MaxLength="10" Width="30%" ToolTip="電話(8碼內)" placeholder="電話(8碼)"></asp:TextBox>#
                                <asp:TextBox ID="ContactPhone_3" runat="server" MaxLength="10" Width="18%" ToolTip="分機(8碼內)" placeholder="分機"></asp:TextBox>
                            </td>
                            <td class="bluecol_need">行動電話</td>
                            <td class="whitecol">
                                <asp:TextBox ID="ContactMobile_1" runat="server" MaxLength="10" Width="18%" ToolTip="手機號碼前4碼" placeholder="手機前4碼(0開頭)"></asp:TextBox>-
                                <asp:TextBox ID="ContactMobile_2" runat="server" MaxLength="10" Width="30%" ToolTip="手機號碼後6碼" placeholder="手機後6碼"></asp:TextBox>
                            </td>
                        </tr>
                        <tr id="trContactPhone_2024_N2" runat="server">
                            <td class="whitecol"></td>
                            <td class="whitecol">
                                <asp:Label ID="lab_ContactPhone_m1" runat="server" Text="(【辦公室電話】、【行動電話】至少須擇一填寫)" ForeColor="Red"></asp:Label>
                            </td>
                            <td class="whitecol"></td>
                            <td class="whitecol">
                                <asp:Label ID="lab_ContactMobile_m2" runat="server" Text="(【辦公室電話】、【行動電話】至少須擇一填寫)" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">電子郵件 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="ContactEmail" runat="server" MaxLength="64" Width="80%"></asp:TextBox>
                                <br />
                                <asp:RegularExpressionValidator ID="chkContactEmail1" runat="server" ControlToValidate="ContactEmail" Display="None" ErrorMessage="電子郵件輸入錯誤" ValidationExpression="\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"></asp:RegularExpressionValidator>
                            </td>
                            <td class="bluecol">傳真 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="ContactFax" runat="server" MaxLength="64" Width="70%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">訓練職能</td>
                            <td class="whitecol">
                                <asp:DropDownList ID="ClassCate" runat="server"></asp:DropDownList>
                                <asp:RequiredFieldValidator ID="fill35" runat="server" ControlToValidate="ClassCate" Display="None" ErrorMessage="請選擇訓練職能"></asp:RequiredFieldValidator>
                            </td>
                            <td class="bluecol_need">報名繳費方式</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="EnterSupplyStyle" runat="server" CssClass="font" RepeatLayout="Flow" Width="90%">
                                    <asp:ListItem Value="1" Selected="True">1.報名時應先繳全額訓練費用，待結訓審核通過後核撥補助款</asp:ListItem>
                                    <asp:ListItem Value="2">2.報名時應先繳50%訓練費用，待結訓審核通過後核撥補助款</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="ContentTR" runat="server">
                            <td class="bluecol_need">課程大綱 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="Content" runat="server" Columns="77" Rows="8" TextMode="MultiLine" Width="70%"></asp:TextBox>
                                <asp:RequiredFieldValidator ID="fill36" runat="server" ControlToValidate="Content" Display="None" ErrorMessage="請輸入課程大綱"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr>
                            <td id="tdContent" colspan="4" align="center" class="table_title">課程大綱 <a href="#TableLay5">(移至班別資料)</a> </td>
                        </tr>
                        <tr>
                            <td colspan="4">
                                <table id="Table5A" border="0" cellspacing="1" cellpadding="1">
                                    <tr>
                                        <td class="bluecol" width="8%">日期 </td>
                                        <td class="bluecol" width="6%">授課時段 </td>
                                        <td class="bluecol" width="6%">授課時間 </td>
                                        <td class="bluecol" width="6%">時數 </td>
                                        <td class="bluecol" width="8%" id="td_EHour_h1" runat="server">
                                            <asp:Label ID="th_EHOUR" runat="server" Text="技檢訓練/"></asp:Label>
                                            <span id="Sp_th_AIAHOUR_WNLHOUR" runat="server">
                                                <br />
                                                <asp:Label ID="th_AIAHOUR" runat="server" Text="AI應用/"></asp:Label>
                                                <br />
                                                <asp:Label ID="th_WNLHOUR" runat="server" Text="職場續航"></asp:Label></span></td>
                                        <td class="bluecol" width="8%">課程進度／內容 </td>
                                        <td class="bluecol" width="6%">學／術科 </td>
                                        <td class="bluecol" width="5%">上課地點 </td>
                                        <td class="bluecol" width="3%" id="td_cbFARLEARN_h" runat="server">遠距教學</td>
                                        <td class="bluecol" width="3%">室外教學 </td>
                                        <td class="bluecol" width="6%">授課師資 </td>
                                        <td class="bluecol" width="6%">助教 </td>
                                        <td class="bluecol" width="6%">功能 </td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol" align="center">
                                            <asp:TextBox ID="STrainDate" runat="server" MaxLength="10" Columns="8"></asp:TextBox>
                                            <span runat="server">
                                                <img style="cursor: pointer" id="date3" alt="" align="top" src="../../images/show-calendar.gif" runat="server" width="30" height="30" /></span>
                                        </td>
                                        <td id="tdTPERIOD28" runat="server" class="whitecol" align="center" nowrap="nowrap">
                                            <asp:CheckBox ID="TPERIOD28_1" runat="server" Text="早上" ToolTip="7:00-13:00" />
                                            <br />
                                            <asp:CheckBox ID="TPERIOD28_2" runat="server" Text="下午" ToolTip="13:00-18:00" />
                                            <br />
                                            <asp:CheckBox ID="TPERIOD28_3" runat="server" Text="晚上" ToolTip="18:00-22:00" />
                                        </td>
                                        <td class="whitecol" align="center" nowrap="nowrap">
                                            <asp:DropDownList ID="ddlpnH1" runat="server"></asp:DropDownList>：
                                            <asp:DropDownList ID="ddlpnM1" runat="server"></asp:DropDownList><br />
                                            ~<br />
                                            <asp:DropDownList ID="ddlpnH2" runat="server"></asp:DropDownList>：
                                            <asp:DropDownList ID="ddlpnM2" runat="server"></asp:DropDownList><br />
                                            <asp:TextBox ID="PName" runat="server" onfocus="this.blur()" Visible="False"></asp:TextBox>
                                        </td>
                                        <td class="whitecol" align="center">
                                            <asp:TextBox ID="PHour" runat="server" MaxLength="5" Width="70%"></asp:TextBox></td>
                                        <td class="whitecol" align="left" id="td_EHour_h2" runat="server" nowrap="nowrap"><%--技檢訓練時數--%>
                                            <span id="SpEHOUR" runat="server">技:<asp:TextBox ID="EHOUR" runat="server" Width="66%" MaxLength="5" placeholder="時數"></asp:TextBox></span>
                                            <span id="Sp_AIAHOUR_WNLHOUR" runat="server">
                                                <br />
                                                AI:<asp:TextBox ID="AIAHOUR" runat="server" Width="66%" MaxLength="5" placeholder="時數"></asp:TextBox>
                                                <br />
                                                職:<asp:TextBox ID="WNLHOUR" runat="server" Width="66%" MaxLength="5" placeholder="時數"></asp:TextBox></span>
                                        </td>
                                        <td class="whitecol" align="center">
                                            <asp:TextBox ID="PCont" runat="server" Columns="20" Rows="5" TextMode="MultiLine" placeholder="(請論述課程主題及其細項大綱內容)"></asp:TextBox></td>
                                        <td class="whitecol" align="center">
                                            <asp:DropDownList ID="Classification1" runat="server" AutoPostBack="True">
                                                <asp:ListItem Value="0">請選擇</asp:ListItem>
                                                <asp:ListItem Value="1">學科</asp:ListItem>
                                                <asp:ListItem Value="2">術科</asp:ListItem>
                                            </asp:DropDownList>
                                        </td>
                                        <td class="whitecol" align="center"><%--上課地點--%>
                                            <asp:DropDownList ID="PTID1" runat="server"></asp:DropDownList>
                                            <asp:DropDownList ID="PTID2" runat="server"></asp:DropDownList>
                                        </td>
                                        <td class="whitecol" align="center" id="td_cbFARLEARN_d" runat="server">
                                            <asp:CheckBox ID="cbFARLEARN" runat="server" ToolTip="遠距教學" />
                                            <asp:Label ID="lab_cbFARLEARN" runat="server"></asp:Label>
                                        </td>
                                        <td class="whitecol" align="center">
                                            <asp:CheckBox ID="cbOUTLEARN" runat="server" ToolTip="室外教學" />
                                        </td>
                                        <td class="whitecol" align="center">
                                            <input id="OLessonTeah1Value" type="hidden" name="OLessonTeah1Value" runat="server" />
                                            <asp:TextBox ID="OLessonTeah1" runat="server" onfocus="this.blur()" Columns="5" ToolTip="點選兩下跳出視窗選擇教師" Width="90%" placeholder="(點擊)"></asp:TextBox>
                                        </td>
                                        <td class="whitecol" align="center">
                                            <input id="OLessonTeah2Value" type="hidden" name="OLessonTeah2Value" runat="server" />
                                            <asp:TextBox ID="OLessonTeah2" runat="server" onfocus="this.blur()" Columns="5" ToolTip="點選兩下跳出視窗選擇助教" Width="90%" placeholder="(點擊)"></asp:TextBox>
                                        </td>
                                        <td class="whitecol" align="center">
                                            <asp:Button ID="Button1" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_M"></asp:Button></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4">
                                <table id="Datagrid3Table" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                    <tr>
                                        <td>
                                            <asp:DataGrid ID="Datagrid3" Width="100%" runat="server" AutoGenerateColumns="False" Style="text-align: left">
                                                <ItemStyle BackColor="White"></ItemStyle>
                                                <EditItemStyle Wrap="False"></EditItemStyle>
                                                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                <AlternatingItemStyle BackColor="#F5F5F5" />
                                                <Columns>
                                                    <asp:TemplateColumn HeaderText="日期">
                                                        <ItemStyle Width="8%" HorizontalAlign="Center"></ItemStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="STrainDateLabel" runat="server"></asp:Label>
                                                        </ItemTemplate>
                                                        <FooterStyle Wrap="False"></FooterStyle>
                                                        <EditItemTemplate>
                                                            <asp:TextBox ID="STrainDateTxt" runat="server" MaxLength="10" Columns="8"></asp:TextBox>
                                                            <span runat="server">
                                                                <img id="Img2" style="cursor: pointer" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30" /></span>
                                                        </EditItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="授課時段">
                                                        <ItemStyle Width="6%" Wrap="false" HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <asp:CheckBox ID="TPERIOD28_1t" runat="server" Text="早上" ToolTip="7:00-13:00" Enabled="false" /><br />
                                                            <asp:CheckBox ID="TPERIOD28_2t" runat="server" Text="下午" ToolTip="13:00-18:00" Enabled="false" /><br />
                                                            <asp:CheckBox ID="TPERIOD28_3t" runat="server" Text="晚上" ToolTip="18:00-22:00" Enabled="false" />
                                                        </ItemTemplate>
                                                        <EditItemTemplate>
                                                            <asp:CheckBox ID="TPERIOD28_1e" runat="server" Text="早上" ToolTip="7:00-13:00" /><br />
                                                            <asp:CheckBox ID="TPERIOD28_2e" runat="server" Text="下午" ToolTip="13:00-18:00" /><br />
                                                            <asp:CheckBox ID="TPERIOD28_3e" runat="server" Text="晚上" ToolTip="18:00-22:00" />
                                                        </EditItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="授課時間">
                                                        <ItemStyle Width="8%" Wrap="false" HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <asp:Label ID="PNameLabel" runat="server"></asp:Label>
                                                        </ItemTemplate>
                                                        <EditItemTemplate>
                                                            <asp:DropDownList ID="Eddlh1" runat="server"></asp:DropDownList>：
                                                            <asp:DropDownList ID="Eddlm1" runat="server"></asp:DropDownList><br />
                                                            ~<br />
                                                            <asp:DropDownList ID="Eddlh2" runat="server"></asp:DropDownList>：
                                                            <asp:DropDownList ID="Eddlm2" runat="server"></asp:DropDownList>
                                                            <asp:TextBox ID="PNameTxt" runat="server" onfocus="this.blur()" Visible="False"></asp:TextBox>
                                                        </EditItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="時數">
                                                        <ItemStyle Width="6%" Wrap="false" HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <asp:Label ID="PHourLabel" runat="server"></asp:Label>
                                                        </ItemTemplate>
                                                        <EditItemTemplate>
                                                            <asp:TextBox ID="PHourTxt" runat="server" Width="70%" MaxLength="5"></asp:TextBox>
                                                        </EditItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="技檢訓練時數">
                                                        <HeaderStyle Width="6%" Wrap="false" HorizontalAlign="Left" />
                                                        <ItemStyle Wrap="false" HorizontalAlign="Left" />
                                                        <HeaderTemplate>
                                                            <asp:Label ID="lb_EHOUR" runat="server" Text="技檢訓練/"></asp:Label>
                                                            <span id="Sp_lb_AIAHOUR_WNLHOUR" runat="server">
                                                                <br />
                                                                <asp:Label ID="lb_AIAHOUR" runat="server" Text="AI應用/"></asp:Label>
                                                                <br />
                                                                <asp:Label ID="lb_WNLHOUR" runat="server" Text="職場續航"></asp:Label></span>
                                                        </HeaderTemplate>
                                                        <ItemTemplate>
                                                            技:<asp:Label ID="EHourLabel" runat="server"></asp:Label>
                                                            <span id="Sp_AIAHOUR_WNLHOURLabel" runat="server">
                                                                <br />
                                                                AI:<asp:Label ID="AIAHOURLabel" runat="server"></asp:Label>
                                                                <br />
                                                                職:<asp:Label ID="WNLHOURLabel" runat="server"></asp:Label>
                                                            </span>
                                                        </ItemTemplate>
                                                        <EditItemTemplate>
                                                            技:<asp:TextBox ID="EHourTxt" runat="server" Width="66%" MaxLength="5" placeholder="時數"></asp:TextBox>
                                                            <span id="Sp_AIAHOUR_WNLHOURTxt" runat="server">
                                                                <br />
                                                                AI:<asp:TextBox ID="AIAHOURTxt" runat="server" Width="66%" MaxLength="5" placeholder="時數"></asp:TextBox>
                                                                <br />
                                                                職:<asp:TextBox ID="WNLHOURTxt" runat="server" Width="66%" MaxLength="5" placeholder="時數"></asp:TextBox>
                                                            </span>
                                                        </EditItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="課程進度／內容">
                                                        <HeaderStyle Width="8%" />
                                                        <ItemStyle Width="8%" Wrap="false" HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <asp:TextBox ID="PContText" runat="server" onfocus="this.blur()" TextMode="MultiLine" Rows="5" Enabled="False"></asp:TextBox>
                                                        </ItemTemplate>
                                                        <EditItemTemplate>
                                                            <asp:TextBox ID="PContEdit" runat="server" TextMode="MultiLine" Rows="5"></asp:TextBox>
                                                        </EditItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="學／術科">
                                                        <ItemStyle Width="6%" Wrap="false" HorizontalAlign="Center" />
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
                                                        <ItemStyle Width="9%" HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <asp:DropDownList ID="drpPTID" runat="server" Enabled="False"></asp:DropDownList>
                                                        </ItemTemplate>
                                                        <EditItemTemplate>
                                                            <asp:DropDownList ID="drpPTIDEdit1" runat="server"></asp:DropDownList>
                                                            <asp:DropDownList ID="drpPTIDEdit2" runat="server"></asp:DropDownList>
                                                        </EditItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="遠距教學">
                                                        <%--<HeaderTemplate>遠距教學<input onclick="chackAll(7);" type="checkbox" name="Choose1" id="Choose1" /></HeaderTemplate>--%>
                                                        <ItemStyle Width="6%" HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <asp:CheckBox ID="cb_FARLEARNi" runat="server" Enabled="False" />
                                                        </ItemTemplate>
                                                        <EditItemTemplate>
                                                            <asp:CheckBox ID="cb_FARLEARNe" runat="server" />
                                                        </EditItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="室外教學">
                                                        <ItemStyle Width="6%" HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <asp:CheckBox ID="cb_OUTLEARNi" runat="server" Enabled="False" />
                                                        </ItemTemplate>
                                                        <EditItemTemplate>
                                                            <asp:CheckBox ID="cb_OUTLEARNe" runat="server" />
                                                        </EditItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="授課師資">
                                                        <ItemStyle Width="6%" Wrap="false" HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <input id="Tech1Value" type="hidden" name="Tech1Value" runat="server" />
                                                            <asp:TextBox ID="Tech1Text" runat="server" onfocus="this.blur()" Columns="5" Enabled="False"></asp:TextBox>
                                                        </ItemTemplate>
                                                        <EditItemTemplate>
                                                            <input id="Tech1ValueEdit" type="hidden" runat="server" />
                                                            <asp:TextBox ID="Tech1Edit" runat="server" onfocus="this.blur()" Columns="5" ToolTip="點選兩下跳出視窗選擇教師" placeholder="(點擊)"></asp:TextBox>
                                                        </EditItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="助教">
                                                        <ItemStyle Width="6%" Wrap="false" HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <input id="Tech2Value" type="hidden" name="Tech2Value" runat="server" />
                                                            <asp:TextBox ID="Tech2Text" runat="server" onfocus="this.blur()" Columns="5" Enabled="False"></asp:TextBox>
                                                        </ItemTemplate>
                                                        <EditItemTemplate>
                                                            <input id="Tech2ValueEdit" type="hidden" runat="server" />
                                                            <asp:TextBox ID="Tech2Edit" runat="server" onfocus="this.blur()" Columns="5" ToolTip="點選兩下跳出視窗選擇助教" placeholder="(點擊)"></asp:TextBox>
                                                        </EditItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="功能">
                                                        <ItemStyle Width="6%" HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <asp:Button ID="Button6" runat="server" Text="修改" CausesValidation="False" CommandName="edit" CssClass="asp_button_M"></asp:Button>
                                                            <asp:Button ID="Button7" runat="server" Text="刪除" CausesValidation="False" CommandName="del" CssClass="asp_button_M"></asp:Button>
                                                        </ItemTemplate>
                                                        <EditItemTemplate>
                                                            <asp:Button ID="Button10" runat="server" Text="儲存" CausesValidation="False" CommandName="save" CssClass="asp_button_M"></asp:Button>
                                                            <asp:Button ID="Button11" runat="server" Text="取消" CausesValidation="False" CommandName="cancel" CssClass="asp_button_M"></asp:Button>
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
                            <td colspan="4">
                                <table id="Table3" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                                    <tr>
                                        <td colspan="4" align="center" class="table_title" width="100%">上課時間 <a href="#TableLay5">(移至班別資料)</a> </td>
                                    </tr>
                                    <tr>
                                        <td align="center" class="bluecol" width="10%">星期 </td>
                                        <td align="center" class="bluecol" width="80%">時間 </td>
                                        <td align="center" class="bluecol" width="10%">功能 </td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol" align="center">
                                            <asp:DropDownList ID="Weeks" runat="server"></asp:DropDownList></td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="txtTimes" runat="server" Columns="70" Width="80%"></asp:TextBox>
                                            <%--color="red"--%><span style="color: #FF0000">輸入範例，半型字體例如：18:00~21:00，多筆以 ; 做分隔</span></td>
                                        <td class="whitecol" align="center">
                                            <asp:Button ID="Button29" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_M"></asp:Button></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4">
                                <table id="DataGrid1Table" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                    <tr>
                                        <td>
                                            <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AutoGenerateColumns="False">
                                                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                <AlternatingItemStyle BackColor="#F5F5F5" />
                                                <Columns>
                                                    <asp:TemplateColumn HeaderText="星期">
                                                        <ItemStyle Width="10%" HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <asp:Label ID="Weeks1" runat="server"></asp:Label>
                                                        </ItemTemplate>
                                                        <EditItemTemplate>
                                                            <asp:DropDownList ID="Weeks2" runat="server"></asp:DropDownList>
                                                        </EditItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="上課時段">
                                                        <ItemStyle Width="80%" CssClass="whitecol" />
                                                        <ItemTemplate>
                                                            <asp:Label ID="Times1" runat="server"></asp:Label>
                                                            <%--<asp:Label ID="labmsgTimes1" runat="server" ForeColor="Red"></asp:Label>--%>
                                                        </ItemTemplate>
                                                        <EditItemTemplate>
                                                            <asp:TextBox ID="Times2" runat="server" Columns="70" Width="80%"></asp:TextBox>
                                                        </EditItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="功能">
                                                        <ItemStyle Width="10%" HorizontalAlign="Center" />
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
                                <table id="tbPlanDepot" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                    <tr>
                                        <td class="bluecol" width="16%">&nbsp; 轄區重點產業 </td>
                                        <td class="whitecol" width="84%">
                                            <asp:DropDownList Style="z-index: 0" ID="ddlDEPOT15" runat="server"></asp:DropDownList>
                                            <a href="#TableLay5">(移至班別資料)</a> </td>
                                    </tr>
                                    <%--<tr><td class="bluecol">&nbsp; 新興產業 </td><td class="whitecol">
                                        <asp:DropDownList Style="z-index: 0" ID="ddlKID06" runat="server"></asp:DropDownList></td></tr><tr><td class="bluecol">&nbsp; 重點服務業 </td>
                                        <td class="whitecol"><asp:DropDownList Style="z-index: 0" ID="ddlKID10" runat="server"></asp:DropDownList></td></tr>--%>
                                    <tr id="trKID19" runat="server">
                                        <td class="bluecol">&nbsp; 政府政策性產業 </td>
                                        <td class="whitecol">
                                            <asp:DropDownList Style="z-index: 0" ID="ddlKID19" runat="server"></asp:DropDownList></td>
                                    </tr>
                                    <tr id="trKID18" runat="server">
                                        <td class="bluecol">&nbsp; 新南向政策 </td>
                                        <td class="whitecol">
                                            <asp:DropDownList Style="z-index: 0" ID="ddlKID18" runat="server"></asp:DropDownList></td>
                                    </tr>
                                    <tr id="trKID20" runat="server">
                                        <td class="bluecol">&nbsp; 政府政策性產業 </td>
                                        <td class="whitecol">
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
                                                    <td class="whitecol">
                                                        <asp:CheckBoxList ID="CBLKID20_6" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="bluecol">進階政策性產業類別</td>
                                                    <td class="whitecol">
                                                        <asp:CheckBoxList ID="CBLKID22" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr id="trKID25" runat="server">
                                        <td class="bluecol">&nbsp; 政府政策性產業 </td>
                                        <td class="whitecol">
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
                                                <tr id="trCBLKID25_7" runat="server">
                                                    <td class="bluecol">AI加值應用</td>
                                                    <td class="whitecol">
                                                        <asp:CheckBoxList ID="CBLKID25_7" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                                    </td>
                                                </tr>
                                                <tr id="trCBLKID25_8" runat="server">
                                                    <td class="bluecol">職場續航</td>
                                                    <td class="whitecol">
                                                        <asp:CheckBoxList ID="CBLKID25_8" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="bluecol">進階政策性產業類別</td>
                                                    <td class="whitecol">
                                                        <asp:CheckBoxList ID="CBLKID22B" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>

                                    <%--五大信賴產業推動方案,六大區域產業及生活圈,智慧國家2.0綱領
                                        ,新南向政策推動計畫,國家人才競爭力躍升方案,AI新十大建設推動方案
                                        ,台灣AI行動計畫2.0,智慧機器人產業推動方案,臺灣2050淨零轉型,*/,--%>
                                    <tr id="trKID26" runat="server">
                                        <td class="bluecol">&nbsp; 政府政策性產業 </td>
                                        <td class="whitecol">
                                            <table border="0" cellspacing="1" cellpadding="1" width="100%">
                                                <tr>
                                                    <td class="bluecol" width="22%">五大信賴產業推動方案</td>
                                                    <td class="whitecol" width="78%">
                                                        <asp:CheckBoxList ID="CBLKID26_1" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="bluecol">六大區域產業及生活圈</td>
                                                    <td class="whitecol">
                                                        <asp:CheckBoxList ID="CBLKID26_2" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="bluecol">智慧國家2.0綱領</td>
                                                    <td class="whitecol">
                                                        <asp:CheckBoxList ID="CBLKID26_3" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="bluecol">新南向政策推動計畫</td>
                                                    <td class="whitecol">
                                                        <asp:CheckBoxList ID="CBLKID26_4" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="bluecol">國家人才競爭力躍升方案</td>
                                                    <td class="whitecol">
                                                        <asp:CheckBoxList ID="CBLKID26_5" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="bluecol">AI新十大建設推動方案</td>
                                                    <td class="whitecol">
                                                        <asp:CheckBoxList ID="CBLKID26_6" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                                    </td>
                                                </tr>
                                                <tr id="tr2" runat="server">
                                                    <td class="bluecol">台灣AI行動計畫2.0</td>
                                                    <td class="whitecol">
                                                        <asp:CheckBoxList ID="CBLKID26_7" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                                    </td>
                                                </tr>
                                                <tr id="tr3" runat="server">
                                                    <td class="bluecol">智慧機器人產業推動方案</td>
                                                    <td class="whitecol">
                                                        <asp:CheckBoxList ID="CBLKID26_8" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="bluecol">臺灣2050淨零轉型</td>
                                                    <td class="whitecol">
                                                        <asp:CheckBoxList ID="CBLKID26_9" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>

                                    <%--<HeaderTemplate>遠距教學<input onclick="chackAll(7);" type="checkbox" name="Choose1" id="Choose1" /></HeaderTemplate>--%>
                                    <tr id="trCBLKID60" runat="server">
                                        <td class="bluecol">產業別(管考) </td>
                                        <td class="whitecol">
                                            <asp:CheckBoxList ID="CBLKID60" runat="server" RepeatColumns="4" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4">
                                <table id="Datagrid4headTable" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                    <tr>
                                        <td colspan="4" align="center" class="table_title" width="100%">包班事業單位資料 </td>
                                    </tr>
                                    <tr>
                                        <td align="center" class="bluecol" width="36%">企業名稱 </td>
                                        <td align="center" class="bluecol" width="26%">服務單位統一編號 </td>
                                        <td align="center" class="bluecol" width="26%">保險證號 </td>
                                        <td align="center" class="bluecol" width="12%">功能 </td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">
                                            <asp:TextBox ID="txtUname" runat="server" Columns="40" MaxLength="50"></asp:TextBox></td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="txtIntaxno" runat="server" Columns="9" MaxLength="10"></asp:TextBox></td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="txtUbno" runat="server" Columns="9" MaxLength="9"></asp:TextBox></td>
                                        <td class="whitecol" align="center">
                                            <asp:Button ID="btnAddBusPackage" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_M"></asp:Button></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4">
                                <table id="Datagrid4Table" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                    <tr>
                                        <td>
                                            <asp:DataGrid ID="Datagrid4" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False">
                                                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                <AlternatingItemStyle BackColor="#F5F5F5" />
                                                <Columns>
                                                    <asp:TemplateColumn HeaderText="企業名稱">
                                                        <HeaderStyle Width="36%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="slsbUname" runat="server"></asp:Label>
                                                        </ItemTemplate>
                                                        <EditItemTemplate>
                                                            <asp:TextBox ID="ttxtUname" runat="server"></asp:TextBox>
                                                        </EditItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="服務單位統一編號">
                                                        <HeaderStyle Width="26%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="slabIntaxno" runat="server"></asp:Label>
                                                        </ItemTemplate>
                                                        <EditItemTemplate>
                                                            <asp:TextBox ID="ttxtIntaxno" runat="server"></asp:TextBox>
                                                        </EditItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="保險證號">
                                                        <HeaderStyle Width="26%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="slabUbno" runat="server"></asp:Label>
                                                        </ItemTemplate>
                                                        <EditItemTemplate>
                                                            <asp:TextBox ID="ttxtUbno" runat="server"></asp:TextBox>
                                                        </EditItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="功能">
                                                        <HeaderStyle Width="12%"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center" />
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
                            <td>
                                <table id="DataGrid2Table" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                    <tr>
                                        <td>
                                            <table id="TableCost18" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                                <tr>
                                                    <td width="100%" class="whitecol">
                                                        <asp:Label ID="Label4a" runat="server" ForeColor="Black" Font-Bold="True"></asp:Label>
                                                        <asp:Button ID="btnPDF20171226" runat="server" Text="參考資料" CssClass="asp_button_M" />&nbsp;&nbsp;
                                                        <asp:Label ID="Label4b" runat="server" ForeColor="Black" Font-Bold="True">如超出材料費編列比上限，請填寫超出原因。</asp:Label>
                                                        <%----%>
                                                        <asp:Label ID="LabMsgFIX1" runat="server">&nbsp;</asp:Label><br />
                                                        <table class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                                                            <tr>
                                                                <td class="bluecol_need" width="16%">固定費用總額</td>
                                                                <td class="whitecol" width="34%">
                                                                    <asp:TextBox ID="FIXSUMCOST" runat="server" MaxLength="7" Width="44%"></asp:TextBox></td>
                                                                <td class="bluecol_need" width="18%">固定費用總額單一人時成本</td>
                                                                <td class="whitecol" width="32%">
                                                                    <asp:TextBox ID="ACTHUMCOST" runat="server" MaxLength="7" Width="44%" onfocus="this.blur()"></asp:TextBox></td>
                                                            </tr>
                                                            <tr>
                                                                <td id="tdFIXExceeDesc" class="bluecol">超出人時成本原因說明：</td>
                                                                <td class="whitecol" colspan="3">
                                                                    <asp:HiddenField ID="hid_FIXExceeDesc" runat="server" />
                                                                    <asp:TextBox ID="FIXExceeDesc" runat="server" MaxLength="200" Width="70%"></asp:TextBox></td>
                                                            </tr>
                                                            <tr>
                                                                <td class="bluecol_need">材料費用總額</td>
                                                                <td class="whitecol"><%--材料費上限--%>
                                                                    <asp:TextBox ID="METSUMCOST" runat="server" MaxLength="7" Width="44%"></asp:TextBox></td>
                                                                <td class="bluecol_need">材料費占比</td>
                                                                <td class="whitecol">
                                                                    <asp:TextBox ID="METCOSTPER" runat="server" MaxLength="7" Width="44%" onfocus="this.blur()"></asp:TextBox><asp:Label ID="lab_iPERC100" runat="server" Text="%"></asp:Label></td>
                                                            </tr>
                                                            <tr>
                                                                <td id="tdMETExceeDesc" class="bluecol">超出材料費比率上限原因說明：</td>
                                                                <td class="whitecol" colspan="3">
                                                                    <asp:HiddenField ID="hid_METExceeDesc" runat="server" />
                                                                    <asp:TextBox ID="METExceeDesc" runat="server" MaxLength="200" Width="70%"></asp:TextBox></td>
                                                            </tr>
                                                        </table>
                                                    </td>
                                                </tr>
                                            </table>
                                            <table id="TableCost6" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                                <tr>
                                                    <td colspan="8" width="100%" class="table_title">一人份材料明細 </td>
                                                </tr>
                                                <tr>
                                                    <td colspan="8" class="whitecol" width="100%">
                                                        <asp:Label ID="Label7" runat="server" CssClass="font">匯入明細</asp:Label>
                                                        <input id="File1" type="file" size="50" name="File1" runat="server" accept=".csv,.xls" />
                                                        <asp:Button ID="BtnImport1" runat="server" CausesValidation="False" Text="匯入明細" CssClass="asp_button_M" />(必須為csv格式)
                                                        <asp:HyperLink ID="HyperLink1" runat="server" CssClass="font" NavigateUrl="../../Doc/PersonCost_Imp2.zip">下載匯入格式檔</asp:HyperLink>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="bluecol" width="6%"></td>
                                                    <td class="bluecol" width="8%">項次 </td>
                                                    <td class="bluecol" width="14%">品名 </td>
                                                    <td class="bluecol" width="12%">規格 </td>
                                                    <td class="bluecol" width="8%">單位 </td>
                                                    <td class="bluecol" width="10%">每人數量 </td>
                                                    <td class="bluecol">用途說明 </td>
                                                    <td class="bluecol" width="8%">功能 </td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:Button ID="BtnDelDG6" runat="server" Text="勾選刪除" CssClass="asp_button_M" CommandName="BtnDelDG6" /></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:TextBox ID="tItemNo6" runat="server" Columns="5" MaxLength="5"></asp:TextBox></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:TextBox ID="tCName6" runat="server" Columns="10" MaxLength="30"></asp:TextBox></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:TextBox ID="tStandard6" runat="server" Columns="10" MaxLength="300"></asp:TextBox></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:TextBox ID="tUnit6" runat="server" Columns="5" MaxLength="5"></asp:TextBox></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:TextBox ID="tPerCount6" runat="server" Columns="5" MaxLength="5"></asp:TextBox></td>
                                                    <td class="whitecol">
                                                        <asp:TextBox ID="tPurpose6" runat="server" Columns="35" MaxLength="300" Width="90%"></asp:TextBox></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:Button ID="btnAddCost6" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_S"></asp:Button></td>
                                                </tr>
                                                <tr>
                                                    <td colspan="8">
                                                        <table id="DataGrid6Table" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                                            <tr>
                                                                <td>
                                                                    <asp:DataGrid ID="DataGrid6" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" BorderColor="Gray">
                                                                        <AlternatingItemStyle BackColor="WhiteSmoke" />
                                                                        <Columns>
                                                                            <asp:TemplateColumn HeaderText="勾選刪除">
                                                                                <HeaderTemplate>勾選刪除</HeaderTemplate>
                                                                                <ItemTemplate>
                                                                                    <asp:CheckBox ID="CheckBoxDG6" runat="server" />
                                                                                </ItemTemplate>
                                                                                <HeaderStyle VerticalAlign="Middle" HorizontalAlign="Center" Width="6%" />
                                                                                <ItemStyle VerticalAlign="Middle" HorizontalAlign="Center" CssClass="whitecol" />
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="項次">
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="eItemNo6" runat="server" Columns="5" MaxLength="5"></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="lItemNo6" runat="server"></asp:Label>
                                                                                </ItemTemplate>
                                                                                <HeaderStyle Width="6%" />
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="品名">
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="eCName6" runat="server" Columns="10" MaxLength="30"></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="lCName6" runat="server"></asp:Label>
                                                                                </ItemTemplate>
                                                                                <HeaderStyle Width="14%" />
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="規格">
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="eStandard6" runat="server" Columns="10" MaxLength="300"></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="lStandard6" runat="server"></asp:Label>
                                                                                </ItemTemplate>
                                                                                <HeaderStyle Width="16%" />
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="單位">
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="eUnit6" runat="server" Columns="3" MaxLength="5"></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="lUnit6" runat="server"></asp:Label>
                                                                                </ItemTemplate>
                                                                                <HeaderStyle Width="6%" />
                                                                                <ItemStyle HorizontalAlign="Center" />
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="每人數量">
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="ePerCount6" runat="server" Columns="3" MaxLength="5"></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="lPerCount6" runat="server"></asp:Label>
                                                                                </ItemTemplate>
                                                                                <HeaderStyle Width="6%" />
                                                                                <ItemStyle HorizontalAlign="Center" />
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="訓練人數">
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="eTNum6" runat="server" Columns="3"></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="lTNum6" runat="server"></asp:Label>
                                                                                </ItemTemplate>
                                                                                <HeaderStyle Width="6%" />
                                                                                <ItemStyle HorizontalAlign="Center" />
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="總數量">
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="eTotal6" runat="server" Columns="3"></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="lTotal6" runat="server"></asp:Label>
                                                                                </ItemTemplate>
                                                                                <HeaderStyle Width="6%" />
                                                                                <ItemStyle HorizontalAlign="Center" />
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="用途說明">
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="ePurPose6" runat="server" Columns="15" MaxLength="300"></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="lPurPose6" runat="server"></asp:Label>
                                                                                </ItemTemplate>
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="功能">
                                                                                <EditItemTemplate>
                                                                                    <asp:Button ID="btnUpd6" runat="server" CausesValidation="False" CommandName="UPD6" CssClass="asp_button_M" Text="更新" />
                                                                                    <asp:Button ID="btnCls6" runat="server" CausesValidation="False" CommandName="CLS6" CssClass="asp_button_M" Text="取消" />
                                                                                </EditItemTemplate>
                                                                                <ItemTemplate>
                                                                                    <asp:HiddenField ID="Hid_DataKey" runat="server" />
                                                                                    <asp:Button ID="btnDel6" runat="server" CausesValidation="False" CommandName="DEL6" CssClass="asp_button_M" Text="刪除" />
                                                                                    <asp:Button ID="btnEdt6" runat="server" CausesValidation="False" CommandName="EDT6" CssClass="asp_button_M" Text="修改" />
                                                                                </ItemTemplate>
                                                                                <ItemStyle HorizontalAlign="Center" Width="8%" />
                                                                            </asp:TemplateColumn>
                                                                        </Columns>
                                                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                                    </asp:DataGrid>
                                                                </td>
                                                            </tr>
                                                        </table>
                                                    </td>
                                                </tr>
                                            </table>
                                            <table id="TableCost7" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                                <tr>
                                                    <td class="table_title" colspan="8">共同材料明細 </td>
                                                </tr>
                                                <tr>
                                                    <td colspan="8" class="whitecol">
                                                        <asp:Label ID="Label8" runat="server" CssClass="font">匯入明細</asp:Label>
                                                        <input id="File2" type="file" size="50" name="File2" runat="server" accept=".csv,.xls" />
                                                        <asp:Button ID="BtnImport2" runat="server" CausesValidation="False" Text="匯入明細" CssClass="asp_button_M" />(必須為csv格式)
                                                        <asp:HyperLink ID="HyperLink2" runat="server" CssClass="font" NavigateUrl="../../Doc/CommonCost_Imp2.zip">下載匯入格式檔</asp:HyperLink>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="bluecol" width="6%"></td>
                                                    <td class="bluecol" width="8%">項次 </td>
                                                    <td class="bluecol" width="14%">品名 </td>
                                                    <td class="bluecol" width="12%">規格 </td>
                                                    <td class="bluecol" width="8%">單位 </td>
                                                    <td class="bluecol" width="10%">使用數量 </td>
                                                    <td class="bluecol">用途說明 </td>
                                                    <td class="bluecol" width="8%">功能 </td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">
                                                        <asp:Button ID="BtnDelDG7" runat="server" Text="勾選刪除" CssClass="asp_button_M" CommandName="BtnDelDG7" /></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:TextBox ID="tItemNo7" runat="server" Columns="3" MaxLength="5"></asp:TextBox></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:TextBox ID="tCName7" runat="server" Columns="10" MaxLength="30"></asp:TextBox></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:TextBox ID="tStandard7" runat="server" Columns="10" MaxLength="300"></asp:TextBox></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:TextBox ID="tUnit7" runat="server" Columns="3" MaxLength="5"></asp:TextBox></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:TextBox ID="tAllCount7" runat="server" Columns="5" MaxLength="5"></asp:TextBox></td>
                                                    <td class="whitecol">
                                                        <asp:TextBox ID="tPurPose7" runat="server" Columns="35" MaxLength="300" Width="90%"></asp:TextBox></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:Button ID="btnAddCost7" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_M"></asp:Button></td>
                                                </tr>
                                                <tr>
                                                    <td colspan="8">
                                                        <table id="DataGrid7Table" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                                            <tr>
                                                                <td>
                                                                    <asp:DataGrid ID="DataGrid7" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" BorderColor="Gray">
                                                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                                                        <Columns>
                                                                            <asp:TemplateColumn HeaderText="勾選刪除">
                                                                                <HeaderTemplate>勾選刪除</HeaderTemplate>
                                                                                <ItemTemplate>
                                                                                    <asp:CheckBox ID="CheckBoxDG7" runat="server" />
                                                                                </ItemTemplate>
                                                                                <HeaderStyle VerticalAlign="Middle" HorizontalAlign="Center" Width="6%" />
                                                                                <ItemStyle VerticalAlign="Middle" HorizontalAlign="Center" CssClass="whitecol" />
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="項次">
                                                                                <HeaderStyle Width="8%"></HeaderStyle>
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
                                                                                <HeaderStyle Width="12%"></HeaderStyle>
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="lStandard7" runat="server"></asp:Label>
                                                                                </ItemTemplate>
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="eStandard7" runat="server" Columns="10" MaxLength="300"></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="單位">
                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                <ItemStyle HorizontalAlign="Center" />
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="lUnit7" runat="server"></asp:Label>
                                                                                </ItemTemplate>
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="eUnit7" runat="server" Columns="3" MaxLength="5"></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="使用數量">
                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                <ItemStyle HorizontalAlign="Center" />
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="lAllCount7" runat="server"></asp:Label>
                                                                                </ItemTemplate>
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="eAllCount7" runat="server" Columns="3" MaxLength="5"></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="訓練人數">
                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                <ItemStyle HorizontalAlign="Center" />
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="lTNum7" runat="server"></asp:Label>
                                                                                </ItemTemplate>
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="eTNum7" runat="server" Columns="3"></asp:TextBox>
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
                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                <ItemStyle HorizontalAlign="Center" />
                                                                                <ItemTemplate>
                                                                                    <asp:HiddenField ID="Hid_DataKey" runat="server" />
                                                                                    <asp:Button ID="btnDel7" runat="server" Text="刪除" CausesValidation="False" CommandName="DEL7" CssClass="asp_button_M"></asp:Button>
                                                                                    <asp:Button ID="btnEdt7" runat="server" Text="修改" CausesValidation="False" CommandName="EDT7" CssClass="asp_button_M"></asp:Button>
                                                                                </ItemTemplate>
                                                                                <EditItemTemplate>
                                                                                    <asp:Button ID="btnUpd7" runat="server" Text="更新" CausesValidation="False" CommandName="UPD7" CssClass="asp_button_M"></asp:Button>
                                                                                    <asp:Button ID="btnCls7" runat="server" Text="取消" CausesValidation="False" CommandName="CLS7" CssClass="asp_button_M"></asp:Button>
                                                                                </EditItemTemplate>
                                                                            </asp:TemplateColumn>
                                                                        </Columns>
                                                                    </asp:DataGrid>
                                                                </td>
                                                            </tr>
                                                        </table>
                                                    </td>
                                                </tr>
                                            </table>
                                            <table id="TableCost8" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                                <tr>
                                                    <td class="table_title" colspan="8">教材明細 </td>
                                                </tr>
                                                <tr>
                                                    <td colspan="8" class="whitecol">
                                                        <asp:Label ID="Labeld8" runat="server" CssClass="font">匯入明細</asp:Label>
                                                        <input id="File3" type="file" size="50" name="File3" runat="server" accept=".csv,.xls" />
                                                        <asp:Button ID="BtnImport8" runat="server" CausesValidation="False" Text="匯入明細" Width="88px" CssClass="asp_button_M" />(必須為csv格式)
                                                        <asp:HyperLink ID="HyperLink3" runat="server" CssClass="font" NavigateUrl="../../Doc/SheetCost_Imp2.zip">下載匯入格式檔</asp:HyperLink>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="bluecol" width="6%"></td>
                                                    <td class="bluecol" width="8%">項次 </td>
                                                    <td class="bluecol" width="14%">品名 </td>
                                                    <td class="bluecol" width="12%">規格 </td>
                                                    <td class="bluecol" width="8%">單位 </td>
                                                    <td class="bluecol" width="10%">使用數量 </td>
                                                    <td class="bluecol">用途說明 </td>
                                                    <td class="bluecol" width="8%">功能 </td>
                                                </tr>
                                                <tr>
                                                    <td colspan="8" class="whitecol">
                                                        <asp:Label ID="Labexptitle1" runat="server" Text="填寫範例" ForeColor="Red"></asp:Label></td>
                                                </tr>
                                                <tr>
                                                    <td class="bluecol"></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:TextBox ID="exptItemNo8" runat="server" Columns="5" MaxLength="5" ReadOnly="True" ForeColor="#CCCCCC">1</asp:TextBox></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:TextBox ID="exptCName8" runat="server" Columns="10" MaxLength="30" ReadOnly="True" ForeColor="#CCCCCC">書籍/講義</asp:TextBox></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:TextBox ID="exptStandards8" runat="server" Columns="10" MaxLength="300" ReadOnly="True" ForeColor="#CCCCCC">書名出版社/講義</asp:TextBox></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:TextBox ID="exptUnit8" runat="server" Columns="5" MaxLength="5" ReadOnly="True" ForeColor="#CCCCCC">本/冊</asp:TextBox></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:TextBox ID="exptAllCount8" runat="server" Columns="5" MaxLength="5" ReadOnly="True" ForeColor="#CCCCCC">30</asp:TextBox></td>
                                                    <td class="whitecol">
                                                        <asp:TextBox ID="exptPurPose8" runat="server" Columns="35" MaxLength="300" ReadOnly="True" ForeColor="#CCCCCC" Width="90%">學科教學使用</asp:TextBox></td>
                                                    <td class="whitecol" style="text-align: center;"></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">
                                                        <asp:Button ID="BtnDelDG8" runat="server" Text="勾選刪除" CssClass="asp_button_M" CommandName="BtnDelDG8" /></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:TextBox ID="tItemNo8" runat="server" Columns="5" MaxLength="5"></asp:TextBox></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:TextBox ID="tCName8" runat="server" Columns="10" MaxLength="30"></asp:TextBox></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:TextBox ID="tStandards8" runat="server" Columns="10" MaxLength="300"></asp:TextBox></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:TextBox ID="tUnit8" runat="server" Columns="5" MaxLength="5"></asp:TextBox></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:TextBox ID="tAllCount8" runat="server" Columns="5" MaxLength="5"></asp:TextBox></td>
                                                    <td class="whitecol">
                                                        <asp:TextBox ID="tPurPose8" runat="server" Columns="35" MaxLength="300" Width="90%"></asp:TextBox></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:Button ID="btnAddCost8" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_M"></asp:Button></td>
                                                </tr>
                                                <tr>
                                                    <td colspan="8">
                                                        <table id="DataGrid8Table" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                                            <tr>
                                                                <td>
                                                                    <asp:DataGrid ID="DataGrid8" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" BorderColor="Gray">
                                                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                                                        <Columns>
                                                                            <asp:TemplateColumn HeaderText="勾選刪除">
                                                                                <HeaderTemplate>勾選刪除</HeaderTemplate>
                                                                                <ItemTemplate>
                                                                                    <asp:CheckBox ID="CheckBoxDG8" runat="server" />
                                                                                </ItemTemplate>
                                                                                <HeaderStyle VerticalAlign="Middle" HorizontalAlign="Center" Width="6%" />
                                                                                <ItemStyle VerticalAlign="Middle" HorizontalAlign="Center" CssClass="whitecol" />
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="項次">
                                                                                <HeaderStyle Width="8%"></HeaderStyle>
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
                                                                                <HeaderStyle Width="12%"></HeaderStyle>
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="lStandards8" runat="server"></asp:Label>
                                                                                </ItemTemplate>
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="eStandards8" runat="server" Columns="10" MaxLength="300"></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="單位">
                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                <ItemStyle HorizontalAlign="Center" />
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="lUnit8" runat="server"></asp:Label>
                                                                                </ItemTemplate>
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="eUnit8" runat="server" Columns="3" MaxLength="5"></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="使用數量">
                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                <ItemStyle HorizontalAlign="Center" />
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="lAllCount8" runat="server"></asp:Label>
                                                                                </ItemTemplate>
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="eAllCount8" runat="server" Columns="3" MaxLength="5"></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="訓練人數">
                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                <ItemStyle HorizontalAlign="Center" />
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="lTNum8" runat="server"></asp:Label>
                                                                                </ItemTemplate>
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="eTNum8" runat="server" Columns="3"></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="用途說明" ItemStyle-CssClass="whitecol">
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="lPurPose8" runat="server"></asp:Label>
                                                                                </ItemTemplate>
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="ePurPose8" runat="server" Columns="35" MaxLength="300" Width="90%"></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="功能">
                                                                                <ItemStyle Width="8%" HorizontalAlign="Center"></ItemStyle>
                                                                                <ItemTemplate>
                                                                                    <asp:HiddenField ID="Hid_DataKey" runat="server" />
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
                                                        </table>
                                                    </td>
                                                </tr>
                                            </table>
                                            <table id="TableCost9" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                                <tr>
                                                    <td class="table_title" colspan="8">其他明細 </td>
                                                </tr>
                                                <tr>
                                                    <td colspan="8" class="whitecol">
                                                        <asp:Label ID="Labeld9" runat="server" CssClass="font">匯入明細</asp:Label>
                                                        <input id="File4" type="file" size="50" name="File4" runat="server" accept=".csv,.xls" />
                                                        <asp:Button ID="BtnImport9" runat="server" CausesValidation="False" Text="匯入明細" CssClass="asp_button_M" />(必須為csv格式)
                                                        <asp:HyperLink ID="HyperLink4" runat="server" CssClass="font" NavigateUrl="../../Doc/OtherCost_Imp2.zip">下載匯入格式檔</asp:HyperLink>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="bluecol" style="width: 6%;"></td>
                                                    <td class="bluecol" style="width: 8%;">項次 </td>
                                                    <td class="bluecol" style="width: 14%;">項目 </td>
                                                    <td class="bluecol" style="width: 12%;">規格 </td>
                                                    <td class="bluecol" style="width: 8%;">單位 </td>
                                                    <td class="bluecol" style="width: 10%;">使用數量 </td>
                                                    <td class="bluecol">用途說明 </td>
                                                    <td class="bluecol" style="width: 8%;">功能 </td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">
                                                        <asp:Button ID="BtnDelDG9" runat="server" Text="勾選刪除" CssClass="asp_button_M" CommandName="BtnDelDG9" /></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:TextBox ID="tItemNo9" runat="server" Columns="5" MaxLength="5"></asp:TextBox></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:TextBox ID="tCName9" runat="server" Columns="10" MaxLength="30"></asp:TextBox></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:TextBox ID="tStandards9" runat="server" Columns="10" MaxLength="300"></asp:TextBox></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:TextBox ID="tUnit9" runat="server" Columns="5" MaxLength="5"></asp:TextBox></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:TextBox ID="tAllCount9" runat="server" Columns="5" MaxLength="5"></asp:TextBox></td>
                                                    <td class="whitecol">
                                                        <asp:TextBox ID="tPurpose9" runat="server" Columns="35" MaxLength="300" Width="90%"></asp:TextBox></td>
                                                    <td class="whitecol" style="text-align: center;">
                                                        <asp:Button ID="btnAddCost9" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_M"></asp:Button></td>
                                                </tr>
                                                <tr>
                                                    <td colspan="8">
                                                        <table id="DataGrid9Table" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                                            <tr>
                                                                <td>
                                                                    <asp:DataGrid ID="DataGrid9" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" BorderColor="Gray">
                                                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                                                        <Columns>
                                                                            <asp:TemplateColumn HeaderText="勾選刪除">
                                                                                <HeaderTemplate>勾選刪除</HeaderTemplate>
                                                                                <ItemTemplate>
                                                                                    <asp:CheckBox ID="CheckBoxDG9" runat="server" />
                                                                                </ItemTemplate>
                                                                                <HeaderStyle VerticalAlign="Middle" HorizontalAlign="Center" Width="6%" />
                                                                                <ItemStyle VerticalAlign="Middle" HorizontalAlign="Center" CssClass="whitecol" />
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="項次">
                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="lItemNo9" runat="server"></asp:Label>
                                                                                </ItemTemplate>
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="eItemNo9" runat="server" Columns="5" MaxLength="5"></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="項目">
                                                                                <HeaderStyle Width="14%"></HeaderStyle>
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="lCName9" runat="server"></asp:Label>
                                                                                </ItemTemplate>
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="eCName9" runat="server" Columns="10" MaxLength="30"></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="規格">
                                                                                <HeaderStyle Width="12%"></HeaderStyle>
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="lStandards9" runat="server"></asp:Label>
                                                                                </ItemTemplate>
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="eStandards9" runat="server" Columns="10" MaxLength="300"></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="單位">
                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                <ItemStyle HorizontalAlign="Center" />
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="lUnit9" runat="server"></asp:Label>
                                                                                </ItemTemplate>
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="eUnit9" runat="server" Columns="3" MaxLength="5"></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="使用數量">
                                                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                                                <ItemStyle HorizontalAlign="Center" />
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="lAllCount9" runat="server"></asp:Label>
                                                                                </ItemTemplate>
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="eAllCount9" runat="server" Columns="3" MaxLength="5"></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="訓練人數">
                                                                                <HeaderStyle Width="10%"></HeaderStyle>
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="lTNum9" runat="server"></asp:Label>
                                                                                </ItemTemplate>
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="eTNum9" runat="server" Columns="3"></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="用途說明" ItemStyle-CssClass="whitecol">
                                                                                <ItemTemplate>
                                                                                    <asp:Label ID="lPurPose9" runat="server"></asp:Label>
                                                                                </ItemTemplate>
                                                                                <EditItemTemplate>
                                                                                    <asp:TextBox ID="ePurPose9" runat="server" Columns="35" MaxLength="300" Width="90%"></asp:TextBox>
                                                                                </EditItemTemplate>
                                                                            </asp:TemplateColumn>
                                                                            <asp:TemplateColumn HeaderText="功能">
                                                                                <ItemStyle Width="8%" HorizontalAlign="Center"></ItemStyle>
                                                                                <ItemTemplate>
                                                                                    <asp:HiddenField ID="Hid_DataKey" runat="server" />
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
                                                        </table>
                                                    </td>
                                                </tr>
                                            </table>
                                            <table id="TableNote2" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                                <tr>
                                                    <td>
                                                        <table id="tbtNote2" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                                            <tr>
                                                                <td class="table_title" colspan="2">其他說明 </td>
                                                            </tr>
                                                            <tr>
                                                                <td colspan="2" class="whitecol">
                                                                    <asp:TextBox ID="tNote2" runat="server" Columns="77" Rows="5" Width="70%" TextMode="MultiLine" ToolTip="其他說明(欄位字數為1000)"></asp:TextBox>
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
                    <table id="TableLay7" border="0" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <%--<asp:Label ID="Label4c" runat="server" ForeColor="Black">(請先確認訓練時數、人數、固定費用總額) </asp:Label>--%>
                            <td class="bluecol_need" style="width: 16%;">經費來源 </td>
                            <td class="whitecol" style="width: 84%;">
                                <table id="FundingsourceTable" border="0" cellspacing="1" cellpadding="1" width="100%" class="whitecol">
                                    <tr>
                                        <td>訓練費用總額：新臺幣
                                            <asp:TextBox ID="tTrainingCost" runat="server" onfocus="this.blur()" Columns="7">0</asp:TextBox>元
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>每人總訓練費用：新臺幣
                                            <asp:TextBox ID="tPerPersonCost" runat="server" onfocus="this.blur()" Columns="7">0</asp:TextBox>元
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>政府補助金額：新臺幣(每期費用)
                                            <asp:TextBox ID="DefGovCost" runat="server" onfocus="this.blur()" Columns="7">0</asp:TextBox>元/每班人數
                                            <asp:Label ID="TNum1" runat="server"></asp:Label>=
                                            <asp:Label ID="Total1" runat="server"></asp:Label>元
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>學員負擔金額：新臺幣(每期費用)
                                            <asp:TextBox ID="DefStdCost" runat="server" onfocus="this.blur()" Columns="7">0</asp:TextBox>元/每班人數
                                            <asp:Label ID="TNum3" runat="server"></asp:Label>=
                                            <asp:Label ID="Total3" runat="server"></asp:Label>元
                                            <asp:CustomValidator ID="CustomValidator1" runat="server" Display="None" ErrorMessage="請輸入經費來源" ClientValidationFunction="CheckDef"></asp:CustomValidator>
                                            <asp:CustomValidator ID="Customvalidator7" runat="server" Display="None" ErrorMessage="經費來源-每人的政府補助金額與學員負擔金額應為整數數字" ClientValidationFunction="CheckDef2"></asp:CustomValidator>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                    <table id="TableLay8" border="0" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol_need" width="16%">經費分類代碼 </td>
                            <td class="whitecol" width="84%">
                                <asp:TextBox ID="GCIDName" runat="server" onfocus="this.blur()" Columns="88" Width="77%"></asp:TextBox>
                                <input id="btn_GCID" onclick="Get_GovClass('GCIDName');" value="..." type="button" runat="server" class="button_b_Mini" />
                                <input id="GCIDValue" type="hidden" runat="server" />
                                <input id="GCID1Value" type="hidden" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練費用編列說明</td>
                            <td class="whitecol">
                                <table width="100%">
                                    <tr>
                                        <td width="70%">
                                            <asp:TextBox ID="Note" runat="server" Columns="88" Rows="8" TextMode="MultiLine" Width="95%"></asp:TextBox></td>
                                        <td width="30%">
                                            <br />
                                            <asp:Button ID="Button21b" runat="server" Text="匯出EXCEL" CausesValidation="False" CssClass="asp_Export_M"></asp:Button></td>
                                    </tr>
                                </table>
                                <%--color="red"--%><span style="color: #FF0000"><asp:Label ID="Labmsg3" runat="server"></asp:Label></span>
                            </td>
                        </tr>
                    </table>
                    <table id="TableLay9" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td class="whitecol">
                                <table cellspacing="1" cellpadding="1" border="0" width="100%">
                                    <tr>
                                        <td class="bluecol_need" width="16%">教學方法 </td>
                                        <td class="whitecol">
                                            <asp:CheckBoxList ID="cblTMethod" runat="server" CssClass="font" RepeatLayout="Flow">
                                                <asp:ListItem Value="01">講授教學法（運用敘述或講演的方式，傳遞教材知識的一種教學方法，提供相關教材或講義）</asp:ListItem>
                                                <asp:ListItem Value="02">討論教學法（指團體成員齊聚一起，經由說、聽和觀察的過程，彼此溝通意見，由講師帶領達成教學目標）</asp:ListItem>
                                                <asp:ListItem Value="03">演練教學法（由講師的帶領下透過設備或教材，進行練習、表現和實作，親自解說示範的技能或程序的一種教學方法）</asp:ListItem>
                                                <asp:ListItem Value="99">其他教學方法：</asp:ListItem>
                                            </asp:CheckBoxList>
                                            <asp:TextBox runat="server" ID="TMethodOth" Width="77%" MaxLength="100"></asp:TextBox><br />
                                            <asp:Label ID="Label1" runat="server" ForeColor="red">(若選"其他教學方法"，需填寫輸入，上限100個字)</asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need" width="16%">訓練需求調查</td>
                                        <td class="whitecol">
                                            <table cellspacing="1" cellpadding="1" width="100%" border="0">
                                                <tr>
                                                    <td>產業人力需求調查：<asp:Label ID="Label5" runat="server" ForeColor="red">(填寫輸入，上限1000個字)</asp:Label></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">
                                                        <asp:TextBox ID="tPOWERNEED1" runat="server" Width="70%" Height="80px" TextMode="MultiLine" MaxLength="1000" placeholder="(應論述調查期間、區域範圍、調查對象、產業發展趨勢及該產業之訓練需求)"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td>區域人力需求調查：<asp:Label ID="Label6" runat="server" ForeColor="red">(填寫輸入，上限1000個字)</asp:Label></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">
                                                        <asp:TextBox ID="tPOWERNEED2" runat="server" Width="70%" Height="80px" TextMode="MultiLine" MaxLength="1000" placeholder="(依產業人力需求調查結果，進行區域性的人力需求調查，應論述調查期間、區域範圍、調查對象及該產業於該區域之訓練需求)"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td>訓練需求概述：<asp:Label ID="Label9" runat="server" ForeColor="red">(填寫輸入，上限200個字)</asp:Label>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">
                                                        <asp:TextBox ID="tPOWERNEED3" runat="server" Width="70%" Height="80px" TextMode="MultiLine" MaxLength="200"></asp:TextBox></td>
                                                </tr>
                                                <tr id="trPOLICYREL_t" runat="server">
                                                    <td>與政策性產業課程之關聯性概述：<asp:Label ID="Label11" runat="server" ForeColor="red">(填寫輸入，上限500個字)</asp:Label>
                                                    </td>
                                                </tr>
                                                <tr id="trPOLICYREL_c" runat="server">
                                                    <td class="whitecol">
                                                        <asp:HiddenField ID="Hid_tPOLICYREL" runat="server" />
                                                        <asp:TextBox ID="tPOLICYREL" runat="server" Width="70%" Height="80px" TextMode="MultiLine" MaxLength="500" placeholder="(有勾選政府政策性產業，須填寫「與政策性產業課程之關聯性概述」，若無勾選政府政策性產業，則不儲存此欄資料)"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td>
                                                        <asp:CheckBox ID="cbPOWERNEED4" runat="server" />課程須符合目的事業主管機關相關規定：<asp:Label ID="Label10" runat="server" ForeColor="red">(填寫輸入，上限200個字)</asp:Label>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">
                                                        <asp:TextBox ID="tPOWERNEED4" runat="server" Width="70%" Height="80px" TextMode="MultiLine" MaxLength="200" placeholder="(如為目的事業主管機關已定有訓練課程、時數、參訓人員資格認定及程序等相關規定者，應依其規定辦理，並加以說明規定內容。)"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(是否瞭解區域產業需求)<br />
                                                        <asp:Label ID="lab_REPORTE" runat="server" Text="本課程是否應報請主管機關核備" ForeColor="Red"></asp:Label>
                                                        <asp:RadioButton ID="rbl_REPORTE_Y" runat="server" Text="是" GroupName="rbl_REPORTE" />
                                                        <asp:RadioButton ID="rbl_REPORTE_N" runat="server" Text="否" GroupName="rbl_REPORTE" /></td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need" width="16%">訓練目標</td>
                                        <td>
                                            <table cellspacing="1" cellpadding="1" width="100%" border="0">
                                                <tr>
                                                    <td class="bluecol_need" style="text-align: center;">職能級別：(單選) </td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">
                                                        <asp:RadioButtonList ID="rblFuncLevel" runat="server" RepeatLayout="Flow">
                                                            <asp:ListItem Value="01">級別1(能夠在可預計及有規律的情況中，在密切監督及清楚指示下，執行常規性及重複性的工作。且通常不需要特殊訓練、教育及專業知識與技術)</asp:ListItem>
                                                            <asp:ListItem Value="02">級別2(能夠在大部分可預計及有規律的情況中，在經常性監督下，按指導進行需要某些判斷及理解性的工作。需具備基本知識、技術)</asp:ListItem>
                                                            <asp:ListItem Value="03">級別3(能夠在部分變動及非常規性的情況中，在一般監督下，獨立完成工作。需要一定程度的專業知識與技術及少許的判斷能力)</asp:ListItem>
                                                            <asp:ListItem Value="04">級別4(能夠在經常變動的情況中，在少許監督下，獨立執行涉及規劃設計且需要熟練技巧的工作。需要具備相當的專業知識與技術，及作判斷及決定的能力)</asp:ListItem>
                                                            <asp:ListItem Value="05">級別5(能夠在複雜變動的情況中，在最少監督下，自主完成工作。需要具備應用、整合、系統化的專業知識與技術及策略思考與判斷能力)</asp:ListItem>
                                                            <asp:ListItem Value="06">級別6(能夠在高度複雜變動的情況中，應用整合的專業知識與技術，獨立完成專業與創新的工作。需要具備策略思考、決策及原創能力)</asp:ListItem>
                                                        </asp:RadioButtonList>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td colspan="2">
                                            <table class="font" id="tbTeacherDesc_AB" style="width: 100%;" cellspacing="1" cellpadding="1" border="0">
                                                <tr>
                                                    <td colspan="6" class="whitecol">
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
                                                                                        <asp:TextBox ID="TeacherDesc" runat="server" Width="80%" Height="80px" TextMode="MultiLine" onfocus="this.blur()"></asp:TextBox>
                                                                                        <input id="btn_TCTYPEA" type="button" value="..." runat="server" class="button_b_Mini" />
                                                                                    </ItemTemplate>
                                                                                </asp:TemplateColumn>
                                                                            </Columns>
                                                                        </asp:DataGrid>
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </div>
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
                                                                                        <asp:TextBox ID="TeacherDesc" runat="server" Width="80%" Height="80px" TextMode="MultiLine" onfocus="this.blur()"></asp:TextBox>
                                                                                        <input id="btn_TCTYPEB" type="button" value="..." runat="server" class="button_b_Mini" />
                                                                                    </ItemTemplate>
                                                                                </asp:TemplateColumn>
                                                                            </Columns>
                                                                        </asp:DataGrid>
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </div>
                                                        <%--Fundingsource--%>                                                        <%--<asp:Label ID="ProLicense" runat="server"></asp:Label>--%>                                                        <%--<asp:Label ID="ProLicense" runat="server"></asp:Label>--%>                                                        <%--<td class="bluecol_need">授課教師 - 遴選辦法說明</td>--%>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="table_title" align="center" colspan="2">受訓資格</td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need" width="16%">學歷</td>
                                        <td class="whitecol" width="84%">
                                            <asp:DropDownList ID="Degree" runat="server" AppendDataBoundItems="True"></asp:DropDownList>(含以上)
                                            <asp:RequiredFieldValidator ID="fill5" runat="server" ControlToValidate="Degree" Display="None" ErrorMessage="請選擇受訓資格「學歷」"></asp:RequiredFieldValidator>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need">年齡 </td>
                                        <td class="whitecol">
                                            <asp:RadioButton ID="rdoAge1" runat="server" Checked="True" GroupName="GroupAge" />
                                            <asp:Label ID="l_Age" runat="server">年滿15歲以上</asp:Label>
                                            <asp:RadioButton ID="rdoAge2" runat="server" GroupName="GroupAge" />
                                            <asp:Label ID="l_Age2a" runat="server">應符合相關法規須年滿</asp:Label>
                                            <asp:TextBox ID="txtAge1" runat="server" Width="40px" MaxLength="2"></asp:TextBox>
                                            <asp:Label ID="l_Age2b" runat="server">以上</asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need">學員資格</td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="CapAll" runat="server" TextMode="MultiLine" Columns="88" Rows="6" Width="70%"></asp:TextBox>
                                            <asp:HiddenField ID="Hid_D25_8_CapAll_MSG" runat="server" />
                                            <asp:HiddenField ID="Hid_CapAll" runat="server" />
                                            <br />
                                            <span class="red-style1">※申請政策性產業 -「職場續航」之課程，學員資格欄位將由系統自動帶入，不提供訓練單位增修。</span>
                                        </td>
                                    </tr>
                                </table>
                                <table class="font" id="Table_2" style="width: 100%;" cellspacing="1" cellpadding="1" border="0">
                                    <tr>
                                        <td class="table_title" align="center" colspan="2">裝備與設施</td>
                                        <%--<asp:TextBox ID="TeacherDesc_A" runat="server" Width="60%" Height="80px" TextMode="MultiLine"></asp:TextBox>
<input id="btn_TCTYPEA" type="button" value="..." runat="server" class="button_b_Mini" />--%>
                                    </tr>
                                    <tr>
                                        <td colspan="2">
                                            <table id="Table_2_D" style="width: 100%;" cellspacing="1" cellpadding="1" border="0">
                                                <tr id="TRA1" runat="server">
                                                    <td rowspan="2" width="14%" class="bluecol">學科場地 </td>
                                                    <td id="TD11" runat="server" width="14%" class="bluecol">容納人數 </td>
                                                    <td width="72%" class="whitecol">
                                                        <asp:TextBox ID="T2Dnum2" runat="server" Width="20%" onfocus="this.blur()"></asp:TextBox></td>
                                                </tr>
                                                <tr id="TRA2" runat="server">
                                                    <td class="bluecol">硬體設施說明 </td>
                                                    <td class="whitecol">
                                                        <asp:TextBox ID="HwDesc2" runat="server" Width="70%" onfocus="this.blur()" Height="80px" TextMode="MultiLine"></asp:TextBox></td>
                                                </tr>
                                                <tr id="TRB1" runat="server">
                                                    <td rowspan="2" class="bluecol">術科場地&nbsp; </td>
                                                    <td class="bluecol">容納人數 </td>
                                                    <td class="whitecol">
                                                        <asp:TextBox ID="T2Dnum3" runat="server" Width="20%" onfocus="this.blur()"></asp:TextBox></td>
                                                </tr>
                                                <tr id="TRB2" runat="server">
                                                    <td class="bluecol">硬體設施說明 </td>
                                                    <td class="whitecol">
                                                        <asp:TextBox ID="HwDesc3" runat="server" Width="70%" onfocus="this.blur()" Height="78px" TextMode="MultiLine"></asp:TextBox></td>
                                                </tr>
                                                <%--<tr id="trOtherDesc23" runat="server"><td colspan="2" class="bluecol">其他器材設備 </td>
                                                    <td><asp:TextBox ID="OtherDesc23" runat="server" onfocus="this.blur()"   TextMode="MultiLine" Rows="3" Width="70%"></asp:TextBox></td></tr>--%>
                                                <tr>
                                                    <td colspan="2" class="bluecol">其他設施說明 </td>
                                                    <td class="whitecol">
                                                        <asp:TextBox ID="txtOthFacDesc23" runat="server" Width="70%" TextMode="MultiLine" Rows="6" MaxLength="3000"></asp:TextBox></td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                </table>
                                <table class="font" id="Table_4_D" style="width: 100%;" cellspacing="1" cellpadding="1" border="0">
                                    <tr>
                                        <td class="table_title" align="center" colspan="2">訓練績效評估</td>
                                        <%--<td class="bluecol_need">授課助教 - 遴選辦法說明</td>--%>
                                    </tr>
                                    <tr>
                                        <td bgcolor="#ffcccc" class="bluecol" width="16%">1.反應評估 </td>
                                        <td class="whitecol" width="84%">
                                            <asp:CheckBox ID="chk_RecDesc" runat="server" />
                                            （是評量學員對訓練的觀感，以量化評量的方式來設計課後評量表，衡量學員對於訓練的反應，例如:設計滿意度調查機制瞭解學員感受包括知識、學習後關聯性、行政作業、課程是否值得推薦等）<br />
                                            <asp:TextBox ID="RecDesc" runat="server" Width="60%" MaxLength="500"></asp:TextBox>(滿意度調查機制)
                                        </td>
                                    </tr>
                                    <tr>
                                        <td bgcolor="#ffcccc" class="bluecol" width="16%">2.學習評估 </td>
                                        <td class="whitecol" width="84%">
                                            <asp:CheckBox ID="chk_LearnDesc" runat="server" />
                                            （是評量學員因為參與訓練而改變態度、增進知識技能的程度。在此階段是關於學員在課程加強知識或是技巧的延伸，學習的評量則可經由課前測驗與課後測驗來達成，即可判斷訓練課程的成效。例如:考試或報告機制）<br />
                                            <asp:TextBox ID="LearnDesc" runat="server" Width="60%" MaxLength="500"></asp:TextBox>(考試或報告機制)
                                        </td>
                                    </tr>
                                    <tr>
                                        <td bgcolor="#ffcccc" class="bluecol" width="16%">3.行為評估 </td>
                                        <td class="whitecol" width="84%">
                                            <asp:CheckBox ID="chk_ActDesc" runat="server" />
                                            （是評量學員因參與訓練而產生工作行為上的改變程度。經過3到6個月的訓練後，可對學員與其主管以問卷、面談、直接觀察、360度績效考評、目標設定等調查方法來評量，評量學員是否真的依照訓練的結果改變工作的模式。例如:課後行動計畫調查機制）<br />
                                            <asp:TextBox ID="ActDesc" runat="server" Width="60%" MaxLength="500"></asp:TextBox>(課後行動計畫調查機制)
                                        </td>
                                    </tr>
                                    <tr>
                                        <td bgcolor="#ffcccc" class="bluecol" width="16%">4.成果評估 </td>
                                        <td class="whitecol" width="84%">
                                            <asp:CheckBox ID="chk_ResultDesc" runat="server" />
                                            （是評量因為參與訓練而產生的最後結果，如銷售額提升、成本降低、績效提升等，同時也是回應到參與訓練的理由。例如:(1)提升較高的客戶滿意度、(2)提高產值、(3)提高銷售額、(4)增加更多的新客戶、(5)降低更多的成本、(6)提高利潤）<br />
                                            <asp:TextBox ID="ResultDesc" runat="server" Width="60%" MaxLength="500"></asp:TextBox>(工作行動調查機制)
                                        </td>
                                    </tr>
                                    <tr>
                                        <td bgcolor="#ffcccc" class="bluecol" width="16%">5.其它機制 </td>
                                        <td class="whitecol" width="84%">
                                            <asp:CheckBox ID="chk_OtherDesc" runat="server" /><br />
                                            <asp:TextBox ID="OtherDesc" runat="server" Width="60%" MaxLength="500"></asp:TextBox>
                                        </td>
                                    </tr>
                                </table>
                                <table class="font" id="Table_5_D" style="width: 100%;" cellspacing="1" cellpadding="1" border="0">
                                    <tr>
                                        <td class="table_title" align="center" colspan="2">促進學習機制</td>
                                        <%--<asp:TextBox ID="TeacherDesc_B" runat="server" Width="60%" Height="80px" TextMode="MultiLine"></asp:TextBox><input id="btn_TCTYPEB" type="button" value="..." runat="server" class="button_b_Mini" />--%>
                                    </tr>
                                    <tr>
                                        <td width="16%" bgcolor="#ffcccc" class="bluecol_need">是否為iCAP課程</td>
                                        <td width="84%" class="whitecol">
                                            <br />
                                            <table width="100%" id="tb_ISiCAPCOUR" runat="server">
                                                <tr>
                                                    <td width="20%" class="whitecol">
                                                        <%--<td bgcolor="#FFCC00">二、裝備與設施 </td>--%>
                                                        <asp:RadioButton ID="RB_ISiCAPCOUR_Y" runat="server" Text="是,請填寫" GroupName="RB_ISiCAPCOUR" />
                                                        <br />
                                                        <br />
                                                        <asp:RadioButton ID="RB_ISiCAPCOUR_N" runat="server" Text="否" Checked="true" GroupName="RB_ISiCAPCOUR" />
                                                    </td>
                                                    <td class="whitecol" valign="top">
                                                        <%--<td class="bluecol_need" style="text-align: center;" colspan="2">訓練績效評估 </td>--%>
                                                        <asp:Label ID="lab_iCAPCOURDESC" runat="server" Text="iCAP課程相關說明"></asp:Label><br />
                                                        <asp:TextBox ID="iCAPCOURDESC" runat="server" Width="80%" Height="80px" TextMode="MultiLine" MaxLength="500"></asp:TextBox>

                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td bgcolor="#ffcccc" class="bluecol_need">招訓方式</td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="Recruit" runat="server" Width="80%" Height="80px" TextMode="MultiLine"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td bgcolor="#ffcccc" class="bluecol_need">遴選方式</td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="Selmethod" runat="server" Width="80%" Height="80px" TextMode="MultiLine"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td bgcolor="#ffcccc" class="bluecol_need">學員激勵辦法 </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="Inspire" runat="server" Width="80%" Height="80px" TextMode="MultiLine"></asp:TextBox></td>
                                    </tr>
                                </table>
                                <table class="font" id="Table_7b" width="100%" cellpadding="1" cellspacing="1">
                                    <tr>
                                        <td class="table_title" align="center" colspan="2">其他</td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="16%">學員是否可依個人需求參加政府機關辦理相關證照考試或技能檢定 </td>
                                        <td class="whitecol" width="84%">
                                            <asp:CheckBox ID="TGovExamCY" runat="server" CssClass="group-check-tgov" />是。<asp:TextBox ID="GOVAGENAME" runat="server" MaxLength="50" Width="28%" placeholder="(政府機關名稱)"></asp:TextBox>&nbsp;<asp:TextBox ID="TGovExamName" runat="server" MaxLength="50" Width="28%" placeholder="(證照或檢定名稱)"></asp:TextBox><br />
                                            <asp:CheckBox ID="TGovExamCN" runat="server" CssClass="group-check-tgov" />否。(包含非政府機關辦理相關證照或檢定)<br />
                                            <asp:CheckBox ID="TGovExamCG" runat="server" CssClass="group-check-tgov" />本課程結訓後須參加環境部辦理之淨零綠領人才培育課程測驗；測驗成績達及格，即可申請本方案補助。
                                        </td>
                                    </tr>
                                </table>
                                <table class="font" id="Table_8" style="width: 100%;" cellspacing="1" cellpadding="1" border="0">
                                    <tr>
                                        <td class="bluecol_need" width="16%">備註 </td>
                                        <td>
                                            <table class="font" id="Table_8b" width="100%" cellpadding="1" cellspacing="1">
                                                <tr>
                                                    <td class="whitecol">&nbsp;&nbsp;&nbsp;<asp:CheckBox ID="chkMEMO8C1" runat="server" />
                                                        <asp:Label ID="lbMEMO8" runat="server"></asp:Label><br />
                                                        &nbsp;&nbsp;(課程內容類似職業安全衛生教育訓練且不報請主管機關核備者，應點選此項，避免民眾誤解可作為時數認列)。<br />
                                                        <br />
                                                        &nbsp;&nbsp;&nbsp;<asp:CheckBox ID="chkMEMO8C2" runat="server" />
                                                        <asp:TextBox ID="txtMemo8" runat="server" Width="60%" MaxLength="500"></asp:TextBox><br />
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">&nbsp;</td>
                                                </tr>
                                            </table>
                                            <input id="hidmemo8" type="hidden" name="hidmemo8" runat="server" />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need" width="16%">專長能力標籤 </td>
                                        <td>
                                            <%-- <script type="text/javascript">
                                                $(document).ready(function () {
                                                    //re-binding autocomplete input event
                                                    autocomplete_init();
                                                    //pageLoad();
                                                });
                                                function pageLoad() {
                                                    //re-binding autocomplete input event
                                                    //autocomplete_init();
                                                    //bindAbilityEdit();
                                                    //unblockUI();
                                                }
                                            </script>--%>
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
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol">
                                <div align="center">
                                    <table width="100%" cellpadding="1" cellspacing="1">
                                        <tr>
                                            <td class="whitecol" width="20%"></td>
                                            <td class="whitecol" width="60%" align="left">
                                                <span style="color: #FF0000">
                                                    <asp:Label ID="LabMsg7" runat="server"></asp:Label></span>
                                                <%--<font color="red"></font>--%>
                                            </td>
                                            <td class="whitecol" width="20%"></td>
                                        </tr>
                                    </table>
                                </div>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td class="whitecol">
                    <div align="center">
                        &nbsp;<asp:Button ID="Button8" runat="server" Text="1.草稿儲存" CausesValidation="False" CssClass="asp_Export_M"></asp:Button>
                        &nbsp;<asp:Button ID="btnAdd" runat="server" Text="2.基本儲存" CssClass="asp_Export_M"></asp:Button>
                        &nbsp;<asp:Button ID="Button24" runat="server" Text="回上一頁" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                        &nbsp;<asp:Button ID="BtnSAVE2" runat="server" Text="3.正式儲存" CssClass="asp_Export_M"></asp:Button>
                        <%--RB_ISiCAPCOUR--%>                        <%----%>
                        <asp:ValidationSummary ID="ValidationSummary1" runat="server" DisplayMode="List" ShowSummary="False" ShowMessageBox="True"></asp:ValidationSummary>
                    </div>
                </td>
            </tr>
        </table>
        <input id="LayerState" type="hidden" runat="server" />
        <%--<td class="bluecol_need" style="text-align: center;" colspan="2">其他 </td>--%>
        <input id="time_spent" onfocus="this.blur()" maxlength="256" size="5" type="hidden" name="time_spent" runat="server" />
        <input id="isBlack" type="hidden" name="isBlack" runat="server" />
        <input id="Blackorgname" type="hidden" name="Blackorgname" runat="server" />
        <input id="hTPlanID54" type="hidden" name="hTPlanID54" runat="server" />
        <input id="upt_PlanX" name="upt_PlanX" type="hidden" runat="server" />
        <asp:HiddenField ID="Hid_ComIDNO" runat="server" />
        <asp:HiddenField ID="Hid_PSNO28" runat="server" />
        <asp:HiddenField ID="Hid_PDF20171226" runat="server" />
        <asp:HiddenField ID="Hid_THours" runat="server" />
        <asp:HiddenField ID="Hid_TNum" runat="server" />
        <%--<asp:HiddenField ID="Hid_FIXExceeDesc" runat="server" />--%>
        <asp:HiddenField ID="Hid_PERC100" runat="server" />
        <asp:HiddenField ID="Hid_TotalCost1" runat="server" />
        <asp:HiddenField ID="Hid_sender1" runat="server" />
        <asp:HiddenField ID="Hid_CheckAddTime" runat="server" />
        <asp:HiddenField ID="Hid_ChkTDescH1" runat="server" />
        <asp:HiddenField ID="Hid_ChkTDescH2" runat="server" />
        <asp:HiddenField ID="Hid_ChkTDescH3" runat="server" />
        <asp:HiddenField ID="Hid_ChkTEACHHOURS1" runat="server" />
        <asp:HiddenField ID="Hid_sisyphus" runat="server" />
        <asp:HiddenField ID="Hid_GOVAGENAME_og" runat="server" />
        <asp:HiddenField ID="Hid_TGovExamName_og" runat="server" />
        <asp:HiddenField ID="Hid_REPORTE" runat="server" />
        <asp:HiddenField ID="Hid_TGovExam" runat="server" />

        <asp:HiddenField ID="hid_TrainDescTable_guid1" runat="server" />
        <asp:HiddenField ID="hid_PersonCostTable_guid1" runat="server" />
        <asp:HiddenField ID="hid_CommonCostTable_guid1" runat="server" />
        <asp:HiddenField ID="hid_SheetCostTable_guid1" runat="server" />
        <asp:HiddenField ID="hid_OtherCostTable_guid1" runat="server" />
        <asp:HiddenField ID="hid_planONCLASS_guid1" runat="server" />
        <asp:HiddenField ID="hid_PLAN_BUSPACKAGE_guid1" runat="server" />

        <asp:HiddenField ID="hid_AddressSciPTID" runat="server" />
        <asp:HiddenField ID="hid_AddressSciPTID2" runat="server" />
        <asp:HiddenField ID="hid_AddressTechPTID" runat="server" />
        <asp:HiddenField ID="hid_AddressTechPTID2" runat="server" />
        <asp:HiddenField ID="hid_MAX_iACTHUMCOST" runat="server" />
        <asp:HiddenField ID="Hid_OJT22071401" runat="server" />
        <asp:HiddenField ID="Hid_USE_CBLKID60_TP28" runat="server" />
        <asp:HiddenField ID="hfScrollToAnchor" runat="server" />
    </form>
</body>
</html>
