// 說明: 取得「縣市鄉鎮名稱」、「郵遞區碼」
// 範例1: onclick="getZip('zip.htm', 'txtAllName', 'txtZipId')"
// 範例2: onclick="getZip('zip.htm', 'txtAllName', 'txtZipId', 'txtCityId')"       //(可再多傳回「縣市代碼」)
function getZip(page, CityZipNameField, ZipIdField) {
    var CityIdField = "";
    if (arguments.length > 3) {
        CityIdField = arguments[3];
    }
    var mywin = window.open(page + "?city_id_field=" + CityIdField + "&zip_id_field=" + ZipIdField + "&all_name_field=" + CityZipNameField, "ZipWin", "toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=no,width=680,height=560");
    mywin.focus();
}

// 取得「縣市名稱」、「縣市代碼」、「郵遞區碼」、「鄉鎮名稱」
// 範例: onclick="getZipAll('zip.htm', 'txtCityId', 'txtCityName', 'txtZipId', 'txtZipName')"
function getZipAll(page, CityIdField, CityNameField, ZipIdField, ZipNameField) {
    var mywin = window.open(page + "?city_id_field=" + CityIdField + "&city_name_field=" + CityNameField + "&zip_id_field=" + ZipIdField + "&zip_name_field=" + ZipNameField, "ZipWin", "toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=no,width=680,height=560");
    mywin.focus();
}

function wopen(url, name, width, height, k) {
    //k:scrollbars
    if (!isNaN(k)) {
        if (k == 0) { k = "no" } else { k = "yes" }
    }
    LeftPosition = (screen.width) ? (screen.width - width) / 2 : 0;
    TopPosition = (screen.availHeight) ? (screen.availHeight - height) / 2 : 0;
    window.open(url, name, 'top=' + TopPosition + ',left=' + LeftPosition + ',width=' + width + ',height=' + height + ',resizable=yes,scrollbars=' + k + ',status=no');
}

// 說明: 取得縣市代碼、名稱
// 範例: onclick="getCity('zip.htm', 'txtCityId', 'txtCityName')"
function getCity(page, CityIdField, CityNameField) {
    var mywin = window.open(page + "?city_id_field=" + CityIdField + "&city_name_field=" + CityNameField + "&CityOnly=1", "CityWin", "toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=no,width=680,height=560");
    mywin.focus();
}

//取得標準職業分類代碼
function getJobId(page, JobIdField, JobNameField) {
    var mywin = window.open(page + "?job_id_field=" + JobIdField + "&job_name_field=" + JobNameField, "JobIdWin", "toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=no,width=800,height=660");
    mywin.focus();
}

//取得技能類別代碼
function getExamId(page, ExamIdField, ExamNameField) {
    var mywin = window.open(page + "?exam_id_field=" + ExamIdField + "&exam_name_field=" + ExamNameField, "ExamIdWin", "toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=no,width=800,height=660");
    mywin.focus();
}

//取得科系所代碼
function getDeptId(page, DeptIdField, DeptNameField) {
    var mywin = window.open(page + "?dept_id_field=" + DeptIdField + "&dept_name_field=" + DeptNameField, "JobIdWin", "toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=800,height=660");
    mywin.focus();
}

//顯示信件內容
function openMail(url) {
    wopen(url, 'mail', 700, 600, 1);
}

//開啟職類視窗
function openTrain(TMID) {
    wopen('../../Common/TrainJob.aspx?field=TB_career_id&TMID=' + TMID, 'TrainJob', 700, 560, 0);
}

//開啟業別視窗
function openTrain2(TMID) {
    wopen('../../Common/TrainJob.aspx?type=2&field=TB_career_id&TMID=' + TMID, 'Train2', 700, 560, 0);
}

//開啟通俗職類視窗
function openCjob(cjobValue) {
    wopen('../../Common/sendCjob.aspx?field=txtCJOB_NAME&cjobValue=' + cjobValue, 'sendCjob', 700, 360, 0);
}

//開啟 檢定職類與考試級別 視窗
function openExamC(NMGP, NMXM, NMLV, VLXM, VLLV, BTN1) {
    var uu1 = '../../Common/SendExamC.aspx?pg=1';
    uu1 += '&NMGP=' + NMGP + '&NMXM=' + NMXM + '&NMLV=' + NMLV;
    uu1 += '&VLXM=' + VLXM + '&VLLV=' + VLLV + '&BTN1=' + BTN1;
    wopen(uu1, 'SendExamC', 700, 360, 0);
}

//開啟班級查詢視窗
function openClass(page) {
    wopen(page, 'Class', 1300, 700, 1);
}

function openOrg(page) {
    wopen(page, '訓練機構', 800, 680, 1);
}

//開啟報表列印視窗
function openPrint(page) {
    wopen(page, '列印報表', 800, 680, 1);
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
        wopen(page + '?STDate=' + STDate + '&FTDate=' + FTDate + '&NowDate=' + NowDate + '&ValueField=' + obj + '&Button=' + btn, '', 420, 260);
    }
}

function PublicCalendar(obj) {
    var Today = new Date();
    var Years = Today.getFullYear();
    var Months = Today.getMonth() + 1;
    var Days = Today.getDate();
    var NowDate = '';
    NowDate = Years + '/' + Months + '/' + Days;
    openCalendar(obj, '1900/1/1', '2100/12/31', NowDate)
}

function showObj(obj) {
    if (document.getElementById(obj)) {
        if (document.getElementById(obj).style.display == 'none') {
            document.getElementById(obj).style.display = 'inline';
        }
        else {
            document.getElementById(obj).style.display = 'none';
        }
    }
}