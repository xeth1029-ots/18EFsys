var ie5 = (document.all && document.getElementById) ? true : false; // ie5+
var ns4 = (document.layers && !document.getElementById) ? true : false;
var ie4 = (document.all) ? true : false;
var ns6 = (document.getElementById && !document.all) ? true : false;
var ypos = 0;
var xpos = 0;

function MM_swapImgRestore() { //v3.0
    var i, x, a = document.MM_sr; for (i = 0; a && i < a.length && (x = a[i]) && x.oSrc; i++) x.src = x.oSrc;
}

function MM_preloadImages() { //v3.0
    var d = document; if (d.images) {
        if (!d.MM_p) d.MM_p = new Array();
        var i, j = d.MM_p.length, a = MM_preloadImages.arguments; for (i = 0; i < a.length; i++)
            if (a[i].indexOf("#") != 0) { d.MM_p[j] = new Image; d.MM_p[j++].src = a[i]; }
    }
}

function MM_findObj(n, d) { //v4.01
    var p, i, x; if (!d) d = document; if ((p = n.indexOf("?")) > 0 && parent.frames.length) {
        d = parent.frames[n.substring(p + 1)].document; n = n.substring(0, p);
    }
    if (!(x = d[n]) && d.all) x = d.all[n]; for (i = 0; !x && i < d.forms.length; i++) x = d.forms[i][n];
    for (i = 0; !x && d.layers && i < d.layers.length; i++) x = MM_findObj(n, d.layers[i].document);
    if (!x && d.getElementById) x = d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
    var i, j = 0, x, a = MM_swapImage.arguments; document.MM_sr = new Array; for (i = 0; i < (a.length - 2) ; i += 3)
        if ((x = MM_findObj(a[i])) != null) { document.MM_sr[j++] = x; if (!x.oSrc) x.oSrc = x.src; x.src = a[i + 2]; }
}

//
//取得中英文字串的總length；
//
function getLength(for_check) {
    var I, cnt = 0;
    for (I = 0; I < for_check.length; I++) {
        if (escape(for_check.charAt(I)).length >= 4) {
            cnt += 2;
        } else {
            cnt++;
        }
    }
    return cnt;
}

//
//檢查 是否為合法日期
//
function isDate(strDate) {

    if (strDate.length != 10) return false;

    var year, month, day;
    year = strDate.substring(0, 4);
    month = strDate.substring(5, 7);
    day = strDate.substring(8, 10);

    if (year.length <= 0 || month == "" || day == "") return false;
    if (isNaN(year)) return false;
    if (isNaN(month)) return false;

    if (isNaN(day)) return false;

    if (strDate.substring(4, 5) != "-" || strDate.substring(7, 8) != "-") return false;

    if (month.charAt(0) == '0') month = month.substring(1, 2);
    if (day.charAt(0) == '0') day = day.substring(1, 2);

    year = parseInt(year);
    month = parseInt(month);
    day = parseInt(day);

    if (year < 1900) return false;
    if (month < 1 || month > 13) return false;
    if (day < 1 || day > 31) return false;

    if (month == 2) {
        //
        // 1900 不是閏年，而 2000 是閏年
        //
        if ((year / 4) != Math.floor(year / 4) || year == 1900) {
            if (day > 28) return false;
        }
        else {
            if (day > 29) return false;
        }
    }

    //
    // 檢查是否為小月
    // 
    if (month == 2 || month == 4 || month == 6 || month == 9 || month == 11) {
        if (day > 30) return false;
    }

    return true;
}

/* 取得兩個日期之間所相差的天數
 * 參數日期格式必須為「MM/DD/YYYY」或者「MM-DD-YYYY」，可先使用YMDD2MDDY()作轉換
 * @param   date1	   開始日期字串1
 * @param   date2	   結束日期字串2
 * @return  integer
 */
function getDiffDay(date1, date2) {
    var laterdate, earlierdate;
    var difference;

    earlierdate = Date.parse(date1);
    laterdate = Date.parse(date2);

    difference = laterdate - earlierdate;

    return Math.floor(difference / 1000 / 60 / 60 / 24);
}

/* 將「年月日」的日期格式轉換為「月日年」字串
 * 傳入的日期格式可為「YYYY/MM/DD」或「YYYY-MM-DD」
 * @param   mydate	  欲轉換的日期字串
 * @return  string
 */
function YMDD2MDDY(mydate) {
    var pos;
    var separator = "-";
    var newdate;

    pos = mydate.indexOf("-");
    if (pos == -1) {
        pos = mydate.indexOf("/");
        separator = "/";
    }

    // IF 不是正確的日期格式 THEN 傳回NaN
    if (pos == -1) {
        return NaN;
    }

    newdate = mydate.substring(pos + 1) + separator + mydate.substring(0, pos);

    return newdate;

}
/* 判斷是否為英文字
 * @param   value	   欲檢查的值
 * @return  boolean
 */
function isEng(value) {
    var pattern = /^[A-Za-z\- ]+$/;
    return pattern.test(value);
}

/* 判斷是否為正整數或負整數，例如: +25, -33, 77 皆符合條件
 * @param   value	   欲檢查的值
 * @return  boolean
 */
function isInt(value) {
    var pattern = /^(\+|\-)?\d+$/;
    return pattern.test(value);
}

/* 判斷是否為全為數字，例如: 25, 002 皆符合條件 (有含正號、負號的整數不符合條件)
 * @param   value	   欲檢查的值
 * @return  boolean
 */
function isUnsignedInt(value) {
    var pattern = /^\d+$/;
    return pattern.test(value);
}


/* 檢查指定表單欄位物件的值是否為空值或空白字串，支援所有的表單物件型態 checkbox,
 * radio, select, text, textarea, hidden
 * @param   obj		 表單欄位物件
 * @return  boolean
 */
function isEmpty(obj) {
    var oform = document.forms[0];
    var result = false;
    var obj_name;

    if (typeof (obj) == "string") {
        obj_name = obj;
        obj = oform.elements[obj_name];
        if (typeof (obj) != "object") {
            var valuelist = getCheckBoxListValue(obj_name);
            if (parseInt(valuelist, 10) == 0) {
                return true;
            } else {
                return false;
            }
        }
    }

    if (typeof (obj) != "object") {
        alert("isEmpty: Please give the \"string\" id name to check empty.");
        return false;
    }

    switch (obj.type) {
        case "select-multiple":
        case "select-one":
            result = !isSelected(obj);
            break;
        case "hidden":
        case "password":
        case "text":
        case "textarea":
            result = isBlank(obj);
            break;
        case "checkbox":
        case "radio":
        default:
            result = !isChecked(obj);
    }

    return result;
}

