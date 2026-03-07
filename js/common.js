/* 去除掉指定字串的前、後空白字元
 * @param   strText	 想要去除前後空白字元的字串
 * @return  string
 */
function trim(strText) {
    strText = ltrim(strText);
    strText = rtrim(strText);

    return strText;
}

/* 去除掉指定字串的左側空白字元
 * @param   strText	 想要去除左側空白字元的字串
 * @return  string
 */
function ltrim(strText) {
    if (strText != undefined) {
        while (strText.substring(0, 1) == ' ')
            strText = strText.substring(1, strText.length);
    }

    return strText;
}

/* 去除掉指定字串的右側空白字元
 * @param   strText	 想要去除右側空白字元的字串
 * @return  string
 */
function rtrim(strText) {
    if (strText != undefined) {
        while (strText.substring(strText.length - 1, strText.length) == ' ')
            strText = strText.substring(0, strText.length - 1);
    }
    return strText;
}

/* 設定指定表單欄位物件的值，支援所有的表單物件型態
 * checkbox, radio, select, text, textarea, hidden,
 * pXsswXrd, button, submit, reset
 * (唯有file型態的物件因為安全性的關係，無法設定值)
 * @param   myobj	   表單欄位物件or表單欄位物件的字串(格式為"formname.elementname")
 * @param   value	   欲設定的值
 * @return  none
 */
function setValue(myobj, value) {
    var obj;

    /*IF 傳入的myobj是字串(格式為formname.elementname) THEN 取得此物件的指標*/
    if (typeof (myobj) == "string") {
        var form_name = myobj.substring(0, myobj.indexOf("."));
        var element_name = myobj.substr(myobj.indexOf(".") + 1);
        obj = document.forms[form_name].elements[element_name];
    } else {
        obj = myobj;
    }

    if (typeof (obj) != "object") {
        if (typeof (myobj) == "string") {
            alert("Cannot set \"" + value + "\" to the " + myobj + " field.");
        } else {
            alert("Cannot set \"" + value + "\" to the undefined form field.");
        }
        return false;
    }
    switch (obj.type) {
        case "checkbox":
            setCheckboxValue(obj, value);
            break;
        case "select-multiple":
        case "select-one":
            setSelectValue(obj, value);
            break;
        case "hidden":
        case "text":
        case "textarea":
        case "password":
        case "file":
        case "button":
        case "submit":
        case "reset":
            obj.value = decodeHtml(value); /* 若文字有含html編碼「&#dddd;」，會將其還原*/

            break;
        case "radio":
        default:
            var flagobjtagNameTable = false; /*tagName 為TABLE*/
            if (obj.tagName) {
                if (obj.tagName == "TABLE") {
                    flagobjtagNameTable = true;
                }
            }

            if (flagobjtagNameTable) {
                /*tagName 為TABLE*/
                setRadioValue1(obj, value);
            } else {
                /*tagName 不為TABLE*/
                setRadioValue(obj, value);
            }
    }
}

function DisabledRadio(obj, flag) {
    for (i = 0; i < obj.length; i++) {
        obj[i].disabled = flag;
    }
}

/* 設定radio表單欄位物件的值，也就是說，選取符合的項目
 * @param   obj		 radio表單欄位物件
 * @param   value	   欲設定的值
 * @return  none
 */
function setRadioValue1(obj, value) {
    /*RadioButtonList TABLE*/
    if (obj.tagName == "TABLE") {
        for (var i = 0; i < obj.rows.length; i++) {
            for (var j = 0; j < obj.rows[i].cells.length; j++) {
                if (obj.rows[i].cells[j].childNodes[0]) {
                    if (obj.rows[i].cells[j].childNodes[0].value == value) {
                        obj.rows[i].cells[j].childNodes[0].checked = true;
                        break;
                    }
                }
            }
        }
    }
}

function setCheckBoxList1(obj, value) {
    /*debugger; Checkboxlist*/
    if (obj.cells != undefined && typeof (obj.cells.length) == "number") {
        var objid = obj.id;
        for (var i = 0; i < value.length; i++) {
            if (value.substr(i, 1) == "1") {
                if (document.getElementById(objid + "_" + i)) {
                    document.getElementById(objid + "_" + i)
                        .checked = true;
                }
                /*obj.cells[i].firstChild.checked = true;*/
            }
            if (value.substr(i, 1) == "0") {
                if (document.getElementById(objid + "_" + i)) {
                    document.getElementById(objid + "_" + i)
                        .checked = false;
                }
                /*obj.cells[i].firstChild.checked = false;*/
            }
        }
        /*return "";*/
    }
}

/* 設定radio表單欄位物件的值，也就是說，選取符合的項目
 * @param   obj		 radio表單欄位物件
 * @param   value	   欲設定的值
 * @return  none
 */
function setRadioValue(obj, value) {
    /*debugger; RadioButtonList*/
    if (typeof (obj.length) == "number") {
        for (var i = 0; i < obj.length; i++) {
            if (obj[i].value == value) {
                obj[i].checked = true;
                break;
            }
        }
        return "";
    }
    if (obj.childNodes != undefined && typeof (obj.childNodes.length) == "number") {
        var vv1 = 0;
        for (var i = 0; i < obj.childNodes.length; i++) {
            if (vv1 >= value.length) {
                break;
            }
            if (obj.childNodes[i].checked != undefined) {
                obj.childNodes[i].checked = false;
                if (value.substr(vv1, 1) == "1") {
                    obj.childNodes[i].checked = true;
                }
                vv1 += 1; /*從0開始*/
            }
        }
        return "";
    }
    /*Checkboxlist*/
    if (obj.cells != undefined && typeof (obj.cells.length) == "number") {
        var objid = obj.id;
        for (var i = 0; i < value.length; i++) {
            if (value.substr(i, 1) == "1") {
                if (document.getElementById(objid + "_" + i)) {
                    document.getElementById(objid + "_" + i)
                        .checked = true;
                }
                /*obj.cells[i].firstChild.checked = true;*/
            }
            if (value.substr(i, 1) == "0") {
                if (document.getElementById(objid + "_" + i)) {
                    document.getElementById(objid + "_" + i)
                        .checked = false;
                }
                /*obj.cells[i].firstChild.checked = false;*/
            }
        }
        return "";
    } else {
        alert(obj);
        alert(obj.value);
        if (obj.value == value) {
            obj.checked = true;
        }
    }
}

/** 設定radio表單欄位物件的值，不選擇
 * @param   obj		 radio表單欄位物件
 * @return  none */
function setRadioNoChoice(obj) {
    /*RadioButtonList*/
    if (typeof (obj.length) == "number") {
        for (var i = 0; i < obj.length; i++) {
            if (obj[i].checked) {
                obj[i].checked = false;
                break;
            }
        }
        return "";
    }
}

/* 設定checkbox表單欄位物件的值，也就是說，選取符合的項目
 * @param   obj		 checkbox表單欄位物件
 * @param   value	   欲設定的值
 * @return  none
 */
function setCheckboxValue(obj, value) {
    if (typeof (obj.length) == "number") {
        for (var i = 0; i < obj.length; i++) {
            if (obj[i].value == value) {
                obj[i].checked = true;
                break;
            }
        }
    } else {
        if (obj.value == value) {
            obj.checked = true;
        }
    }
}

/* 設定select表單欄位物件的值，也就是說，選取符合的項目
 * @param   obj		 select表單欄位物件
 * @param   value	   欲設定的值
 * @return  none
 */
function setSelectValue(obj, value) {
    if (obj.options) {
        for (var i = 0; i < obj.options.length; i++) {
            if (obj.options[i].value == value) {
                obj.options[i].selected = true;
                break;
            }
        }
    }
}

/* 取得指定表單欄位物件的值，支援所有的表單物件型態
 * checkbox, radio, select, text, textarea, hidden,
 * pXsswXrd, file, button, submit, reset
 * @param   myobj	   表單欄位物件or表單欄位物件的字串(格式為"formname.elementname")
 * @return  string	  若有多個值的話，每個值之間以,隔開
 */
function getValue(myobj) {
    var obj;
    var result = "";

    /*DropDownList IF 傳入的myobj是字串(格式為formname.elementname) THEN 取得此物件的指標*/
    if (typeof (myobj) == "string") {
        var pos = myobj.indexOf(".");
        if (pos > 0) {
            var form_name = myobj.substring(0, pos);
            var element_name = myobj.substr(pos + 1);
            obj = document.forms[form_name].elements[element_name];
        } else {
            return getValueByName(myobj);
        }
    } else {
        obj = myobj;
    }

    if (typeof (obj) != "object") {
        if (typeof (myobj) == "string") {
            alert("Cannot get the value of " + myobj + " field.");
        } else {
            alert("Cannot get the value of undefined form field.");
        }
        return false;
    }

    switch (obj.type) {
        case "checkbox":
            result = getCheckboxValue(obj);
            break;
        case "select-multiple":
        case "select-one":
            result = getSelectValue(obj);
            break;
        case "hidden":
        case "text":
        case "textarea":
        case "password":
        case "file":
        case "button":
        case "submit":
        case "reset":
            result = obj.value;
            break;
        case "radio":
        default:
            result = getRadioValue(obj);
            if (result == "") {
                result = getSPANValue(obj);
            }
    }

    return result;
}

/* 取得指定表單欄位物件的值，支援所有的表單物件型態(.NET OBJ)
 * checkbox, radio, select, text, textarea, hidden,
 * pXsswXrd, file, button, submit, reset
 * @param   myobj	   表單欄位物件or表單欄位物件的字串(格式為"formname.elementname")
 * @return  string	  若有多個值的話，每個值之間以,隔開
 */
function getSPANValue(obj) {
    var result = "";
    if (obj.children) {
        if (typeof (obj.children.length) == "number") {
            for (var i = 0; i < obj.children.length; i++) {
                switch (obj.children[i].type) {
                    case "checkbox":
                        if (obj.children[i].checked) {
                            result += "1";
                        } else {
                            result += "0";
                        }
                        break;
                    case "radio":
                        if (obj.children[i].checked) {
                            result = obj.children[i].value;
                        }
                        break;
                }
            }
            if (result == '') {
                for (var i = 0; i < obj.children.length; i++) {
                    if (obj.children[i].firstChild) {
                        switch (obj.children[i].firstChild.type) {
                            case "checkbox":
                                if (obj.children[i].firstChild.checked) {
                                    result += "1";
                                } else {
                                    result += "0";
                                }
                                break;
                            case "radio":
                                if (obj.children[i].firstChild.checked) {
                                    result = obj.children[i].value;
                                }
                                break;
                        }
                    }
                }
            }

        }
    }
    return result;
}

/*設定*/
function setSPANValue(obj, value) {
    var vi = 0;
    if (obj.children) {
        if (typeof (obj.children.length) == "number") {
            for (var i = 0; i < obj.children.length; i++) {
                switch (obj.children[i].type) {
                    case "checkbox":
                        var result = value.substr(vi, 1);
                        if (result == "1") {
                            obj.children[i].checked = true;
                        } else {
                            obj.children[i].checked = false;
                        }
                        vi += 1;
                        break;
                    case "radio":
                        var result = value.substr(vi, 1);
                        if (result == "1") {
                            obj.children[i].checked = true;
                        } else {
                            obj.children[i].checked = false;
                        }
                        vi += 1;
                        break;
                }
            }
        }
    }
}

/*取得*/
function getValueByName(myobj_name) {
    var oform = document.forms[0];
    var obj = null;
    var result = "";
    var item_name;

    if (typeof (myobj_name) != "string") {
        alert("getValueByName: Please give the \"string\" id name to get the values.");
        return false;
    }

    obj = oform.elements[myobj_name];
    if (typeof (obj) == "object") {
        result = getValue(obj);
    } else {
        for (var i = 0; i < oform.elements.length; i++) {
            obj = oform.elements[i];
            item_name = obj.name;
            if (item_name.indexOf(myobj_name + ":") == 0 && obj.type == "checkbox") {
                result = getCheckBoxListValue(myobj_name);
                break;
            }
        }
    }

    return result;
}

/* 取得在radio表單欄位物件上，使用者所選取的值
 * js查無資料要顯示錯誤。
 * @param   obj		 radio表單欄位物件
 * @return  string
 */
function getRadioValue(obj) {
    var result = "";
    if (!obj) {
        /*alert("error:getRadioValue()");*/
        return result;
    }

    if (typeof (obj.length) == "number") {
        for (var i = 0; i < obj.length; i++) {
            if (obj[i].checked) {
                result = obj[i].value;
                break;
            }
        }
    } else {
        if (obj.checked) {
            result = obj.value;
        }
    }

    return result;
}

/* 取得在checkbox表單欄位物件上，使用者所選取的值
 * @param   obj		 checkbox表單欄位物件
 * @return  string	  若有多個值的話，每個值之間以,隔開
 */
function getCheckboxValue(obj) {
    var result = "";

    if (typeof (obj.length) == "number") {
        for (var i = 0; i < obj.length; i++) {
            if (obj[i] && obj[i].checked) {
                if (result != "") {
                    result += ",";
                }
                result += obj[i].value;
            }
        }
    } else {
        if (obj.checked) {
            if (result != "") {
                result += ",";
            }
            result += obj.value;
        }
    }
    return result;
}

/* 取得在checkbox表單欄位物件上，使用者所選取的值
 * @param   obj		 checkbox表單欄位物件
 * @return  Array	   傳回所選取的值的陣列
 */
function getCheckboxValueArray(obj) {
    var result = Array();
    var index = 0;

    if (typeof (obj.length) == "number") {
        for (var i = 0; i < obj.length; i++) {
            if (obj[i].checked) {
                result[index++] = obj[i].value;
            }
        }
    } else {
        if (obj.checked) {
            result[index++] = obj[i].value;
        }
    }

    return result;
}

/* 在ASP.NET網頁上，CheckBoxList所產生的表單欄位物件，取得使用者選取的狀態 (以01表示選取的狀態)
 * @param   objname		CheckBoxList控制項名稱
 * @return  string		傳回所選取的值的字串。例如：101，代表的是1、3項有被選取，第2項沒被選取
 */
function getCheckBoxListValue(objname) {
    var oform = document.forms[0];
    var obj = null;
    var result = "";
    var item_name;
    var index;
    var values = Array();
    var ttype2 = "";

    if (typeof (objname) == "string") {
        ttype2 = typeof (objname);
    }
    if (typeof (objname) == "object") {
        ttype2 = typeof (objname);
    }

    if (ttype2 == "") {
        alert("getCheckBoxListValue: Please give the \"string\" id name to get the index.");
        return false;
    }

    if (ttype2 == "object") {
        /*if (!objname.children) return null;*/
        return getSPANValue(objname); /*return null;*/
    }

    /*var item_names = '';,ttype2 == "string"*/
    for (var i = 0; i < oform.elements.length; i++) {
        obj = oform.elements[i];
        item_name = obj.name;
        /*if (item_names != '') item_names += ',';,item_names += obj.name;alert(item_name);*/
        var type1 = "";
        var chkObject = false;
        if (item_name.indexOf(objname + "$") == 0) {
            type1 = "$";
            chkObject = true;
        }
        if (item_name.indexOf(objname + ":") == 0) {
            type1 = ":";
            chkObject = true;
        }
        if (obj.type == "checkbox" && chkObject) {
            index = parseInt(item_name.substring(item_name.indexOf(type1) + 1), 10)
            if (obj.checked) {
                values[index] = 1;
            } else {
                values[index] = 0;
            }
        }
    }
    /*alert(item_names);*/
    for (var i = 0; i < values.length; i++) {
        if (values[i] == 1) {
            result += "1";
        } else {
            result += "0";
        }
    }
    return result;
}

/* 在ASP.NET網頁上，檢查CheckBoxList在指定的索引項目上，是否有被選取 (索引從0開始)
 * @param   objname		CheckBoxList控制項名稱
 * @return  bool		若指定的索引項目有被選取，則傳回true，則否為false
 */
function hasCheckBoxListIndex(objname, index) {
    var result = false;
    var value = getCheckBoxListValue(objname);
    if (value.length > index && value.charAt(index) == '1') {
        result = true;
    }

    return result;
}

/* 取得在select表單欄位物件上，使用者所選取的值
 * @param   obj		 select表單欄位物件
 * @return  string	  若有多個值的話，每個值之間以,隔開
 */
function getSelectValue(obj) {
    var result = "";

    for (var i = 0; i < obj.options.length; i++) {
        if (obj.options[i].selected) {
            result += "," + obj.options[i].value;
        }
    }

    if (result != "") {
        result = result.substr(1);
    }

    return result;
}

/* 取得在select表單欄位物件上，使用者所選取的值
 * @param   obj		 select表單欄位物件
 * @return  Array	   傳回所選取的值的陣列
 */
function getSelectValueArray(obj) {
    var result = Array();
    var index = 0;

    for (var i = 0; i < obj.options.length; i++) {
        if (obj.options[i].selected) {
            result[index++] = obj[i].value;
        }
    }

    return result;
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

/* 檢查指定的radio或checkbox表單欄位物件是否被使用者選取過
 * @param   obj		 radio或checkbox表單欄位物件
 * @return  boolean
 */
function isChecked(obj) {
    var result = false;
    if (obj) {
        if (typeof (obj.length) == "number") {
            for (var i = 0; i < obj.length; i++) {
                if (obj[i].checked) {
                    result = true;
                    break;
                }
            }
        } else {
            if (obj.checked) {
                result = true;
            }
        }
    }
    return result;
}

/* 檢查指定的select表單欄位物件是否被使用者選取過
 * (若使用者選取的項目值為空字串，則仍傳回未被選取)
 * @param   obj		 select表單欄位物件
 * @return  boolean
 */
function isSelected(obj) {
    var result = false;

    for (var i = 0; i < obj.options.length; i++) {
        if (obj.options[i].selected && obj.options[i].value != "") {
            result = true;
            break;
        }
    }

    return result;
}

/* 檢查指定的text或textarea或hidden表單欄位物件是否已輸入值
 * (若欄位值為空白字串，則仍傳回未被輸入)
 * @param   obj		 select表單欄位物件
 * @return  boolean
 */
function isBlank(obj) {
    return isSpace(obj.value);
}

/* 檢查指定的字串是否為空字串或為空白字串
 * @param   value	   欲檢查的值
 * @return  boolean
 */
function isSpace(value) {
    return (trim(value) == "") ? true : false;
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

/* 判斷是否為正整數，例如: +25, 77 皆符合條件
 * @param   value	   欲檢查的值
 * @return  boolean
 */
function isPositiveInt(value) {
    var pattern = /^(\+)?\d+$/;
    return pattern.test(value);
}

/*檢查*/
function isCellPhone(value) {
    var pattern = /^[09]{2}[0-9]{8}$/;
    return pattern.test(value);
}

/* 判斷是否為負整數，例如: -33, -006 皆符合條件
 * @param   value	   欲檢查的值
 * @return  boolean
 */
function isNegativeInt(value) {
    var pattern = /^\-\d+$/;
    return pattern.test(value);
}

/* 判斷是否為正浮點數或負浮點數，例如: +25.7, -33.7, 77.7 皆符合條件
 * @param   value	   欲檢查的值
 * @return  boolean
 */
function isFloat(value) {
    var pattern = /^(\+|\-)?\d+\.\d+$/;
    return pattern.test(value);
}

/*最多到小數點第1位*/
function isFloat1(value) {
    var pattern = /^(\+|\-)?\d+\.{0,1}\d{0,1}$/;
    return pattern.test(value);
}

/*最多到小數點第二位*/
function isFloat2(value) {
    var pattern = /^(\+|\-)?\d+\.{0,1}\d{0,2}$/;
    return pattern.test(value);
}

/* 判斷是否為正浮點數，例如: +25.7, 77.7 皆符合條件
 * @param   value	   欲檢查的值
 * @return  boolean
 */
function isPositiveFloat(value) {
    var pattern = /^(\+)?\d+\.\d+$/;
    return pattern.test(value);
}

/* 判斷是否為負浮點數，例如: -33.7, -006.7 皆符合條件
 * @param   value	   欲檢查的值
 * @return  boolean
 */
function isNegativeFloat(value) {
    var pattern = /^\-\d+\.\d+$/;
    return pattern.test(value);
}

/* 選擇或不選擇所有的相同名稱的checkbox項目
 * @param   obj		 checkbox表單欄位物件
 * @param   value	   true/false (全選/全不選)
 * @return  none
 */
function checkAll(obj, value) {
    if (typeof (obj.length) == "number") {
        for (var i = 0; i < obj.length; i++) {
            obj[i].checked = value;
        }
    } else {
        obj.checked = value;
    }
}

/* 將不是目前使用者點選的checkbox項目狀態全部設為不選
 * 模擬出radio只能單選的效果，但是又能取消選取
 * 用法: checkOnlyOne(form1.chkLevel, this.value)
 * @param   obj		 checkbox表單欄位物件
 * @param   value	   目前點選的checkbox值
 * @return  none
 */
function checkOnlyOne(obj, value) {
    if (typeof (obj.length) == "number") {
        for (var i = 0; i < obj.length; i++) {
            if (obj[i].value != value) {
                obj[i].checked = false;
            }
        }
    } else {
        if (obj.value != value) {
            obj.checked = false;
        }
    }
}

/* 選擇符合條件值的checkbox項目
 * @param   obj		 checkbox表單欄位物件
 * @param   value	   條件值
 * @return  none
 */
function checkValue(obj, value) {
    setValue(obj, value);
}

/* 不選擇符合條件值的checkbox項目
 * @param   obj		 checkbox表單欄位物件
 * @param   value	   條件值
 * @return  none
 */
function uncheckValue(obj, value) {
    if (typeof (obj.length) == "number") {
        for (var i = 0; i < obj.length; i++) {
            if (obj[i].value == value) {
                obj[i].checked = false;
                break;
            }
        }
    } else {
        if (obj.value == value) {
            obj.checked = false;
        }
    }
}

/* 選擇或不選擇所有的清單項目
 * @param   obj		 select表單欄位物件
 * @param   value	   true/false (全選/全不選)
 * @return  none
 */
function selectAll(obj, value) {
    for (var i = 0; i < obj.options.length; i++) {
        obj.options[i].selected = value;
    }
}

/*選擇全部 obj:string hidobj:string*/
function SelectAll(obj, hidobj) {
    var num = getCheckBoxListValue(obj)
        .length; /*長度*/
    var myallcheck = document.getElementById(obj.concat('_0')); /*第1個*/

    if (document.getElementById(hidobj)
        .value != getCheckBoxListValue(obj)
            .charAt(0)) {
        document.getElementById(hidobj)
            .value = getCheckBoxListValue(obj)
                .charAt(0);
        for (var i = 1; i < num; i++) {
            var mycheck = document.getElementById(obj.concat('_', i));
            mycheck.checked = myallcheck.checked;
        }
    } else {
        for (var i = 1; i < num; i++) {
            if ('0' == getCheckBoxListValue(obj)
                .charAt(i)) {
                document.getElementById(hidobj)
                    .value = getCheckBoxListValue(obj)
                        .charAt(i);
                var mycheck = document.getElementById(obj.concat('_', i));
                myallcheck.checked = mycheck.checked;
                break;
            }
        }
    }
}

/* 選擇符合條件值的清單項目
 * @param   obj		 select表單欄位物件
 * @param   value	   項目值
 * @return  none
 */
function selectValue(obj, value) {
    setValue(obj, value);
}

/* 不選擇符合條件值的清單項目
 * @param   obj		 select表單欄位物件
 * @param   value	   項目值
 * @return  none
 */
function unselectValue(obj, value) {
    for (var i = 0; i < obj.options.length; i++) {
        if (obj.options[i].value == value) {
            obj.options[i].selected = false;
            break;
        }
    }
}

/* 判斷是有符合條件值的清單項目名稱
 * @param   obj		 select表單欄位物件
 * @param   text		項目名稱
 * @return  boolean
 */
function hasOptionByText(obj, text) {
    var result = false;
    for (var i = 0; i < obj.options.length; i++) {
        if (obj.options[i].text == text) {
            result = true;
            break;
        }
    }

    return result;
}

/* 判斷是有符合條件值的清單項目值
 * @param   obj		 select表單欄位物件
 * @param   value	   條件值
 * @return  boolean
 */
function hasOptionByValue(obj, value) {
    var result = false;

    for (var i = 0; i < obj.options.length; i++) {
        if (obj.options[i].value == value) {
            result = true;
            break;
        }
    }

    return result;
}

/* 加入一新的清單項目至指定的清單物件
 * @param   obj		 select表單欄位物件
 * @param   text		項目名稱
 * @param   value	   項目值
 * @return  none
 */
function addOption(obj, text, value) {
    var newOpt;

    newOpt = document.createElement("OPTION");
    newOpt.text = text;
    newOpt.value = value;
    obj.add(newOpt);
}

/* 移除符合條件值的清單項目名稱
 * @param   obj		 select表單欄位物件
 * @param   text		項目名稱
 * @return  boolean
 */
function delOptionByText(obj, text) {
    var result = false;

    for (var i = 0; i < obj.options.length; i++) {
        if (obj.options[i].text == text) {
            obj.options.remove(i);
            result = true;
            break;
        }
    }

    return result;
}

/* 移除符合條件值的清單項目值
 * @param   obj		 select表單欄位物件
 * @param   value	   項目值
 * @return  boolean
 */
function delOptionByValue(obj, value) {
    var result = false;

    for (var i = 0; i < obj.options.length; i++) {
        if (obj.options[i].value == value) {
            obj.options.remove(i);
            result = true;
            break;
        }
    }

    return result;
}

/* 從指定的清單xfrom，搬移選擇的項目至另一清單xto
 * @param   xfrom	   來源select表單欄位物件
 * @param   xto		 目的select表單欄位物件
 * @return  none
 */
function moveOption(xfrom, xto) {
    for (var i = 0; i < xfrom.options.length;) {
        if (xfrom.options[i].selected) {
            addOption(xto, xfrom.options[i].text, xfrom.options[i].value);
            xfrom.options.remove(i);
        } else {
            i++;
        }
    }
}

/* 從指定的清單xfrom，搬移所有的項目至另一清單xto
 * @param   xfrom	   來源select表單欄位物件
 * @param   xto		 目的select表單欄位物件
 * @return  none
 */
function moveAllOption(xfrom, xto) {
    for (var i = 0; i < xfrom.options.length;) {
        addOption(xto, xfrom.options[i].text, xfrom.options[i].value);
        xfrom.options.remove(i);
    }
}

/* 檢查指定email格式是否正確
 * @param   email	   欲檢查的email
 * @return  boolean
 */
function checkEmail(email) {
    /*var filter = /^.+@.+\..{2,3}$/;*/
    var filter = /^\w+[\w\.\-]*[\w\-]+@.+\..{2,3}$/;

    if (filter.test(email)) {
        return true;
    } else {
        return false;
    }
}

/* 檢查指定檔案名稱格式是否正確 (在Windows系統上，有些特別的字元不可為檔案名稱的一部份)
 * @param   filename	欲檢查的檔名
 * @return  boolean
 */
function checkFileName(filename) {
    var re = /[\\\/\:\*\?\"\'\<\>\|\[\]]/gi;
    return !re.test(filename);
}

/* 傳回不可為檔案名稱的一部份的特別字元)
 * @return  string
 */
function strFileName() {
    return "\\ / : * ? \" ' < > | [ ]";
}

/* 在完整檔案路徑下，取得檔案名稱 (即不包含路徑資訊)
 * @param   myPath	  完整檔案路徑
 * @return  string
 */
function getFileName(myPath) {
    var i;

    i = myPath.lastIndexOf("\\");
    if (i == -1) {
        i = myPath.lastIndexOf("/");
    }
    if (i != -1) {
        return myPath.substring(i + 1, myPath.length);
    }
    return "";
}

/* 判斷指定的值是否為空字串或空白字串
 * @param   value	   欲檢查的字串
 * @param   filename	欲檢查的檔名
 * @return  boolean
 */
function checkEmpty(value) {
    var result = false;

    if (trim(value) == "") {
        result = true;
    }

    return result;
}

/* 取得字串長度 (一個中文字算二個字元)
 * @param   myStr	   欲檢查的字串
 * @return  integer
 */
function getStrLen(myStr) {
    var myLength = 0;

    for (var i = 0; i < myStr.length; i++) {
        myLength++;
        if (myStr.charCodeAt(i) > 127) {
            myLength++;
        }
    }
    return myLength;
}

/* 判斷指定的字串是否超過最大長度限制值，若超過則傳回true
 * @param   value	   欲檢查的字串
 * @return  boolean
 */
function checkMaxLen(value, Length) {
    var actualLen = getStrLen(value);

    return (actualLen > Length);
}

/* 判斷指定的字串是否超過最大長度限制值，若超過則傳回true
 * @param   value	   欲檢查的字串
 * @return  boolean
 */
function checkMaxLen2(value, Length) {
    return (value.length > Length);
}

/* 判斷指定的字串是否小於最小長度限制值，若小於則傳回true
 * @param   value	   欲檢查的字串
 * @return  boolean
 */
function checkMinLen(value, Length) {
    var actualLen = getStrLen(value);

    return (actualLen < Length);
}

/* 將指定的數字四捨五入到指定的小數位數
 * 例如: getRound(2.35, 4),傳回 2.3500,getRound(2.35, 1),傳回 2.4
 * @param   value	   欲檢查的值
 * @return  string
 */
function getRound(number, noOfPlaces) {
    if (isNaN(number)) {
        alert("Please enter a valid number");
        Number = 0;
    }

    val = (Math.round(number * Math.pow(10, noOfPlaces))) / Math.pow(10, noOfPlaces);
    val = val.toString();
    ind = val.indexOf(".");
    if (ind == -1) {
        val = val.toString() + ".";
        for (i = 0; i < noOfPlaces; i++)
            val = val + "0";
    } else {
        len = val.length;
        x = len - ind - 1;
        if (x < noOfPlaces) {
            for (i = 0; i < (noOfPlaces - x) ; i++)
                val = val + "0";
        }
    }

    return (val);
}

/* 檢查西元日期是否正確(含民國) (格式可為 YYYY-MM-DD 或 YYYY/MM/DD)
 * @param   value	   欲檢查的日期字串
 * @return  boolean
 */
function checkDate(DateString) {
    var separator = "/";
    if (DateString.indexOf("/") != -1) {
        separator = "/";
    } else if (DateString.indexOf("-") != -1) {
        separator = "-";
    }
    return CheckDatefmt(DateString, separator);
}

/* 檢查中華民國日期是否正確 (格式可為 YYY-MM-DD 或 YYY/MM/DD)
 * 為了避免使用者輸入西元年份，若年份欄位超出3碼者，視為不正確的中華民國日期
 * @param   value	   欲檢查的日期字串
 * @return  boolean
 */
function checkRocDate(mydate) {
    var separator = "/";

    /*將日期格式固定為 YYYY/MM/DD or YYYY-MM-DD (即MM, DD皆為兩位數)*/
    DateString = formatDate(mydate);

    if (DateString.indexOf("/") != -1) {
        separator = "/";
    } else if (DateString.indexOf("-") != -1) {
        separator = "-";
    }

    var idx = DateString.indexOf(separator);
    var y = DateString.substring(0, idx) /*年*/

    /*使用者可能輸入西元年份*/
    if (y.length > 3) {
        return false;
    }
    y = parseInt(y, 10);
    if (isNaN(y)) {
        return false;
    }
    var md = DateString.substring(idx + 1, DateString.length);
    y = y + 1911; /*將民國年轉為西元年*/
    var ymd = y + separator + md;

    return CheckDatefmt(ymd, separator);
}

/* 檢查日期是否正確，可自定分隔字元，但仍必須符合日期格式=4位年+分隔+2位月+分隔+2位日期
 * @param   DateString  欲檢查的日期字串
 * @param   chrFmt	  分隔字元
 * @return  boolean
 */
function CheckDatefmt(DateString, chrFmt) {
    if (DateString.length > 10 || DateString.length < 6) return false;
    var y, m, d;
    var idx = DateString.indexOf(chrFmt)
    y = DateString.substring(0, idx) /*年*/
    DateString = DateString.substring(idx + 1, DateString.length)
    var idx = DateString.indexOf(chrFmt)
    m = DateString.substring(0, idx); /*月*/
    d = DateString.substring(idx + 1, DateString.length); /*日*/

    if (m.substring(0, 1) == '0') m = m.substring(1, m.length);
    if (d.substring(0, 1) == '0') d = d.substring(1, d.length);
    /*alert("y="+y); alert("m="+m); alert("d="+d);*/
    var CharNum = "0123456789";
    /*判別是否皆為數字*/
    for (var i = 0; i < y.length; i++) {
        var str = y.substring(i, i + 1);
        if (CharNum.indexOf(str) < 0) return false;
    }

    for (var i = 0; i < m.length; i++) {
        var str = m.substring(i, i + 1);
        if (CharNum.indexOf(str) < 0) return false;
    }

    for (var i = 0; i < d.length; i++) {
        var str = d.substring(i, i + 1);
        if (CharNum.indexOf(str) < 0) return false;
    }

    y = parseInt(y, 10);
    m = parseInt(m, 10);
    d = parseInt(d, 10);
    if (isNaN(y)) return false;
    if (isNaN(m)) return false;
    if (isNaN(d)) return false;

    /*if (y < 100 && y > 70) y += 1900;
    if (y < 70) y += 2000;
    if (y > 2070 || y < 1970) return false;*/
    if (y > 1 && y < 222) {
        y += 1900; /*民國改為西元年判斷*/
    }
    if (y < 1900 || y > 2100) return false;
    if (m < 1 || m > 12) return false;
    if (d < 1 || d > 31) return false;

    var isleap = ((y % 100) && !(y % 4)) || !(y % 400);

    switch (m) {
        case 1:
        case 3:
        case 5:
        case 7:
        case 8:
        case 10:
        case 12:
            return true;
        case 4:
        case 6:
        case 9:
        case 11:
            if (d > 30) return false;
            else return true;
        case 2:

            if (isleap) {
                if (d > 29) {
                    return false;
                } else {
                    return true;
                }
            } else {
                if (parseInt(d, 10) > 28) {
                    return false;
                }
            }
            return true;
        default:
            return false;
    }
}

/* 比較兩個日期的大小
 * 參數日期格式必須為「MM/DD/YYYY」或者「MM-DD-YYYY」，可先使用YMDD2MDDY()作轉換
 * @param   date1	   日期字串1
 * @param   date2	   日期字串2
 * @return  integer	 -1 if date1 is less than date2
 *					   1 if date1 is greater than date2
 *					   0 if they are equal.
 */
function compareDate(date1, date2) {
    var myStartDate, myEndDate;
    var re = /\-/g;

    myStartDate = Date.parse(date1);
    if (isNaN(myStartDate)) {
        return NaN;
    }

    myEndDate = Date.parse(date2);
    if (isNaN(myEndDate)) {
        return NaN;
    }

    if (myStartDate < myEndDate) {
        return -1;
    } else if (myStartDate > myEndDate) {
        return 1;
    } else {
        return 0;
    }

}
/* 將日期的「月」、「日」格式化為兩位數的 YYYYY/MM/DD 或 YYYY-MM-DD 格式
 * @param   DateString  日期字串
 * @return  string
 */
function formatDate(DateString) {
    var separator = "/";
    var idx;
    var temp;
    var y, m, d;
    if (DateString.indexOf("/") != -1) {
        separator = "/";
    } else if (DateString.indexOf("-") != -1) {
        separator = "-";
    }

    /* Year*/
    temp = DateString;
    idx = temp.indexOf(separator);
    y = temp.substring(0, idx);
    temp = temp.substring(idx + 1);

    /* Month*/
    idx = temp.indexOf(separator);
    m = temp.substring(0, idx);
    if (m.length == 1) {
        m = "0" + m;
    }

    /* Day*/
    d = temp.substring(idx + 1);
    if (d.length == 1) {
        d = "0" + d;
    }

    return (y + separator + m + separator + d);
}

/* 將傳入的民國日期 (其格式可為YY/MM/DD或YY-MM-DD)
 * 格式化為西元年的日期
 */
function getAdDate(DateString) {
    var separator = "/";
    if (DateString.indexOf("/") != -1) {
        separator = "/";
    } else if (DateString.indexOf("-") != -1) {
        separator = "-";
    }
    var idx = DateString.indexOf(separator);
    var y = DateString.substring(0, idx) /*年*/
    y = parseInt(y, 10);
    if (isNaN(y)) return "";
    var md = DateString.substring(idx + 1, DateString.length);
    y = y + 1911; /*將民國年轉為西元年*/
    var ymd = y + separator + md;
    return ymd;
}

/* 將傳入的西元年日期 (其格式可為YYYY/MM/DD或YYYY-MM-DD)
 * 格式化為民國年日期
 */
function getRocDate(DateString) {
    var separator = "/";
    if (DateString.indexOf("/") != -1) {
        separator = "/";
    } else if (DateString.indexOf("-") != -1) {
        separator = "-";
    }

    var idx = DateString.indexOf(separator);
    var y = DateString.substring(0, idx) /*年*/
    y = parseInt(y, 10);
    if (isNaN(y)) return "";
    var md = DateString.substring(idx + 1, DateString.length);
    y = y - 1911; /*將西元年轉為民國年*/
    var ymd = y + separator + md;
    return ymd;
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

/* 將指定的日期字串加/減天數
 * 傳回的日期格式為「YYYY/MM/DD」
 * @param   date1	   指定的日期字串
 * @param   day		 欲加/減的天數
 * @return  string
 */
function addDateByDayOld(mydate, day) {
    var lngDate;
    var newdate;
    var result;

    lngDate = Date.parse(mydate);
    lngDate += day * 24 * 60 * 60 * 1000;

    newdate = new Date(lngDate);

    result = newdate.getFullYear() + "/";
    if ((newdate.getMonth() + 1) < 10) {
        result += "0";
    }
    result += (newdate.getMonth() + 1) + "/";

    if (newdate.getDate() < 10) {
        result += "0";
    }
    result += newdate.getDate();

    return result;
}

/* 將指定的日期字串加/減天數
 * 傳回的日期格式為「YYYY/MM/DD」
 * @param   date1	   指定的日期字串
 * @param   day		 欲加/減的天數
 * @return  string
 */
function addDateByDay(mydate, day) {
    const newDate = new Date(mydate);
    newDate.setDate(newDate.getDate() + day); /*return newDate;*/
    var result = "";
    result = result.concat(newDate.getFullYear(), "/");
    result = result.concat((newDate.getMonth() + 1) < 10 ? "0" : "", (newDate.getMonth() + 1), "/");
    result = result.concat(newDate.getDate() < 10 ? "0" : "", newDate.getDate());
    return result;
}

/* 將指定的日期字串加/減月數
 * 傳回的日期格式為「YYYY/MM/DD」
 * @param   mydate	  指定的日期字串
 * @param   value	   欲加/減的月數
 * @return  string
 */
function addDateByMonth(mydate, value) {
    const newDate = new Date(mydate);
    const myvalue = parseInt(value, 10);

    var y, m, d;
    y = newDate.getFullYear();
    m = (newDate.getMonth() + 1);
    d = newDate.getDate();
    m += myvalue;

    if (m > 12) {
        y += 1;
    }
    m %= 12;
    if (m == 0) {
        m = 12;
    }

    var result = "";
    result = result.concat(y, "/");
    result = result.concat(m < 10 ? "0" : "", m, "/");
    result = result.concat(d < 10 ? "0" : "", d);
    if (!checkDate(result)) {
        result = getLastDay(y + "/" + m + "/1");
    }
    return result;
}

/* 將民國年加上月數，ex:"92/01/31"加上"1"個月，結果為"92/02/28"
 * 注：當大月加上某月後，變成小月，日期會為小月的最後一天，而不是再下個月的第一天
 * @param   datestr	欲加月數的日期
 * @param   addstr	 想加幾月
 * @return  string	 加總後的結果字串
 */
function addMonth(datestr, addstr) {
    var mydate;
    var newdate;

    if (!checkRocDate(datestr))
        return false;

    mydate = getAdDate(datestr);
    newdate = addDateByMonth(mydate, addstr);

    return getRocDate(newdate);
}

/* 取得指定日期下，此月份的最後一天的日期
 * 傳回的日期格式為「YYYY/MM/DD」
 * @param   mydate	  指定的日期字串
 * @return  string
 */
function getLastDay(mydate) {
    const newDate = new Date(mydate);
    var y, m, d;
    var result;
    y = newDate.getFullYear();
    m = (newDate.getMonth() + 1);
    d = 1;

    if (m == 12) {
        y++;
        m = 1
    } else {
        m++;
    }

    var lngDate = Date.parse(y + "/" + m + "/" + d);
    lngDate -= 1000;
    var newdate = new Date(lngDate);

    result = newdate.getFullYear() + "/";
    if ((newdate.getMonth() + 1) < 10) {
        result += "0";
    }
    result += (newdate.getMonth() + 1) + "/";

    if (newdate.getDate() < 10) {
        result += "0";
    }
    result += newdate.getDate();
    return result;
}

/* 將「年月日」的日期格式轉換為「月日年」字串
 * 傳入的日期格式可為「YYYY/MM/DD」或「YYYY-MM-DD」
 * @param   mydate	  欲轉換的日期字串
 * @return  string
 */
function YMDD2MDDY(mydate) {
    var i_pos;
    var separator = "-";
    var newdate;
    if (!mydate) {
        return newdate;
    }
    i_pos = mydate.indexOf("-");
    if (i_pos == -1) {
        i_pos = mydate.indexOf("/");
        separator = "/";
    }
    /*IF 不是正確的日期格式 THEN 傳回NaN*/
    if (i_pos == -1) {
        return NaN;
    }
    newdate = mydate.substring(i_pos + 1) + separator + mydate.substring(0, i_pos);
    return newdate;
}

/* 檢查輸入的身分證字號是否正確
 * @param   IDString	欲檢查的身分證字號
 * @return  boolean
 */
function checkId(IDString) {
    var ID1 = IDString.toUpperCase();
    if (IDString.length != 0) {
        IDString = IDString.toUpperCase()
    }
    /*if (IDString.length != 10) { ErrString = ErrString + "身分證字號字數不對。" + unescape('%0D') }*/
    if (ID1.length != 10) return false; /*alert("身分證字號字數不對 !");*/
    var IDdigit = new Array(10);
    for (var i = 0; i < 10; i++) {
        IDdigit[i] = ID1.charAt(i);
    }
    var CharEng = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    IDdigit[0] = CharEng.indexOf(IDdigit[0]);
    if (IDdigit[0] == -1) return false; /*alert("身分證字號第一位為錯誤英文字母 !");*/
    if (IDdigit[1] != 1 && IDdigit[1] != 2) return false; /*alert("身分證字號無法辨識性別 !");*/

    var Array1 = new Array(26);
    Array1[0] = 1;
    Array1[1] = 10;
    Array1[2] = 19;
    Array1[3] = 28;
    Array1[4] = 37;
    Array1[5] = 46;
    Array1[6] = 55;
    Array1[7] = 64;
    Array1[8] = 39;
    Array1[9] = 73;
    Array1[10] = 82;
    Array1[11] = 2;
    Array1[12] = 11;
    Array1[13] = 20;
    Array1[14] = 48;
    Array1[15] = 29;
    Array1[16] = 38;
    Array1[17] = 47;
    Array1[18] = 56;
    Array1[19] = 65;
    Array1[20] = 74;
    Array1[21] = 83;
    Array1[22] = 21;
    Array1[23] = 3;
    Array1[24] = 12;
    Array1[25] = 30;
    var result = Array1[IDdigit[0]];
    for (var i = 1; i < 10; i++) {
        var Number = "0123456789";
        IDdigit[i] = Number.indexOf(IDdigit[i]);
        if (IDdigit[i] == -1) {
            /*alert("身分證字號錯誤 !");*/
            return false;
        } else {
            result += IDdigit[i] * (9 - i);
        }
    }
    result += 1 * IDdigit[9];
    /*alert("result=="+result);*/
    if (result % 10 != 0) {
        /*alert("身分證字號錯誤 !");*/
        return false;
    } else {
        return true;
    }
}

/* 檢查輸入的 居留證字號 是否正確
 * @param   IDString	欲檢查的 居留證字號
 * @return  boolean
 */
function checkId2(IDString) {
    /*檢查 居留證字號(身分證字號)*/
    var ID1 = IDString.toUpperCase();
    if (ID1.length != 10) return false;
    if (isNaN(ID1.substr(2, 8)) || (ID1.substr(0, 1) < "A" || ID1.substr(0, 1) > "Z") || (ID1.substr(1, 1) < "A" || ID1.substr(1, 1) > "Z")) {
        return false;
    }
    var head = "ABCDEFGHJKLMNPQRSTUVXYWZIO";
    var id = (head.indexOf(ID1.substr(0, 1)) + 10) + '' + ((head.indexOf(ID1.substr(1, 1)) + 10) % 10) + '' + ID1.substr(2, 8)
    var s = parseInt(id.substr(0, 1), 10) +
        parseInt(id.substr(1, 1), 10) * 9 +
        parseInt(id.substr(2, 1), 10) * 8 +
        parseInt(id.substr(3, 1), 10) * 7 +
        parseInt(id.substr(4, 1), 10) * 6 +
        parseInt(id.substr(5, 1), 10) * 5 +
        parseInt(id.substr(6, 1), 10) * 4 +
        parseInt(id.substr(7, 1), 10) * 3 +
        parseInt(id.substr(8, 1), 10) * 2 +
        parseInt(id.substr(9, 1), 10) +
        parseInt(id.substr(10, 1), 10);
    /*判斷是否可整除*/
    if ((s % 10) != 0) return false;
    /*居留證號碼正確*/
    return true;
}

/*
 * '4:居留證2(外來人口統一證號)／新式統一證號
 * @param   studIdNumber 檢查 居留證字號
 * @return  boolean
 */
function checkId4(studIdNumber) {
    if (studIdNumber.length != 10) return false;
    /*9碼為數字,第1個字母是否為英文字母*/
    if (isNaN(studIdNumber.substr(1, 9)) || !/^[A-Z]$/.test(studIdNumber.substr(0, 1))) return false;
    /*按照轉換後權數的大小進行排序*/
    var idHeader = "ABCDEFGHJKLMNPQRSTUVXYWZIO";
    /*這邊把身分證字號轉換成準備要對應的*/
    studIdNumber = (idHeader.indexOf(studIdNumber.substring(0, 1)) + 10)
        + '' + studIdNumber.substr(1, 9);
    /*開始進行身分證數字的相乘與累加，依照順序乘上1987654321*/
    var intS = parseInt(studIdNumber.substr(0, 1))
        + parseInt(studIdNumber.substr(1, 1)) * 9
        + parseInt(studIdNumber.substr(2, 1)) * 8
        + parseInt(studIdNumber.substr(3, 1)) * 7
        + parseInt(studIdNumber.substr(4, 1)) * 6
        + parseInt(studIdNumber.substr(5, 1)) * 5
        + parseInt(studIdNumber.substr(6, 1)) * 4
        + parseInt(studIdNumber.substr(7, 1)) * 3
        + parseInt(studIdNumber.substr(8, 1)) * 2
        + parseInt(studIdNumber.substr(9, 1)) * 1
        + parseInt(studIdNumber.substr(10, 1)); /*(檢查碼)*/
    /*證號OK*/
    if ((intS % 10) == 0) return true;
    return false;
}

/* ** 營利事業統一編號檢查程式
 * 可至 http:/ /www.etax.nat.gov.tw/ 查詢營業登記資料
 * @since 2006/07/19
 * 統一編號檢查實作(統編 規則) 'isValidTWBID isTWBID TWBID COMIDNO true:OK false:NG
 */
function isTWBID(value) {
    var pattern = /^[0-9]{8}$/;
    return pattern.test(value);
}
/**
 * 統一編號檢查實作(統編 規則) 'isValidTWBID isTWBID TWBID COMIDNO true:OK false:NG
 * @param {any} twbid
 */
function isValidTWBID(twbid) {
    var result = false;
    var weight = "12121241";
    var type2 = false; /*第七個數是否為七*/
    if (isTWBID(twbid)) {
        var tmp = 0;
        var sum = 0;
        for (i = 0; i < 8; i++) {
            tmp = (twbid.charAt(i) - '0') * (weight.charAt(i) - '0');
            sum += parseInt(tmp / 10, 10) + (tmp % 10); /*取出十位數和個位數相加*/
            if (i == 6 && twbid.charAt(i) == '7') {
                type2 = true;
            }
        }
        if (type2) {
            if ((sum % 5) == 0 || ((sum + 1) % 5) == 0) { /*如果第七位數為7*/
                result = true;
            }
        } else {
            if ((sum % 5) == 0) {
                result = true;
            }
        }
    }
    return result;
}

/* 檢查指定的數值，其值的限制是否符合指定的最小or最大值
 * @param   my_value	指定的數值 (可為整數or浮點數)
 * @param   my_min	  最小值。若不想有最小值限制，可給空字串""
 * @param   my_max	  最大值。若不想有最大值限制，可給空字串""
 * @return  boolean
 */
function checkRange(my_value, my_min, my_max) {
    var result;
    var now_value, min_value, max_value;

    now_value = parseFloat(my_value);
    if (min_value == "") {
        min_value = now_value;
    } else {
        min_value = parseFloat(my_min);
    }

    if (max_value == "") {
        max_value = now_value;
    } else {
        max_value = parseFloat(my_max);
    }

    result = true;
    if ((now_value < min_value) || (now_value > max_value)) {
        result = false;
    }

    return result;
}

/* 檢查指定的日期，其值的限制是否符合指定的最小or最大值
 * @param   my_value	指定的日期
 * @param   my_min	  最小值。若不想有最小值限制，可給空字串""
 * @param   my_max	  最大值。若不想有最大值限制，可給空字串""
 * @return  boolean
 */
function checkDateRange(my_value, my_min, my_max) {
    var result;
    var now_value, min_value, max_value;

    now_value = YMDD2MDDY(my_value);
    if (min_value == "") {
        min_value = now_value;
    } else {
        min_value = YMDD2MDDY(my_min);
    }

    if (max_value == "") {
        max_value = now_value;
    } else {
        max_value = YMDD2MDDY(my_max);
    }

    result = true;
    if ((compareDate(now_value, min_value) < 0) || (compareDate(now_value, max_value) > 0)) {
        result = false;
    }

    return result;
}

/* 檢查指定的字串，其長度的限制是否符合指定的最小or最大長度
 * @param   my_value	指定的字串
 * @param   my_min	  最小長度。若不想有最小長度限制，可給空字串""
 * @param   my_max	  最大長度。若不想有最大長度限制，可給空字串""
 * @return  boolean
 */
function checkLength(my_value, my_min, my_max) {
    var result;
    var now_value, min_value, max_value;

    now_value = getStrLen(my_value);
    if (min_value == "") {
        min_value = now_value;
    } else {
        min_value = parseInt(my_min, 10);
    }

    if (max_value == "") {
        max_value = now_value;
    } else {
        max_value = parseInt(my_max, 10);
    }

    result = true;
    if ((now_value < min_value) || (now_value > max_value)) {
        result = false;
    }

    return result;
}

/* 從一組同樣名稱但序號為1開始的checkbox，依其勾選與否的狀態組成一01的字串值
 * @param   itemname	checkbox的ID名稱
 * @param   value	   位元值
 * @return  boolean
 */
function getBits(itemname, count) {
    var result = "";

    for (var i = 1; i <= count; i++) {
        if (document.getElementById(itemname + i)
            .checked) {
            result += "1";
        } else {
            result += "0";
        }
    }
    return result;
}

/* 依據一個由01組成的字串，來設定一組同樣名稱但序號為1開始的checkbox是否要勾選
 * @param   itemname	checkbox的ID名稱
 * @param   value	   位元值
 * @return  boolean
 */
function setBits(itemname, value) {
    var count = value.length;
    var status;

    for (var i = 1; i <= count; i++) {
        if (value.charAt(i - 1) == "1") {
            document.getElementById(itemname + i)
                .checked = true;
        } else {
            document.getElementById(itemname + i)
                .checked = false;
        }
    }
}

/* 將有含html編碼的文字&#dddd;，還原回其所代表unicode碼的字符
 * @param   strMessage  欲解碼的文字
 * @return  string
 */
function decodeHtml(strMessage) {
    var lngStartPos;
    var lngEndPos;
    var intCharCode;
    var HtmlDecode = new String(strMessage);

    do {
        lngStartPos = 0
        lngEndPos = 0
        lngStartPos = HtmlDecode.indexOf("&#");
        if (lngStartPos != -1) {
            lngEndPos = HtmlDecode.indexOf(";", lngStartPos + 2);

            if (lngEndPos > 0) {
                intCharCode = parseInt(HtmlDecode.substring(lngStartPos + 2, lngEndPos), 10);
                if (!isNaN(intCharCode)) {
                    HtmlDecode = HtmlDecode.substring(0, lngStartPos) + String.fromCharCode(intCharCode) + HtmlDecode.substring(lngEndPos + 1, HtmlDecode.length);
                }
            }
        }
    } while (lngStartPos != -1 && lngEndPos != -1);

    return HtmlDecode;
}

/* 取得目前網址列的參數值
 * @param   name		參數名稱
 * @return  string
 */
function getParamValueReg(name) {
    var querystring;
    var values;
    var result = "";

    if (location.search.length > 1) {
        querystring = unescape(location.search + "&");

        var re = new RegExp("[\?|\&]" + name + "=(.+)\&");
        values = querystring.match(re);
        if (values != null) {

            result = values[1];
            if (result.indexOf("&") != -1) {
                result = result.substring(0, result.indexOf("&"));
            }
        }
    }
    return result;
}

/* 取得目前網址列的參數值
 * @param   name		參數名稱
 * @return  string
 */
function getParamValue(name) {
    var paramValue = "";
    if (location.search.length <= 1) {
        return paramValue;
    }
    var paramArr = window.location.search.substring(1)
        .split("&");
    for (var i = 0; i < paramArr.length; i++) {
        var param = paramArr[i].split("=");
        if (param[0] == name) {
            paramValue = unescape(param[1]);
            break;
        }
    }
    return paramValue;
}

/* 取得 RadioButtonList 值 */
function getRBLValue(strObjID) {
    var i = 0;
    var strRtn = '';

    while (document.getElementById(strObjID + '_' + i) && i < 10000) {
        if (document.getElementById(strObjID + '_' + i).checked) {
            strRtn = document.getElementById(strObjID + '_' + i).value;
            break;
        }
        i += 1;
    }
    return strRtn;
}

/* 取得 CheckBoxList 值 */
function getCBLValue(strObjID) {
    var i = 0;
    var strRtn = '';

    while (document.getElementById(strObjID + '_' + i) && i < 10000) {
        if (document.getElementById(strObjID + '_' + i).checked) {
            if (strRtn != '') { strRtn += ","; }
            strRtn += "'" + i.toString() + "'"; /*break;*/
        }
        i += 1;
    }
    return strRtn;
}

/* 判斷欄位值 */
function chkValue(strFlag, strName, obj1, obj2, obj3) {
    var strMsg = '';

    if (!obj1) {
        strMsg += '查無 ' + strName + ' 物件!\n';
        return strMsg;
    }

    switch (strFlag) {
        case 'empty': /*判斷必填欄位*/
            if (obj1.value == undefined) {
                strMsg += 'empty查無 ' + strName + ' 物件!\n';
                return strMsg;
            }
            break;
        case 'select': /*判斷必選下拉選單*/
            if (obj1.value == undefined) {
                strMsg += 'select查無 ' + strName + ' 物件!\n';
                return strMsg;
            }
            break;
    }

    switch (strFlag) {
        case 'empty': /*判斷必填欄位*/
            if (isBlank(obj1)) strMsg = '請輸入' + strName + '!\n';
            break;

        case 'select': /*判斷必選下拉選單*/
            if (isBlank(obj1)) strMsg = '請選擇' + strName + '!\n';
            break;

        case 'int': /*判斷非必填數字欄位*/
            if (obj1.value != "") {
                if (isBlank(obj1)) strMsg = '請輸入' + strName + '!\n';
                else if (!isUnsignedInt(obj1.value)) strMsg = strName + '格式有誤，請輸入整數!\n';
            }
            break;

        case 'int_must': /*判斷必填數字欄位*/
            if (isBlank(obj1)) strMsg = '請輸入' + strName + '!\n';
            else if (!isUnsignedInt(obj1.value)) strMsg = strName + '格式有誤，請輸入整數!\n';
            break;

        case 'int12_must': /*判斷必填數字欄位(且不可超過12個月)*/
            if (isBlank(obj1)) strMsg = '請輸入' + strName + '!\n';
            else if (!isUnsignedInt(obj1.value)) strMsg = strName + '格式有誤，請輸入整數!\n';
            if (strMsg == '' && parseInt(obj1.value, 10) > 12) {
                strMsg = strName + '格式有誤，請輸入不可大於12月!\n';
            }
            break;
        case 'float':
            if (!isBlank(obj1)) {
                if (!isUnsignedInt(obj1.value)) {
                    if (!isPositiveFloat(obj1.value)) {
                        /*msg += '單價必須為數字\n';*/
                        strMsg = strName + '格式有誤，請輸入數字!\n';
                    } else {
                        if (obj1.value.indexOf('.') < obj1.value.length - 3) {
                            /*msg += '單價只能輸入到小數點第二位\n';*/
                            strMsg = strName + '格式有誤，只能輸入到小數點第二位!\n';
                        }
                    }
                }
            }
            break;
        case 'float_must':
            if (isBlank(obj1)) strMsg = '請輸入' + strName + '!\n';
            else {
                if (!isUnsignedInt(obj1.value)) {
                    if (!isPositiveFloat(obj1.value)) {
                        /*msg += '單價必須為數字\n';*/
                        strMsg = strName + '格式有誤，請輸入數字!\n';
                    } else {
                        if (obj1.value.indexOf('.') < obj1.value.length - 3) {
                            /*msg += '單價只能輸入到小數點第二位\n';*/
                            strMsg = strName + '格式有誤，只能輸入到小數點第二位!\n';
                        }
                    }
                }
            }
            break;

        case 'date': /*判斷非必填日期欄位*/
            /*if (!isBlank(obj1) && !checkRocDate(obj1.value)) strMsg = strName + '格式有誤，請輸入正確日期格式!\n';*/
            if (!isBlank(obj1) && !checkDate(obj1.value)) strMsg = strName + '格式有誤，請輸入正確日期格式!\n';
            break;

        case 'date_must': /*判斷必填日期欄位*/
            if (isBlank(obj1)) strMsg = '請輸入' + strName + '!\n';
            else if (!checkDate(obj1.value)) strMsg = strName + '格式有誤，請輸入正確日期格式!\n';
            /*else if (!checkRocDate(obj1.value)) strMsg = strName + '格式有誤，請輸入正確日期格式!\n';*/
            break;

        case 'yearrange': /*判斷非必填年度*/
            if (!isBlank(obj1) && !isBlank(obj2)) {
                if (obj1.value > obj2.value) strMsg += strName + '起年' + '不可大於' + strName + '迄年!\n';
            }
            break;

        case 'yearrange_must': /*判斷必填年度*/
            if (isBlank(obj1)) strMsg += '請選擇' + strName + '起年!\n';
            if (isBlank(obj2)) strMsg += '請選擇' + strName + '迄年!\n';

            if (!isBlank(obj1) && !isBlank(obj2)) {
                if (Number(obj1.value) > Number(obj2.value)) strMsg += strName + '起年' + '不可大於' + strName + '迄年!\n';
            }
            break;

        case 'yearmonth_must': /*判斷必填年度月份*/
            if (isBlank(obj1)) strMsg += '請選擇' + strName + '年度!\n';
            if (isBlank(obj2)) strMsg += '請選擇' + strName + '起月!\n';
            if (isBlank(obj3)) strMsg += '請選擇' + strName + '迄月!\n';

            if (!isBlank(obj2) && !isBlank(obj3)) {
                if (Number(obj2.value) > Number(obj3.value)) strMsg += strName + '起月' + '不可大於' + strName + '迄月!\n';
            }
            break;

        case 'daterange': /*判斷非必填日期區間*/
            if (!isBlank(obj1) && !checkDate(obj1.value)) strMsg += strName + '起始日期格式有誤，請輸入正確日期格式!\n';
            if (!isBlank(obj2) && !checkDate(obj2.value)) strMsg += strName + '結束日期格式有誤，請輸入正確日期格式!\n';
            /*if (!isBlank(obj1) && !checkRocDate(obj1.value)) strMsg += strName + '起始日期格式有誤，請輸入正確日期格式!\n';*/
            /*if (!isBlank(obj2) && !checkRocDate(obj2.value)) strMsg += strName + '結束日期格式有誤，請輸入正確日期格式!\n';*/
            if (checkDate(obj1.value) && checkDate(obj2.value)) {
                /*if (getDiffDay(getAdDate(obj1.value), getAdDate(obj2.value)) < 0) strMsg += strName + '結束日期不可小於起始日期!\n';*/
                if (getDiffDay(YMDD2MDDY(obj1.value), YMDD2MDDY(obj2.value)) < 0) strMsg += strName + '結束日期不可小於起始日期!\n';
            }
            break;

        case 'daterange_must': /*判斷必填日期區間*/
            if (isBlank(obj1)) strMsg += '請輸入' + strName + '起始日期!\n';
            else if (!checkDate(obj1.value)) strMsg += strName + '起始日期格式有誤，請輸入正確日期格式!\n';
            /*else if (!checkRocDate(obj1.value)) strMsg += strName + '起始日期格式有誤，請輸入正確日期格式!\n';*/

            if (isBlank(obj2)) strMsg += '請輸入' + strName + '結束日期!\n';
            else if (!checkDate(obj2.value)) strMsg += strName + '結束日期格式有誤，請輸入正確日期格式!\n';
            /*else if (!checkRocDate(obj2.value)) strMsg += strName + '結束日期格式有誤，請輸入正確日期格式!\n';*/
            if (checkDate(obj1.value) && checkDate(obj2.value)) {
                /*if (getDiffDay(getAdDate(obj1.value), getAdDate(obj2.value)) < 0) strMsg += strName + '結束日期不可小於起始日期!\n';*/
                if (getDiffDay(YMDD2MDDY(obj1.value), YMDD2MDDY(obj2.value)) < 0) strMsg += strName + '結束日期不可小於起始日期!\n';
            }
            /* if (checkRocDate(obj1.value) && checkRocDate(obj2.value)) { if (getDiffDay(getAdDate(obj1.value), getAdDate(obj2.value)) < 0) strMsg += strName + '結束日期不可小於起始日期!\n'; } */
            break;
        case 'mobile': /*判斷行動電話*/
            if (isBlank(obj1)) {
                strMsg = '請輸入' + strName + '!\n';
            } else if (!isUnsignedInt(obj1.value)) strMsg = strName + '格式有誤，請輸入數字!\n';
            else if (obj1.value.length < 10 || obj1.value.length > 10) {
                strMsg = strName + '請輸入十碼!\n';
            }
            break;

            /*非常用新增區↓↓↓*/
        case 'chgpass': /*判斷修改密碼*/
            if (isBlank(obj1)) strMsg += '請輸入舊密碼!\n';
            if (isBlank(obj2)) strMsg += '請輸入新密碼!\n';
            else if (obj2.value.length < 12 || obj2.value.length > 16) strMsg += '新密碼長度應為12~16碼，且需至少為英數字混合!\n';
            else if (!chkPassFmt(obj2.value)) strMsg += '新密碼長度應為12~16碼，且需至少為英數字混合!\n';

            if (isBlank(obj3)) strMsg += '請輸入確認新密碼!\n';
            if (obj2.value != obj3.value) strMsg += '新密碼與確認新密碼輸入的內容不同!\n';
            break;

        case 'auth': /*判斷帳號4~16英數字*/
            if (isBlank(obj1)) strMsg += '請輸入' + strName + '!\n';
            else if (obj1.value.length < 4 || obj1.value.length > 16) strMsg += strName + '請輸入4~16位英數字!\n';
            else if (!isIntEng(obj1.value)) strMsg += strName + '請輸入正確4~16位英數字!\n';
            break;

        case 'pass': /*判斷密碼12~16英數字*/
            if (isBlank(obj1)) strMsg += '請輸入' + strName + '!\n';
            else if (obj1.value.length < 12 || obj1.value.length > 16) strMsg += strName + '長度應為12~16碼，且需至少為英數字混合!\n';
            else if (!chkPassFmt(obj1.value)) strMsg += strName + '長度應為12~16碼，且需至少為英數字混合!\n';
            break;

        case 'email':
        case 'mail': /*判斷電子郵件格式*/
            if (isBlank(obj1)) strMsg = '請輸入' + strName + '!\n';
            else if (obj1.value.indexOf('@') < 0 || obj1.value.split('@')
                .length != 2) strMsg = strName + '格式有誤!\n';
            break;

        case 'idno': /*身分證號必填&格式*/
            if (isBlank(obj1)) strMsg = '請輸入' + strName + '!\n';
            else if (!checkNID(obj1.value)) strMsg = strName + '格式有誤!\n';
            break;

        case 'radio':
        case 'radio_must':
            if (getRBLValue(obj1.id) == "") {
                strMsg = '請點選' + strName + '!\n';
            }
            break;
        case 'checkboxlist_must':
        case 'checkboxlist':
            if (parseInt(getCheckBoxListValue(obj1.id), 10) == 0) {
                strMsg = '請勾選' + strName + '!\n';
            }
            break;
        default:
            strMsg += strFlag + ' 查無 ' + strName + ' 物件!\n';
            break;
    }

    return strMsg;
}

/*設定*/
function PL_focusState(xObj) {
    var obj1 = document.getElementById(xObj);
    var msgX1 = document.getElementById(xObj)
        .attributes["placeholder"].value;
    if (obj1.value == msgX1) {
        obj1.value = "";
    }
    obj1.style.color = "#000000";
}

/*設定*/
function PL_blurState(xObj) {
    var obj1 = document.getElementById(xObj);
    var msgX1 = document.getElementById(xObj).attributes["placeholder"].value;
    obj1.style.color = "#000000";
    if (obj1.value == "") {
        obj1.value = msgX1;
        obj1.style.color = "#666666";
    }
}

/*設定*/
function INPUT_readOnly(tObj) {
    tObj.style.backgroundColor = '#cccccc';
    tObj.style.color = '#666666';
    tObj.readOnly = true;
}

/*設定*/
function INPUT_readOnly2(tObj, tiMsg) {
    tObj.style.backgroundColor = '#cccccc';
    tObj.style.color = '#666666';
    tObj.readOnly = true;
    tObj.setAttribute('title', tiMsg);
}

/*設定*/
function INPUT_readOnlyU(tObj) {
    tObj.style.backgroundColor = '';
    tObj.style.color = '';
    tObj.readOnly = false;
    tObj.removeAttribute('title');
}

/*檢查*/
function checkTextLength(obj, length) {
    /*限定textbox的欄位長度*/
    if (obj.value.length > length) {
        obj.value = obj.value.substring(0, length);
        alert("限欄位長度不能大於" + length + "個字元(含空白字元)，超出字元將自動截斷");
    }
}

/*清除日曆物件資料*/
function clearDate(objId) {
    var myObj = document.getElementById(objId);

    if (myObj) {
        myObj.value = "";
    }
}

/* * checkzip23  地址-郵遞區號 後3碼 /後2碼 問題檢測
 * @param   flag_must 必須輸入
 * @param   tt1 中文名稱
 * @param   oID1	㯗位ID
 * @return  string l_msg 錯誤訊息
 */
function checkzip23(flag_must, tt1, oID1) {
    /*var tt1 = '通訊地址'; var oID1 = 'ZipCode1_B3';*/
    let voID1 = '#' + oID1; /*var flag_must = false;*/
    let l_msg = "";

    if (isEmpty(oID1)) {
        if (flag_must) {
            l_msg += "請輸入" + tt1 + "郵遞區號(3+3/3+2)，後3碼或2碼!!\n";
            return l_msg;
        }
        return l_msg;
    }
    if ($(voID1)
        .val()
        .length != 2 && $(voID1)
            .val()
            .length != 3) {
        l_msg += tt1 + "郵遞區號(3+3/3+2)，後3碼或2碼，長度只能為2或3!!\n";
        return l_msg;
    }
    if ($(voID1)
        .val()
        .length == 2) {
        if (!isInt(getValue(oID1))) {
            l_msg += tt1 + "郵遞區號後2碼.請輸入數字!!\n";
            return l_msg;
        }
        if (parseInt($(voID1)
            .val(), 10) < 1) {
            l_msg += tt1 + '郵遞區號後2碼必須為數字，得輸入 01~99 \n';
            return l_msg;
        }
        if (parseInt($(voID1)
            .val(), 10) > 99) {
            l_msg += tt1 + '郵遞區號後2碼必須為數字，得輸入 01~99 \n';
            return l_msg;
        }
    }
    if ($(voID1)
        .val()
        .length == 3) {
        if (!isInt(getValue(oID1))) {
            l_msg += tt1 + "郵遞區號後3碼.請輸入數字!!\n";
            return l_msg;
        }
        if (parseInt($(voID1)
            .val(), 10) < 1) {
            l_msg += tt1 + '郵遞區號後3碼必須為數字，得輸入 001~999 \n';
            return l_msg;
        }
        if (parseInt($(voID1)
            .val(), 10) > 999) {
            l_msg += tt1 + '郵遞區號後3碼必須為數字，得輸入 001~999 \n';
            return l_msg;
        }
    }
    return l_msg;
}

/*(僅數字)驗證數字*/
function isValidNumber(value) {
    return typeof value === 'number' && !Number.isNaN(value);
}

/*(僅數字)驗證數字可以整除0.5 */
function isDivisibleByHalf(num) {
    /*Check if the remainder when dividing by 0.5 is 0*/
    return num % 0.5 === 0;
}

/*文字轉數字 並且防呆*/
function convertToNumber(text) {
    /* Convert the input to a number type*/
    const num = Number(text);
    /* Check if it's a valid, finite number*/
    if (Number.isFinite(num)) {
        return num; /* Return the converted number*/
    } else {
        /* Handle cases where conversion failed (e.g., "abc", undefined, null, empty string)
        * You can return a default value, throw an error, or return null/undefined
        * console.warn(`Warning: "${text}" could not be converted to a finite number.`);
        * null; Or 0, or throw new Error("Invalid number input"); */
        return 0;
    }
}

/**
 * 驗證給定字串是否為有效數字。支援整數、小數、正數和負數。
 * @param {string} str - 需要驗證的字串。* @returns {boolean} 如果是有效數字則回傳 true，否則回傳 false。
 */
function isValidNumberUsingRegex(str) {
    /* 檢查輸入是否為字串類型，若不是則直接回傳 false*/
    if (typeof str !== 'string') { return false; }
    /*
    * 定義一個正規表達式來匹配數字,這個正規表達式表示：
    * ^      : 字串的開頭
    * [+-]?  : 可選的正號 (+) 或負號 (-)
    * \d+    : 一個或多個數字 (0-9)
    * (\.\d+)? : 可選的小數部分 (由一個點 '.' 和一個或多個數字組成)
    * $      : 字串的結尾
    */
    const numberRegex = /^[+-]?\d+(\.\d+)?$/;

    /* 使用 test() 方法來檢查字串是否符合正規表達式*/
    return numberRegex.test(str);
}
