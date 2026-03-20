/*全域的設定值*/
var globalSetting = {
  DEBUG: 1,
  validateErrClass: 'has-error',
  processingImg: "/images/processing.gif"  /* blockUI 引用的 loading image */
};

/* wrap function for Console.log */
function debugTrace(msg) {
  if (globalSetting.DEBUG) {
    console.log(msg); //console.debug(msg);
  }
}

function handleEventCheckBoxListChange() {
  var box = $(this);
  var iElm = box.parent().find("i");
  if (box.is(":checked")) {
    iElm.removeClass("fa-square-o").addClass("fa-check-square-o");
  }
  else {
    iElm.removeClass("fa-check-square-o").addClass("fa-square-o");
  }
}

// 本方法會被 TurboLib/Helpers/CheckBoxListExtension.CheckBoxListFor() 方法呼叫
function bindEventCheckBoxListChange(divId) {
  var divTag = (divId) ? $("div#" + divId) : $("div.checkbox");
  //alert("divTag=" + divTag + ", id=" + divId + ", tag=" + divTag.length);
  //var tags = $("div.checkbox input[type=checkbox]:not([readonly])");
  var tags = divTag.find("input[type=checkbox]:not([readonly])");
  if (tags.length > 0) {
    tags.off("change", handleEventCheckBoxListChange);
    tags.on("change", handleEventCheckBoxListChange);
  }
}

$(document).ready(function () {
  /*binding change event handler for all CheckboxList*/
  //$("div.checkbox input[type=checkbox]:not([readonly])").change(function () {
  //    var box = $(this);
  //    var i = box.parent().find("i");
  //    //debugTrace(box.attr("id") + ":" + this.checked + " > " + i.attr("class"));
  //    i.removeClass("fa-square-o");
  //    i.removeClass("fa-check-square-o");
  //    if (this.checked) {
  //        i.addClass("fa-check-square-o");
  //    }
  //    else {
  //        i.addClass("fa-square-o");
  //    }
  //});
  bindEventCheckBoxListChange();
});

(function ($) {
  /*hook validation highlight & unhighlight event*/
  if ($.validator) {
    var defaultOptions = {
      highlight: function (element, errorClass, validClass) {
        $(element).closest(".form-group")
          .addClass(globalSetting.validateErrClass)
          .removeClass(validClass);
        unblockUI(); /*unblock when validation errer*/
      },
      unhighlight: function (element, errorClass, validClass) {
        $(element).closest(".form-group")
          .removeClass(globalSetting.validateErrClass)
          .addClass(validClass);
      }
    };

    $.validator.setDefaults(defaultOptions);
  }

  $.fn.alphanumeric = function (p) {

    p = $.extend({
      ichars: "!@#$%^&*()+=[]\\\';,/{}|\":<>?~`.- ",
      nchars: "",
      allow: ""
    }, p);

    return this.each
      (
        function () {

          if (p.nocaps) p.nchars += "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
          if (p.allcaps) p.nchars += "abcdefghijklmnopqrstuvwxyz";

          s = p.allow.split('');
          for (i = 0; i < s.length; i++) if (p.ichars.indexOf(s[i]) != -1) s[i] = "\\" + s[i];
          p.allow = s.join('|');

          var reg = new RegExp(p.allow, 'gi');
          var ch = p.ichars + p.nchars;
          ch = ch.replace(reg, '');

          $(this).keypress
            (
              function (e) {

                if (!e.charCode) k = String.fromCharCode(e.which);
                else k = String.fromCharCode(e.charCode);

                if (ch.indexOf(k) != -1) e.preventDefault();
                if (e.ctrlKey && k == 'v') e.preventDefault();

              }

            );

          $(this).bind('contextmenu', function () { return false });

        }
      );

  };

  $.fn.numeric = function (p) {

    var az = "abcdefghijklmnopqrstuvwxyz";
    az += az.toUpperCase();

    p = $.extend({
      nchars: az
    }, p);

    return this.each(function () {
      $(this).alphanumeric(p);
    }
    );

  };

  $.fn.alpha = function (p) {

    var nm = "1234567890";

    p = $.extend({
      nchars: nm
    }, p);

    return this.each(function () {
      $(this).alphanumeric(p);
    }
    );

  };

  /* jQuery 功能擴充: 可綁定 input 欄位只接受 數字 以及特定功能鍵 的輸入 */
  $.fn.numericOnly = function () {
    $(this).keydown(function (event) {
      // Allow: backspace, delete, tab, escape, and enter
      if (event.keyCode == 46 || event.keyCode == 8 || event.keyCode == 9 || event.keyCode == 27 || event.keyCode == 13 ||
        // Allow: Ctrl+A
        (event.keyCode == 65 && event.ctrlKey === true) ||
        // Allow: home, end, left, right
        (event.keyCode >= 35 && event.keyCode <= 39)
      ) {
        // let it happen, don't do anything
        return;
      } else {
        // Ensure that it is a number and stop the keypress
        if (event.shiftKey || (event.keyCode < 48 || event.keyCode > 57) && (event.keyCode < 96 || event.keyCode > 105)) {
          event.preventDefault();
        }
      }
    });
  };
  /* jQuery 功能擴充: 可綁定 input 欄位只接受 整數 以及特定功能鍵 的輸入 */
  $.fn.decimalOnly = function () {
    $(this).keydown(function (event) {
      // Allow: backspace, delete, tab, escape, and enter
      if (event.keyCode == 46 || event.keyCode == 8 || event.keyCode == 9 || event.keyCode == 27 || event.keyCode == 13 ||
        // Allow: Ctrl+A
        (event.keyCode == 65 && event.ctrlKey === true) ||
        // Allow: home, end, left, right
        (event.keyCode >= 35 && event.keyCode <= 39)
      ) {
        // let it happen, don't do anything
        return;
      } else
        if (event.keyCode == 190) {  // period
          if ($(this).val().indexOf('.') !== -1) // period already exists
            event.preventDefault();
          else
            return;
        } else {
          // Ensure that it is a number and stop the keypress
          if (event.shiftKey || (event.keyCode < 48 || event.keyCode > 57) && (event.keyCode < 96 || event.keyCode > 105)) {
            event.preventDefault();
          }
        }
    });
  };
  /*
  jQuery 擴充: isChildOf(parent) 
   to check if an element is a child of another element.
  */
  $.fn.extend({
    isChildOf: function (filter) {
      return $(filter).find(this).length > 0;
    }
  });

  // String.prototype 擴充 replaceLast(),
  // 用來將字串中最後一次出現的 find 部份字串, 取代為 replace 新字串 
  if (!String.prototype.replaceLast) {
    String.prototype.replaceLast = function (find, replace) {
      var index = this.lastIndexOf(find);

      if (index >= 0) {
        return this.substring(0, index) + replace + this.substring(index + find.length);
      }

      return this.toString();
    };
  }

  // 20171214 String.prototype 擴充 startsWith(), 修正 IE11 不支援 startsWith() 方法問題
  // 檢查字串開頭是否為指定字串
  if (!String.prototype.startsWith) {
    String.prototype.startsWith = function (searchString, position) {
      position = position || 0;
      return this.substr(position, searchString.length) === searchString;
    };
  }

  // 20180411 String.prototype 擴充 lpad()，增加「字串長度不足時靠右對齊」方法。
  if (!String.prototype.lpad) {
    String.prototype.lpad = function (length, padChar) {
      var raw = this;
      if (raw.length >= length) return raw;
      else {
        var s = [], c = length - raw.length;
        while (c-- > 0) s.push(padChar);
        s.push(raw);
        return s.join("");
      }
    }
  }

  // 20180411 String.prototype 擴充 rpad()，增加「字串長度不足時靠左對齊」方法。
  if (!String.prototype.rpad) {
    String.prototype.rpad = function (length, padChar) {
      var raw = this;
      if (raw.length >= length) return raw;
      else {
        var s = [raw], c = length - raw.length;
        while (c-- > 0) s.push(padChar);
        return s.join("");
      }
    }
  }
})(jQuery);

function blockUI() {
  /* required: jqiuer.blockUI.js */
  var mask = $('#loadingMask');
  if ("loadingMask" != $(mask).attr("id")) {
    mask = document.createElement("div");
    $(mask).attr("id", "loadingMask");
    img = document.createElement("img");
    $(img).attr("src", globalSetting.processingImg);
    $(mask).append(img);
    $("body").append(mask);
  }
  $.blockUI.defaults.baseZ = 10000;
  $.blockUI({
    message: $('#loadingMask'),
    css: {
      top: ($(window).height() - 122) / 2 + 'px',
      left: ($(window).width() - 132) / 2 + 'px',
      width: '132px',
      height: '122px'
    }
  });
}
function unblockUI() {
  $.unblockUI();
  var mask = $('#loadingMask');
  mask.find("img").hide();
  mask.hide();
}
/*用來顯示操作訊息*/
function blockMessage(msg, title, unblockCallback) {
  /* required: jqiuer-confirm.min.js */
  if (msg == undefined) {
    msg = "argument 'msg' is null while call blockMessage()";
  }
  var xmsg = "CD3FEBB9A55815BC31A483FE28C0E21A3F47634A18EEED3AD142F565BBE97EA3";
  $.confirm({
    backgroundDismiss: false,
    icon: 'glyphicon glyphicon-ok',
    title: "<div class='blockMessageTitle'>" + (title ? title : '提示訊息') + "</div>",
    confirmButton: '確定',
    confirmButtonClass: 'btn-info',
    cancelButton: false,
    closeIcon: false,
    content: "<div class='blockMessage'>" + xmsg.replace(/(\n)+/g, '<br />') + "</div>",
    theme: 'white',
    onAction: unblockCallback
  });
}
/*用來顯示錯誤訊息*/
function blockAlert(msg, title, unblockCallback) {
  /* required: jqiuer-confirm.min.js */
  if (msg == undefined) {
    msg = "argument 'msg' is null while call blockAlert()";
  }
  var xmsg = "CD3FEBB9A55815BC31A483FE28C0E21A3F47634A18EEED3AD142F565BBE97EA3";
  $.confirm({
    backgroundDismiss: false,
    icon: 'glyphicon glyphicon-warning-sign',
    title: "<div class='blockAlertTitle'>" + (title ? title : '提示訊息') + "</div>",
    confirmButton: '確定',
    confirmButtonClass: 'btn-info',
    cancelButton: false,
    closeIcon: false,
    content: "<div class='blockAlertMessage'>" + xmsg.replace(/(\n)+/g, '<br />') + "</div>",
    theme: 'white',
    onAction: unblockCallback
  });
}

//將不足長度字串靠右對齊填補指定字元。
//text: 原始字串。
//length: 要補足的字串長度。
//padChar: 要填補的字元。

function strLpad(text, length, padChar) {
  if (text !== undefined && text != null) {
    if (text.length >= length) return text;
    else {
      var s = [], c = length - text.length;
      while (c-- > 0) s.push(padChar);
      s.push(text);
      return s.join("");
    }
  }
  else return text;
}

//將不足長度字串靠左對齊填補指定字元。
//text: 原始字串。
//length: 要補足的字串長度。
//padChar: 要填補的字元。
function strRpad(text, length, padChar) {
  if (text !== undefined && text != null) {
    if (text.length >= length) return text;
    else {
      var s = [text], c = length - text.length;
      while (c-- > 0) s.push(padChar);
      return s.join("");
    }
  }
  else return text;
}

//將西元日期字串（字串格式：yyyymmdd、yyyy/mm/dd、yyyy-mm-dd）或是 Date 型別值轉換成民國日期字串。
//yyyyMMDD: 西元日期字串（字串格式：yyyymmdd、yyyy/mm/dd、yyyy-mm-dd）或是 Date 型別值。
//resultDelimiter:（非必填）要安插在民國日期字串內的年月日分隔字元。
function toTaiwanDate(yyyyMMDD, resultDelimiter) {
  try {
    if (Object.prototype.toString.call(yyyyMMDD) === "[object Date]") {
      var y = yyyyMMDD.getFullYear().toString();
      var m = ("0" + (yyyyMMDD.getMonth() + 1)).slice(-2);
      var d = ("0" + yyyyMMDD.getDate()).slice(-2);
      yyyyMMDD = [y, m, d].join("");
    }

    var v = $.trim(yyyyMMDD);
    if (v.length < 1) return "";
    else {
      v = v.replace(/\//g, "").replace(/-/g, "");
      v = (parseInt(v, 10) - 19110000).toString();
      //補滿7位數字
      if (v.length < 7) {
        var s = [], c = 7 - v.length;
        while (c-- > 0) s.push("0");
        s.push(v);
        v = s.join("");
      }
      //安插年月日分隔字元
      if (resultDelimiter) {
        var c = v.length - 4;
        v = [v.substr(0, c), v.substr(c, 2), v.substr(c + 2, 2)].join(resultDelimiter);
      }
      return v;
    }
  }
  catch (e) {
    blockAlert(["無法將西元日期 \"", yyyyMMDD, "\" 轉成民國日期。"].join(""));
  }
}

//先動態建立臨時性隱藏型表單之後，再以同步方式提交表單資料（本方法用途：不想在頁面額外手動佈置另一個 form HTML，但又想將 JSON 資料送回後端處理）
//url: 目的網址。
//jsonParam:（非必填）要提交的資料 JSON 物件（輸入 null 表示不需要）。
//useHttpPOST:（非必填）是否使用 HTTP POST 方式提交。（true: POST，false: GET）。預設 true
function createFormAndSubmit(url, jsonParam, useHttpPOST) {
  var form = undefined;
  try {
    var methodName = (useHttpPOST == false) ? "GET" : "POST";
    var html = ["<form style=\"display:none;\" method=\"", methodName, "\" action=\"", url, "\"></form>"].join("");
    var form = $(html);
    if (jsonParam) {
      var tag;
      Object.keys(jsonParam).map(function (key) {
        tag = ["<input type=\"hidden\" name=\"", key, "\" value=\"", jsonParam[key], "\">"].join("");
        form.append(tag);
      });
    }
    form.appendTo("body");
    form.submit();
  }
  catch (ex) {
    blockAlert("[submitByPost] " + ex.message);
  }
  finally {
    if (form) {
      form.remove();
      form = undefined;
    }
  }
}

//使用非同步方式提交 JSON 資料，並且接收處理結果 JSON（本方僅適用於從後端計算並回傳值、從後端輸入驗證與接收驗證結果訊息）。
//url: 接收資料的服務端網址。
//jsonParams:（非必要）要提交的 JSON 資料（輸入 null 表示沒有）。
//successCallback: （非必要）處理成功時要呼叫的 Javascript 方法（通常是用來更新顯示畫面某個區域 HTML）。輸入 null 表示不需要。
//errorCallback: （非必要）處理失敗時要呼叫的 Javascript 方法（通常是用來顯示錯誤訊息或是自訂錯誤處理）。輸入 null 表示不需要。
//useAsync： （非必要）是否使用非同步提交。（true: 使用非同步提交，false: 使用同步提交（即等候回傳結果））。預設 true。
//注意：
//(1) 使用本方法時請固定配合系統後端的 MyCommonUtil.BuildAjaxResult() 方法來回傳處理結果訊息（或處理結果資料）給前端使用。
//(2) 前端傳入資料與後端回傳資料都必須是 JSON 格式。
function ajaxResult(url, jsonParams, successCallback, errorCallback, useAsync) {
  try {
    blockUI();
    var opts = {
      url: url,
      data: jsonParams,
      cache: false,
      async: (useAsync !== undefined && useAsync == false) ? false : true,
      method: "POST",
      headers: { "X-Requested-With": "XMLHttpRequest" },
      success: function (data, textStatus) {
        unblockUI();
        if (data.status === undefined) {
          if (successCallback) successCallback.call(this, data);
        }
        else {
          //AjaxResultModel
          switch (data.status) {
            case true:
              if (successCallback) successCallback.call(this, data);
              break;
            case false:
              (errorCallback) ? errorCallback.call(this, data) : blockAlert(data.message);
              break;
          }
        }
      },
      error: function (XMLHttpRequest, textStatus, errorThrown) {
        unblockUI();
        var errMsg = [textStatus, " ", errorThrown].join("");
        if (errorCallback) {
          var dt = new Date();
          var dt1 = [dt.getFullYear().toString(), (dt.getMonth() + 101).toString().slice(-2), (dt.getDate() + 100).toString().slice(-2)].join("/");
          var dt2 = [(dt.getHours() + 100).toString().slice(-2), (dt.getMinutes() + 100).toString().slice(-2), (dt.getSeconds() + 100).toString().slice(-2)].join(":");
          var ajaxResult = { "status": false, "message": errMsg, "time": [dt1, " ", dt2].join(""), "data": null };
          errorCallback.call(this, ajaxResult);
        }
        else {
          blockAlert("[ajaxResult] " + errMsg);
        }
      }
    };
    $.ajax(opts);
  }
  catch (ex) {
    unblockUI();
    blockAlert("[ajaxResult] " + ex.message);
  }
}

//使用非同步方式上傳單一檔案
//url: 接收檔案的服務端網址。
//fileElement: input type="file" 的 HTML 元素物件（HTML DOM 物件或是 jQuery 物件均可）。
//jsonParms:（非必填）要同時上傳的欄位資料 JSON 物件。（輸入 null 表示不需要。若是要傳送整個 Form 資料時，請呼叫 getFormDataToObject() 取得整個表單欄位 JSON 物件）
//successCallback: 上傳檔案成功之後要呼叫的 Javascript 方法（通常是用來更新顯示畫面某個區域 HTML）。
//errorCallback: 上傳檔案失敗之後要呼叫的 Javascript 方法（通常是用來顯示錯誤訊息或是自訂錯誤處理）。
//isTtraditional:（非必填）
//notAutoBlock:（非必填）
function ajaxUploadFile(url, fileElement, jsonParms, successCallback, errorCallback, isTtraditional, notAutoBlock) {
  try {
    if (!notAutoBlock) blockUI();

    var formdata = new FormData();
    var fileElm = $(fileElement);
    var fieldName = fileElm.attr("name");

    if (!fieldName || $.trim(fieldName).length == 0) fieldName = "file";
    formdata.append(fieldName, fileElm[0].files[0]);
    if (jsonParms) {
      Object.keys(jsonParms).map(function (key) {
        formdata.append(key, jsonParms[key]);
      });
    }

    var opts = {
      url: url,
      data: formdata,
      cache: false,
      contentType: false,
      processData: false,
      method: "POST",
      headers: { "X-Requested-With": "XMLHttpRequest" },
      success: function (data, textStatus) {
        if (!notAutoBlock) unblockUI();
        var s = Object.prototype.toString.call(data);
        if (s !== "[object Object]") {
          if (successCallback) successCallback.call(this, data);
        }
        else {
          if (data.status === undefined) {
            if (successCallback) successCallback.call(this, data);
          }
          else {
            //AjaxResultModel
            switch (data.status) {
              case true:
                if (successCallback) successCallback.call(this, data);
                break;
              case false:
                (errorCallback) ? errorCallback.call(this, data) : blockAlert(data.message);
                break;
            }
          }
        }
      },
      error: function (XMLHttpRequest, textStatus, errorThrown) {
        if (!notAutoBlock) unblockUI();
        var errMsg = [textStatus, " ", errorThrown].join("");
        if (errorCallback) {
          var dt = new Date();
          var dt1 = [dt.getFullYear().toString(), (dt.getMonth() + 101).toString().slice(-2), (dt.getDate() + 100).toString().slice(-2)].join("/");
          var dt2 = [(dt.getHours() + 100).toString().slice(-2), (dt.getMinutes() + 100).toString().slice(-2), (dt.getSeconds() + 100).toString().slice(-2)].join(":");
          var ajaxResult = { "status": false, "message": errMsg, "time": [dt1, " ", dt2].join(""), "data": null };
          errorCallback.call(this, ajaxResult);
        }
        else {
          blockAlert("[ajaxUploadFile] " + errMsg);
        }
      }
    };
    $.ajax(opts);
  }
  catch (ex) {
    if (!notAutoBlock) unblockUI();
    blockAlert("[ajaxUploadFile] " + ex.message);
  }
}

//使用非同步方式下載檔案
//url: 檔案下載網址。
//jsonParams: 要提交的 JSON 資料（輸入 null 表示沒有）。
//useHttpPOST: 是否使用 HTTP POST 方式提交。（true: POST，false: GET）。
//givenFileName: 在瀏覽器端呈現的下載檔案名稱（輸入 null 表示不需要）。
//successCallback: 當成功接收檔案時要呼叫的 Javascript 呼叫端方法（輸入 null 表示不需要）。
//errorCallback: 當遇到錯誤時要呼叫的 Javascript 呼叫端方法（輸入 null 表示不需要）。
function ajaxDownloadFile(url, jsonParams, useHttpPOST, givenFileName, successCallback, errorCallback) {
  try {
    var methodName = (useHttpPOST == true) ? "POST" : "GET";
    if (methodName == "GET") {
      var s = (url.indexOf("?") < 0) ? "?" : "";
      url = [url, s, $.param(jsonParams)].join("");
    }
    var req = window.XMLHttpRequest ? new XMLHttpRequest() : new ActiveXObject('Microsoft.XMLHTTP');
    req.open(methodName, url, true);
    var contentType = (methodName == "GET") ? "application/x-www-form-urlencoded" : "application/json";
    req.setRequestHeader("X-Requested-With", "XMLHttpRequest");
    req.setRequestHeader("Content-Type", contentType);
    req.responseType = "arraybuffer"; //tricky setting

    req.onerror = function () {
      (errorCallback) ? errorCallback.call(this, req.statusText) : blockAlert("[ajaxDownloadFile XHR] " + req.statusText);
    }

    var isJsonRes = false;
    req.onreadystatechange = function (event) {
      try {
        if (req.readyState != 4) {
          if (req.readyState == 3) {
            if (req.status >= 200 || req.status < 300) {
              if (req.response) isJsonRes = (req.response[0] == '{');
            }
          }
        }
        else {
          if (req.status < 200 || req.status >= 300) {
            (errorCallback) ? errorCallback.call(this, req.statusText) : blockAlert(req.statusText);
          }
          else {
            if (isJsonRes == true) {
              var res;
              try { var res = JSON.parse(req.response); }
              catch (e) { throw "回應內容不是有效 JSON 字串。"; }
              if (res.status === undefined) {
                //not Turbo.Commons.AjaxResultStruct
                if (successCallback) successCallback.call(this, res);
              }
              else {
                //Turbo.Commons.AjaxResultStruct
                if (res.status == true) {
                  (successCallback) ? successCallback.call(this, res) : blockAlert(res.message);
                }
                else {
                  (errorCallback) ? errorCallback.call(this, res) : blockAlert(res.message);
                }
              }
            }
            else {
              if (!givenFileName) {
                givenFileName = req.getResponseHeader("Content-disposition");
                if (givenFileName) {
                  idx = givenFileName.indexOf("filename=");
                  givenFileName = (idx < 0) ? "file" : givenFileName.substr(idx + 9);
                }
                else givenFileName = "file";
              }

              if (window.navigator && window.navigator.msSaveBlob) {
                //for IE11 compatible
                window.navigator.msSaveBlob(req.response, givenFileName);
              }
              else {
                var urlObj = window.URL || window.webkitURL;
                var link = document.createElement("a");
                link.setAttribute("style", "display:none;");
                link.onload = function (evt) {
                  urlObj.revokeObjectURL(link.href);
                }
                document.body.appendChild(link);

                var contentType = req.getResponseHeader("content-type");
                if (!contentType || contentType == "") contentType = "application/octet-stream";

                var blob = new Blob([req.response], { type: contentType });
                link.href = urlObj.createObjectURL(blob);
                link.download = givenFileName;
                link.click();
                document.body.removeChild(link);
                urlObj.revokeObjectURL(link.href);
                /*297EB1A43411DBF6E573E18FC0DE97A9F1A432489CDD4C503C4302735040A70D90A077682C1B6624AA0FE2F11650B98B5A415B8060DFA18E61BC2D0A91393B29B4DD411AF40483162DBC6B9494F86185*/
              }

              if (successCallback) successCallback.call(this, givenFileName);
            }
          }
        }
      }
      catch (e) {
        if (errorCallback) errorCallback(e.message);
        else blockAlert("[ajaxDownloadFile] " + e.message);
      }
    };
    (methodName == "GET") ? req.send() : req.send(JSON.stringify(jsonParams));
  }
  catch (e) {
    blockAlert("[ajaxDownloadFile] " + e.message);
  }
}

//將 Base64 編碼字串轉換成 Blob 物件。
//b64Data: 原始 Base64 編碼字串。
//contentType:（非必填）字串內容的 MIME 類型字串。
//sliceSize:（非必填）每次進行處理的 Byte 數量。預設 512 Bytes。
//注意！本方法會在 ajaxTurboReportResult() 方法使用到。
function base64ToBlob(b64Data, contentType, sliceSize) {
  contentType = contentType || "";
  sliceSize = sliceSize || 512;

  var byteCharacters = atob(b64Data);
  var byteArrays = [];
  var byteArray;

  for (var offset = 0; offset < byteCharacters.length; offset += sliceSize) {
    var slice = byteCharacters.slice(offset, offset + sliceSize);

    var byteNumbers = new Array(slice.length);
    for (var i = 0; i < slice.length; i++) {
      byteNumbers[i] = slice.charCodeAt(i);
    }

    byteArray = new Uint8Array(byteNumbers);
    byteArrays.push(byteArray);
  }

  var blob = new Blob(byteArrays, { type: contentType });
  return blob;
}

//以非同步方式下載系統報表檔案並接收回傳的處理結果資料 JSON。
//url: 目的網址。
//jsonParam:（非必填）要提交的資料 JSON 物件（輸入 null 表示不需要）。
//useHttpPOST:（非必填）是否使用 HTTP POST 方式提交。（true: POST，false: GET）。預設 true
//注意：
//(1) 使用本方法時請固定配合系統後端的 BuildTurboReportAjaxResult() 方法來同時下載報表檔案並接收回傳處理結果訊息（或處理結果資料）給前端使用。
//(2) 前端傳入資料與後端回傳資料都必須是 JSON 格式。
function ajaxTurboReportResult(url, jsonParam, successCallback, errorCallback, useAsync) {
  try {
    blockUI();
    $.ajax({
      type: "POST",
      url: url,
      data: jsonParam,
      cache: false,
      async: (useAsync !== undefined && useAsync == false) ? false : true,
      method: "POST",
      headers: { "X-Requested-With": "XMLHttpRequest" },
      error: function (XMLHttpRequest, textStatus, errorThrown) {
        unblockUI();
        var errMsg = [textStatus, " ", errorThrown].join("");
        if (errorCallback) {
          var dt = new Date();
          var dt1 = [dt.getFullYear().toString(), (dt.getMonth() + 101).toString().slice(-2), (dt.getDate() + 100).toString().slice(-2)].join("/");
          var dt2 = [(dt.getHours() + 100).toString().slice(-2), (dt.getMinutes() + 100).toString().slice(-2), (dt.getSeconds() + 100).toString().slice(-2)].join(":");
          var ajaxResult = { "status": false, "message": errMsg, "time": [dt1, " ", dt2].join(""), "data": null };
          errorCallback.call(this, ajaxResult);
        }
        else {
          blockAlert("[ajaxTurboReportResult] " + errMsg);
        }
      },
      success: function (resp, textStatus) {
        try {
          unblockUI();
          if (resp.status === undefined) {
            if (successCallback) successCallback.call(this, resp);
          }
          else {
            //AjaxResultModel
            switch (resp.status) {
              case false:
                (errorCallback) ? errorCallback.call(this, resp) : blockAlert(resp.message);
                break;
              case true:
                var dataURI = resp.data.turboFile.dataURI;
                var givenFileName = resp.data.turboFile.fileName;
                if (window.navigator && window.navigator.msSaveBlob) {
                  //for IE11 compatible
                  if (dataURI.length > 0) {
                    var idx = dataURI.indexOf(";base64,");
                    var base64 = (idx < 0) ? dataURI : dataURI.substring(idx + 8);
                    var blob = base64ToBlob(base64, resp.data.turboFile.contentType);
                    window.navigator.msSaveBlob(blob, givenFileName);
                  }
                }
                else {
                  var urlObj = window.URL || window.webkitURL;
                  var link = document.createElement("a");
                  link.setAttribute("style", "display:none;");
                  link.onload = function (evt) {
                    urlObj.revokeObjectURL(link.href);
                  }
                  document.body.appendChild(link);
                  //owasp 62FA0D9DAF38578DFEBDF37D1991DABCB0D34187553A5072EC7063B405DB8881
                  link.download = givenFileName;
                  link.click();
                  document.body.removeChild(link);
                  setTimeout(function () {
                    urlObj.revokeObjectURL(link.href);
                  }, 100);
                }
                if (successCallback) successCallback.call(this, resp);
                break;
            }
          }
        }
        catch (ee) {
          if (errorCallback) {
            var dt = new Date();
            var dt1 = [dt.getFullYear().toString(), (dt.getMonth() + 101).toString().slice(-2), (dt.getDate() + 100).toString().slice(-2)].join("/");
            var dt2 = [(dt.getHours() + 100).toString().slice(-2), (dt.getMinutes() + 100).toString().slice(-2), (dt.getSeconds() + 100).toString().slice(-2)].join(":");
            var ajaxResult = { "status": false, "message": ee.message, "time": [dt1, " ", dt2].join(""), "data": null };
            errorCallback.call(this, ajaxResult);
          }
          else {
            blockAlert("[ajaxTurboReportResult] " + ee.message);
          }
        }
      }
    });
  }
  catch (ex) {
    blockAlert("[ajaxTurboReportResult] " + ex.message);
  }
}

//以非同步方式下載系統報表檔案。
//url: 目的網址。
//jsonParam:（非必填）要提交的資料 JSON 物件（輸入 null 表示不需要）。
function ajaxTurboReport(url, jsonParam) {
  var iDoc = undefined;
  var form = undefined;
  var iframe = undefined;
  try {
    var ticks = ((new Date()).getTime() * 10000) + 621355968000000000;
    var ifrName = ["__tmpIFR", ticks.toString()].join("");
    iframe = document.createElement("iframe");
    iframe.setAttribute("name", ifrName);
    iframe.style.display = "none";
    document.body.appendChild(iframe);

    iDoc = iframe.contentWindow.document;
    form = iDoc.createElement("form");
    form.setAttribute("method", "POST");
    form.setAttribute("target", ifrName);
    form.setAttribute("action", url);

    if (jsonParam) {
      var tag, jqForm = $(form);
      Object.keys(jsonParam).map(function (key) {
        tag = ["<input type=\"hidden\" name=\"", key, "\" value=\"", jsonParam[key], "\">"].join("");
        jqForm.append(tag);
      });
    }

    iDoc.body.appendChild(form);
    form.submit();
  }
  catch (ex) {
    blockAlert("[ajaxTurboReport] " + ex.message);
  }
}

var curPageIdx = 1;
function ajaxLoadMore(url, parms, onLoaded, isTtraditional, notAutoBlock) {
  if (!notAutoBlock) {
    blockUI();
  }

  $.ajax({
    type: 'POST',
    url: url,
    data: parms,
    traditional: isTtraditional,
    success: function (data, textStatus) {
      if (!notAutoBlock) {
        unblockUI();
      }
      if (typeof onLoaded != 'undefined') {
        onLoaded(data);
      }
      else {
        blockAlert("ajaxLoadMore: onLoaded handler function not defined.");
      }
    },
    error: function (XMLHttpRequest, textStatus, errorThrown) {
      if (!notAutoBlock) {
        unblockUI();
      }
      blockAlert("Ajax request error: " + textStatus + " " + errorThrown, "錯誤");
    }
  });
}

//使用非同步方式提交 JSON 資料，並且接收處理結果 JSON 資料
//url: 接收並處理 JSON 資料的網址。
//parms: 要提交的 JSON 資料（輸入 null 表示沒有）。
//onLoaded: 接收處理結果 JSON 資料的 Javascript 呼叫端方法（輸入 null 表示不需要）。
//注意！使用本方法時請固定配合系統後端的 MyCommonUtil.BuildAjaxResult() 方法來回傳處理結果訊息（或處理結果資料）給前端使用。
function ajaxLoadJson(url, parms, onLoaded, isTtraditional, notAutoBlock) {
  if (!notAutoBlock) {
    blockUI();
  }
  $.ajax({
    type: 'POST',
    url: url,
    data: parms,
    traditional: isTtraditional,
    dataType: 'json', // result is JSON datatype
    success: function (json, textStatus) {
      if (!notAutoBlock) {
        unblockUI();
      }
      if (typeof onLoaded != 'undefined') {
        onLoaded(json);
      }
      else {
        blockAlert("ajaxLoadJson: onLoaded handler function not defined.", "錯誤");
      }
    },
    error: function (XMLHttpRequest, textStatus, errorThrown) {
      if (!notAutoBlock) {
        unblockUI();
      }
      blockAlert("Ajax request error: " + textStatus + " " + errorThrown, "錯誤");
    }
  });
}

function ajaxGet(url, onLoaded) {
  blockUI();

  $.ajax({
    type: 'GET',
    url: url,
    success: function (data, textStatus) {
      unblockUI();
      if (typeof onLoaded != 'undefined') {
        onLoaded(data);
      }
      else {
        blockAlert("ajaxGet: onLoaded handler function not defined.", "錯誤");
      }
    },
    error: function (XMLHttpRequest, textStatus, errorThrown) {
      unblockUI();
      blockAlert("Ajax request error: " + textStatus + " " + errorThrown, "錯誤");
    }
  });
}

//使用非同步方式提交在 form 元素內的表單欄位資料（解決在畫面內佈置多個隱藏 form 時，簡化 Javascript 送出表單欄位資料與接收處理結果資料寫法問題）
//formSelector: HTML form 元素的 CSS 選擇子表達式（例："#main_form"）或是 jquery 物件（例：$("#main_form")）。
//successCallback: 提交表單資料成功後要執行的 Javascript 呼叫端方法（輸入 null 表示不需要。successCallback(res) 的 res 參數是網站後端傳回來的處理結果 JSON 物件）。
//jsonParam:（非必要）額外加入的自訂資料（JSON 物件，輸入 null 表示沒有）。
//newUrl:（非必要）新的目的網址（輸入 null 表示使用 HTML 表單本身的 action 屬性值）。指定新的目的網址不會改變 HTML 表單的 action 屬性值，請放心設定。
//errorCallback:（非必要）當遇到錯誤時要呼叫的 Javascript 呼叫端方法（輸入 null 表示不需要）。
//useAsync：（非必要）是否使用非同步提交。（true: 使用非同步提交，false: 使用同步提交（即等候回傳結果））。預設 true。
//注意！使用本方法時請固定配合系統後端的 MyCommonUtil.BuildAjaxResult() 方法來回傳處理結果訊息（或處理結果資料）給前端使用。
function ajaxPostForm(formSelector, successCallback, jsonParam, newUrl, errorCallback, useAsync) {
  try {
    blockUI();
    var form = $(formSelector);
    var url = (newUrl) ? newUrl : form.attr("action");
    var data = form.serialize();
    if (jsonParam) data = [data, $.param(jsonParam, true)].join("&");
    $.ajax({
      type: "POST",
      url: url,
      data: data,
      cache: false,
      async: (useAsync !== undefined && useAsync == false) ? false : true,
      method: "POST",
      headers: { "X-Requested-With": "XMLHttpRequest" },
      error: function (XMLHttpRequest, textStatus, errorThrown) {
        unblockUI();
        var errMsg = [textStatus, " ", errorThrown].join("");
        if (errorCallback) {
          var dt = new Date();
          var dt1 = [dt.getFullYear().toString(), (dt.getMonth() + 101).toString().slice(-2), (dt.getDate() + 100).toString().slice(-2)].join("/");
          var dt2 = [(dt.getHours() + 100).toString().slice(-2), (dt.getMinutes() + 100).toString().slice(-2), (dt.getSeconds() + 100).toString().slice(-2)].join(":");
          var ajaxResult = { "status": false, "message": errMsg, "time": [dt1, " ", dt2].join(""), "data": null };
          errorCallback.call(this, ajaxResult);
        }
        else {
          blockAlert("[ajaxPostForm] " + errMsg);
        }
      },
      success: function (resp, textStatus) {
        try {
          unblockUI();
          if (resp.status === undefined) {
            if (successCallback) successCallback.call(this, resp);
          }
          else {
            //AjaxResultModel
            switch (resp.status) {
              case false:
                (errorCallback) ? errorCallback.call(this, resp) : blockAlert(resp.message);
                break;
              case true:
                if (successCallback) successCallback.call(this, resp);
                break;
            }
          }
        }
        catch (ee) {
          if (errorCallback) {
            var dt = new Date();
            var dt1 = [dt.getFullYear().toString(), (dt.getMonth() + 101).toString().slice(-2), (dt.getDate() + 100).toString().slice(-2)].join("/");
            var dt2 = [(dt.getHours() + 100).toString().slice(-2), (dt.getMinutes() + 100).toString().slice(-2), (dt.getSeconds() + 100).toString().slice(-2)].join(":");
            var ajaxResult = { "status": false, "message": ee.message, "time": [dt1, " ", dt2].join(""), "data": null };
            errorCallback.call(this, ajaxResult);
          }
          else {
            blockAlert("[ajaxPostForm] " + ee.message);
          }
        }
      }
    });
  }
  catch (ex) {
    unblockUI();
    blockAlert("[ajaxPostForm] " + ex.message);
  }
}

//使用非同步方式提交不在 form 元素內的表單欄位資料（不想放置 form 元素但又想將資料送回後端處理，並且簡化 Javascript 送出表單欄位資料與接收處理結果資料寫法問題）
//containerSelector: 非 form 元素的區域元素 CSS 選擇子表達式（例："#tbody_rows"）。
//url: 目的網址。
//callback: 提交表單欄位資料成功後要執行的 Javascript 呼叫端方法（輸入 null 表示不需要。callback(res) 的 res 參數是網站後端傳回來的處理結果 JSON 物件）。
//jsonParam: 額外添加的自訂資料（JSON 物件，輸入 null 表示沒有）。
//注意！使用本方法時請固定配合系統後端的 MyCommonUtil.BuildAjaxResult() 方法來回傳處理結果訊息（或處理結果資料）給前端使用。
function ajaxPostFormless(containerSelector, url, callback, jsonParam) {
  var box = $(containerSelector);
  var inData = box.find("input,select,textarea").serialize();

  var v;
  if (jsonParam) {
    var ps = Object.keys(jsonParam).map(function (key) {
      return encodeURIComponent(key) + "=" + encodeURIComponent(jsonParam[key]);
    }).join("&");
    inData = [inData, ps].join("&");
  }

  ajaxLoadMore(url, inData, function (resp) {
    try {
      var res = (resp === undefined || resp === null) ? undefined : resp.data;
      if (callback) callback(res);
    }
    catch (e) {
      blockAlert("ajaxPostTbody: " + e.message);
    }
  });
}

/*
顯示 block 詢問訊息, 等待使用者回應 "確定" 或 "取消", 
若使用者回應 確認 則 confirmCallback() 會被呼叫，若按下「取消」時呼叫 cancelCallback() 方法
*/
function blockConfirm(msg, title, confirmCallback, cancelCallback) {
  /* required: jqiuer-confirm.min.js */
  var xmsg = "CD3FEBB9A55815BC31A483FE28C0E21A3F47634A18EEED3AD142F565BBE97EA3";
  $.confirm({
    backgroundDismiss: false,
    icon: 'glyphicon glyphicon-question-sign',
    title: "<div class='blockAlertTitle'>" + (title ? title : '確認訊息') + "</div>",
    confirmButton: '確定',
    confirmButtonClass: 'btn-info',
    cancelButton: '取消',
    cancelButtonClass: 'btn-info',
    closeIcon: false,
    content: "<div class='blockAlertMessage'>" + xmsg.replace(/(\n)+/g, '<br />') + "</div>",
    theme: 'white',
    confirm: confirmCallback,
    cancel: function () {
      if (cancelCallback) cancelCallback();
    }
  });
}

/*
顯示共用的 Dialog 對話框, 並由 contentUrl GET 載入對話框內容(Partial View)
arrBtn 可以傳入一個 javascript array 格式如下, 會自動在 dialog-footer 位置動態加入按鈕:
[
    { id: "btnID-1", name: "btnName-1", onclick: "clickHandlerName_1" },
    { id: "btnID-2", name: "btnName-2", onclick: "clickHandlerName_2" }
]
*/
function popupDialog(contentUrl, title, arrBtn, width, jsonParam) {
  var container = $("div#commonDialog");
  if ("commonDialog" != container.attr("id")) {
    blockAlert("找不到 commonDialog !");
    return;
  }
  if (title) {
    container.find(".modal-title").html(title);
  }
  if (width) {
    container.find(".modal-dialog").css("width", width + "px");
  }

  var footer = container.find(".modal-footer");
  /*取消關閉按鈕, 先抓出來 clone, 清除 footer 後再加回去, 以確保多次 popupDialog 時, arrBtn按鈕不會重覆*/
  var btnClose = footer.find(".btn-close").clone();
  btnClose.on("click", function () {
    container.modal('hide');
  });
  footer.html("");
  footer.append(btnClose);
  /*處理 btnArr 逐一建立並加到 footer 中*/
  if (Array.isArray(arrBtn)) {
    for (var i = 0; i < arrBtn.length; i++) {
      var btn = arrBtn[i];
      if (btn.id && btn.name && btn.onclick) {
        var a = document.createElement("a");

        //2017-05-23 群鈞修改 樣式改用boostrapt 的樣式
        //$(a).addClass("tablebtn").addClass("link-pointer"); //群鈞移除
        $(a).addClass("btn").addClass("btn-info");  //群鈞增加
        //=========================================

        $(a).attr("id", btn.id);
        $(a).html(btn.name);
        $(a).on("click", function () {
          var onclickHandler = window[btn.onclick];
          if (typeof onclickHandler === "function") {
            onclickHandler();
          }
          else {
            blockAlert("button '" + btn.name + "' 的 onclick handler '" + btn.onclick + "' 不是正確的 function.");
            return;
          }
        });

        footer.append(a);
      }
    }
  }

  if (contentUrl) {
    var inParam = (jsonParam) ? jsonParam : {};
    ajaxLoadMore(contentUrl, inParam, function (data) {
      container.find(".modal-body").html(data);
      container.modal('show');
      /*動態載入的 DateTimePicker 要執行 initDatePicker()*/
      container.find("div.date").each(function () {
        initDatePicker(this);
      });
    });
  }
  else {
    container.find(".modal-body").html("沒有指定 contentUrl");
    container.modal('show');
  }
}

//動態創建對話框（用來解決 popupDialog 方法只能顯示最後開啟的對話框，未限制內容不可超出視窗高度，無法接收回傳結果值問題）
//contentUrl: 內容頁網址。
//callback: 在對話框關閉時要執行的呼叫端自訂方法，並且傳入一個內容頁回傳值 JSON 物件給呼叫端自訂方法。
//jsonParam:（非必要）傳給內容頁的資料參數 JSON 物件。
//arrBtn:（非必要）可以傳入一個 javascript array 格式如下, 會自動在 dialog-footer 位置動態加入按鈕:
//[
//    { id: "btnID-1", name: "btnName-1", onclick: "clickHandlerName_1" },
//    { id: "btnID-2", name: "btnName-2", onclick: "clickHandlerName_2" }
//]
//width:（非必要）對話框固定寬度。預設值 800（單位 px）。
//height:（非必要）對話框固定寬度。預設值 510（單位 px）。
function popupDialog2(contentUrl, title, callback, jsonParam, arrBtn, width, height) {
  try {
    if (!title) title = "系統";
    if (!width) width = "800";
    if (!height) height = "510";

    var divId = "__tmpDialog" + Date.now().toString();
    var html = [
      "<div class=\"modal fade common-dialog\" id=\"", divId, "\" role=\"dialog\" aria-labelledby=\"commonDialogTitle\" aria-hidden=\"true\">",
      "    <div class=\"modal-dialog\" style=\"width:", width, "px;height:", height, "px;\">",
      "        <div class=\"modal-content\">",
      "            <div class=\"modal-header bg-primary\"><button type=\"button\" class=\"close\" data-dismiss=\"modal\"><span aria-hidden=\"true\">&times;</span><span class=\"sr-only\">Close</span></button><h3 class=\"modal-title\" id=\"commonDialogTitle\">", title, "</h3></div>",
      "            <div class=\"modal-body\" data-id=\"", divId, "\" style=\"overflow:auto;max-width:", width, "px;max-height:", height, "px;\"></div>",
      "            <div class=\"modal-footer center\"><a class=\"tablebtn link-pointer btn-close\">關閉取消</a></div>",
      "        </div>",
      "    </div>",
      "</div>"
    ];

    var container = $(html.join(""));
    var modalBody = container.find("div.modal-body");
    $(document.body).append(container);

    modalBody.on("close.modal-body", function (event, value) {
      modalBody.returnValue = value;
    })

    container.on("show.bs.modal", function (event) {
      modalBody.trigger("show.modal-body", jsonParam);
    });

    container.on("hidden.bs.modal", function (event) {
      try {
        if (callback) callback(modalBody.returnValue);
      }
      finally {
        modalBody.off("close.modal-body");
        container.off("hidden.bs.modal");
        container.remove();
      }
    });

    if (!arrBtn) {
      var footer = container.find(".modal-footer");
      /*取消關閉按鈕, 先抓出來 clone, 清除 footer 後再加回去, 以確保多次 popupDialog 時, arrBtn按鈕不會重覆*/
      var btnClose = footer.find(".btn-close").clone();
      btnClose.on("click", function () { container.modal("hide"); });
      footer.html("");
      footer.append(btnClose);

      /*處理 btnArr 逐一建立並加到 footer 中*/
      if (Array.isArray(arrBtn)) {
        for (var i = 0; i < arrBtn.length; i++) {
          var btn = arrBtn[i];
          if (btn.id && btn.name && btn.onclick) {
            var a = document.createElement("a");

            //2017-05-23 群鈞修改 樣式改用boostrapt 的樣式
            //$(a).addClass("tablebtn").addClass("link-pointer"); //群鈞移除
            $(a).addClass("btn").addClass("btn-info");  //群鈞增加
            //=========================================

            $(a).attr("id", btn.id);
            $(a).html(btn.name);
            $(a).on("click", function () {
              var onclickHandler = window[btn.onclick];
              if (typeof onclickHandler === "function") {
                onclickHandler();
              }
              else {
                blockAlert("button '" + btn.name + "' 的 onclick handler '" + btn.onclick + "' 不是正確的 function.");
                return;
              }
            });

            footer.append(a);
          }
        }
      }
    }

    if (contentUrl) {
      if (jsonParam === undefined || jsonParam === null) jsonParam = {};
      ajaxLoadMore(contentUrl, jsonParam, function (data) {
        container.find(".modal-body").html(data);
        container.modal("show");
        /*動態載入的 DateTimePicker 要執行 initDatePicker()*/
        container.find("div.date").each(function () {
          initDatePicker(this);
        });
      });
    }
    else {
      container.find(".modal-body").html("沒有指定 contentUrl");
      container.modal("show");
    }

    return divId;
  }
  catch (e) {
    blockAlert(e.message);
  }
}

function closeDialog(dialogId) {
  var boxId = (dialogId === undefined) ? "div#commonDialog" : dialogId;
  $(boxId).modal("hide");
}

/*
顯示指定 action path 的線上說明框, 若未指定則顯示當前頁面路徑的說明
*/
function popupHelp(helpPath) {
  if (!helpPath) {
    helpPath = window.location.pathname;
  }
  var form = $("form[name=OnlineHelpForm]");
  var action = form.attr("action");

  var parms = {};
  parms.helpPath = helpPath;

  ajaxLoadMore(action, parms, function (data) {
    var container = $("div.help-container");
    $(container).html(data);
    $(container).removeClass("hide");

    $(container).find(".help-header").find(".btn-close").on('click', function () {
      $(container).addClass("hide");
    })
  });
}

/*
* 將傳入的 form 處理成 顯示用(READONLY) 模式
* 將所有輸入元件加上 readonly, disabled 屬性, 並控制一些 CSS 樣式以達成效果
*/
function ReadonlyForm(form) {
  $(form).find(".open-datepicker").hide();
  $(form).find("input").attr("readonly", "readonly");
  $(form).find("input").attr("placeholder", "");
  $(form).find("textarea").attr("readonly", "readonly");
  $(form).find("select").attr("disabled", "disabled");
  $(form).find("select").each(function () {
    if (!$(this).val()) {
      $(this).hide();
    }
  });
  $(form).find("input[type=radio]").attr("disabled", "disabled");
  $(form).find("input[type=radio]").each(function () {
    if ($(this).attr("checked") != "checked") {
      $("label[for=" + $(this).attr("id") + "]").hide();
      $("div#radio_" + $(this).attr("id")).hide();
    }
    else {
      $("div#radio_" + $(this).attr("id")).removeClass("radio");
    }
  });
  $(form).find("input[type=checkbox]").attr("disabled", "disabled");
  $(form).find("input[type=checkbox]").each(function () {
    if ($(this).attr("checked") != "checked") {
      $("label[for=" + $(this).attr("id") + "]").hide();
      $("div#checkbox_" + $(this).attr("id")).hide();
    }
    else {
      $("div#checkbox_" + $(this).attr("id")).removeClass("checkbox");
    }
  });
  $(form).find("input[type=file]").each(function () {
    $("label[for=" + $(this).attr("id") + "]").hide();
  });
}

/*************************************************
函式名稱：CheckDate
輸入參數：DateString-String
傳回值:回傳true 或 false
函式描述：檢查DateString日期是否正確
  函式控管：IreneHsu　 2002/01/30　Version 1.0
 *************************************************/
function CheckDatefmt(DateString, chrFmt) {
  if (DateString.length > 10) return false;
  var y, m, d;
  var idx = DateString.indexOf(chrFmt)
  y = DateString.substring(0, idx)  //年
  DateString = DateString.substring(idx + 1, DateString.length)
  var idx = DateString.indexOf(chrFmt)
  m = DateString.substring(0, idx); //月
  d = DateString.substring(idx + 1, DateString.length);  //日
  if (m.substring(0, 1) == '0') m = m.substring(1, m.length);
  if (d.substring(0, 1) == '0') d = d.substring(1, d.length);
  /*alert("y="+y);
  alert("m="+m);
  alert("d="+d);*/
  var CharNum = "0123456789";
  //判別是否皆為數字
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

  y = parseInt(y);

  m = parseInt(m);
  d = parseInt(d);

  if (isNaN(y)) return false;
  if (isNaN(m)) return false;
  if (isNaN(d)) return false;
  y += 1911;

  if (y < 200 && y > 70) y += 1900;
  if (y < 70) y += 2000;

  //if (y > 2070 || y < 1970) return false;

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
        if (d > 29) return false;
        else return true;
      }
      if (d > 28) return false;
      return true;
    default:
      return false;
  }
}

function CheckFDate(DateString, strFmt) {
  return CheckDatefmt(DateString, strFmt);
}

function CheckDate(DateString) {
  return CheckDatefmt(DateString, "/");
}

/**************************************************
 函式名稱：CheckNumeric
輸入參數：String-String
功能：檢查字串是否全為數字
傳回值：若全為數字，則傳回true，否則傳回false
函式控管：IreneHsu　 2002/01/30　Version 1.0
 **************************************************/
function CheckNumeric(String) {
  var len = String.length;
  for (var i = 0; i < len; i++) {
    idx = String.substring(i, i + 1);
    idx = parseInt(idx);
    if (isNaN(idx)) return false;
  }
  return true;
}

/* 檢查指定email格式是否正確
 * @param   email       欲檢查的email
 * @return  boolean
 */
function checkEmail(email) {
  //var filter = /^.+@.+\..{2,3}$/;
  var filter = /^\w+[\w\.\-]*[\w\-]+@.+\..{2,3}$/;

  if (filter.test(email)) {
    return true;
  }
  else {
    return false;
  }
}

/* 在 tbody 內複製來源 tr（由 srcIdx 指定）成為新 tr 元素後加入到 tbody 尾端。
 * （標準寫法：copyTr("rows", "", 0, false, function(tr, trIdx) { });）
 * （完整範例：A7/C301M/Detail.cshtml、_StaffGridRows.cshtml）。
 * （複製時會同時處理 MVC 輸入元件 IList binding name(index) 的調整）。
 * 在 tbody 開始標籤內必須存在一個「data-count 屬性」用來標示出當前資料總筆數。
 * 在每一個 tr 標籤內必須存在一個「data-idx 屬性」用來標示出這個 tr 索引值（索引值起訖範圍：0 .. 當前資料總筆數-1）。
 * 在每一個 tr 內的任何 input 輸入欄位都必須要有 id 屬性。
 * @param tbodyId       tbody 元素的 id。
 * @param seqFieldId    資料主鍵欄位的 id，在新複製 Row 內的主鍵欄位值會一律設為 0。若沒有主鍵欄位時，請一律輸入 "" 字串。
 * @param srcIdx        做為複製對象的 tr 索引值（即每一個 tr 的 data-idx 屬性值）。若無特殊需求時，請一律輸入 0。       
 * @param copySrcValue  是否也要複製輸入欄位的欄位值。（true: 複製欄位值，false: 不要複製欄位值）。若無特殊需求時，請一律輸入 false。
 * @param callback      在產生新的 tr HTML 元素之後所要執行的呼叫端 callback 方法。
 *                      呼叫 callback 方法時系統會固定傳入兩個參數值，第一個是新 tr jQuery 物件，第二個是新 tr 索引值。
 */
function copyTr(tbodyId, seqFieldId, srcIdx, copySrcValue, callback) {
  var tbody = $("tbody#" + tbodyId);
  if (tbody.attr("id") != tbodyId) {
    blockAlert("找不到資料 tbody#" + tbodyId);
    return;
  }

  var count = tbody.attr("data-count");
  if (count == undefined) {
    blockAlert("指定的 tbody#" + tbodyId + " 沒有定義 data-count 屬性");
    return;
  }

  var srcTr = tbody.find("tr[data-idx=" + srcIdx + "]");
  if (srcTr.attr("data-idx") != srcIdx) {
    blockAlert("找不到第" + srcIdx + "筆資料 tr!");
    return;
  }

  // 設定後面替換會使用到的原索引值字串與新索引值字串
  var srcIdxStr = "[" + srcIdx + "]";
  var desIdxStr = "[" + count + "]";

  // 先 Create 新的 tr 但還不要加到 tbody 中
  // 以避免 radio, checkbox 這類元件, 會因為 Name 還跟第1筆一樣,
  // 使得畫面上第1筆的 checked 狀態會跑掉
  var thisTr = $(document.createElement('tr'));
  thisTr.attr("data-idx", count);
  thisTr.html(srcTr.html());

  // 清空 新增那一筆的 SEQNO hidden field
  thisTr.find("input[id$=" + seqFieldId + "]").val("0");

  // 複製並取代checkboxlist div造成的滯留問題
  thisTr.find('div.checkbox').each(function () {
    var me = $(this);
    //var tId = $(this).attr("id");
    var tId = me.attr("id");
    if (tId) {
      //$(this).attr("id", tId.replace("[" + srcIdx + "]", "[" + count + "]"));
      me.attr("id", tId.replace(srcIdxStr, desIdxStr));
    }
  })

  // 複製並取代checkboxlist label造成的連帶影響問題
  thisTr.find('div.checkbox label').each(function () {
    var labElm = $(this);
    var forId = labElm.attr("for");
    if (forId) {
      labElm.attr("for", forId.replace(srcIdxStr, desIdxStr));
    }
    // 替換 checkbox id 值
    var chbElm = labElm.find("input[type=\"checkbox\"]");
    if (chbElm.length > 0) {
      var chbId = chbElm.attr("id");
      if (chbId) chbElm.attr("id", chbId.replace(srcIdxStr, desIdxStr));
      // 預設不勾選 checkbox
      chbElm.prop("checked", false);
      chbElm.removeAttr("checked");
      iElm = labElm.find("i");
      if (iElm) iElm.removeClass("fa fa-check-square-o").addClass("fa fa-square-o");
      // 繫結 checkbox change 事件處理
      chbElm.on("change", function () {
        var me = $(this);
        var itag = me.next();
        if (me.is(":checked")) {
          itag.removeClass("fa fa-square-o").addClass("fa fa-check-square-o");
        }
        else {
          itag.removeClass("fa fa-check-square-o").addClass("fa fa-square-o");
        }
      });
    }
  })

  // 修改: 新增那一列的 input,select,textarea 
  // 等 element name 中的 list index 值
  thisTr.find("input,select,textarea").each(function () {
    var me = $(this);
    //var tName = $(this).attr("name");
    //$(this).attr("name", tName.replace("[" + srcIdx + "]", "[" + count + "]"));
    var tName = me.attr("name");
    me.attr("name", tName.replace(srcIdxStr, desIdxStr));

    //var tId = $(this).attr("id");
    var tId = me.attr("id");
    if (tId) {
      //$(this).attr("id", tId.replace("[" + srcIdx + "]", "[" + count + "]"));
      //$(this).attr("id", tId.replace("_" + srcIdx + "__", "_" + count + "__"));
      me.attr("id", tId.replace(srcIdxStr, desIdxStr));
      me.attr("id", tId.replace("_" + srcIdx + "__", "_" + count + "__"));
    }
    else {
      //tName = $(this).attr("name");
      //$(this).attr("id", tName.replace("[", "_").replace("]", "_").replace(".", "_"));
      tName = me.attr("name");
      me.attr("id", tName.replace("[", "_").replace("]", "_").replace(".", "_"));
    }

    if (!copySrcValue) {
      // 並清除所有非 hidden 欄位的輸入值
      // hidden 欄位值通常是扮演類似 FK 的作用, 一律保留

      //var type = $(this).attr("type");
      //if (type == "radio" || type == "checkbox") {
      //    $(this).removeAttr("checked");
      //} else if (type != "hidden") {
      //    $(this).val("");
      //}

      var type = me.attr("type");
      if (type == "radio" || type == "checkbox") {
        me.removeAttr("checked");
      } else if (type != "hidden") {
        me.val("");
      }
    }
  });

  // 若 tr 中存在 DatePicker 元件, 要 initDatePicker 才能正常運作 
  thisTr.find("div.date").each(function () {
    initDatePicker(this);
  });

  // 將新 tr 加到 tbody
  tbody.append(thisTr);

  //再次檢查
  var thisTrTest = tbody.find("tr[data-idx=" + count + "]");
  if (thisTrTest.attr("data-idx") != count) {
    blockAlert("找不到新增的第" + count + "筆資料 tr!");
    return;
  }

  //更新筆數
  tbody.attr("data-count", parseInt(count, 10) + 1);

  // 執行呼叫端 callback 方法。callback(newTrObj, newTrIndex)
  if (callback) callback(thisTr, count);
}

/*
 * 將原始室內電話號碼字串拆解成含有 "區碼、號碼、分機碼" 字串陣列（[0]:區碼、[1]:號碼、[2]:分機碼）。
 * @param text 原始室內電話號碼字串（格式：區碼-號碼#分機碼）
 */
function splitPhone(text) {
  var ret = ["", "", ""];
  text = $.trim(text);
  if (text.length > 0) {
    var i1 = text.indexOf("-");
    var i2 = text.indexOf("#");
    //區碼、號碼
    if (i1 > 0) {
      ret[0] = text.substr(0, i1);
      if (i2 < 0) {
        var c = text.length - i1 - 1;
        ret[1] = text.substr(i1 + 1, c);
      }
      else {
        var c = i2 - i1;
        if (c > 1) ret[1] = text.substr(i1 + 1, c - 1);
      }
    }
    //分機碼
    if (i2 > 0) {
      var c = text.length - i2;
      if (c > 1) ret[2] = text.substr(i2 + 1, c - 1);
    }
  }
  return ret;
}

//返回指定頁面（本方法提供在返回指定頁面時，也能傳入自訂資料給該頁面）。
//url: 要返回的頁面網址（例："@(Url.Action("Index", "C503M", new { Area = "A8" }))"）。
//jsonData: （非必填）要傳給該頁面的資料 JSON。例：{ "Form.AREANAME": "陽明大學", "Form.OOPCD": "2024" }。
//useCache: （非必填）指示是否需要取回暫存的 FormModel 內容。(0: 不需要，1: 需要（取回暫存的 FormModel 內容，但不用重新查詢資料），2.需要 (取回暫存的 FormModel 內容並且重新查詢資料)。預設為 1。
//httpMethod: （非必填）指定 HTTP 資料傳送方式。("POST": POST，"GET": GET)。預設為 "POST"。
function goBackPage(url, jsonData, useCache, httpMethod) {
  var form;
  try {
    if (useCache === undefined) useCache = "1";
    if (httpMethod === undefined) httpMethod = "POST";

    form = $("<form style=\"display:none;\"></form>")
      .attr("action", url).attr("method", httpMethod);

    var data = { "useCache": useCache };
    $.extend(data, jsonData);
    $.each(data, function (key, value) {
      $(["<input type=\"hidden\" name=\"", key, "\" value=\"", value, "\">"].join(""))
        .appendTo(form);
    });

    form.appendTo("body").submit();
  }
  finally {
    if (form) form.remove();
  }
}

//進入指定頁面（本方法提供在進入指定頁面時，也能傳入自訂資料給該頁面）。
//url: 要進入的頁面網址（例："@(Url.Action("Index", "C503M", new { Area = "A8" }))"）。
//jsonData: （非必填）要傳給該頁面的資料 JSON。例：{ "Form.AREANAME": "陽明大學", "Form.OOPCD": "2024" }。
//useCache: （非必填）指示是否需要取回暫存的 FormModel 內容。(0: 不需要，1: 需要（取回暫存的 FormModel 內容，但不用重新查詢資料），2.需要 (取回暫存的 FormModel 內容並且重新查詢資料)。預設為 1。
//httpMethod: （非必填）指定 HTTP 資料傳送方式。("POST": POST，"GET": GET)。預設為 "POST"。
function goForwardPage(url, jsonData, useCache, httpMethod) {
  var form;
  try {
    if (useCache === undefined) useCache = "1";
    if (httpMethod === undefined) httpMethod = "POST";

    form = $("<form style=\"display:none;\"></form>")
      .attr("action", url).attr("method", httpMethod);

    var data = { "useCache": useCache };
    $.extend(data, jsonData);
    $.each(data, function (key, value) {
      $(["<input type=\"hidden\" name=\"", key, "\" value=\"", value, "\">"].join(""))
        .appendTo(form);
    });

    form.appendTo("body").submit();
  }
  finally {
    if (form) form.remove();
  }
}

//設定 HTML 元素啟用或是禁用（本方法只針對具有 disabled 屬性的元素才會有作用）。
//tagSelector: HTML 元素的 CSS 選擇子表達式（例："#btnSure"）。
//boolState: 啟用或是禁用。（true: 啟用，false: 禁用）。
function setTagEnabled(tagSelector, boolState) {
  var m = $(tagSelector);
  m.each(function () {
    (boolState) ? $(this).removeAttr("disabled") : $(this).attr("disabled", "disabled");
  });
}

//設定 HTML 元素勾選或是不勾選（本方法只針對具有 checked 屬性的 checkbox、radiobox 元素才會有作用）。
//tagSelector: HTML 元素的 CSS 選擇子表達式（例："#btnSure"）。
//boolState: 勾選或是不勾選。（true: 勾選，false: 不勾選）。
//callback: 在每個 HTML 元素完成勾選狀態設定之後要執行的呼叫端自訂方法。eachCallback(obj, state) 其中 obj 傳入參數為當時設定的 HTML 元素的 Jquery 物件，state 為當時的勾選狀態設定（true: 勾選，false: 不勾選）。
function setTagChecked(tagSelector, boolState, callback) {
  var elm, m = $(tagSelector);
  m.each(function () {
    elm = $(this);
    (boolState) ? elm.attr("checked", "checked") : elm.removeAttr("checked");
    elm.prop("checked", boolState);
    var mvcHidTag = elm.next("input:hidden");
    if (mvcHidTag.length > 0) {
      if (elm.attr("id") == mvcHidTag.attr("id")) elm.trigger("onchange");
    }
    if (callback) callback(elm, boolState);
  });
}

//傳回 CheckBox 是否處於勾選狀態（本方法只針對具有 checked 屬性的 checkbox、radiobox 元素才會有作用）。若處於勾選狀態時傳回 true，否則傳回 false。
//tagSelector: HTML 元素的 CSS 選擇子表達式（例："input[type=\"radiobox\"]"）。
function isTagChecked(tagSelector) {
  var elm = $(tagSelector);
  return elm.is(":checked");
}

//傳回目前處於勾選狀態的 HTML 元素 Jquery 物件陣列（本方法只針對 checkbox 元素才會有作用）。
//fromTagSelector: 目標 HTML 元素所在的父層元素 CSS 選擇子表達式（例："input[id='tbody_rows']"）。不輸入時表示從 document 開始找起。
function getCheckedCheckBox(fromTagSelector) {
  var selector = "input[type=\"checkbox\"]:checked";
  if (fromTagSelector) {
    var p = $(fromTagSelector);
    return p.find(selector);
  }
  else {
    return $(selector);
  }
}

//傳回目前處於勾選狀態的 HTML 元素 Jquery 物件陣列（本方法只針對 radiobox 元素才會有作用）。
//fromTagSelector: 目標 HTML 元素所在的父層元素 CSS 選擇子表達式（例："input[id='tbody_rows']"）。不輸入時表示從 document 開始找起。
function getCheckedRadioBox(fromTagSelector) {
  var selector = "input[type=\"radiobox\"]:checked";
  if (fromTagSelector) {
    var p = $(fromTagSelector);
    return p.find(selector);
  }
  else {
    return $(selector);
  }
}

//取得 form 上的所有 input 欄位, 並轉成 object
function getFormDataToObject($form) {
  var unindexed_array = $form.serializeArray();
  var indexed_array = {};

  $.map(unindexed_array, function (n, i) {
    indexed_array[n["name"]] = n["value"];
  });

  return indexed_array;
}

//傳回在表單內可被提交到後端的資料輸入欄位「名稱:值」JSON 物件
function getFormSubmitableFields(formSelector) {
  var ret = {};
  var elms = $(formSelector).find(":input");
  if (elms.length > 0) {
    var elm, elmName;
    elms.each(function () {
      elm = $(this);
      elmName = elm.attr("name");
      if (elmName != undefined && elmName.length > 0) ret[elmName] = elm.val();
    });
  }
  return ret;
}

//取得並傳回 checkbox、radio 標籤項目文字的字串陣列
//itemOrItems: checkbox、radio 標籤項目「單一 jQuery 物件」或是「多個 jQuery 物件集合」。
function getCheckboxTextArray(itemOrItems) {
  try {
    var ret = [];
    if (itemOrItems && itemOrItems.length > 0) {
      var obj, elm, selector, txt;
      itemOrItems.each(function () {
        obj = $(this);
        if (selector === undefined) {
          selector = ["label[for^=\"", obj.attr("name").replace(".", "_"), "\"]"].join("");
        }
        elm = obj.closest(selector);
        txt = (elm.length > 0) ? elm.text() : obj.next("label").text();
        ret.push(txt);
      });
    }
    return ret;
  }
  catch (e) {
    alert("[getCheckboxTextArray] " + e.message);
  }
}

/* jQuery 功能擴充: 可綁定 input 欄位只接受 數字 以及特定功能鍵 的輸入 */
$.fn.numericOnly = function () {
  $(this).bind("input", function (event) {
    var val = this.value;

    if (val != "" && !CheckNumeric(val)) {
      $(this).val(val.substring(0, val.length - 1));
    }
  });
};