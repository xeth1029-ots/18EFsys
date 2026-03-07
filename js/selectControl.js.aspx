<%@ Page Language="C#" %>
var _DEBUG_ = false;

/* ----------------------------------------------------------------------
*         JavaScript  Function  selectControl                         
*                                                                                         
*  透過Ajax呼叫 /REST/SelOptions/list?stmId=stmId       
*  取得下拉選項清單, 並填入由select_obj_id指定的   
*  <select>標籤下, 須要 jQuery API                                 
*-----------------------------------------------------------------------
* stmId                 SqlMap 中配置的 statement ID編號                     
* select_obj_id    要填充的<select>標籤ID                                
* textField            在資料清單中作為Option TEXT 顯示的欄位名稱
* valueField          在資料清單中作為Option VALUE 的欄位名稱
* hintOption         (此參數可不傳)  第一個下拉選項提示字串 
*                           如:   請選擇
* selectedValue    (此參數可不傳)     預設選擇的項目 value               
* params               (此參數可不傳)   要傳遞的參數，採用數組的形式傳遞   
*                           如： [["orgID","101"],["deptCode","62"]]   
* onOptLoaded     callback function, 當所有 options 產生完畢之後, 若這個參數存在, 會呼叫這個 function, 
*                            並傳入3個參數: select_obj_id, selectedValue, AJAX 回應的 data object
*-----------------------------------------------------------------------*
* 範例:   selectControl('ajaxTPlanList', 'PlanID', 'PlanName', 'TPlanID', '請選擇', '02');	  															
---------------------------------------------------------------------*/
function selectControl(stmId, select_obj_id, textField, valueField, hintOption, selectedValue, params, onOptLoaded) {
    var selObj = document.getElementById(select_obj_id);

    // 一律先清空下拉項目
    selObj.length = 0;
    if (hintOption) {
        var oOption = document.createElement("OPTION");
        oOption.text = hintOption;
        oOption.value = "";
        selObj.add(oOption);
    }

    var sUrlBase = 'http://<%= Request.Url.Authority + Request.ApplicationPath %>';
    if(!sUrlBase.match("/$") ) {
        sUrlBase += "/";
    }
    var sUrl = sUrlBase + "REST/SelOptions/list?stmId=" + stmId;

	jQuery(eval(params)).each(function(index, item) {
	    sUrl += "&" + encodeURIComponent(item[0]) + "=" + encodeURIComponent(item[1]);
	});

	if (_DEBUG_) {
	    alert("selectControl:\n " + sUrl );
	}

    var idx = 0;
	jQuery.getJSON(
                sUrl,
                {},
	    	    function (result) {
	    	        //alert("AA");
	    	        if (result.status == '1') {
	    	            var data = eval("result." + stmId + ";");
	    	            jQuery(data).each(function (index, item) {
                            idx ++;
	    	                var oOption = document.createElement("OPTION");
	    	                oOption.text = eval("item." + textField);
	    	                oOption.value = eval("item." + valueField);
	    	                selObj.add(oOption);
                            if(selectedValue == oOption.value) 
                            {
                                selObj.selectedIndex = idx;
                            }
	    	            });
                        if(typeof onOptLoaded == "function") {
                            onOptLoaded(select_obj_id, selectedValue, data);
                        }
	    	        }
	    	        else {
	    	            alert("selectControl error: " + result.message);
	    	        }
	    	    });
    }
