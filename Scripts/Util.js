//function chkSize(str, size)=檢查字符串長度  
//已通過測試  
function chkSize(str, size)
{
	if(str == null)	
		return false;
	else
	{
		
		if( str.toString().length == parseInt(size, 10) )
		 	return true;
		else
			return false;
	}				
}

//function chkInt(str)=檢查是否為整數
//已通過測試
function chkInt(str)
{	
	if( (str != null) && (str != "") )
	{
		var Letters = "0123456789";
		for (var i=0; i< str.length; i++){
  			var CheckChar = str.charAt(i);
   			if (Letters.indexOf(CheckChar) == -1){
     			return false;
    		}
    	}
   	 	return true;
  	}
  	else
  		return false;
}


//function chkNumber(str)=檢查是否為數字
//已測試通過
function chkNumber(str)
{	
	if( (str != null) && (str != "") )
	{				
  		var strNumber = str.toString();
  		var Letters = "0123456789.";
  		
  		for (var j=0; j< strNumber.length; j++)
  		{
    		var CheckChar = strNumber.charAt(j);
     		if (Letters.indexOf(CheckChar) == -1)
     		{
       			return false;
      		}      		
      	}
      	return true;
    }
    else 
    	return false;	
}

//前面補0//number為要添加的0的個數
//已通過測試
function addZeroF(str, size)
{
	var strNo = str.toString();
	if(strNo.length < parseInt(size, 10))
	{
		var strZero = "";
		
		for(var i = 0; i<(parseInt(size, 10) - strNo.length); i++)
		{
			strZero += "0";
		}
		strNo = strZero + strNo;
		return strNo;		
	}
	else if( strNo.length == parseInt(size, 10) )
		return strNo;
	else
		
		return strNo;
}

//轉換為大寫
//已通過測試
function convertToUpp(str)
{
	var string = str.toUpperCase();
	return string;
}

//轉換為小寫
//已通過測試
function convertToLow(str)
{
	var string  = str.toLowerCase();
	return string;
}

//=============================
//檢核欄位是否numeric（含小數點） 
//Check whether the data type of column is numeric (include decimal point)
var LegalNumeric = "0123456789.,-"
function isNumeric(s)
{
    var i;
    var j=0;
    var c;
    if (isempty(s)) return false;
    for (i = 0; i < s.length; i++)
    {
     c = s.charAt(i);
     if (LegalNumeric.indexOf(c) == -1) return false;
     if (c == '.') j++;
     if (j>1) return false;
     if (c=='-' && i!=0) return false;	 //檢核為負數時 If the value is a negative number
    }
    return true;
}

//==================
//檢核欄位是否empty //
//Check whether the field is empty
function isEmpty(control)
{
	var s = control.value;
	var result = false;
	if (s == null){result = true}
	if ((s.length==0) || (s=="")) result = true;
	var s1=0;
	var s2=0;
	for (var i=0 ; i < s.length ; i++)
	{
		if(s.charAt(i)==" ") {s1++;}
		else {s2++;}
	}
	if (s2>0) result = false
	else result = true;
	
	if(result) ChangeControlStyle(control,true);
	
	return result;
}

//for string  of trim.--------------
//去左空白
function ltrim ( s ){
	return s.replace( /^\s*/, "" )
}
//去右空白
function rtrim ( s ){
	return s.replace( /\s*$/, "" );
}
//去空白
function trim ( s ){
	return rtrim(ltrim(s));
}
var m_seperatorCode = "⊿";


// 檢查是否為合法西元日期 {yyyy-MM-dd || yyyy/MM/dd}
function isDate(dateval){
    var arr = new Array();
    
    if(dateval.indexOf("-") != -1){
        arr = dateval.toString().split("-");
    }else if(dateval.indexOf("/") != -1){
        arr = dateval.toString().split("/");
    }else{
        return false;
    }
    
    //yyyy-mm-dd || yyyy/mm/dd
    if(arr[0].length==4){
        var date = new Date(arr[0],arr[1]-1,arr[2]);
        if(date.getFullYear()==arr[0] && date.getMonth()==arr[1]-1 && date.getDate()==arr[2]){
            return true;
        } 
    }
    //dd-mm-yyyy || dd/mm/yyyy
    if(arr[2].length==4){
        var date = new Date(arr[2],arr[1]-1,arr[0]);
        if(date.getFullYear()==arr[2] && date.getMonth()==arr[1]-1 && date.getDate()==arr[0]){
            return true;
        }
    }
    //mm-dd-yyyy || mm/dd/yyyy
    if(arr[2].length==4){
        var date = new Date(arr[2],arr[0]-1,arr[1]);
        if(date.getFullYear()==arr[2] && date.getMonth()==arr[0]-1 && date.getDate()==arr[1]){
            return true;
        }
    } 
    
    return false;
} 




//validate date format
//檢查日期格式符合規範(YYYY[0,curYear]MM[1,12]DD[1, 31];YYYY[0,curYear]/MM[1,12]/DD[1, 31])
function chkDate(fmtDate){

	if( (fmtDate == null) || (fmtDate == "") ){
	  	return false;
  	}else{
  		fmtDate = fmtDate.toString();
		var re;
		re = /\//g;
		fmtDate = fmtDate.replace(re, "");
			
		var bValid = chkInt(fmtDate);
		
		if(bValid == false){
			return false;
		}else{
			var curDate, curMonth, curYear, curDay, nYear;
			var str, numYear, numMonth, numDate
			
			if( (fmtDate.length != 8))	{		
				return false;		
			}else{								
				
				numYear = parseInt(fmtDate.substr(0, 4), 10);

				numMonth = parseInt(fmtDate.substr(4, 2), 10);

				numDate = parseInt(fmtDate.substr(6, 2), 10);

				if( numYear<1 || numMonth<1 || numMonth>12
					|| numDate<1 || numDate>31)	{
					return false;
				}else{
					
					if((numMonth==2) && (numDate==29)){
						if( (((numYear % 4) == 0) && 
							((numYear % 100) !==0)) || ((numYear % 400) == 0) )	{
							return true;
						}else{
							return false;
                  		}
					}

					if ((numMonth<=7) && ((numMonth % 2)==0) && (numDate>=31)){
					  return false;
					}
					if ((numMonth>=8) && ((numMonth % 2)==1) && (numDate>=31)){
						return false;
					}
					if ((numMonth==2) && (numDate==30))	{
					   return false;
					}
				}
				return true;
			}
		}
	}
}


function chtDateToAD( strDate ) {
	var yr, mon, day, dateAD;
	
	if( !chkInt(strDate) ) {
		return "";
	}
	else if(strDate.length == 7) {		// yyyMMdd
		yr = strDate.substr(0, 3);
		mon = strDate.substr(3, 2);
		day = strDate.substr(5, 2);
	} 
	else if(strDate.length == 6) {		// yyMMdd
		yr = strDate.substr(0, 2);
		mon = strDate.substr(2, 2);
		day = strDate.substr(4, 2);
	} 
	else {
		return "";
	}
	
	dateAD = "" + (1911 + parseInt(yr)) + mon + day;
	return dateAD;
}

function formatDate(strDate){
	var strResult, numYear, numMonth, numDate
	
	if (strDate.indexOf("/") > 0){
		return strDate;
	}else{
		numYear = strDate.substr(0, 4);
		numMonth = strDate.substr(4, 2);
		numDate = strDate.substr(6, 2);
		
		strResult = numYear + "/" + numMonth + "/" + numDate;
		return strResult;
	}
}

function dateIsAfter(date1,date2) {
	var d1 = date1.replace("/") * 1;
	var d2 = date2.replace("/") * 1;
	

	if (new Date(date1) > new Date(date2)) {
		return true;
	} else {
		return false;
	}

}

function Trim(str)
{
	//window.alert(GetFirstPos(str));
	
	//window.alert(GetLastPos(str));
	
	return str.substring(GetFirstPos(str),GetLastPos(str))

}

function GetFirstPos(str)
{
	var iStart=0;
	
	for(var i=0;i<str.length;i++)
	{
		
		if(str.substring(i,i+1)!=" ")
		{
			break;
		}
		else
		{
			//window.alert("aa");
			iStart++;
		}
	}
	
	return iStart;
}

function GetLastPos(str)
{
	var iEnd=str.length;
	
	for(var i=str.length-1;i>0;i--)
	{
		if(str.substring(i,i+1)!=" ")
		{
			return iEnd;
		}
		else
		{
			iEnd=i;
		}
	}
	
	return iEnd;
	
}
function AddByHour(dtOrigin, ihour){
	var dtTarget = new Date (dtOrigin);
	
	var MinMilli = 1000 * 60;
	var HrMilli = MinMilli * 60;


	if ( dtTarget == "NaN") {
		return dtOrigin;
	}else{
		var iTarget = dtTarget.getTime();
		iTarget += ihour * HrMilli;
		dtTarget = new Date (iTarget);
		return dtTarget;
	}
}
function getTimeZone(){
	var dtToday = new Date();
	var tz = dtToday.getTimezoneOffset();
	tz = tz / 60;
	return tz;
}



function openWin(sUrl, formId, width, height, top, left) {
	var putItHere = null; 
  	putItHere = window.open(sUrl,formId,"width=" + width + ",height=" + height + ",top=" + top + ",left=" + left);
}

function openWinScroll(sUrl, formId, width, height, top, left) {
	var putItHere = null; 
  	putItHere = window.open(sUrl,formId,"scrollbars=yes,width=" + width + ",height=" + height + ",top=" + top + ",left=" + left);
    putItHere.focus();
}


function openHiddenWin(sUrl, formId) {
	var putItHere = null; 
  	putItHere = window.open(sUrl,formId,'width=0,height=0,top=-100,left=-100');
}

function isIP(obj) {
		if (isEmpty(obj)) 
			return false;
		
		var strIP = obj.value
		var re=/^(\d+)\.(\d+)\.(\d+)\.(\d+)$/g 
		if(re.test(strIP))
		{
			if( RegExp.$1 <256 && RegExp.$2<256 && RegExp.$3<256 && RegExp.$4<256) 
				return true;
		}
		return false;
}

function mousemove(obj)
{
	obj.className ="mouseMove";
}

function mouseout(obj,className)
{
	obj.className =className;
}

function CheckStrIsCaret(control)
{
	var str = control.value;
	var charStr = "~!@#$%^&*()_-+={[]}|\\:;\"\'/.,<>?`";
	
	for (i = 0; i < str.length; i++)
	{ 
		if ( charStr.indexOf(str.charAt(i)) > -1 )
		{
		   ChangeControlStyle(control,true);
		   return false; 
		}
	}
	
	return true;
}

function checkEnglishAndNumber(control)
{
	var txt= control.value
	re = /\W/;
	if (re.test(txt))
	{
		ChangeControlStyle(control,true);
	    return false;
	}
	else
	{
	   return true;
	}  

}

function ReplaceChar(string,character){
	var str = string;
	var strchar = character;
	var restr = "";
	for(var i=0; i< str.length ;i++)
	{
		for (var j=0; j < strchar.length; j++)
		
		if (str.charCodeAt(i) != strchar.charCodeAt(j))
		{
			restr += String.fromCharCode(str.charCodeAt(i))
		}
	}
	return restr;
}

function CheckTextBox(control, minLength, msxLength)
{
	var str = control.value;
	if (ReplaceChar(str," ").length >= minLength)
	{
		var strLength = StringLength(str);

		if (strLength < minLength || strLength > msxLength)
		{
			ChangeControlStyle(control,true);
			return false;
		}
		else return true;
	}
	else
	{
		return false;
	}
	return true;
}

function StringLength(sValue)
{	
	var str = sValue;
	
	var iLen = 0;
	
	if (str != "")
	{
		for (var i = 0; i < str.length; i++ )
		{					
			if (sValue.charCodeAt(i) >= 33 && sValue.charCodeAt(i) <= 126) iLen ++;
			else iLen += 2;
		}
	}
	return iLen;
}

function ChangeControlStyle(control,isError)
{
   if(isError) control.style.backgroundColor = 'ffccff';
   else control.style.backgroundColor = '';
}

function ChangeALLStyleToDefault() {
	var myInput = document.getElementsByTagName('INPUT');
	for (var i = 0; i < myInput.length; i++) {
		if (myInput[i].type.toLowerCase() == 'text') ChangeControlStyle(myInput[i], false);
		else if (myInput[i].type.toLowerCase() == 'radio') myInput[i].style.backgroundColor='';
	}
	var mySelect = document.getElementsByTagName('SELECT');
	
	for (var i = 0; i < mySelect.length; i++) ChangeControlStyle(mySelect[i], false);
	
	if(document.getElementById("lblMsg") != null) document.getElementById("lblMsg").innerHTML = "";
}


