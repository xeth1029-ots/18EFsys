<%@ Page CodeBehind="zipcodeJob.aspx.vb" Language="vb" AutoEventWireup="false" Inherits="TIMS.zipcodeJob" %>
<HTML>
	<HEAD>
		<title>縣市鄉鎮分類表</title>
		<meta http-equiv="Content-Language" content="zh-tw">
		<meta http-equiv="Content-Type" content="text/html; charset=big5">
		<script language="JavaScript">
var city_id_field, city_name_field;
var zip_id_field, zip_name_field;
var all_id_field, all_name_field;

var zip_list = new Array;

function getParamValue(name) {
    var querystring;
    var values;
    var result = "";
    
    if (location.search.length>1) {
        querystring = unescape(location.search+"&");
        
        var re = new RegExp("[\?|\&]"+name+"=(.+)\&");
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

function formatCityRow(city_id, city_name) {
	var result;
	result = "<font color=\"#003399\" size=\"2\">【"+city_id+"】"+city_name+"</font>";
	return result;
}

function formatZipRow(city_id, city_name, zip_id, zip_name) {
	var result;
	result = "<font color=\"#003399\" size=\"2\">【"+zip_id+"】"+zip_name+"</font>";
	return result;
}

function city_onclick(city_id, city_name) {
    var result;
    result = "javascript:getZipTable('"+city_id+"')";
    return result;
}

function zip_onclick(city_id, city_name, zip_id, zip_name) {
    var result;
    result = "javascript:return_value('"+city_id+"','"+city_name+"','"+zip_id+"','"+zip_name+"')";
    return result;
}

function return_value(city_id, city_name, zip_id, zip_name) {
    //alert(city_id +":"+ city_name+":"+ zip_id+":"+ zip_name);
    
    if (city_id_field != "") {
        opener.document.getElementById(city_id_field).value = city_id;
    }
    if (city_name_field != "") {
        opener.document.getElementById(city_name_field).value = city_name;
    }
    if (zip_id_field != "") {
        opener.document.getElementById(zip_id_field).value = zip_id;
    }
    if (zip_name_field != "") {
        opener.document.getElementById(zip_name_field).value = zip_name;
    }
    if (all_id_field != "") {
        opener.document.getElementById(all_id_field).value = city_id+zip_id;
    }
    if (all_name_field != "") {
        var all_name = '('+zip_id+')'+city_name;
        if (city_name != zip_name) {
            all_name += zip_name;
        }
        
        opener.document.getElementById(all_name_field).value = all_name;
    }
    window.close();
}

function window_onload() {
    city_id_field = getParamValue("city_id_field");
    city_name_field = getParamValue("city_name_field");
    zip_id_field = getParamValue("zip_id_field");
    zip_name_field = getParamValue("zip_name_field");
    all_id_field = getParamValue("all_id_field");
    all_name_field = getParamValue("all_name_field");

	var city_id = document.getElementById("QueryCityId").value;
	if (city_id == "") {
		showCityTable(city_id);
	} else {
		showZipTable(city_id);
	}
}

function showCityTable(city_id) {
	var objTable = document.getElementById("CityTable");
	var objOddRow = objTable.rows(0);
	var objEvenRow = objTable.rows(1);
    var intCellCount = objTable.rows(0).cells.length;
	var objRow, objCell;
	var intRowCount = 0;
	var i = 0;

    hideZipTable();
    displayCityTable();
    
    while (objTable.rows.length > 2) {
        objTable.deleteRow(2);
    }
        
    for (var city_id in zip_list) {
		if (i%intCellCount==0) {
			objRow = objTable.insertRow();
			for (var j=0; j<intCellCount; j++) {
				objCell = objRow.insertCell();
			}
			intRowCount++;
			
			if (intRowCount%2==0) {
				objRow.bgColor = objEvenRow.bgColor;
			} else {
				objRow.bgColor = objOddRow.bgColor;
			}
		}
		
		objRow.cells(i%intCellCount).innerHTML = "<a href=\"" + city_onclick(city_id, zip_list[city_id]['city_name']) + "\">" +
		                                         formatCityRow(city_id,zip_list[city_id]['city_name']) +
		                                         "</a>";   
    
        i++;
    }
}

function showZipTable(city_id) {
    var objTable = document.getElementById("ZipTable");
	var objOddRow = objTable.rows(0);
	var objEvenRow = objTable.rows(1);
    var intCellCount = objTable.rows(0).cells.length;
	var objRow, objCell;
	var intRowCount = 0;
	var i = 0;

    hideCityTable();
    displayZipTable();
    
    document.getElementById("lblCityName").innerHTML = zip_list[city_id]['city_name'];
    
    while (objTable.rows.length > 2) {
        objTable.deleteRow(2);
    }
    
    for (var zip_id in zip_list[city_id]) {
        if (zip_id != "city_name") {
    		if (i%intCellCount==0) {
    			objRow = objTable.insertRow();
    			for (var j=0; j<intCellCount; j++) {
    				objCell = objRow.insertCell();
    				objCell.innerHTML = "&nbsp;";
    			}
    			intRowCount++;
    			
    			if (intRowCount%2==0) {
    				objRow.bgColor = objEvenRow.bgColor;
    			} else {
    				objRow.bgColor = objOddRow.bgColor;
    			}
    		}
    		
    		objRow.cells(i%intCellCount).innerHTML = "<a href=\"" + zip_onclick(city_id, zip_list[city_id]['city_name'], zip_id, zip_list[city_id][zip_id]) + "\">" +
    		                                         formatZipRow(city_id, zip_list[city_id]['city_name'], zip_id, zip_list[city_id][zip_id]) +
    		                                         "</a>";   
        
            i++;
        }
    }
}

function getZipTable(city_id, city_name) {
	document.getElementById("QueryCityId").value = city_id;
	document.forms[0].submit();
}

function displayCityTable() {
    var objTable = document.getElementById("CityTableBox");
    objTable.style.display = "inline";
}

function hideCityTable() {
    var objTable = document.getElementById("CityTableBox");
    objTable.style.display = "none";
}

function displayZipTable(city_id) {
    var objTable = document.getElementById("ZipTableBox");
    objTable.style.display = "inline";
}

function hideZipTable(city_id) {
    var objTable = document.getElementById("ZipTableBox");
    objTable.style.display = "none";
}
		</script>
	</HEAD>
	<body bgColor="#e6efff" onload="javascript:window_onload();">
		<form id="form1" method="post" runat="server">
			<asp:literal id="CityList" runat="server"></asp:literal><asp:literal id="ZipList" runat="server"></asp:literal><input id="QueryCityId" type="hidden" runat="server">
			<center>
				<table id="CityTableBox" borderColor="#0000cc" width="90%" border="2">
					<tr>
						<td>
							<table width="100%" align="center">
								<tr>
									<td align="center"><strong><font face="標楷體" color="#000066" size="4">縣 市 名 稱</font></strong></td>
								</tr>
								<tr>
									<td align="center"><font color="#ff0000" size="2">※請先按下縣市名稱後，再選擇下一層分類選單 ※</font></td>
								</tr>
							</table>
							<table id="CityTable" cellSpacing="2" cellPadding="6" width="100%" align="center" border="0">
								<tr style="DISPLAY: none" align="center" bgColor="#ccccff">
									<td width="33%" bgColor="#ccccff">
										<div align="left"><A onclick='open_win("city_2.asp?city_code=1&amp;city_name=台北市&amp;city_id=01","");return false'
												href="#"><font color="#003399" size="2">【01】台北市</font> </A>
										</div>
									</td>
									<td width="33%" bgColor="#ccccff">
										<div align="left"><A onclick='open_win("city_2.asp?city_code=2&amp;city_name=高雄市&amp;city_id=02","");return false'
												href="#"><font color="#003399" size="2">【02】高雄市</font> </A>
										</div>
									</td>
									<td width="34%" bgColor="#ccccff">
										<div align="left"><A onclick='open_win("city_2.asp?city_code=3&amp;city_name=基隆市&amp;city_id=11","");return false'
												href="#"><font color="#003399" size="2">【11】基隆市</font> </A>
										</div>
									</td>
								</tr>
								<tr style="DISPLAY: none" align="center" bgColor="#ddddff">
									<td width="33%" bgColor="#ddddff">
										<div align="left"><A onclick='open_win("city_2.asp?city_code=4&amp;city_name=新竹市&amp;city_id=12","");return false'
												href="#"><font color="#003399" size="2">【12】新竹市</font> </A>
										</div>
									</td>
									<td width="33%" bgColor="#ddddff">
										<div align="left"><A onclick='open_win("city_2.asp?city_code=5&amp;city_name=台中市&amp;city_id=13","");return false'
												href="#"><font color="#003399" size="2">【13】台中市</font> </A>
										</div>
									</td>
									<td width="34%" bgColor="#ddddff">
										<div align="left"><A onclick='open_win("city_2.asp?city_code=6&amp;city_name=嘉義市&amp;city_id=14","");return false'
												href="#"><font color="#003399" size="2">【14】嘉義市</font> </A>
										</div>
									</td>
								</tr>
							</table>
							<center>【<A title="關閉視窗" href="javascript:window.close();">關閉</A>】&nbsp; 【<A title="清除縣市鄉鎮" href="javascript:return_value('','','','');">清除</A>】
							</center>
						</td>
					</tr>
				</table>
				<table id="ZipTableBox" borderColor="#0000cc" width="98%" border="2">
					<tr>
						<td>
							<table width="100%" align="center">
								<tr>
									<td align="center"><strong><font face="標楷體" color="#000066" size="4">鄉 鎮 名 稱--<span id="lblCityName">宜蘭縣</span>
											</font></strong>
									</td>
								</tr>
								<tr>
									<td align="center"><font color="#ff0000" size="2">※請選擇鄉鎮區名稱※</font></td>
								</tr>
							</table>
							<table id="ZipTable" cellSpacing="2" cellPadding="6" width="100%" align="center" border="0">
								<tr style="DISPLAY: none" align="center" bgColor="#ccccff">
									<td width="33%">
										<div align="left"><A href="javascript:send('32032','宜蘭縣不限(032)')"><font color="#003399" size="2">【032】不限</font>
											</A>
										</div>
									</td>
									<td width="33%">
										<div align="left"><A href="javascript:send('32260','宜蘭縣宜蘭市(260)')"><font color="#003399" size="2">【260】宜蘭市</font>
											</A>
										</div>
									</td>
									<td width="34%">
										<div align="left"><A href="javascript:send('32261','宜蘭縣頭城鎮(261)')"><font color="#003399" size="2">【261】頭城鎮</font>
											</A>
										</div>
									</td>
								</tr>
								<tr style="DISPLAY: none" align="center" bgColor="#ddddff">
									<td width="33%">
										<div align="left"><A href="javascript:send('32262','宜蘭縣礁溪鄉(262)')"><font color="#003399" size="2">【262】礁溪鄉</font>
											</A>
										</div>
									</td>
									<td width="33%">
										<div align="left"><A href="javascript:send('32263','宜蘭縣壯圍鄉(263)')"><font color="#003399" size="2">【263】壯圍鄉</font>
											</A>
										</div>
									</td>
									<td width="34%">
										<div align="left"><A href="javascript:send('32264','宜蘭縣員山鄉(264)')"><font color="#003399" size="2">【264】員山鄉</font>
											</A>
										</div>
									</td>
								</tr>
							</table>
							<center>【<A title="關閉視窗" href="javascript:window.close();">關閉</A>】&nbsp; 【<A title="清除縣市鄉鎮" href="javascript:return_value('','','','');">清除</A>】&nbsp;【<A title="返回縣市名稱清單" href="javascript:showCityTable('');">回上頁</A>】
							</center>
						</td>
					</tr>
				</table>
			</center>
		</form>
	</body>
</HTML>
