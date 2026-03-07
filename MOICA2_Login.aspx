<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MOICA2_Login.aspx.vb" Inherits="WDAIIP.MOICA2_Login" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
	<title>職業訓練資訊管理系統</title>
	<link rel="Shortcut Icon" href="./css/wdalogo.ico" type="image/x-icon" />
	<meta http-equiv="X-UA-Compatible" content="IE=EDGE" charset="utf-8" />
	<meta name="google-site-verification" content="1vPcd6IrtN7KTwB1nKCGSn16VWUsYEy6Z1-dMEDeaos" />
	<link href="css/css.css" rel="stylesheet" type="text/css" />
	<link href="css/style.css" rel="stylesheet" type="text/css" />
	<style type="text/css">
		html { display: none; }
		BODY { background-color: #ffffff; }
	</style>
	<script src="js/jquery-1.6.2.js" type="text/javascript"></script>
	<script type="text/javascript" src="js/errorcode.js"></script>
	<script language="javascript" type="text/javascript">
		//alert(top.location);alert(self.location);
		if (self == top) { document.documentElement.style.display = 'block'; }
		else { top.location = self.location; }
		if (parent.document.frames != undefined && parent.document.frames.length != 0) {
			top.location.replace(self.location);
		}
		//alert("openhttps:" + openhttps.value);
		var urlNG1 = "vm-tims";
		var url = window.location.href.toLowerCase(); //alert(url);
		if (url.indexOf("https:") == -1
			&& url.indexOf("localhost") == -1
			&& url.indexOf(urlNG1) == -1) {
			var Usehttps = false;
			var openhttps = document.getElementById("openhttps");
			if (openhttps) {
				if (openhttps.value != "0") { Usehttps = true; }
			}
			else {
				Usehttps = true;
			}
			if (Usehttps) {
				url = url.replace("http:", "https:");
				window.location.replace(url);
			}
		}
	</script>
	<% 
		Server.ScriptTimeout = 50
	%>
	<script language="javascript" type="text/javascript">
		/*
		var gflag_json2 = false;
		if (typeof (JSON) == 'undefined') {
		//alert("瀏覽器不支援JSON!!"); return null;
		//如果瀏覽器不支援JSON則載入json2.js  
		$.getScript('js/json2.js');
		gflag_json2 = true;
		}
		*/

		//依據不同的瀏覽器，取得 XMLHttpRequest 物件
		/*
		function xmlhttpObj2() {
		if (window.ActiveXObject) {
		try {
		return new ActiveXObject("Msxml2.XMLHTTP");
		} catch (e) {
		try {
		return new ActiveXObject("Microsoft.XMLHTTP");
		} catch (e2) {
		return null;
		}
		}
		} else if (window.XMLHttpRequest) {
		return new XMLHttpRequest();
		} else {
		return null;
		}
		}
		*/

		var postTarget;
		var timeoutId;

		function postData(target, data) {
			//var http = document.getElementById("httpObject");
			//var http = new XMLHttpRequest();
			/*
			var http = xmlhttpObj2();
			http.open('POST', target, false);
			http.send(data);
			return http.responseText;
			*/
			http = document.getElementById("http");
			http.url = target;
			http.actionMethod = "POST";
			var code = http.sendRequest(data);
			if (code != 0) return null;
			return http.responseText;
		}

		function checkFinish() {
			if (postTarget) {
				postTarget.close();
				alert("尚未安裝元件");
				return false;
			}
		}

		//簽章
		function makeSignature() {
			//var httpObject = document.getElementById("httpObject"); //var type1 = "Firefox";
			var ua = window.navigator.userAgent;
			if (ua.indexOf("MSIE") != -1 || ua.indexOf("Trident") != -1) {
				//type1 = "is IE, use ActiveX" //is IE, use ActiveX
				var tbsPackage = getTbsPackage();
				//httpObject.innerHTML = '<OBJECT id="http" width=1 height=1 style="LEFT: 1px; TOP: 1px" type="application/x-httpcomponent" VIEWASTEXT></OBJECT>';
				document.getElementById("httpObject").innerHTML = '<OBJECT id="http" width=1 height=1 style="LEFT: 1px; TOP: 1px" type="application/x-httpcomponent" VIEWASTEXT></OBJECT>';
				var data = postData("http://localhost:61161/sign", "tbsPackage=" + tbsPackage);
				if (!data) alert("尚未安裝元件");
				else setSignature(data);
			}
			else {
				//debugger;
				postTarget = window.open("http://localhost:61161/popupForm", "Signing", "height=200, width=200, left=100, top=20");
				timeoutId = setTimeout(checkFinish, 3500);
			}
			//alert(type1);
			return true;
		}

		//組合JSON
		function getTbsPackage() {
			var txtpass = document.getElementById("txtpass");
			var nonce = document.getElementById("nonce"); // 'string
			var tbsData = {};
			tbsData["tbs"] = "TBS"; //document.getElementsByName("tbs")[0].value;
			tbsData["hashAlgorithm"] = "SHA256"; //document.getElementsByName("hashAlgorithm")[0].value;
			tbsData["withCardSN"] = "false"; //document.getElementsByName("withCardSN")[0].value;
			tbsData["pin"] = txtpass.value; //document.getElementsByName("pin")[0].value;
			tbsData["nonce"] = nonce.value; //document.getElementsByName("nonce")[0].value;
			tbsData["func"] = "MakeSignature";
			tbsData["signatureType"] = "PKCS7";
			var json = JSON.stringify(tbsData);
			return json;
		}

		function setSignature(signature) {
			var ret = JSON.parse(signature);
			var credential = document.getElementById("credential");
			var returnCode = document.getElementById("returnCode");
			var btnSubmit2 = document.getElementById("btnSubmit2");
			credential.value = ret.signature;
			returnCode.value = ret.ret_code;
			//debugger;
			//if (credential.value == "") {alert("憑證資訊不可為空!!"); return false;}
			if (ret.ret_code != 0) {
				alert(MajorErrorReason(ret.ret_code));
				return false;
			}
			btnSubmit2.click();
			return true;
		}

		function receiveMessage(event) {
			if (console) console.debug(event);
			//安全起見，這邊應填入網站位址檢查
			if (event.origin != "http://localhost:61161")
				return;
			try {
				//debugger;
				var ret = JSON.parse(event.data);
				if (ret.func) {
					if (ret.func == "getTbs") {
						clearTimeout(timeoutId);
						var json = getTbsPackage()
						postTarget.postMessage(json, "*");
					} else if (ret.func == "sign") {
						setSignature(event.data);
					}
				} else {
					if (console) console.error("no func");
				}
			} catch (e) {
				//errorhandle
				if (console) console.error(e);
			}
		}

		if (window.addEventListener) {
			window.addEventListener("message", receiveMessage, false);
		} else {
			//for IE8
			window.attachEvent("onmessage", receiveMessage);
		}
		//for IE8
		var console = console || { "log": function () { }, "debug": function () { }, "error": function () { } };
	</script>
	<script language="javascript" type="text/javascript">
		function CheckCard() {
			//var Credential = null; // 'object
			//var rcode = null; // 'int
			//var TBSData = null; // 'string
			//var nonce = document.getElementById("nonceid"); // 'string
			var msg = "";
			var txtname = document.getElementById("txtname");
			var txtpass = document.getElementById("txtpass");
			txtname.value = txtname.value.toUpperCase();

			if (txtname.value == "") {
				msg += "請輸入正確的個人身分證號，不可為空。\n";
			}
			if (txtpass.value == "") {
				msg += "請輸入正確的PIN 密碼，不可為空。\n";
			}
			if (!checkId(txtname.value)) {
				msg += "請輸入正確的個人身分證號格式。\n";
			}
			if (msg != "") {
				alert(msg);
				return false;
			}
			//return makeSignature(); //簽章
			makeSignature(); //簽章 
			return false; //簽章
		}

		/* 檢查輸入的身份證字號是否正確
		* @param   IDString	欲檢查的身份證字號
		* @return  boolean */
		function checkId(IDString) {
			var ErrString = "";
			var ID1 = IDString.toUpperCase();
			if (IDString.length != 0) { IDString = IDString.toUpperCase() }
			if (IDString.length != 10) { ErrString = ErrString + "身分證字號字數不對。" + unescape('%0D') }
			if (ID1.length != 10) return false; //alert("身分證字號字數不對 !");
			var IDdigit = new Array(10);
			for (var i = 0; i < 10; i++) { IDdigit[i] = ID1.charAt(i); }
			var CharEng = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
			IDdigit[0] = CharEng.indexOf(IDdigit[0]);
			if (IDdigit[0] == -1) return false; //alert("身分證字號第一位為錯誤英文字母 !");
			if (IDdigit[1] != 1 && IDdigit[1] != 2) return false; //alert("身分證字號無法辨識性別 !");

			var Array1 = new Array(26);
			Array1[0] = 1; Array1[1] = 10; Array1[2] = 19;
			Array1[3] = 28; Array1[4] = 37; Array1[5] = 46;
			Array1[6] = 55; Array1[7] = 64; Array1[8] = 39;
			Array1[9] = 73; Array1[10] = 82; Array1[11] = 2;
			Array1[12] = 11; Array1[13] = 20; Array1[14] = 48;
			Array1[15] = 29; Array1[16] = 38; Array1[17] = 47;
			Array1[18] = 56; Array1[19] = 65; Array1[20] = 74;
			Array1[21] = 83; Array1[22] = 21; Array1[23] = 3;
			Array1[24] = 12; Array1[25] = 30;
			var result = Array1[IDdigit[0]];
			for (var i = 1; i < 10; i++) {
				var Number = "0123456789";
				IDdigit[i] = Number.indexOf(IDdigit[i]);
				if (IDdigit[i] == -1) {
					//alert("身分證字號錯誤 !");
					return false;
				} else {
					result += IDdigit[i] * (9 - i);
				}
			}
			result += 1 * IDdigit[9];
			//alert("result=="+result);
			if (result % 10 != 0) {
				//alert("身分證字號錯誤 !");
				return false;
			}
			else {
				return true;
			}
		}
	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<div style="display: none">
		<asp:Button ID="btnSubmit2" runat="server" Text="Button" />
		<span id="httpObject" runat="server"></span>
		<input id="nonce" type="hidden" runat="server" name="nonce" autocomplete="off" />
		<input id="openhttps" type="hidden" name="openhttps" runat="server" autocomplete="off" />
		<input id="credential" type="hidden" name="credential" runat="server" autocomplete="off">
		<input id="returnCode" type="hidden" name="returnCode" runat="server" autocomplete="off">
	</div>
	<table id="tbMoica1" runat="server" border="0" cellspacing="0" cellpadding="0" width="1000" style="position: absolute; left: 0px; top: 0px; margin-left: 0px; margin-top: 0px;">
		<tr>
			<td colspan="5">
				<img alt="" src="images/i2/login_01.jpg" width="1000" height="242">
			</td>
		</tr>
		<tr>
			<td colspan="5" align="center">
				<br style="line-height: 20px">
				<font style="color: #24498f; font-weight: bold">自然人憑證登入</font>
				<table cellspacing="3" cellpadding="1">
					<tr>
						<td style="z-index: 0" align="right">身 分 證 號： </td>
						<td>&nbsp;
							<asp:TextBox Style="z-index: 0" ID="txtname" runat="server" AutoCompleteType="Disabled" MaxLength="11"></asp:TextBox>
						</td>
					</tr>
					<tr>
						<td align="right">PIN 碼： </td>
						<td>&nbsp;
							<asp:TextBox Style="z-index: 0" ID="txtpass" runat="server" TextMode="Password" AutoCompleteType="Disabled" MaxLength="33"></asp:TextBox>
						</td>
					</tr>
					<tr>
						<td colspan="2" align="center"><a id="A3" title="HiPKI Local Server 瀏覽器元件安裝" class="l" href="#" onclick="window.open('https://gpkiapi.nat.gov.tw/PKCS7Verify/','_self');">瀏覽器元件安裝</a>
							<asp:ImageButton Style="z-index: 0" ID="bt_submit" runat="server" ImageUrl="images/i2/button/login_send.bmp"></asp:ImageButton><font face="新細明體">&nbsp;</font>
							<asp:ImageButton Style="z-index: 0" ID="bt_reset" runat="server" ImageUrl="images/i2/button/login_rest.bmp"></asp:ImageButton>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td rowspan="2">
				<%--style="padding: 0px; margin: 0px; border-collapse: 0; border-spacing: 0px; empty-cells: 0; caption-side: 0; table-layout: 0;"--%>
				<img alt="" src="images/i2/login_05.jpg" width="229" height="157">
			</td>
			<td style="z-index: 0" colspan="3">
				<img alt="" src="images/i2/login_06b.jpg" width="564" height="105">
			</td>
			<td rowspan="2">
				<img alt="" src="images/i2/login_07.jpg" width="207" height="157">
			</td>
		</tr>
		<tr>
			<td style="z-index: 0" colspan="3">
				<img alt="" src="images/i2/login_08.jpg" width="564" height="46">
			</td>
		</tr>
		<tr>
			<td colspan="5">
				<table width="100%" height="76" cellpadding="0" cellspacing="0" style="background-image: url('images/i2/bottom_blue.jpg')">
					<tr>
						<td align="center">
							<asp:Label ID="labBottomContent" CssClass="bottom_content" runat="server"></asp:Label>&nbsp; </td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<img alt="" src="images/i2/間距.gif" width="229" height="1">
			</td>
			<td>
				<img alt="" src="images/i2/間距.gif" width="113" height="1">
			</td>
			<td>
				<img alt="" src="images/i2/間距.gif" width="291" height="1">
			</td>
			<td>
				<img alt="" src="images/i2/間距.gif" width="160" height="1">
			</td>
			<td>
				<img alt="" src="images/i2/間距.gif" width="207" height="1">
			</td>
		</tr>
	</table>
	<div style="position: absolute; top: 508px; left: 1px;" id="div12" runat="server">
		<%--ForeColor="White"--%>
		<asp:Label ID="Labmsg1" runat="server" Text="Labx" ForeColor="White"></asp:Label>
	</div>
	<div style="position: absolute; top: 377px; left: 233px" id="div1" runat="server">
		<table id="tbBtnGroup1" runat="server">
			<tr>
				<td>&nbsp;&nbsp; </td>
			</tr>
			<tr>
				<td>
					<asp:Button ID="Btn_3" runat="server" Text="學員填寫期末滿意度" CssClass="home_button_b_M" />
					<asp:Button ID="Btn_7" runat="server" Text="學員受訓期間意見調查表填寫" CssClass="home_button_b_L" />
					<asp:Button ID="Btn_X" runat="server" Text="X" CssClass="home_button_b_S" />
				</td>
			</tr>
		</table>
		<table id="tbBtnGroup2" runat="server" width="600">
			<tr>
				<td colspan="2" align="center">&nbsp;&nbsp;</td>
			</tr>
			<tr>
				<td align="right">驗 證 碼： </td>
				<td style="text-align: left;">&nbsp;
					<asp:TextBox ID="txtvnum" runat="server" Columns="10" MaxLength="4" AutoCompleteType="Disabled"></asp:TextBox>
					<asp:Image ID="Image2" runat="server" ImageUrl="Common/ValidateCode.aspx" onclick="RefreshImage('Image2');" Style="vertical-align: middle; cursor: pointer; border: 1px solid #cccccc;" title="點選可更換驗證碼!!"></asp:Image>
				</td>
			</tr>
			<tr>
				<td colspan="2" align="center">&nbsp;&nbsp;
					<asp:ImageButton Style="z-index: 0" ID="ImgBtnSubmit2" runat="server" ImageUrl="images/i2/button/login_send.bmp"></asp:ImageButton>
					<font face="新細明體">&nbsp;</font>
					<asp:ImageButton Style="z-index: 0" ID="ImgBtnBackup2" runat="server" ImageUrl="images/i2/button/login_rest.bmp"></asp:ImageButton>
				</td>
			</tr>
		</table>
	</div>
	<div style="position: absolute; left: -14px; top: 19px; height: 17px; width: 26px;" id="divC" runat="server">
		<a id="A1" title="關閉" class="l" href="#" onclick="window.opener=null; window.open('','_self'); window.close();">關閉</a>
	</div>
	<div style="position: absolute; left: 13px; top: 1px; height: 17px; width: 148px;" id="divC2" runat="server">
		<a id="A2" title="MOICA內政部憑證管理中心" class="l" href="#" onclick="window.open('http://moica.nat.gov.tw/','_self');">MOICA內政部憑證管理中心</a>
	</div>
	<input id="hid_btsubmit1" type="hidden" runat="server" />
	<asp:HiddenField ID="Hid_BtnV1" runat="server" />
	<input id="AltMsg" type="hidden" runat="server">
	<script language="javascript" type="text/javascript">
		var Msg1 = document.getElementById("AltMsg");
		if (Msg1 && Msg1.value != '') {
			alert(Msg1.value);
		}
	</script>
	</form>
</body>
</html>
