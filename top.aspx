<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="top.aspx.vb" Inherits="WDAIIP.top" %>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
	<title>top</title>
	<link rel="Shortcut Icon" href="http://tims.etraining.gov.tw/css/wdalogo.ico" type="image/x-icon" />
	<link href="css/css.css" rel="stylesheet" type="text/css" />
	<meta http-equiv="Content-Type" content="text/html; charset=big5" />
	<script language="javascript" type="text/javascript">
		var ap_name = navigator.appName;
		var ap_vinfo = navigator.appVersion;
		var ap_ver = parseFloat(ap_vinfo.substring(0, ap_vinfo.indexOf('(')));
		var dl_ok = false;

		var time_start = new Date();
		var clock_start = time_start.getTime();

		function init() {
			if (ap_name == "Netscape" && ap_ver >= 3.0) {
				dl_ok = true;
				return true;
			}
		}

		function get_time_spent() {
			var time_now = new Date();

			if (document.getElementById('hidTime').value != '') {
				clock_start = time_now.getTime();
				document.getElementById('hidTime').value = '';
			}
			return ((time_now.getTime() - clock_start) / 1000);
		}

		// show the time user spent on the side //顯示倒數時間
		function show_secs() {
			var i_MM = 20; //5.05; //0.5; //20; //倒數時間(分鐘)
			var s_lastM = '05:00'; //'00:10'; //'05:00'; //最後倒數時間(分鐘)
			var i_total_secs = i_MM * 60 - Math.round(get_time_spent());
			var i_secs_spent = i_total_secs % 60;
			var i_mins_spent = Math.round((i_total_secs - 30) / 60);
			var s_secs_spent = "" + ((i_secs_spent > 9) ? i_secs_spent : "0" + i_secs_spent);
			var s_mins_spent = "" + ((i_mins_spent > 9) ? i_mins_spent : "0" + i_mins_spent);

			// document.getElementById("txtTime").value = s_mins_spent + ":" + s_secs_spent;
			// window.setTimeout('show_secs()', 1000);
			var flagexit = false;
			flagexit = false;
			document.getElementById("txtTime").value = s_mins_spent + ":" + s_secs_spent;
			if (flagexit == false && i_mins_spent <= 0 && i_secs_spent <= 0) { flagexit = true; }
			if (flagexit == false && s_mins_spent == '00' && s_secs_spent == '00') { flagexit = true; }
			if (flagexit == false && document.getElementById("txtTime").value == '00:00') { flagexit = true; }

			if (!flagexit) {
				window.setTimeout('show_secs()', 1000);
				if (document.getElementById("txtTime").value <= s_lastM) {
					document.getElementById("txtTime").style.color = 'red';
					document.getElementById("txtTime").style.fontWeight = 'bolder';
					if (document.getElementById("txtTime").value == s_lastM) { showMsg(); }
					if (document.getElementById("txtTime").value == s_lastM) { showMsg2(); }
				} else {
					document.getElementById("txtTime").style.color = 'black';
					document.getElementById("txtTime").style.fontWeight = '';
					closeDiv();
				}
			}
			else {
				alert('您的登入資訊已經遺失，請重新登入');
				top.location.replace('logout.aspx');
				//top.location.replace(self.location); //top.location.href = 'logout.aspx';
				//document.getElementById("logout").click();
			}
		}

		//MSN提示訊息
		window.onresize = resizeDiv;
		var divTop, divLeft, divWidth, divHeight, docHeight, docWidth, objTimer, i = 0;

		function showMsg2() {
			var msg2 = '您閒置系統已超過１５分鐘\n若未進行資料儲存或點選其他功能，再過５分鐘系統將會自動中斷連線。';
			alert(msg2);
		}

		function showMsg() {
			try {
				divTop = parseInt(document.getElementById("eMeng").style.top, 10)
				divLeft = parseInt(document.getElementById("eMeng").style.left, 10)
				divHeight = parseInt(document.getElementById("eMeng").offsetHeight, 10)
				divWidth = parseInt(document.getElementById("eMeng").offsetWidth, 10)

				docWidth = document.body.clientWidth;
				docHeight = document.body.clientHeight;

				document.getElementById("eMeng").style.top = parseInt(document.body.scrollTop, 10) + docHeight + 10;
				document.getElementById("eMeng").style.left = parseInt(document.body.scrollLeft, 10) + docWidth - divWidth - 17
				document.getElementById("eMeng").style.visibility = "visible"
				objTimer = window.setInterval("moveDiv()", 10)
			}
			catch (err) { }
			/**
			var txt="There was an error on this page.\n\n";
			txt+="Error description: " + err.description + "\n\n";
			txt+="Click OK to continue.\n\n";
			alert(txt);
			**/
		}

		//長寬動作
		function resizeDiv() {
			if (i > 500) closeDiv()
			try {
				divHeight = parseInt(document.getElementById("eMeng").offsetHeight, 10)
				divWidth = parseInt(document.getElementById("eMeng").offsetWidth, 10)
				docWidth = document.body.clientWidth;
				docHeight = document.body.clientHeight;
				document.getElementById("eMeng").style.top = docHeight - divHeight + parseInt(document.body.scrollTop, 10)
				document.getElementById("eMeng").style.left = docWidth - divWidth + parseInt(document.body.scrollLeft, 10) - 17
			}
			catch (err) { }
			/**
			var txt="There was an error on this page.\n\n";
			txt+="Error description: " + err.description + "\n\n";
			txt+="Click OK to continue.\n\n";
			alert(txt);
			**/
		}

		//移動動作
		function moveDiv() {
			try {
				if (parseInt(document.getElementById("eMeng").style.top, 10) <= (docHeight - divHeight + parseInt(document.body.scrollTop, 10))) {
					window.clearInterval(objTimer);
					objTimer = window.setInterval("resizeDiv()", 1);
				}

				divTop = parseInt(document.getElementById("eMeng").style.top, 10);
				document.getElementById("eMeng").style.top = divTop - 3;
			}
			catch (err) { }
			/**
			var txt="There was an error on this page.\n\n";
			txt+="Error description: " + err.description + "\n\n";
			txt+="Click OK to continue.\n\n";
			alert(txt);
			**/
		}

		//關閉動作
		function closeDiv() {
			document.getElementById('eMeng').style.visibility = 'hidden';
			if (objTimer) window.clearInterval(objTimer);
		} 

	</script>
</head>
<body>
	<form id="form1" runat="server">
	<table border="0" cellpadding="0" cellspacing="0" style="background-image: url(./images/i2/top.bmp); width: 1000px; height: 83px">
		<tr>
			<td valign="top">
				<table width="305px" align="right" border="0" cellpadding="0" cellspacing="0">
					<tr>
						<td>
							<asp:HyperLink ID="HyperLink2" ImageUrl="./images/i2/button/top_sch.bmp" runat="server" Target="mainFrame" /><a href="index.htm" target="_parent" title="回首頁" id="a_http" runat="server"><asp:Image ID="img_http" runat="server" ImageUrl="./images/i2/button/top_main.bmp" /></a><asp:HyperLink ID="HyperLink1" runat="server" ImageUrl="./images/i2/button/top_plan.bmp" ToolTip="切換計畫" Target="_parent"></asp:HyperLink><a id="logout" href="logout.aspx" target="_parent" title="登出"><asp:Image ID="Image2" ImageUrl="./images/i2/button/top_logout.bmp" runat="server" /></a> </td>
						<td>&nbsp; </td>
					</tr>
					<tr height="32px">
						<td align="right">
							<input id="txtTime" type="text" value="20:00" onfocus="this.blur()" style="border: 0; font-size: 15px; width: 40px" />&nbsp;&nbsp; </td>
						<td width="20px">&nbsp; </td>
					</tr>
				</table>
				<asp:HiddenField ID="hidTime" runat="server" />
			</td>
		</tr>
	</table>
	<table class="fontmsn" id="eMeng" style="visibility: hidden; border-right: #455690 1px solid; border-top: #a6b4cf 1px solid; z-index: 99999; left: 0px; border-left: #a6b4cf 1px solid; border-bottom: #455690 1px solid; position: absolute; top: 0px; height: 100px; background-color: #c9d3f3" cellspacing="1" cellpadding="1" width="180px" border="0" runat="server">
		<tr>
			<td background="./images/MSNTitle.gif"><font face="新細明體">
				<table class="fontmsn" id="Table4" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>TIMS系統提示： </td>
						<td style="cursor: pointer" onclick="closeDiv();" align="center" width="15">
							<img src="./images/CloseMsn.gif">
						</td>
					</tr>
				</table>
			</font></td>
		</tr>
		<tr>
			<td style="border-right: #b9c9ef 1px solid; padding-right: 10px; border-top: #728eb8 1px solid; padding-left: 10px; font-size: 12px; padding-bottom: 0px; border-left: #728eb8 1px solid; width: 100%; color: #1f336b; padding-top: 5px; border-bottom: #b9c9ef 1px solid;" align="left" background="./images/MsnBack.gif" colspan="1" height="80"><font style="color: red">您閒置系統已超過１５分鐘<br />
				若未進行資料儲存或點選其他功能，再過５分鐘系統將會自動中斷連線。 </font></td>
		</tr>
	</table>
	<div style="position: absolute; top: 3px; left: 1031px;" id="div13" runat="server">
		<%--ForeColor="White"--%>
		<asp:Label ID="Labmsg1" runat="server" Text="Labx" ForeColor="White"></asp:Label>
	</div>
	</form>
</body>
<script language="javascript" type="text/javascript">
	init();
	window.setTimeout('show_secs();', 1);
</script>
</html>
