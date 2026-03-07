<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="indexA1.aspx.vb" Inherits="WDAIIP.indexA1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
	<title>職業訓練業務資訊管理網</title>
	<link rel="Shortcut Icon" href="http://tims.etraining.gov.tw/css/wdalogo.ico" type="image/x-icon" />
	<meta http-equiv="Content-Type" content="text/html; charset=big5">
	<style id="antiClickjack">
		body { display: none !important; }
	</style>
	<script language="javascript" type="text/javascript">
		if (self === top) {
			var antiClickjack = document.getElementById("antiClickjack");
			antiClickjack.parentNode.removeChild(antiClickjack);
		} else {
			top.location = self.location;
		}
		var url = window.location.href;
		if (url.indexOf("https:") == -1) {
			url = url.replace("http:", "https:")
			window.location.replace(url);
		}
	</script>
</head>
<frameset rows="83,*" cols="*" border="0" framespacing="0" frameborder="0">
	<frame src="top.aspx" name="topFrame" scrolling="no" noresize>
	<frameset rows="*,63" cols="*" framespacing="0" frameborder="NO" border="0">
		<frameset id="MainBlock" cols="253,*" frameborder="NO" border="0" framespacing="0">
			<frame src="menu.aspx" name="leftFrame">
			<frameset rows="43,*" frameborder="no" border="0" framespacing="0">
				<frame src="title.aspx" name="titleFrame" scrolling="No" noresize="noresize" />
				<frame src="main.aspx" name="mainFrame" scrolling="auto" noresize="noresize" />
			</frameset>
		</frameset>
		<frame src="bottom.aspx" name="bottomFrame" scrolling="no" noresize>
	</frameset>
	<noframes>
		<body>
			<p>此網頁使用框架，但是您的瀏覽器不支援框架。</p>
		</body>
	</noframes>
</frameset>
</html>
