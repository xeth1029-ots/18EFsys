<%@ Page Language="VB" AutoEventWireup="false" Inherits="WDAIIP.TR_04_001_R_Rpt" CodeBehind="TR_04_001_R_Rpt.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
	<title></title>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
	<style type="text/css">
		<!--
		/* Style Definitions */
		@page Section1 { size: 841.9pt 595.3pt; margin: 0cm 0cm 0cm 0cm; mso-header-margin: 42.55pt; mso-footer-margin: 49.6pt; mso-paper-source: 0; position: absolute; background-image: url('../../images/rptpic/temple/TIMS_2.jpg'); }
		div.Section1 { page: Section1; }
		/*以絕對位置設定當作最下層layer*/
		.Layer_Z { position: absolute; z-index: -1; margin: 0; padding: 0; left: 0px; top: 0px; }
		-->
	</style>
</head>
<body>
	<%--<object style="display: none" id="factory" codebase="../../scriptx/ScriptX.cab#Version=6,2,433,14" classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" viewastext></object>--%>
	<script defer>
	    function print() {
	        window.print();
			//if (!factory.object) {
			//	return
			//} else {
			//	factory.printing.header = ""
			//	factory.printing.footer = ""
			//	factory.printing.portrait = false
			//	factory.printing.Print(true)
			//	window.close();
			//}
		}
	</script>
	<form id="form1" runat="server">
	<div id="div_print" class="Section1" runat="server">
	</div>
	</form>
</body>
</html>
