<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_007_R_Prt2.aspx.vb" Inherits="WDAIIP.SD_15_007_R_Prt2" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
	<title></title>
	<style type="text/css">
		/*<!-- Style Definitions */
		@page Section1 { size: 841.9pt 595.3pt; margin: 0cm 0cm 0cm 0cm; mso-header-margin: 42.55pt; mso-footer-margin: 49.6pt; mso-paper-source: 0; position: absolute; background-repeat: repeat; background-image: url('../../images/rptpic/temple/TIMS_2.jpg'); }
		div.Section1 { page: Section1; }
		.Layer_Z { position: absolute; z-index: -1; /*以絕對位置設定當作最下層layer;*/ margin: 0; padding: 0; left: 0px; top: 0px; }
		/*--> */
	</style>
</head>
<body>
	<%--<object style="display: none" id="factory" codebase="../../scriptx/ScriptX.cab#Version=6,2,433,14" classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" viewastext></object>--%>
	<form id="form1" runat="server">
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr id="trBtn" runat="server">
			<td align="right">
				<asp:Button ID="btnPrt" runat="server" Text="列印" CssClass="asp_Export_M" />
				<asp:Button ID="btnExcel" runat="server" Text="匯出Excel" CssClass="asp_Export_M" />
                <asp:Button ID="btnExpOds1" runat="server" Text="匯出Ods" CssClass="asp_Export_M" />
				<asp:Button ID="btnCancel" runat="server" Text="取消" CssClass="asp_button_S" />
			</td>
		</tr>
		<tr>
			<td>
				<div id="div_print" class="Section1" runat="server">
				</div>
			</td>
		</tr>
	</table>
	</form>
</body>
</html>
