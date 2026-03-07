<%@ Page Language="VB" AutoEventWireup="false" Inherits="WDAIIP.SD_05_002_R_Rpt" CodeBehind="SD_05_002_R_Rpt.aspx.vb" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>學員出缺勤明細表</title>
    <%--<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">--%>
    <%--<link href="../../css/style.css" type="text/css" rel="stylesheet" />--%>
    <%--<style type="text/css">
        @page Section1 { size: 841.9pt 545.3pt; margin: 3.17cm 3.17cm 3.17cm 3.17cm; mso-header-margin: 42.55pt; mso-footer-margin: 49.6pt; mso-paper-source: 0; }
        div.Section1 { page: Section1; }
    </style>--%>
    <style type="text/css">
        table { width: 99%; background-image: url('../../images/rptpic/temple/TIMS_1.jpg'); background-repeat: repeat-y; background-size: 100%; }
        .content-box { background-image: url('../../images/rptpic/temple/TIMS_1.jpg'); background-repeat: repeat-y; background-position: center; background-attachment: fixed; background-size: 100%; }
    </style>
    <script type="text/javascript" language="javascript">
        var isprint = false;
        function PrintRpt() {
            var content_vlue = document.getElementById("div_print_content").innerHTML;
            /**
			var disp_setting="toolbar=yes,location=no,directories=yes,menubar=yes,"; 
			disp_setting+="scrollbars=yes,width=650, height=600, left=100, top=25"; 
			var docprint=window.open("","",disp_setting); 

			docprint.document.open(); 
			docprint.document.write('<html><head><title>【產業人才投資計畫】訓練班別計畫表</title>'); 
			docprint.document.write('<style> table{ width:95%;}@page:left {margin-left: 4cm; margin-right:3cm;}@page:right{margin-left: 3cm; margin-right:4cm;} </style> </head><body onLoad="self.print()"><center>');          
			docprint.document.write(content_vlue);          
			docprint.document.write('</center></body></html>'); 
			docprint.document.close(); 
			docprint.focus();   
			**/
            //document.write('<IMG src="../../images/rptpic/temple/TIMS_3.jpg" style="z-index:-1;position:absolute;top:0px;left:0px;;;display:inline" />');
            //document.write('<body  background="../../images/rptpic/temple/TIMS_2.jpg"   onLoad="self.print();window.history.go(-1);">');                 
            document.write('<html><head><style>table{ width:95%;} @page:left {margin-left: 4cm; margin-right:3cm;} @page:right{margin-left: 3cm; margin-right:4cm;} </style><title>學員出缺勤明細表</title> </head>');
            document.write('<body onLoad="self.print();window.history.go(-1);">');
            document.write(content_vlue);
            document.write('</body></html>');
            document.close();
            //window.history.go(-1);   
            // history.go(-1);
        }

        function Show_Background() {
            //debugger;
            var tb_01 = document.getElementById('tb_01');
            tb_01.style.display = "none";
            window.print();
            //if (document.frames != undefined) { document.frames[0].ShowFrame(document.body.innerHTML); }            
        }

        function Set_FrameHeight(hh) {
            //debugger;
            var tb_01 = document.getElementById('tb_01');
            var screenH = 942;
            for (i = 0; i < Math.ceil(hh / screenH) ; i++) {
                document.body.innerHTML = document.body.innerHTML + "<IMG src='../../images/rptpic/temple/TIMS_1.jpg' style='z-index:-1;position:absolute;top:" + (i * screenH) + "px;left:0px;display:inline' />";
            }
            Auto_Print();
            tb_01.style.display = "inline";
        }

        function Auto_Print() {
            window.print();
            isprint = true;
        }

        function bufferTime() {
            setTimeout("goBack()", 500);
        }

        function goBack() {
            if (isprint == true) {
                isprint = false;
                bufferTime();
            }
            else {
                window.history.go(0);
            }
        }
    </script>

</head>
<body>
    <form id="form1" method="post" runat="server">
       <%--  <table width="99%" border="0" cellpadding="0" cellspacing="0"><tr id="trBtn" runat="server"><td align="right" class="whitecol" colspan="2"><asp:Button ID="btnExport" runat="server" Text="匯出明細" CssClass="asp_button_M" /><asp:Button ID="btnPrt" runat="server" Text="列印" CssClass="asp_button_M" /><asp:Button ID="btnCancel" runat="server" Text="取消" CssClass="asp_button_M" /></td></tr><tr><td><div id="div_print" class="Section1" runat="server"></div></td><td width="2%">&nbsp;</td></tr></table>--%>
        <table id="tb_01">
            <tr id="trBtn" runat="server">
                <td align="right">
                    <!-- <IMG onclick="self.print();" alt="列印報表" src="../../images/rptpic/Print.gif">&nbsp; -->
                    <img onclick="if(!isprint){Show_Background();}else{Auto_Print();} goBack();" alt="列印報表" src="../../images/rptpic/Print.gif">&nbsp;
                    <asp:Button ID="btnExport" runat="server" Text="匯出明細" CssClass="asp_button_M" />
                    <%--<asp:Button ID="btnPrt" runat="server" Text="列印Pdf" CssClass="asp_button_M" />--%>
                    <asp:Button ID="btnCancel" runat="server" Text="取消" CssClass="asp_button_M" />
                </td>
            </tr>
        </table>
        <!---print start-->
        <div id="div_print_content" runat="server"></div>
        <!---print end-->
        <iframe id="ifram1" src="../../RPT.htm" width="100%" height="0"></iframe>

        <asp:HiddenField ID="hidOCID" runat="server" />
        <asp:HiddenField ID="hidTMID" runat="server" />
        <asp:HiddenField ID="hidTPlanID" runat="server" />
        <asp:HiddenField ID="hidRID" runat="server" />
        <asp:HiddenField ID="hidSDate" runat="server" />
        <asp:HiddenField ID="hidEDate" runat="server" />
        <asp:HiddenField ID="hidItem1" runat="server" />
        <asp:HiddenField ID="hidItem2" runat="server" />
        <asp:HiddenField ID="hidItem3" runat="server" />
        <asp:HiddenField ID="hidItem4" runat="server" />
        <asp:HiddenField ID="hidUserID" runat="server" />
        <asp:HiddenField ID="hidprtPageSize" runat="server" />
        
        <asp:HiddenField ID="hid" runat="server" />
    </form>
</body>
</html>
