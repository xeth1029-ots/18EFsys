<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_14_002_R.aspx.vb" Inherits="WDAIIP.SD_14_002_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訓練班別計畫表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
<%--<style type="text/css">
table { width: 99%; background-image: url('../../images/rptpic/temple/TIMS_1.jpg'); background-repeat: repeat-y; background-size: 100%; }
.content-box { background-image: url('../../images/rptpic/temple/TIMS_1.jpg'); background-repeat: repeat-y; background-position: center; background-attachment: fixed; background-size: 100%; }
</style>--%>
    <style type="text/css">
        /*body { background-image: url('../../images/rptpic/temple/TIMS_1.jpg'); background-repeat: no-repeat; background-position: center center; }*/
        div { background-image: url('../../images/rptpic/temple/TIMS_3.jpg'); background-repeat:repeat; background-position: center center; }
    </style>
<!-- MeadCo ScriptX -->
<%--<object style="display: none" id="factory" codebase="../../scriptx/smsx.cab#Version=6,6,440,26" classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" viewastext></object>--%>
    <script type="text/javascript" language="javascript">
        var isprint = false;
        function PrintRpt() {
            var content_vlue = document.getElementById("print_content").innerHTML;
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
            document.write('<html><head><style>table{ width:95%;} @page:left {margin-left: 4cm; margin-right:3cm;} @page:right{margin-left: 3cm; margin-right:4cm;} </style><title>【產業人才投資計畫】訓練班別計畫表</title> </head>');
            document.write('<body  onLoad="self.print();window.history.go(-1);">');
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
                document.body.innerHTML = document.body.innerHTML + "<IMG src='../../images/rptpic/temple/TIMS_3.jpg' style='z-index:-1;position:absolute;top:" + (i * screenH) + "px;left:0px;display:inline' />";
            }
            //document.getElementById('tb_01').style.display="none";
            Auto_Print();
            //self.print();
            tb_01.style.display = "inline";
        }

        function Auto_Print() {
            window.print();
            //debugger;
            //window.print();
            //if (!factory.object) { return; }
            //document.all.factory.printing.header = "";
            //document.all.factory.printing.footer = "";
            //document.all.factory.printing.portrait = true;
            //document.all.factory.printing.Print(true);
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
        /*
        var actionText = document.getElementById("actionText");
        if (actionText) {
            var yReqPDFOUT = new Request("PDFOUT");
            if (yReqPDFOUT != null && yReqPDFOUT == "YB") {
                actionText.innerHTML = "資料讀取中";
            }
        }
        */
    </script>
</head>
<body>
    <!--<img src="../../images/rptpic/temple/TIMS_2.jpg" style="DISPLAY:inline;Z-INDEX:-1;LEFT:0px;POSITION:absolute;TOP:0px">-->
    <form id="form1" method="post" runat="server">
        <table id="tb_01" width="100%">
            <tr>
                <td align="right">
                    <!-- <IMG onclick="self.print();" alt="列印報表" src="../../images/rptpic/Print.gif">&nbsp; -->
                    <img onclick="if(!isprint){Show_Background();}else{Auto_Print();} goBack();" alt="列印報表" src="../../images/rptpic/Print.gif">
                    &nbsp;<asp:ImageButton ID="bt_excel" runat="server" ImageUrl="../../images/rptpic/Excel.gif" AlternateText="匯出Excel"></asp:ImageButton>
                    &nbsp;<asp:ImageButton ID="imgBt_Pdf" runat="server" ImageUrl="../../images/rptpic/pdf15.png" AlternateText="匯出pdf"></asp:ImageButton>
                </td>
            </tr>
        </table>
        <%--<span id="actionText" runat="server"></span><img id="img_waiting" runat="server" src="../../images/waiting.gif" alt="waiting" />--%>
        <!---print start-->
        <div id="print_content" runat="server" style="background-image: url('../../images/rptpic/temple/TIMS_3.jpg'); background-repeat:repeat; background-position: center center;"></div>
        <!---print end-->
        <iframe id="ifram1" src="../../RPT.htm" width="100%" height="0"></iframe>
        <asp:HiddenField ID="Hid_OJT22071401" runat="server" />
    </form>
</body>
</html>
