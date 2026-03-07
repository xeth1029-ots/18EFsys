<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_019_R_1.aspx.vb" Inherits="WDAIIP.SD_05_019_R_1" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>學員成績表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <style>
        @page :left {
            margin-left: 4cm;
            margin-right: 3cm;
        }
        @page :right {
            margin-left: 3cm;
            margin-right: 4cm;
        }
    </style>
    <!-- MeadCo ScriptX -->
    <%--<object id="factory" style="display: none" codebase="../../scriptx/smsx.cab#Version=6,6,440,26" classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" viewastext></object>--%>
    <script language="javascript">
        var isprint = false;

        function PrintRpt() {
            var content_vlue = document.getElementById("print_content").innerHTML;
            alert(content_vlue);
            document.write('<html><head><style>table{ width:95%;} @page:left {margin-left: 4cm; margin-right:3cm;} @page:right{margin-left: 3cm; margin-right:4cm;} </style><title>【產業人才投資計畫】訓練班別計畫表</title> </head>');
            document.write('<body  onLoad="self.print();window.history.go(-1);">');
            document.write(content_vlue);
            document.write('</body></html>');
            document.close();
        }

        function Show_Background() {
            document.getElementById('tb_01').style.display = "none";
            document.frames[0].ShowFrame(document.body.innerHTML);
        }

        function Set_FrameHeight(hh) {
            //for(i=0;i<Math.ceil(hh/842);i++){
            //	document.body.innerHTML=document.body.innerHTML+"<IMG src='../../images/rptpic/temple/TIMS_1.jpg' style='z-index:-1;position:absolute;top:"+(i*842)+"px;left:0px;;;display:inline' />";
            //}
            //for(i=0;i<(Math.ceil(hh/595))+ parseInt(document.getElementById('WaterPage').value) ;i++){
            for (i = 0; i < parseInt(document.getElementById('WaterPage').value) ; i++) {
                document.body.innerHTML = document.body.innerHTML + "<IMG src='../../images/rptpic/temple/TIMS_1.jpg' style='z-index:-1;position:absolute;top:" + (i * 842) + "px;left:0px;;;display:inline' />";
            }
            Auto_Print();
            document.getElementById('tb_01').style.display = "inline";
        }

        function Auto_Print() {
            window.print();
            //if (!factory.object) {
            //    return
            //} else {
            //    document.all.factory.printing.header = ""; //頁首，空白為不印頁首，也就不會佔空間
            //    document.all.factory.printing.footer = ""; //註腳，空白為不印註腳，也就不會佔空間
            //    document.all.factory.printing.leftMargin = 0; //左邊界
            //    document.all.factory.printing.topMargin = 0; //上邊界
            //    document.all.factory.printing.rightMargin = 0; //右邊界
            //    document.all.factory.printing.bottomMargin = 0; //下邊界
            //    document.all.factory.printing.portrait = true; //直印，false:橫印 
            //    document.all.factory.printing.Print(true);
            //    isprint = true;
            //}
        }

        function bufferTime() {
            setTimeout("goBack()", 500);
        }

        function goBack() {
            if (isprint = true) {
                isprint = false;
                bufferTime();
            }
            else {
                window.history.go(0);
            }
        }
    </script>
</head>
<body background="../../images/rptpic/temple/TIMS_1.jpg">
    <form id="form1" method="post" runat="server">
        <table id="tb_01" align="right">
            <tr>
                <td align="right">
                    <input id="doPrint" style="width: 24px; height: 22px" type="hidden" name="doPrint" runat="server">
                    <input id="WaterPage" style="width: 24px; height: 22px" type="hidden" name="WaterPage" runat="server">
                    <asp:LinkButton ID="lb_print" runat="server"></asp:LinkButton>
                    <asp:ImageButton ID="bt_first" runat="server" Visible="False" ImageUrl="../../images/rptpic/First_Disabled.gif"></asp:ImageButton>&nbsp;&nbsp;
                    <asp:ImageButton ID="bt_pre" runat="server" Visible="False" ImageUrl="../../images/rptpic/Previous_Disabled.gif"></asp:ImageButton>&nbsp;&nbsp;
                    <asp:TextBox ID="tb_Page" runat="server" Visible="False" Width="40px" onfocus="this.blur()"></asp:TextBox>&nbsp;
                    <asp:TextBox ID="tb_PageTotal" runat="server" Visible="False" Width="40px" onfocus="this.blur()"></asp:TextBox>&nbsp;
                    <asp:ImageButton ID="bt_next" runat="server" Visible="False" ImageUrl="../../images/rptpic/Next_Disabled.gif"></asp:ImageButton>&nbsp;
                    <asp:ImageButton ID="bt_end" runat="server" Visible="False" ImageUrl="../../images/rptpic/Last_Disabled.gif"></asp:ImageButton>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <img id="printRpt" onclick="if(!isprint){  Show_Background();}else{Auto_Print();}goBack();" alt="列印報表" src="../../images/rptpic/Print.gif">&nbsp;&nbsp;
                    <asp:ImageButton ID="bt_excel" runat="server" Visible="False" ImageUrl="../../images/rptpic/Excel.gif" AlternateText="匯出Excel"></asp:ImageButton>&nbsp;
                </td>
            </tr>
            <tr>
                <td><asp:Label ID="lbmsg" runat="server" ForeColor="Red"></asp:Label></td>
            </tr>
        </table>
        <!---print    start-->
        <div id="print_content" align="left" runat="server"></div>
        <!---print    end-->
        <iframe id="ifram1" src="../../RPT.htm" width="100%" height="0"></iframe>
        <asp:HiddenField ID="Hid_OCID" runat="server" />
    </form>
</body>
</html>