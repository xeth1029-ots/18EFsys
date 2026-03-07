<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_02_002_R_Prt.aspx.vb" Inherits="WDAIIP.CP_02_002_R_Prt" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css"> 
            <!-- 
            /* Style Definitions */ 
            @page Section1 {size:841.9pt 595.3pt; margin:1.0cm 1.0cm 1.0cm 1.0cm; mso-header-margin:42.55pt; mso-footer-margin:49.6pt; mso-paper-source:0;}
            div.Section1 {page:Section1;}
            --> 
	</style>
    <script language="javascript" src="../../js/print.js"></script>
</head>
<body>
    <%--<object style="display: none" id="factory" codebase="../../scriptx/ScriptX.cab#Version=6,2,433,14" classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" viewastext></object>--%>
    <script defer>
        function print() {
            window.print();
            //if (!factory.object) {
            //    return
            //} else {
            //    factory.printing.header = ""
            //    factory.printing.footer = ""
            //    factory.printing.portrait = false
            //    factory.printing.Print(true)
            //    window.close();
            //}
        }
    </script>
    <form id="form1" runat="server">
    <input id="Button1" type="button" value="列印" onclick="PrintPart('div_print',false, '');window.close();" />
    <asp:Panel ID="printrpt" runat="server">
        <div id="div_print" class="Section1" runat="server" style="background-image: url('../../images/rptpic/temple/TIMS_2.jpg'); background-repeat: repeat">
        </div>
    </asp:Panel>
    </form>
</body>
</html>
