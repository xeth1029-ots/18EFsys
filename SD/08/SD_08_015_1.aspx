<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_08_015_1.aspx.vb" Inherits="WDAIIP.SD_08_015_1" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>離退清冊列印</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <style>
        <!-- /* Style Definitions */
        @page Section1 { size: 841.9pt 595.3pt; margin: 3.17cm 3.17cm 3.17cm 3.17cm; mso-header-margin: 42.55pt; mso-footer-margin: 49.6pt; mso-paper-source: 0; }
        div.Section1 { page: Section1; }
        -->
    </style>
</head>
<body>
    <!-- MeadCo ScriptX -->
    <%--<OBJECT id="factory" style="DISPLAY:none" classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="../../scriptx/smsx.cab#Version=6,6,440,26" VIEWASTEXT></OBJECT>--%>
    <script defer>
        function window.onload() {
            window.print();
            //if (!factory.object) {
            //	return
            //} else {
            //	factory.printing.header = ""
            //	factory.printing.footer = ""
            //	factory.printing.portrait = false//橫印
            //	factory.printing.Print(true)
            //	window.close();
            //}
        }
    </script>

    <form id="form1" method="post" runat="server">
        <asp:Panel ID="printrpt" runat="server">
            <div class="Section1" id="div_print" runat="server"></div>
        </asp:Panel>
    </form>
</body>
</html>
