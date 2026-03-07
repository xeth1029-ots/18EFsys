<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_11_001_add16.aspx.vb" Inherits="WDAIIP.SD_11_001_add16" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訓練期末學員滿意度調查表</title>
    <meta http-equiv="Content-Type" content="text/html; charset=big5">
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>

    <script type="text/javascript" language="javascript">
        function printDoc() {
            window.print();
            //if (_isIE) { window.print(); window.close(); }
            //else { window.print(); }

            //if (!factory.object) {
            //    return false;
            //}
            //else {
            //    factory.printing.header = '';
            //    factory.printing.footer = '';
            //    factory.printing.portrait = true;
            //    factory.printing.Print(true);
            //}
        }
    </script>
</head>
<body>
    <!-- MeadCo ScriptX -->
    <%--<object style="display: none" id="factory" codebase="../../scriptx/smsx.cab#Version=6,6,440,26" classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" viewastext></object>--%>
    <form id="form1" method="post" runat="server">
        <table id="tbControl2" runat="server" class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="labStud" runat="server" CssClass="font"></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:Label ID="labName" runat="server" CssClass="font"></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:Label ID="labStatus" runat="server"></asp:Label>
                </td>
            </tr>
        </table>
        <asp:Panel ID="Panel1" runat="server">
            <asp:PlaceHolder ID="PlaceHolder1" runat="server"></asp:PlaceHolder>
        </asp:Panel>
        <table id="tbControl1" runat="server" class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <div align="center" class="whitecol">
                        <asp:Button ID="btnSave1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>&nbsp;
                        <asp:Button ID="btnBack1" runat="server" Text="回上頁" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                    </div>
                </td>
            </tr>
        </table>
        <input id="H_Answercount" type="hidden" runat="server">
        <input id="H_SKID" type="hidden" runat="server">
        <input id="H_Type" type="hidden" runat="server">
        <input id="H_SQID" type="hidden" runat="server">
        <input id="H_Ivalue" type="hidden" runat="server">
        <asp:HiddenField ID="HID_SVID" runat="server" />
        <asp:HiddenField ID="Hid_SOCID" runat="server" />
        <asp:HiddenField ID="Hid_OCID" runat="server" />
        <asp:HiddenField ID="Hid_ProcessType" runat="server" />
        <asp:HiddenField ID="Hid_Stuedntid" runat="server" />
        <input id="Re_ID" type="hidden" runat="server">
    </form>
</body>
</html>
