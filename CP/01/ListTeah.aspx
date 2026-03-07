<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ListTeah.aspx.vb" Inherits="WDAIIP.ListTeah" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>教師</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function returnValue(ILessonTeah, NLessonTeah, ODegreeID, ODegreeValue) {
            if (document.form1.modifytype.value == 'Add') {
                if (getParamValue('fieldname') == '1') {
                    opener.document.form1.OLessonTeah1.value = NLessonTeah;
                    opener.document.form1.OLessonTeah1Value.value = ILessonTeah;
                    opener.document.form1.ODegreeIDValue1.value = ODegreeID;
                    opener.document.form1.ODegreeID1.value = ODegreeValue;
                }
                if (getParamValue('fieldname') == '2') {
                    opener.document.form1.OLessonTeah2.value = NLessonTeah;
                    opener.document.form1.OLessonTeah2Value.value = ILessonTeah;
                    opener.document.form1.ODegreeIDValue2.value = ODegreeID;
                    opener.document.form1.ODegreeID2.value = ODegreeValue;
                }
                window.close();
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <div style="max-height: 500px; overflow-y: auto;">
            <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                <tr>
                    <td>
                        <table class="font" cellspacing="0" border="1" width="100%">
                            <tr>
                                <td align="center" bgcolor="#2aafc0" class="head_navy" width="42%"><font color="#ffffff">內聘</font></td>
                                <td align="center" bgcolor="#2aafc0" class="head_navy" width="42%"><font color="#ffffff">外聘</font></td>
                                <td align="center" rowspan="2" width="16%">
                                    <asp:Button ID="Clear" Text="清除" runat="server" CssClass="asp_button_M"></asp:Button></td>
                            </tr>
                            <tr>
                                <td valign="top">
                                    <asp:RadioButtonList ID="LessonTeahList1" runat="server" RepeatDirection="Horizontal" RepeatColumns="3" CssClass="font" AutoPostBack="True"></asp:RadioButtonList></td>
                                <td valign="top">
                                    <asp:RadioButtonList ID="LessonTeahList2" runat="server" RepeatDirection="Horizontal" RepeatColumns="3" CssClass="font" AutoPostBack="True"></asp:RadioButtonList></td>
                            </tr>
                        </table>
                        <input id="modifytype" type="hidden" name="modifytype" runat="server" />
                        <input id="fieldname" type="hidden" name="fieldname" runat="server" />
                    </td>
                </tr>
            </table>
        </div>
        <p>
        </p>
    </form>
</body>
</html>
