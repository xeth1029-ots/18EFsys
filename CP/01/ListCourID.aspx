<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ListCourID.aspx.vb" Inherits="WDAIIP.ListCourID" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>當日課程</title>
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
        function returnValue(CourID1, CourIDValue1) {
            if (document.form1.modifytype.value == 'Add') {
                if (getParamValue('fieldname') == '') {
                    opener.document.form1.CourID.value = CourID1;
                    opener.document.form1.CourIDValue.value = CourIDValue1;
                }
                else {
                    eval('opener.document.form1.' + getParamValue('fieldname')).value = CourID1;
                    eval('opener.document.form1.' + getParamValue('hiddenname')).value = CourIDValue1;
                }
                window.close();
            } else if (document.form1.modifytype.value == 'Edit') {
                opener.document.form1[document.form1.fieldname.value].value = CourID1;
                opener.document.form1.CourIDValue.value = CourIDValue1;
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
                                <td align="center" class="head_navy" colspan="2" width="84%"><font color="#ffffff">課程</font></td>
                                <td valign="top" align="center" rowspan="2" width="16%">
                                    <br />
                                    <asp:Button ID="Clear" Text="清除" runat="server" CssClass="asp_button_M"></asp:Button></td>
                            </tr>
                            <tr>
                                <td valign="top" colspan="2">
                                    <asp:RadioButtonList ID="CourList1" runat="server" RepeatDirection="Horizontal" RepeatColumns="2" CssClass="font" AutoPostBack="True"></asp:RadioButtonList></td>
                            </tr>
                        </table>
                        <input id="modifytype" type="hidden" name="modifytype" runat="server">
                        <input id="fieldname" type="hidden" name="fieldname" runat="server">
                    </td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
