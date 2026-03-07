<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="LessonTeah1.aspx.vb" Inherits="WDAIIP.LessonTeah1" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>教師一</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script language="javascript" type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        //if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script language="javascript" type="text/javascript">
        function returnValue(ILessonTeah, NLessonTeah) {
            var flagRole = 1; //一般規則 //Add
            //編輯規則
            var modifytype = document.getElementById('modifytype');
            if (modifytype && modifytype.value == 'Edit') { flagRole = 2; }
            var v_fieldname = getParamValue('fieldname');
            if (v_fieldname != '') { eval('opener.document.form1.' + v_fieldname).value = NLessonTeah; }
            var v_hiddenname = getParamValue('hiddenname');
            if (v_hiddenname != '') { eval('opener.document.form1.' + v_hiddenname).value = ILessonTeah; }
            if (v_fieldname == '') {
                var o_OLessonTeah1 = opener.document.form1.OLessonTeah1;
                var o_OLessonTeah1Value = opener.document.form1.OLessonTeah1Value;
                if (o_OLessonTeah1) { o_OLessonTeah1.value = NLessonTeah; }
                if (o_OLessonTeah1Value) { o_OLessonTeah1Value.value = ILessonTeah; }
            }
            window.close();
        }

        function returnValue2(ILessonTeah, NLessonTeah) {
            var flagRole = 1; //一般規則 //Add2
            //編輯規則
            var modifytype = document.getElementById('modifytype');
            if (modifytype && modifytype.value == 'Edit') { flagRole = 2; }
            var v_fieldname = getParamValue('fieldname');
            if (v_fieldname != '') { eval('opener.document.form1.' + v_fieldname).value = NLessonTeah; }
            var v_hiddenname = getParamValue('hiddenname');
            if (v_hiddenname != '') { eval('opener.document.form1.' + v_hiddenname).value = ILessonTeah; }
            if (getParamValue('fieldname') == '') {
                var o_OLessonTeah2 = opener.document.form1.OLessonTeah2;
                var o_OLessonTeah2Value = opener.document.form1.OLessonTeah2Value;
                if (o_OLessonTeah2) { o_OLessonTeah2.value = NLessonTeah; }
                if (o_OLessonTeah2Value) { o_OLessonTeah2Value.value = ILessonTeah; }
            }
            window.close();
        }

        function Search_click() {
            var txtSearch1 = document.getElementById('txtSearch1');
            var btnSearch = document.getElementById('btnSearch');
            var hbtnSearch = document.getElementById('hbtnSearch');
            if (txtSearch1) {
                if (event.keyCode == 13) {
                    if (btnSearch) { btnSearch.disabled = true; }
                    if (hbtnSearch) { hbtnSearch.click(); }
                }
            }
        }

        function Search_click2() {
            var btnSearch = document.getElementById('btnSearch');
            var hbtnSearch = document.getElementById('hbtnSearch');
            if (btnSearch) { btnSearch.disabled = true; }
            if (hbtnSearch) { hbtnSearch.click(); }
        }

    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <div style="overflow-y: auto; height: 650px">
            <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                <tr>
                    <td>
                        <table class="table_sch" cellspacing="0" border="0">
                            <tr>
                                <td colspan="3">
                                    <table class="font" cellspacing="0" border="0" width="100%">
                                        <tr>
                                            <td class="bluecol" style="width: 20%">姓名：
                                            </td>
                                            <td class="whitecol" style="width: 20%">
                                                <asp:TextBox ID="txtSearch1" runat="server" Width="100%"></asp:TextBox>
                                            </td>
                                            <td>
                                                <asp:Button ID="btnSearch" runat="server" Text="搜尋" CssClass="asp_button_M"></asp:Button><input id="hbtnSearch" style="display: none" type="button" value="搜尋" name="hbtnSearch" runat="server">
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td align="center" class="head_navy">內聘
                                </td>
                                <td align="center" class="head_navy">外聘
                                </td>
                                <td align="center" class="head_navy">未管理外聘師資
                                </td>
                                <td valign="top" align="center" rowspan="2">
                                    <asp:Button ID="Clear" Text="清除" runat="server" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                            <tr>
                                <td valign="top" class="whitecol">
                                    <asp:RadioButtonList ID="LessonTeahList1" runat="server" RepeatDirection="Horizontal" RepeatColumns="2" CssClass="font" AutoPostBack="True">
                                    </asp:RadioButtonList>
                                </td>
                                <td valign="top" class="whitecol">
                                    <asp:RadioButtonList ID="LessonTeahList2" runat="server" RepeatDirection="Horizontal" RepeatColumns="2" CssClass="font" AutoPostBack="True">
                                    </asp:RadioButtonList>
                                </td>
                                <td valign="top" class="whitecol">
                                    <asp:RadioButtonList ID="LessonTeahList3" runat="server" RepeatDirection="Horizontal" RepeatColumns="2" CssClass="font" AutoPostBack="True">
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </div>
        <input id="modifytype" type="hidden" name="modifytype" runat="server" size="1">
        <input id="fieldname" type="hidden" name="fieldname" runat="server" size="1">
        <input id="hiddenname" type="hidden" name="hiddenname" runat="server" size="1">
        <input id="ExistTech" type="hidden" name="ExistTech" runat="server" size="1">
    </form>
</body>
</html>
