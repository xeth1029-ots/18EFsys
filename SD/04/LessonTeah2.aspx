<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="LessonTeah2.aspx.vb" Inherits="WDAIIP.LessonTeah2" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>教師二</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
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
            var modifytype = document.getElementById('modifytype');
            var flagRole = 1; //一般規則 //Add
            if (modifytype.value == 'Edit') {
                flagRole = 2; //編輯規則
            }
            if (getParamValue('fieldname') == '') {
                opener.document.form1.OLessonTeah2.value = NLessonTeah;
                opener.document.form1.OLessonTeah2Value.value = ILessonTeah;
            }
            else {
                eval('opener.document.form1.' + getParamValue('fieldname')).value = NLessonTeah;
                eval('opener.document.form1.' + getParamValue('hiddenname')).value = ILessonTeah;
            }
            window.close();
        }

        function Search_click() {
            if (document.getElementById('txtSearch1')) {
                if (event.keyCode == 13) {
                    document.getElementById('btnSearch').disabled = true;
                    document.getElementById('hbtnSearch').click();
                }
            }
        }

        function Search_click2() {
            document.getElementById('btnSearch').disabled = true;
            document.getElementById('hbtnSearch').click();
        }

    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <div style="overflow-y: auto; height: 650px">
            <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                <tr>
                    <td>
                        <table class="table_sch" border="0" cellspacing="0">
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
                                                <asp:Button ID="btnSearch" runat="server" Text="搜尋" CssClass="asp_button_M"></asp:Button>
                                                <input id="hbtnSearch" style="display: none" type="button" value="搜尋" name="hbtnSearch" runat="server">
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
                                <td align="center" rowspan="2" valign="top">
                                    <asp:Button ID="Clear" runat="server" Text="清除" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                            <tr>
                                <td valign="top" class="whitecol">
                                    <asp:RadioButtonList ID="LessonTeahList1" AutoPostBack="True" CssClass="font" runat="server" RepeatDirection="Horizontal" RepeatColumns="2">
                                    </asp:RadioButtonList>
                                </td>
                                <td valign="top" class="whitecol">
                                    <asp:RadioButtonList ID="LessonTeahList2" AutoPostBack="True" CssClass="font" runat="server" RepeatDirection="Horizontal" RepeatColumns="2">
                                    </asp:RadioButtonList>
                                </td>
                                <td valign="top" class="whitecol">
                                    <asp:RadioButtonList ID="LessonTeahList3" AutoPostBack="True" CssClass="font" runat="server" RepeatDirection="Horizontal" RepeatColumns="2">
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </div>
        <input type="hidden" runat="server" id="modifytype" name="modifytype" />
        <input type="hidden" runat="server" id="fieldname" name="fieldname" />
        <input type="hidden" runat="server" id="hiddenname" name="hiddenname" />
    </form>
</body>
</html>
