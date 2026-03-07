<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="CP_02_016_R.aspx.vb" Inherits="WDAIIP.CP_02_016_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>各分署各項計畫訓練人數</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }

        function search() {
            var msg = '';
            var start_date = document.getElementById("start_date");
            var end_date = document.getElementById("end_date");
            if (start_date.value == '') msg += '請選擇 結訓日期起始\n';
            if (end_date.value == '') msg += '請選擇 結訓日期迄止\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        //選擇全部
        function SelectAll(obj, hidobj) {
            var num = getCheckBoxListValue(obj).length; //長度
            var myallcheck = document.getElementById(obj + '_' + 0); //第1個
            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
            else {
                for (var i = 1; i < num; i++) {
                    if ('0' == getCheckBoxListValue(obj).charAt(i)) {
                        document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(i);
                        var mycheck = document.getElementById(obj + '_' + i);
                        myallcheck.checked = mycheck.checked;
                        break;
                    }
                }
            }
        }

    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練與需求管理&gt;&gt;統計分析&gt;&gt;各分署各項計畫訓練人數</asp:Label>
                </td>
            </tr>
        </table>

        <table id="FrameTable3" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td class="bluecol_need" width="20%">結訓日期 </td>
                <td class="whitecol" runat="server" colspan="3">
                    <asp:TextBox ID="start_date" runat="server" Width="22%" MaxLength="10"></asp:TextBox>
                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="24" height="24">
                    ～<asp:TextBox ID="end_date" runat="server" Width="22%" MaxLength="10"></asp:TextBox>
                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="24" height="24">
                </td>
            </tr>
            <%--<tr>
            <td class="bluecol">預算來源 </td>
            <td class="whitecol">
            <asp:CheckBoxList ID="BudgetList" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
            </asp:CheckBoxList>
            </td>
            </tr>--%>
            <%--<tr>
            <td class="bluecol">身分別 </td>
            <td class="whitecol">
            <asp:CheckBoxList ID="Identity" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="4">
            </asp:CheckBoxList>
            <input id="Identity_List" type="hidden" value="0" name="Identity_List" runat="server">
            </td>
            </tr>--%>
            <tr>
                <td align="center" colspan="4">
                    <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                </td>
            </tr>

        </table>
    </form>
</body>
</html>
