<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_02_009_R.aspx.vb" Inherits="WDAIIP.CP_02_009_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>年度各項訓練人數統計</title>
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
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }

        function print() {
            var msg = '';

            if (document.form1.syear.selectedIndex == 0) msg += '請選擇年度\n';
            if (document.form1.RIDValue.value == '') msg += '請選擇機構\n';

            if (msg != '') {
                alert(msg);
                return false;
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練查核與績效管理&gt;&gt;公務統計報表&gt;&gt;年度各項訓練人數統計</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td width="20%" class="bluecol_need">年度</td>
                <td class="whitecol">
                    <asp:DropDownList ID="syear" runat="server"></asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">訓練機構</td>
                <td class="whitecol">
                    <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                    <input type="button" value="..." id="Button1" name="Button1" runat="server" class="button_b_Mini">
                    <input id="RIDValue" type="hidden" runat="server" size="1"><br />
                    <span id="HistoryList2" style="display: none; position: absolute">
                        <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                    </span>
                </td>
            </tr>
            <tr>
                <td class="bluecol">預算來源 </td>
                <td class="whitecol">
                    <asp:CheckBoxList ID="BudID" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                    </asp:CheckBoxList>
                </td>
            </tr>
            <%--<tr id="trCheckBox1" runat="server">
                <td class="bluecol">查詢結果</td>
                <td class="whitecol">
                    <asp:CheckBox ID="CheckBox1" runat="server" Text="包含下層單位"></asp:CheckBox>
                </td>
            </tr>--%>
            <tr>
                <td colspan="2" align="center">
                    <asp:Button ID="Button2" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
