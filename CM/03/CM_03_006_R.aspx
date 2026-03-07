<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CM_03_006_R.aspx.vb" Inherits="WDAIIP.CM_03_006_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>CM_03_006_R</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript">
        function CheckPrint() {
            var msg = '';
            if (document.form1.yearlist.selectedIndex == 0) msg += '請選擇年度\n';
            if (document.form1.RIDValue.value == '') msg += '請選擇訓練機構\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        //選擇全部
        function SelectAll(obj, hidobj) {
            var num = getCheckBoxListValue(obj).length;
            var myallcheck = document.getElementById(obj + '_' + 0);

            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>首頁&gt;&gt;訓練需求管理&gt;&gt;統計分析&gt;&gt;<font color="#990000">受訓學員補助身份統計表</font> </td>
                        </tr>
                    </table>
                    <table class="table_sch" id="Table2" cellspacing="1" cellpadding="1" border="0">
                        <tr>
                            <td class="bluecol" width="100">年度 </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="yearlist" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練機構 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                <input id="Button8" type="button" value="..." name="Button5" runat="server">
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="310px">
                                    </asp:Table>
                                </span></td>
                        </tr>
                        <tr>
                            <td class="bluecol">縣市 </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="CityID" runat="server" RepeatColumns="8" RepeatDirection="Horizontal" CssClass="font">
                                </asp:CheckBoxList>
                                <input id="CityHidden" type="hidden" value="0" name="CityHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練計畫 </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="TPlanID" runat="server" RepeatColumns="3" RepeatDirection="Horizontal" CssClass="font">
                                </asp:CheckBoxList>
                                <input id="TPlanHidden" type="hidden" value="0" name="TPlanHidden" runat="server">
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td align="center" colspan="2">
                                <asp:Button ID="Print" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
