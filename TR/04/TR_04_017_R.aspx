<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_04_017_R.aspx.vb" Inherits="WDAIIP.TR_04_017_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>自辦職前就業追蹤成果統計表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script language="javascript" src="../../js/common.js"></script>
    <script>
        function search() {
            var msg = '';
            if (!IsDate(document.form1.STDate1.value)) msg += '開訓日期的起始日不是正確的日期格式\n';
            if (!IsDate(document.form1.STDate2.value)) msg += '開訓日期的結束日不是正確的日期格式\n';
            if (!IsDate(document.form1.FTDate1.value)) msg += '結訓日期的起始日不是正確的日期格式\n';
            if (!IsDate(document.form1.FTDate2.value)) msg += '結訓日期的結束日不是正確的日期格式\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function IsDate(MyDate) {
            if (MyDate != '') {
                if (!checkDate(MyDate))
                    return false;
            }
            return true;
        }

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
                <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            首頁&gt;&gt;訓練與就業需求管理&gt;&gt;就業追蹤報表&gt;&gt;<font color="#990000">自辦職前就業追蹤成果統計表</font>
                        </td>
                    </tr>
                </table>
                <table class="table_sch" id="SearchTable" runat="server" cellspacing="1" cellpadding="1">
                    <tr>
                        <td class="bluecol">
                            轄區
                        </td>
                        <td class="whitecol">
                            <asp:CheckBoxList ID="DistID" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                            </asp:CheckBoxList>
                            <input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server">
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            開訓期間
                        </td>
                        <td bgcolor="#ecf7ff" colspan="3" class="whitecol">
                            <asp:TextBox ID="STDate1" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"><font color="#000000">～</font><font color="#ffffff"> </font>
                            <asp:TextBox ID="STDate2" runat="server" Columns="10"></asp:TextBox><font color="#ffffff"><img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            </font>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            結訓日期
                        </td>
                        <td bgcolor="#ecf7ff" colspan="3" class="whitecol">
                            <asp:TextBox ID="FTDate1" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"><font color="#000000">～</font><font color="#ffffff"> </font>
                            <asp:TextBox ID="FTDate2" runat="server" Columns="10"></asp:TextBox><font color="#ffffff"><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            </font>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            報表類型
                        </td>
                        <td bgcolor="#ecf7ff" colspan="3" class="whitecol">
                            <asp:RadioButtonList ID="PrintStyle" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                <asp:ListItem Value="1" Selected="True">依性別、年齡、教育程度 </asp:ListItem>
                                <asp:ListItem Value="2">依身分別</asp:ListItem>
                                <asp:ListItem Value="3">依訓練職類</asp:ListItem>
                            </asp:RadioButtonList>
                        </td>
                    </tr>
                </table>
                <p align="center">
                    <asp:Button ID="Print" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button></p>
                <p align="center">
                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></p>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
