<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_05_002_R.aspx.vb" Inherits="WDAIIP.CP_05_002_R" %>

 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>CP_05_002_R</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script src="../../js/common.js"></script>
    <script language="javascript">
        function CheckPrint() {
            var msg = '';
            var STYear = document.form1.STDate.value.substring(0, 4);
            var FTYear = document.form1.FTDate.value.substring(0, 4);

            if (document.form1.yearlist.value == '') msg += '請選擇年度\n';
            if (document.form1.planlist.value == '') msg += '請選擇訓練計畫\n';

            if (document.form1.STDate.value == '') msg += '開訓日期不能是空白\n';
            else if (!IsDate(document.form1.STDate.value)) msg += '開訓日期不是正確的日期格式\n';
            else if (STYear != document.form1.yearlist.value) msg += '開訓日期的年度跟所選年度要相同\n';

            if (document.form1.FTDate.value == '') msg += '結訓日期不能是空白\n';
            else if (!IsDate(document.form1.FTDate.value)) msg += '結訓日期不是正確的日期格式\n';
            else if (FTYear != document.form1.yearlist.value) msg += '結訓日期的年度跟所選年度要相同\n';

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
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table class="font" width="740">
        <tr>
            <td class="font">
                首頁&gt;&gt;訓練查核與績效管理&gt;&gt;專案計畫查核&gt;&gt;<font color="#990000">專案計畫查核</font>
            </td>
        </tr>
    </table>
    <br>
    <table class="table_nw" cellspacing="1" cellpadding="1" width="740" border="0">
        <tr>
            <td id="td6" width="100" runat="server" class="bluecol_need">
                年度
            </td>
            <td class="whitecol">
                <asp:DropDownList ID="yearlist" runat="server">
                </asp:DropDownList>
            </td>
            <td id="td7" width="100" runat="server" class="bluecol_need">
                訓練計畫
            </td>
            <td class="whitecol">
                <asp:DropDownList ID="planlist" runat="server">
                </asp:DropDownList>
            </td>
        </tr>
        <tr>
            <td width="100" class="bluecol_need">
                開訓期間
            </td>
            <td colspan="4" class="whitecol">
                <asp:TextBox ID="STDate" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
            </td>
        </tr>
        <tr>
            <td width="100" class="bluecol_need">
                結訓期間
            </td>
            <td colspan="4" class="whitecol">
                <asp:TextBox ID="FTDate" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
            </td>
        </tr>
    </table>
    <table width="740">
        <tr>
            <td>
                <div align="center">
                    <asp:Button ID="Print" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button></div>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
