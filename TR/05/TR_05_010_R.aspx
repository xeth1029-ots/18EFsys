<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_05_010_R.aspx.vb" Inherits="WDAIIP.TR_05_010_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>TR_05_010_R</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript">

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


        function search() {
            var msg = '';
            if (document.form1.STDate1.value != '') {
                if (!checkDate(document.form1.STDate1.value)) msg += '開訓日期的起始日不是正確的日期格式\n';
            }

            if (document.form1.STDate2.value != '') {
                if (!checkDate(document.form1.STDate2.value)) msg += '開訓日期的結束日不是正確的日期格式\n';
            }


            if (document.form1.FTDate1.value != '') {
                if (!checkDate(document.form1.FTDate1.value)) msg += '結訓日期的起始日不是正確的日期格式\n';
            }

            if (document.form1.FTDate2.value != '') {
                if (!checkDate(document.form1.FTDate2.value)) msg += '結訓日期的結束日不是正確的日期格式\n';
            }

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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練需求管理&gt;&gt;統計分析&gt;&gt;綜合動態報表</asp:Label>
                    <%--首頁&gt;&gt;訓練與就業需求管理&gt;&gt;統計分析&gt;&gt;<FONT color="#990000">訓練時數統計分析</FONT>--%>
                </td>
            </tr>
        </table>

        <table id="FrameTable3" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td class="bluecol" width="20%">動態報表 </td>
                <td class="whitecol" width="80%">
                    <uc1:WUC2 runat="server" ID="WUC2" />
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="100">年度
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="Syear" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">轄區
                </td>
                <td class="whitecol">
                    <asp:CheckBoxList ID="DistID" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                    </asp:CheckBoxList>
                    <input id="DistHidden" type="hidden" value="0" runat="server" name="DistHidden">
                </td>
            </tr>
            <tr>
                <td class="bluecol">訓練計畫
                </td>
                <td class="whitecol">
                    <asp:CheckBoxList ID="TPlanID" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3">
                    </asp:CheckBoxList>
                    <input id="TPlanHidden" type="hidden" value="0" runat="server" name="TPlanHidden">
                </td>
            </tr>
            <tr>
                <td class="bluecol">開訓期間
                </td>
                <td class="whitecol" runat="server">
                    <asp:TextBox ID="STDate1" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                    ～
                    <asp:TextBox ID="STDate2" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                </td>
            </tr>
            <tr>
                <td class="bluecol">結訓期間
                </td>
                <td class="whitecol" runat="server">
                    <asp:TextBox ID="FTDate1" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                    ～
					<asp:TextBox ID="FTDate2" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                </td>
            </tr>
            <tr>
                <td class="bluecol">報表格式
                </td>
                <td bgcolor="#ecf7ff" colspan="3" class="whitecol">
                    <asp:RadioButtonList ID="PrintStyle" runat="server" RepeatDirection="Horizontal" CssClass="font">
                        <asp:ListItem Value="2009">舊格式(2009年度含之前)</asp:ListItem>
                        <asp:ListItem Value="2010" Selected="True">新格式(2010年度含之後)</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td colspan="2" align="center">
                    <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>

                </td>
            </tr>
        </table>
    </form>
</body>
</html>
