<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_04_015_R.aspx.vb" Inherits="WDAIIP.TR_04_015_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>TR_04_015_R</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script language="javascript" src="../../js/common.js"></script>
    <script>
        function print() {
            var msg = '';
            //if (document.form1.DistID.selectedIndex == 0) msg += '請選擇轄區中心\n';
            if (document.form1.DistID.selectedIndex == 0) msg += '請選擇轄區分署\n';
            if (document.form1.TPlanID.selectedIndex == 0) msg += '請選擇訓練計畫\n';
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
    <table id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
        <tr>
            <td>
                <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            首頁&gt;&gt;訓練與就業需求管理&gt;&gt;就業追蹤報表&gt;&gt;<font color="#990000">受評單位自評及評鑑委員審核表-就業表現指標</font>
                        </td>
                    </tr>
                </table>
                <table class="table_sch" id="Table2" cellspacing="1" cellpadding="1">
                    <tr>
                        <%--<td class="bluecol_need" width="80">轄區中心</td>--%>
                        <td class="bluecol_need" width="80">轄區分署</td>
                        <td bgcolor="#ecf7ff" colspan="3" class="whitecol">
                            <asp:DropDownList ID="DistID" runat="server">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol_need">
                            訓練計畫
                        </td>
                        <td bgcolor="#ecf7ff" colspan="3" class="whitecol">
                            <asp:DropDownList ID="TPlanID" runat="server">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            結訓期間
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="FTDate1" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            <font color="#000000">～
                                <asp:TextBox ID="FTDate2" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></font>
                        </td>
                    </tr>
                </table>
                <p align="center">
                    <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                </p>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
