<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_008_E.aspx.vb" Inherits="WDAIIP.SD_05_008_E" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>結訓學員資料匯出</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        function search() {
            var msg = '';
            if (isEmpty(document.form1.start_date) && isEmpty(document.form1.end_date)) {
                msg += '請選擇日期範圍!\n';
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
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <%-- <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">結訓學員資料匯出</font>
                        </td>
                    </tr>
                </table>--%>
                    <table class="table_sch" id="Table3" cellpadding="1" cellspacing="1">
                        <tr>
                            <td class="bluecol_need" style="width: 20%">結訓日期
                            </td>
                            <td class="whitecol" runat="server">
                                <asp:TextBox ID="start_date" runat="server" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                ～<asp:TextBox ID="end_date" runat="server" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            </td>
                        </tr>
                        <tr>
                            <%--<td class="bluecol">局/非局屬</td>--%>
                            <td class="bluecol">署/非署屬</td>
                            <td class="whitecol">
                                <%--<asp:ListItem Value="1" Selected="True">局屬</asp:ListItem>--%>
                                <%--<asp:ListItem Value="2">非局屬</asp:ListItem>--%>
                                <asp:RadioButtonList ID="Mode" runat="server" CssClass="font" RepeatDirection="Horizontal" AutoPostBack="True">
                                    <asp:ListItem Value="1" Selected="True">署屬</asp:ListItem>
                                    <asp:ListItem Value="2">非署屬</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練機構
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="DistID" runat="server" RepeatColumns="7" CssClass="font">
                                </asp:CheckBoxList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練計畫
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="TPlan" runat="server" RepeatColumns="3" CssClass="font">
                                </asp:CheckBoxList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">匯出檔案格式</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                                    <asp:ListItem Value="ODS">ODS</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>

                    </table>
                    <p align="center" class="whitecol">
                        <%--結訓學員資料匯出--%>
                        <asp:Button ID="Button1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                        <asp:Button ID="Button2" runat="server" Text="回查詢頁面" CssClass="asp_Export_M"></asp:Button>
                    </p>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
