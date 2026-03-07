<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_01_002.aspx.vb" Inherits="WDAIIP.TR_01_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>TR_01_002</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script src="../../js/common.js"></script>
    <script src="../../js/TIMS.js"></script>
    <script>
        function TicketChange() {
            if (document.form1.TicketMode.selectedIndex == 1) {
                document.getElementById('TICKET_TYPE').style.display = 'inline';
            }
            else {
                document.getElementById('TICKET_TYPE').style.display = 'none';
            }
        }
        function check_data() {
            if (document.form1.TicketMode.selectedIndex == 0) {
                alert('請選擇券別');
                return false;
            }
        }

    </script>
</head>
<body onload="FrameLoad();">
    <form id="form1" method="post" runat="server">
    <font face="新細明體">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                首頁&gt;&gt;訓練與需求管理&gt;&gt;開券資料統計&gt;&gt;<font color="#990000">依對象核券人數統計</font>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch" id="Table2" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" width="80">
                                券別種類
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="TicketMode" runat="server">
                                    <asp:ListItem Value="===請選擇===">===請選擇===</asp:ListItem>
                                    <asp:ListItem Value="職訓券">職訓券</asp:ListItem>
                                    <asp:ListItem Value="學習券">學習券</asp:ListItem>
                                    <asp:ListItem Value="推介單">推介單</asp:ListItem>
                                </asp:DropDownList>
                                <asp:RadioButtonList ID="TICKET_TYPE" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="不區分" Selected="True">不區分</asp:ListItem>
                                    <asp:ListItem Value="1">甲式</asp:ListItem>
                                    <asp:ListItem Value="2">乙式</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                查詢期間
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="SYear" runat="server">
                                </asp:DropDownList>
                                年<asp:DropDownList ID="SMonth" runat="server">
                                </asp:DropDownList>
                                月～
                                <asp:DropDownList ID="FYear" runat="server">
                                </asp:DropDownList>
                                年<asp:DropDownList ID="FMonth" runat="server">
                                </asp:DropDownList>
                                月
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                就服中心
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="Station" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="3">
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                    <p align="center">
                        <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>
                    </p>
                </td>
            </tr>
        </table>
        <table class="font" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Table ID="RecordTable" runat="server" CssClass="font" CellPadding="3" CellSpacing="0" BorderWidth="1px">
                    </asp:Table>
                </td>
            </tr>
        </table>
    </font>
    </form>
</body>
</html>
