<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_04_011.aspx.vb" Inherits="WDAIIP.SD_04_011" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>排課時數查詢</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;課程管理&gt;&gt;排課時數查詢</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td class="bluecol" width="20%">年度
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="yearlist" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">訓練機構
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                    <input id="Org" type="button" value="..." name="Org" runat="server" class="button_b_Mini">
                    <input id="RIDValue" style="width: 3%" type="hidden" name="RIDValue" runat="server">
                    <span id="HistoryList2" style="display: none; position: absolute">
                        <asp:Table ID="HistoryRID" runat="server" Width="100%">
                        </asp:Table>
                    </span>
                    <input id="TPlanid" style="width: 3%" type="hidden" name="TPlanid" runat="server">
                    <input id="Re_ID" style="width: 3%" type="hidden" name="Re_ID" runat="server">
                </td>
            </tr>
            <tr>
                <td class="bluecol">日期區間
                </td>
                <td class="whitecol">
                    <span id="span01" runat="server">
                        <asp:TextBox ID="start_date" runat="server" onfocus="this.blur()" Width="15%"></asp:TextBox>
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">～
                                <asp:TextBox ID="end_date" runat="server" onfocus="this.blur()" Width="15%"></asp:TextBox>
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                    </span>
                </td>
            </tr>
            <tr>
                <td class="whitecol" colspan="2">
                    <div align="center">
                        <asp:Label ID="labPageSize" runat="server" DESIGNTIMEDRAGDROP="30" ForeColor="SlateBlue">顯示列數</asp:Label><asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                        <asp:Button ID="btnQuery" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>&nbsp;
                    </div>
                </td>
            </tr>
            <tr>
                <td align="center" colspan="2">
                    <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="dtPlan" runat="server" Width="100%" CssClass="font" AllowSorting="True" PagerStyle-HorizontalAlign="Left" PagerStyle-Mode="NumericPages" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="WhiteSmoke" />
                                    <HeaderStyle CssClass="head_navy" HorizontalAlign="Center"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號">
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="機構名稱">
                                            <HeaderStyle HorizontalAlign="Left" Width="85%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="CsHours" HeaderText="排課時數">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundColumn>
                                    </Columns>
                                    <PagerStyle Visible="False" HorizontalAlign="Left" ForeColor="Blue" Position="Top" Mode="NumericPages"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                            </td>
                        </tr>
                    </table>
                    <br>
                    <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
                </td>
            </tr>
        </table>

        <%--
<tr><td colspan="2"></td></tr>
<table class="table_nw" cellpadding="1" cellspacing="1" width="100%">
</table>
<table width="100%">
</table>
<table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
</table>
        --%>
    </form>
</body>
</html>
