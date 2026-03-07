<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_04_001.aspx.vb" Inherits="WDAIIP.SYS_04_001" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>年度行事曆</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;系統參數管理&gt;&gt;年度行事曆</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch" id="Table3" cellpadding="1" cellspacing="1" width="100%">
                        <tr>
                            <td class="bluecol_need" style="width: 20%">日期 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="Start_Date" onfocus="this.blur()" runat="server" Width="15%" MaxLength="10"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('Start_Date','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">～
							<asp:TextBox ID="End_Date" Width="15%" runat="server" onfocus="this.blur()" MaxLength="10"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('End_Date','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                <asp:RequiredFieldValidator ID="MustSDate" runat="server" ErrorMessage="請輸入起日期" Display="None" ControlToValidate="Start_Date"></asp:RequiredFieldValidator>
                                <asp:RequiredFieldValidator ID="MustEDate" runat="server" ErrorMessage="請輸入迄日期" Display="None" ControlToValidate="End_Date"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td colspan="2" class="whitecol" align="center">
                                <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                &nbsp;<asp:Button ID="but_seach" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>&nbsp;<asp:Button ID="but_submit" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="2" class="whitecol">&nbsp;<asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" border="0" runat="server" width="100%">
                        <tr>
                            <td>
                                <asp:DataGrid ID="ShowHDay" runat="server" AutoGenerateColumns="False" CssClass="font" AllowPaging="True" Width="100%" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:TemplateColumn HeaderText="日期">
                                            <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="labHolDate" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="事由">
                                            <HeaderStyle HorizontalAlign="Center" Width="75%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="LabReason" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="功能" ItemStyle-HorizontalAlign="Center">
                                            <HeaderStyle Width="10%" />
                                            <ItemStyle HorizontalAlign="Center" Font-Size="Small" />
                                            <ItemTemplate>
                                                <asp:LinkButton ID="lbtDel" runat="server" Text="刪除" CommandName="Del" CssClass="linkbutton"></asp:LinkButton>
                                            </ItemTemplate>
                                            <ItemStyle HorizontalAlign="Center" Width="80px"></ItemStyle>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                            </td>
                        </tr>
                    </table>
                    <asp:ValidationSummary ID="Totalmsg" runat="server" DisplayMode="List" ShowMessageBox="True" ShowSummary="False"></asp:ValidationSummary>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
