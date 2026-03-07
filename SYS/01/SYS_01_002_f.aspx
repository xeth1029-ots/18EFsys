<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_01_002_f.aspx.vb" Inherits="WDAIIP.SYS_01_002_f" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>帳號計畫賦予-查詢</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function closeMe() {
            window.opener = null; window.open('', '_self'); window.close();
            return false;
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
            
            <tr>
                <td>
                    <table id="AccountTable" runat="server" class="table_sch" width="100%">
                        <tr>
                            <td class="bluecol" style="width: 20%">帳號</td>
                            <td colspan="3" class="whitecol">
                                <asp:DropDownList ID="Lis_acc" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol">年度</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="Years" runat="server" AutoPostBack="True"></asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4" class="whitecol">
                                <asp:DataGrid ID="DataGrid1" runat="server" HeaderStyle-CssClass="TD_TD1" Width="100%" AutoGenerateColumns="False" Visible="False" CssClass="font" CellPadding="8">
                                    <AlternatingItemStyle BackColor="WhiteSmoke" />
                                    <HeaderStyle CssClass="head_navy" HorizontalAlign="Center"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn DataField="PlanName" HeaderText="計畫代碼　(補助地方政府)" ItemStyle-Width="40%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構" ItemStyle-Width="35%"></asp:BoundColumn>
                                        <asp:TemplateColumn ItemStyle-Width="25%">
                                            <HeaderTemplate>
                                                備註
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr align="center">
                            <td class="whitecol" colspan="4">
                                <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></td>
                        </tr>
                        <tr align="center">
                            <td class="whitecol" colspan="4">
                                <asp:Button ID="btnClose" runat="server" Text="關閉" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td></td>
            </tr>
        </table>
    </form>
</body>
</html>
