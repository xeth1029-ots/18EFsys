<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_05_001.aspx.vb" Inherits="WDAIIP.SYS_05_001" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>上稿維護-FAQ</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;上稿維護&gt;&gt;FAQ</asp:Label>
                </td>
            </tr>
        </table>
        <div></div>
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
            <tr>
                <td>
                    <table id="Table4" class="table_sch" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" style="width: 20%">功能項目
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="FunctionNameList" runat="server">
                                </asp:DropDownList>
                                <asp:RequiredFieldValidator ID="MustFunctionName" runat="server" Display="Dynamic" ControlToValidate="FunctionNameList" ErrorMessage="請選擇功能項目"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td align="center" colspan="2" class="whitecol">
                                <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue" DESIGNTIMEDRAGDROP="30">顯示列數</asp:Label>
                                <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                <asp:Button ID="bt_search" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>&nbsp;
                                <asp:Button ID="bt_add" runat="server" Text="新增" CssClass="asp_button_M" CausesValidation="False"></asp:Button>&nbsp;
                                <asp:Button ID="reset" runat="server" Text="重新設定" CssClass="asp_button_M" CausesValidation="False"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <table id="Table5" cellspacing="0" cellpadding="0" width="100%" align="left" border="0" class="font">
                        <tr>
                            <td align="center">
                                <asp:DataGrid ID="DataGrid1" runat="server" AllowPaging="True" AutoGenerateColumns="False" Width="100%" CssClass="font" OnItemCommand="DataGrid1_ItemCommand">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號">
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="FunctionName" HeaderText="功能項目">
                                            <HeaderStyle Width="85%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="FunctionCount" HeaderText="筆數">
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="細節">
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="lbtDetail" runat="server" Text="細節" CommandName="FunctionDetail" CssClass="linkbutton"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn Visible="False" DataField="FunID" HeaderText="功能代碼"></asp:BoundColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
