<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="FAQ.aspx.vb" Inherits="WDAIIP.FAQ" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>FAQ</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="css/style.css" type="text/css" rel="stylesheet">
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
        <tr>
            <td>
                <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            首頁&gt;&gt;<font color="#800000"></F>問題集</font>
                        </td>
                    </tr>
                </table>
                <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
                            <asp:Button ID="bt_search" runat="server" Text="功能項目搜尋" CssClass="asp_button_M"></asp:Button>
                            <asp:Button ID="reset" runat="server" Text="重新設定" CssClass="asp_button_M"></asp:Button>
                        </td>
                    </tr>
                </table>
                <table class="font" id="Table5" cellspacing="0" cellpadding="0" width="100%" align="left" border="0">
                    <tr>
                        <td align="center">
                            <div align="center">
                                <asp:DataGrid ID="DataGrid1" runat="server" OnItemCommand="DataGrid1_ItemCommand" CssClass="font" Width="100%" AutoGenerateColumns="False">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號">
                                            <ItemStyle HorizontalAlign="Center" Width="10%"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="FunctionName" HeaderText="功能項目"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="FunctionCount" HeaderText="筆數">
                                            <HeaderStyle Width="50px"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="細節">
                                            <HeaderStyle Width="70px"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <font face="新細明體">
                                                    <asp:Button ID="Detail" runat="server" Text="細節" CommandName="FunctionDetail"></asp:Button></font>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid></div>
                            <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <div align="center">
                                <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label></div>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
