<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="sch.aspx.vb" Inherits="WDAIIP.sch" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head runat="server">
    <title>功能搜尋</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="./css/style.css">
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;功能搜尋</asp:Label>
                </td>
            </tr>
        </table>
        <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td align="center">
                    <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
                        <tr>
                            <td width="20%" class="bluecol">功能類別</td>
                            <td width="30%" class="whitecol"><asp:DropDownList ID="ddlFun" runat="server"></asp:DropDownList></td>
                            <td width="20%" class="bluecol">功能名稱</td>
                            <td width="30%" class="whitecol"><asp:TextBox ID="txtFunName" runat="server" MaxLength="20" Width="80%"></asp:TextBox></td>
                        </tr>
                    </table>
                    <br style="line-height: 15px" />
                    <asp:Button ID="btnSch" runat="server" Text="送出" class="asp_button_M" />
                    <br style="line-height: 10px" />
                    <br style="line-height: 10px" />
                    <asp:Label ID="labMsg" Style="color: red" runat="server">查無資料</asp:Label>
                </td>
            </tr>
            <tr>
                <td>
                    <br />
                    <asp:GridView ID="GridView1" runat="server" AllowPaging="false" Width="100%" AutoGenerateColumns="false" BorderWidth="0" CellPadding="0" CellSpacing="0" CssClass="font">
                        <RowStyle BackColor="White" Height="36px" />
                        <AlternatingRowStyle BackColor="#F5F5F5" Height="36px" />
                        <HeaderStyle CssClass="head_navy" />
                        <Columns>
                            <asp:TemplateField HeaderText="序號" ItemStyle-Width="8%" ItemStyle-HorizontalAlign="Center">
                                <ItemTemplate>
                                    <asp:Label ID="lbFunID" runat="server" Text='<%# DataBinder.Eval(Container.DataItem, "FunID") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="功能名稱" ItemStyle-Width="42%" ItemStyle-HorizontalAlign="left">
                                <ItemTemplate>
                                    &nbsp;&nbsp;<asp:Label ID="lbaName" runat="server" Text='<%# DataBinder.Eval(Container.DataItem, "aName") %>'></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="功能位置" ItemStyle-Width="50%" ItemStyle-HorizontalAlign="left">
                                <ItemTemplate>
                                    &nbsp;&nbsp;<asp:LinkButton ID="lbsysKindName" CommandName="viewcmd" CssClass="LinkSpan" runat="server" Text='<%# DataBinder.Eval(Container.DataItem, "sysKindName") %>'></asp:LinkButton>
                                </ItemTemplate>
                            </asp:TemplateField>
                        </Columns>
                    </asp:GridView>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>