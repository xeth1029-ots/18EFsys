<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="PS_01_001.aspx.vb" Inherits="WDAIIP.PS_01_001" %>

<!DOCTYPE html PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>常用表件設定</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;個人化設定&gt;&gt;常用表件設定</asp:Label>
                </td>
            </tr>
        </table>
        <table class="table_nw" id="FrameTable2" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>

                    <table class="font" id="tb_View" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <table class="table_nw" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td class="bluecol" align="center" width="20%">功能類型
                                        </td>
                                        <td class="whitecol">
                                            <asp:DropDownList ID="list_MainMenu2" runat="server" AutoPostBack="true"></asp:DropDownList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" align="center">功能項目類別
                                        </td>
                                        <td class="whitecol">
                                            <asp:DropDownList ID="list_MainMenu3" runat="server" AutoPostBack="true">
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                </table>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                          <asp:TemplateColumn HeaderText="功能類型">
                                            <ItemStyle Width="10%" BackColor="#f1f9fc"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_MainMenu2" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="功能項目類別">
                                            <ItemStyle Width="15%"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:HiddenField ID="Hid_FunID" runat="server" />
                                                <asp:Label ID="lab_MainMenu3" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="報表功能">
                                            <ItemStyle Width="25%"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_FunName" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="備註">
                                            <ItemStyle Width="25%"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="labMemo" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>

                                        <asp:TemplateColumn HeaderText="選用">
                                            <ItemStyle HorizontalAlign="Center" Width="25%"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:CheckBox ID="chk_Enable" runat="server" Text="選用"></asp:CheckBox>&nbsp;
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol" align="center" colspan="4">
                                <asp:Button ID="btn_Save2" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                                <%--<input id="Count" style="width: 32px; height: 22px" type="hidden" name="Re_ID" runat="server">--%>
                            </td>
                        </tr>
                    </table>

                </td>
            </tr>
        </table>
    </form>
</body>
</html>
