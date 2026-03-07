<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CM_01_002_Detail2.aspx.vb" Inherits="WDAIIP.CM_01_002_Detail2" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>CM_01_002_Detail2</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script src="../../js/common.js"></script>
    <script src="../../js/TIMS.js"></script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="bluecol" style="width: 20%">年度：<asp:Label ID="Label1" runat="server"></asp:Label></td>
                            <td class="bluecol" style="width: 20%">&nbsp;<asp:Label ID="Label2" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td colspan="2"></td>
                        </tr>
                        <tr>
                            <td align="left" colspan="2">
                                <asp:DataGrid ID="DG_Grid1" runat="server" PageSize="200" AllowPaging="True" ShowFooter="True" CssClass="font" Width="100%" AutoGenerateColumns="False" Visible="False" CellPadding="8">
                                    <AlternatingItemStyle HorizontalAlign="Center" BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <FooterStyle HorizontalAlign="Right" BackColor="#E3F8FD"></FooterStyle>
                                    <Columns>
                                        <asp:BoundColumn HeaderText="公務">
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <FooterStyle HorizontalAlign="Center"></FooterStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="就安">
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="就保">
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="訓練計畫" FooterText="合計">
                                            <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="LinkButton1" runat="server" ForeColor="Blue">LinkButton1</asp:LinkButton>
                                            </ItemTemplate>
                                            <FooterStyle HorizontalAlign="Center"></FooterStyle>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="advance_class" HeaderText="年度預定開班數">
                                            <HeaderStyle HorizontalAlign="Center" Width="13%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="年度總預算數(元)">
                                            <HeaderStyle HorizontalAlign="Center" Width="13%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="real_class" HeaderText="實際開班數">
                                            <HeaderStyle HorizontalAlign="Center" Width="13%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="實際開班經費(元)">
                                            <HeaderStyle HorizontalAlign="Center" Width="13%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="結餘金額(元)">
                                            <HeaderStyle HorizontalAlign="Center" Width="13%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                        </asp:BoundColumn>
                                    </Columns>
                                    <PagerStyle HorizontalAlign="Center" Position="Top"></PagerStyle>
                                </asp:DataGrid><br>
                                <asp:DataGrid ID="DataGrid1" runat="server" PageSize="200" AllowPaging="True" ShowFooter="True" CssClass="font" Width="100%" AutoGenerateColumns="False" Visible="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <FooterStyle HorizontalAlign="Right" BackColor="#E3F8FD"></FooterStyle>
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號">
                                            <HeaderStyle Width="4%" />
                                            <ItemStyle HorizontalAlign="Center" />
                                            <FooterStyle HorizontalAlign="Right"></FooterStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="訓練機構" FooterText="合計">
                                            <HeaderStyle Width="16%" />
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="LinkButton2" runat="server" ForeColor="Blue">LinkButton</asp:LinkButton>
                                            </ItemTemplate>
                                            <FooterStyle HorizontalAlign="Center"></FooterStyle>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="advance_class" HeaderText="年度預定開班數">
                                            <HeaderStyle HorizontalAlign="Center" Width="16%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="年度總預算數(元)">
                                            <HeaderStyle HorizontalAlign="Center" Width="16%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="real_class" HeaderText="實際開班數">
                                            <HeaderStyle HorizontalAlign="Center" Width="16%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="實際開班經費(元)">
                                            <HeaderStyle HorizontalAlign="Center" Width="16%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="結餘金額(元)">
                                            <HeaderStyle HorizontalAlign="Center" Width="16%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                        </asp:BoundColumn>
                                    </Columns>
                                    <PagerStyle Visible="False" HorizontalAlign="Center" Position="Top"></PagerStyle>
                                </asp:DataGrid><br>
                                <asp:DataGrid ID="DataGrid2" runat="server" PageSize="200" AllowPaging="True" ShowFooter="True" CssClass="font" Width="100%" AutoGenerateColumns="False" Visible="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <FooterStyle HorizontalAlign="Right" BackColor="#E3F8FD"></FooterStyle>
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號">
                                            <HeaderStyle HorizontalAlign="Center" Width="9%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassCName" HeaderText="班別名稱" FooterText="合計">
                                            <ItemStyle HorizontalAlign="Left" Width="13%"></ItemStyle>
                                            <FooterStyle HorizontalAlign="Center"></FooterStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="開結訓日" DataFormatString="{0:d}">
                                            <HeaderStyle HorizontalAlign="Center" Width="13%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="THours" HeaderText="訓練時數">
                                            <HeaderStyle HorizontalAlign="Center" Width="13%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="計畫人數/每人費用">
                                            <HeaderStyle HorizontalAlign="Center" Width="13%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn Visible="False" HeaderText="就安人數/金額">
                                            <HeaderStyle HorizontalAlign="Center" Width="13%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn Visible="False" HeaderText="就保人數/金額">
                                            <HeaderStyle HorizontalAlign="Center" Width="13%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn Visible="False" HeaderText="公務人數/金額">
                                            <HeaderStyle HorizontalAlign="Center" Width="13%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="計畫總經費">
                                            <HeaderStyle HorizontalAlign="Center" Width="13%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="已核銷總金額">
                                            <HeaderStyle HorizontalAlign="Center" Width="13%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="結餘總金額">
                                            <HeaderStyle HorizontalAlign="Center" Width="13%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                        </asp:BoundColumn>
                                    </Columns>
                                    <PagerStyle Visible="False" HorizontalAlign="Center" Position="Top"></PagerStyle>
                                </asp:DataGrid><br>
                                <div align="center" class="whitecol"><input language="javascript" id="Button1" type="button" value="回上一頁" name="Button1" runat="server" class="button_b_M"></div>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>