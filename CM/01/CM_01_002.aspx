<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CM_01_002.aspx.vb" Inherits="WDAIIP.CM_01_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>CM_01_002</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script src="../../js/common.js"></script>
    <script src="../../js/TIMS.js"></script>
    <script language="javascript">
        function show_dg() {
            //style="display:none;position:absolute; left: 476px; top: 96px;"
            if (document.getElementById('TableLay1').style.display == "none") {
                //document.getElementById('TableLay1').style.display = "inline";
                document.getElementById('TableLay1').style.display = "";
            } else {
                document.getElementById('TableLay1').style.display = "none";
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練經費控管&gt;&gt;歷史經費查詢</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch" id="Table3" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" style="width: 20%">年度</td>
                            <td class="whitecol"><asp:DropDownList ID="syear" runat="server" AutoPostBack="True"></asp:DropDownList></td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <asp:Panel ID="Panel" runat="server" Visible="False" Width="100%">
            <asp:DataGrid ID="DG_Grid1" runat="server" Width="100%" Visible="False" AutoGenerateColumns="False" AllowPaging="True" ShowFooter="True" CellPadding="8">
                <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                <FooterStyle HorizontalAlign="Right" BackColor="#E3F8FD"></FooterStyle>
                <Columns>
                    <asp:TemplateColumn HeaderText="轄區" FooterText="合計">
                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        <ItemTemplate>
                            <asp:LinkButton ID="LinkButton1" runat="server" ForeColor="Blue">LinkButton</asp:LinkButton>
                        </ItemTemplate>
                        <FooterStyle HorizontalAlign="Center"></FooterStyle>
                    </asp:TemplateColumn>
                    <asp:BoundColumn DataField="advance_class" HeaderText="年度預定開班數">
                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn HeaderText="年度總預算數(元)">
                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="real_class" HeaderText="實際開班數">
                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:TemplateColumn HeaderText="實際開班經費(元) ">
                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        <HeaderTemplate>
                            實際開班經費(元)<asp:HyperLink ID="HyperLink1" runat="server" ForeColor="Blue">細目</asp:HyperLink>
                        </HeaderTemplate>
                        <ItemTemplate>
                            <asp:Label ID="Label1" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.real_total") %>'>
                            </asp:Label>
                        </ItemTemplate>
                        <EditItemTemplate>
                            <asp:TextBox ID="TextBox1" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.real_total") %>'>
                            </asp:TextBox>
                        </EditItemTemplate>
                    </asp:TemplateColumn>
                    <asp:BoundColumn HeaderText="結餘金額(元)">
                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                    </asp:BoundColumn>
                </Columns>
                <PagerStyle Visible="False"></PagerStyle>
            </asp:DataGrid>
            <div align="left">實際開班數定義：包含在訓(已開訓)及結訓班數</div>
        </asp:Panel>
        <table id="TableLay1" style="display: none; left: 18%; position: absolute; top: 20%" cellspacing="1" cellpadding="0" border="0" width="60%">
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" ShowFooter="True" CssClass="font" AllowPaging="True" AutoGenerateColumns="False">
                        <AlternatingItemStyle BackColor="lemonchiffon"></AlternatingItemStyle>
                        <ItemStyle HorizontalAlign="Center" BackColor="White"></ItemStyle>
                        <HeaderStyle HorizontalAlign="Center" BackColor="PeachPuff"></HeaderStyle>
                        <FooterStyle HorizontalAlign="Right" BackColor="PeachPuff"></FooterStyle>
                        <Columns>
                            <asp:BoundColumn DataField="distname" HeaderText="轄區" FooterText="合計">
                                <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                <FooterStyle HorizontalAlign="Center"></FooterStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn HeaderText="實際開班已付金額(元)">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Right"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn HeaderText="應付未付金額(元)">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Right"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn Visible="False" DataField="real_total" HeaderText="實際總經費"></asp:BoundColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>