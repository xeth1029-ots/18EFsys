<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_006.aspx.vb" Inherits="WDAIIP.TC_01_006" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>師資別設定</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/common.js"></script>
    <script language="javascript">
        function chk() {
            var msg = '';
            //if (document.form1.DropDownList1.selectedIndex==0) msg+='請選擇內外聘\n';
            //if (document.form1.DropDownList2.selectedIndex==0) msg+='請選擇師資別總數\n';
            if (document.form1.TextBox2.value != '' && !isUnsignedInt(document.form1.TextBox2.value)) msg += '起始基本時數必須為數字\n';
            if (document.form1.TextBox3.value != '' && !isUnsignedInt(document.form1.TextBox3.value)) msg += '終至基本時數必須為數字\n';
            if (msg != '') {
                alert(msg);
                return false;
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;基本資料設定&gt;&gt;師資別設定</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table4" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_nw" id="Table2" width="100%" cellpadding="1" cellspacing="1">
                        <tr>
                            <td class="bluecol" width="20%">內外聘</td>
                            <td class="whitecol" width="30%">
                                <asp:DropDownList ID="DropDownList1" runat="server">
                                    <asp:ListItem Value="0">--請選擇--</asp:ListItem>
                                    <asp:ListItem Value="1">內</asp:ListItem>
                                    <asp:ListItem Value="2">外</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                            <td class="bluecol" width="20%">師資別總類</td>
                            <td class="whitecol" width="30%">
                                <asp:DropDownList ID="DropDownList2" runat="server">
                                    <asp:ListItem Value="0">--請選擇--</asp:ListItem>
                                    <asp:ListItem Value="1">訓練師類</asp:ListItem>
                                    <asp:ListItem Value="2">行政人員類</asp:ListItem>
                                    <asp:ListItem Value="3">外聘類</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">師資別名稱</td>
                            <td class="whitecol"><asp:TextBox ID="TextBox1" runat="server" Columns="30" Width="70%"></asp:TextBox></td>
                            <td class="bluecol">基本時數</td>
                            <td class="whitecol">
                                <asp:TextBox ID="TextBox2" runat="server" Width="30%"></asp:TextBox> ～
                                <asp:TextBox ID="TextBox3" runat="server" Width="30%"></asp:TextBox>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>&nbsp;
                                <asp:Button ID="Button2" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <br />
                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CellPadding="8"  CssClass="font" AutoGenerateColumns="False">
                        <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn HeaderText="序號">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="KindEngage" HeaderText="內外聘">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="CateKind" HeaderText="師資別種類">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="KindName" HeaderText="師資別名稱">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="BaseHours" HeaderText="基本時數">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="HightHours" HeaderText="最高請領時數">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="功能">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                                <ItemTemplate>
                                    <asp:LinkButton ID="lbtEdit" runat="server" Text="修改" CommandName="edit" CssClass="linkbutton"></asp:LinkButton>
                                    <asp:LinkButton ID="lbtDel" runat="server" Text="刪除" CommandName="del" CssClass="linkbutton"></asp:LinkButton>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                    </asp:DataGrid>
                    <p align="center"><asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label></p>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>