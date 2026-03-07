<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_04_005.aspx.vb" Inherits="WDAIIP.SYS_04_005" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>計價種類設定作業</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript">
        function check_search() {
            if (document.form1.Syear.selectedIndex == 0) {
                alert('請選擇年度');
                return false;
            }
        }
        function DropChange() {
            document.getElementById('DataGridTable').style.display = 'none';
        }
        function SelectAll(flag, num) {
            var mytable = document.getElementById('DataGrid1');
            var MyCheckList = mytable.rows(num).cells(2).children(0);
            for (var i = 0; i < MyCheckList.children.length; i = i + 2) {
                MyCheckList.children(i).checked = flag;
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;系統參數管理&gt;&gt;計價種類設定作業</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
            <tr>
                <td>
                    <%--<table id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
					<tr>
						<td>
							首頁&gt;&gt;系統管理&gt;&gt;系統參數管理&gt;&gt;<font color="#990000">計價種類設定作業</font>
						</td>
					</tr>
				</table>--%>
                    <table id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
                        <tr>
                            <td colspan="4" align="center">
                                <font face="新細明體">
                                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                        <tr>
                                            <td>
                                                <asp:DataGrid ID="DataGrid1" runat="server" AutoGenerateColumns="False" CssClass="font" Width="100%" CellPadding="8">
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <Columns>
                                                        <asp:BoundColumn DataField="PlanName" HeaderText="訓練計畫" HeaderStyle-Width="30%"></asp:BoundColumn>
                                                        <asp:TemplateColumn HeaderText="預算種類" HeaderStyle-Width="70%">
                                                            <ItemTemplate>
                                                                <asp:RadioButtonList ID="CateNo" runat="server" RepeatColumns="2" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                                                    <asp:ListItem Value="1">成本加工費法</asp:ListItem>
                                                                    <asp:ListItem Value="2">每人每時單價計價法</asp:ListItem>
                                                                    <asp:ListItem Value="3">每人輔助單價計價法</asp:ListItem>
                                                                    <asp:ListItem Value="4">個人單價計價法</asp:ListItem>
                                                                </asp:RadioButtonList>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                    </Columns>
                                                </asp:DataGrid>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td align="center" class="whitecol">
                                                <asp:Button ID="Button2" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                                            </td>
                                        </tr>
                                    </table>
                                </font>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
