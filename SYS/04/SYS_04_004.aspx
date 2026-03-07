<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_04_004.aspx.vb" Inherits="WDAIIP.SYS_04_004" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>年度-訓練計畫-預算別</title>
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

        function SelectAll(flag) {
            const mytable = document.getElementById('DataGrid1');
            // Handle cases where the table or row doesn't exist
            if (!mytable) { return; }
            for (let i = 0; i < mytable.rows.length; i++) {
                //debugger;
                const MyCheckList = mytable.rows[i].children[0];
                for (let j = 0; j < MyCheckList.children.length; j++) {
                    const child = MyCheckList.children[j];
                    if (child && child.type === 'checkbox') { child.checked = flag; }
                }
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;系統參數管理&gt;&gt;年度-訓練計畫-預算別</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table3" class="table_sch" cellpadding="1" cellspacing="1">
            <tr>
                <td class="bluecol_need" style="width: 20%">年度</td>
                <td class="whitecol">
                    <asp:DropDownList ID="Syear" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">訓練計畫</td>
                <td class="whitecol">
                    <asp:DropDownList ID="TPlanID" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td align="center" colspan="2" class="whitecol">
                    <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
        </table>
        <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid1" runat="server" AutoGenerateColumns="False" CssClass="font" Width="100%" CellPadding="8">
                        <AlternatingItemStyle BackColor="WhiteSmoke" />
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:TemplateColumn HeaderText="全選">
                                <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <HeaderTemplate>
                                    全選<input id="Checkbox1" type="checkbox" runat="server" />
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <input id="Checkbox1c" type="checkbox" runat="server" />
                                    <asp:HiddenField ID="Hid_TPlanID" runat="server" />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn DataField="PlanName" HeaderText="訓練計畫">
                                <HeaderStyle Width="50%" />
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="預算別">
                                <HeaderStyle Width="45%" />
                                <ItemTemplate>
                                    <asp:CheckBoxList ID="BudID" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    </asp:CheckBoxList>
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
    </form>
</body>
</html>
