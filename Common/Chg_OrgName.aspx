<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Chg_OrgName.aspx.vb" Inherits="WDAIIP.Chg_OrgName" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>變更機構名稱</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta http-equiv="Content-Type" content="text/html; charset=big5">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        function checkData() {
            var msg = '';
            var New_OrgName = document.getElementById('New_OrgName');
            if (New_OrgName.value == '') msg += '請輸入欲變更機構名稱\n';
            //else if(!checkDate(document.getElementById('New_OrgName').value)) msg+='申請日期必須為正確的時間格式\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
</head>
<body>
    <!-- <div style="OVERFLOW: scroll;  WIDTH: 650px;  HEIGHT: 500px  ;scrollbar-base-color :#cccfff"> -->
    <form id="form1" method="post" runat="server">
        <table class="table_nw" id="supershow" cellspacing="1" cellpadding="1" border="0" runat="server" width="100%">
            <tr>
                <td class="bluecol" width="20%">
                    <asp:Label ID="Label1" runat="server">原機構名稱</asp:Label></td>
                <td class="whitecol" width="80%">
                    <asp:TextBox ID="Old_OrgName" runat="server" Enabled="False" onfocus="this.blur()" Width="70%"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="bluecol">
                    <asp:Label ID="Label2" runat="server">異動後機構名稱</asp:Label></td>
                <td class="whitecol">
                    <asp:TextBox ID="New_OrgName" runat="server" Width="70%" MaxLength="100"></asp:TextBox></td>
            </tr>
            <tr>
                <td align="center" colspan="2" class="whitecol">
                    <asp:Button ID="bt_save" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                    <asp:Button ID="Cancel" runat="server" Text="取消" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
            <tr>
                <td colspan="2">
                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CellPadding="8" BorderStyle="None" AutoGenerateColumns="False">
                        <%--<FooterStyle  ForeColor="#000066" BackColor="White"></FooterStyle>--%>
                        <%--<SelectedItemStyle  Font-Bold="True" ForeColor="White" BackColor="#669999"></SelectedItemStyle>--%>
                        <EditItemStyle></EditItemStyle>
                        <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                        <%--<ItemStyle  ForeColor="#000066"></ItemStyle>--%>
                        <HeaderStyle Font-Bold="True" CssClass="head_navy" Width="100%"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn Visible="False" DataField="EditID"></asp:BoundColumn>
                            <asp:BoundColumn DataField="Years" HeaderText="年度" HeaderStyle-Width="10%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="OrgName" HeaderText="機構名稱" HeaderStyle-Width="50%">
                                <%--<HeaderStyle Width="100px"></HeaderStyle>--%>
                                <ItemStyle HorizontalAlign="Left"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="ModifyDate" HeaderText="修改日期" HeaderStyle-Width="10%">
                                <ItemStyle HorizontalAlign="Left"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="ModifyName" HeaderText="異動者" HeaderStyle-Width="10%">
                                <ItemStyle HorizontalAlign="Left"></ItemStyle>
                            </asp:BoundColumn>
                            <%--<asp:TemplateColumn Visible="False" HeaderText="功能">
                                <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Button ID="restore" runat="server" Text="還原" CommandName="restore"></asp:Button>
                                </ItemTemplate>
                            </asp:TemplateColumn>--%>
                        </Columns>
                        <PagerStyle HorizontalAlign="Left" ForeColor="#000066" BackColor="White" Mode="NumericPages"></PagerStyle>
                    </asp:DataGrid>
                </td>
            </tr>
            <tr>
                <td align="center" colspan="2" style="color: red">
                    <asp:Label ID="msg_1" runat="server"></asp:Label></td>
            </tr>
        </table>
        <asp:HiddenField ID="Hid_orgid" runat="server" />
    </form>
    <!--</div>-->
</body>
</html>