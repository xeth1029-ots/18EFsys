<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_03_004_Rc.aspx.vb" Inherits="TIMS.SD_03_004_Rc" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_03_004_Rc</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script src="../../js/common.js"></script>
</head>
<body ms_positioning="FlowLayout">
    <form id="form1" method="post" runat="server">
    <font face="新細明體">
        <table id="Table1" cellspacing="1" cellpadding="1" width="300" border="0">
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font">
                        <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:TemplateColumn>
                                <HeaderStyle HorizontalAlign="Center" Width="15px"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <input id="Radio1" type="radio" value="Radio1" runat="server">
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn DataField="ClassCName" HeaderText="班別名稱"></asp:BoundColumn>
                            <asp:BoundColumn Visible="False" DataField="CyclType" HeaderText="CyclType"></asp:BoundColumn>
                            <asp:BoundColumn Visible="False" DataField="LevelType" HeaderText="LevelType"></asp:BoundColumn>
                        </Columns>
                    </asp:DataGrid>
                    <p align="center">
                        <asp:Button ID="Button1" runat="server" Text="送出" CssClass="asp_button_S"></asp:Button></p>
                </td>
            </tr>
        </table>
    </font>
    </form>
</body>
</html>
