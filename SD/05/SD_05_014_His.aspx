<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_014_His.aspx.vb" Inherits="WDAIIP.SD_05_014_His" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>學員出缺勤作業(產投)</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
        <tr>
            <td class="bluecol" width="20%">
                姓名
            </td>
            <td class="whitecol" width="30%">
                <asp:Label ID="Name" runat="server"></asp:Label>
            </td>
            <td class="bluecol" width="20%">
                學號
            </td>
            <td class="whitecol" width="30%">
                <asp:Label ID="StudentID" runat="server"></asp:Label>
            </td>
        </tr>
        <tr>
            <td colspan="4" align="center">
                <asp:DataGrid ID="DataGrid1" runat="server" AutoGenerateColumns="False" CssClass="font" Width="100%" CellPadding="8">
                    <HeaderStyle CssClass="head_navy" Width="50%"></HeaderStyle>
                    <AlternatingItemStyle BackColor="#F5F5F5" />
                    <Columns>
                        <asp:BoundColumn DataField="LeaveDate" HeaderText="未出席日期" DataFormatString="{0:d}"></asp:BoundColumn>
                        <asp:BoundColumn DataField="Hours" HeaderText="請假時數"></asp:BoundColumn>
                    </Columns>
                </asp:DataGrid>
                <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
            </td>
        </tr>
        <tr>
            <td align="center" colspan="4" class="whitecol">
                <input type="button" value="關閉視窗" onclick="window.close();" class="asp_button_M">
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
