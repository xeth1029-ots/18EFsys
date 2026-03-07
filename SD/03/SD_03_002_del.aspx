<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_03_002_del.aspx.vb" Inherits="WDAIIP.SD_03_002_del" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>刪除學員</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript">
        function CheckData() {
            if (document.getElementById('DelResaon').selectedIndex == 0) {
                alert('請選擇原因');
                return false;
            }
            else if (document.getElementById('DelResaon').selectedIndex == 3 && document.getElementById('DelReasonOther').value == '') {
                alert('請輸入說明');
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
            <tr>
                <td class="bluecol" width="20%">學號：</td>
                <td class="whitecol" width="80%"><asp:Label ID="StudentID" runat="server"></asp:Label></td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">姓名：</td>
                <td class="whitecol"><asp:Label ID="Name" runat="server"></asp:Label></td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">刪除學員原因：</td>
                <td class="whitecol">
                    <asp:DropDownList ID="DelResaon" runat="server">
                        <asp:ListItem Value="">===請選擇===</asp:ListItem>
                        <asp:ListItem Value="1">擅打錯誤</asp:ListItem>
                        <asp:ListItem Value="2">資格不符</asp:ListItem>
                        <asp:ListItem Value="3">其他</asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">其他說明：</td>
                <td class="whitecol"><asp:TextBox ID="DelReasonOther" runat="server" TextMode="MultiLine" Rows="6" Columns="60"></asp:TextBox></td>
            </tr>
            <tr>
                <td align="center" colspan="2" class="whitecol"><asp:Button ID="Button1" runat="server" Text="確定刪除" CssClass="asp_button_M"></asp:Button></td>
            </tr>
        </table>
    </form>
</body>
</html>