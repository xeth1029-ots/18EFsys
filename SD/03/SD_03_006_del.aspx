<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_03_006_del.aspx.vb" Inherits="TIMS.SD_03_006_del" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>刪除學員</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script>
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
<body ms_positioning="FlowLayout">
    <form id="form1" method="post" runat="server">
    <font face="新細明體">
        <table id="Table1" cellspacing="1" cellpadding="1" width="740" border="0" class="table_nw">
            <tr>
                <td width="100" class="bluecol">
                    學號：
                </td>
                <td class="whitecol">
                    <asp:Label ID="StudentID" runat="server"></asp:Label>
                </td>
            </tr>
            <tr>
                <td class="bluecol">
                    姓名：
                </td>
                <td class="whitecol">
                    <asp:Label ID="Name" runat="server"></asp:Label>
                </td>
            </tr>
            <tr>
                <td class="bluecol">
                    刪除學員原因：
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="DelResaon" runat="server">
                        <asp:ListItem Value="===請選擇===">===請選擇===</asp:ListItem>
                        <asp:ListItem Value="1">擅打錯誤</asp:ListItem>
                        <asp:ListItem Value="2">資格不符</asp:ListItem>
                        <asp:ListItem Value="3">其他</asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">
                    其他說明：
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="DelReasonOther" runat="server" TextMode="MultiLine" Rows="4" Columns="45"></asp:TextBox>
                </td>
            </tr>
        </table>
        <br />
        <table width="740">
            <tr>
                <td align="center">
                    <asp:Button ID="Button1" runat="server" Text="確定刪除" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
        </table>
    </font>
    </form>
</body>
</html>
