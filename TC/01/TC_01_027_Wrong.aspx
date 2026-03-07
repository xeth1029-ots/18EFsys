<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_027_Wrong.aspx.vb" Inherits="WDAIIP.TC_01_027_Wrong" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>TC_01_027_Wrong</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid1" runat="server" AllowPaging="True" CssClass="font" AutoGenerateColumns="False" Width="100%">
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn DataField="Index" HeaderText=" 第幾筆錯誤" HeaderStyle-Width="20%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="TeacherID" HeaderText="講師代碼" HeaderStyle-Width="20%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="Name" HeaderText="講師姓名" HeaderStyle-Width="20%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="IDNO" HeaderText="身分證號碼" HeaderStyle-Width="20%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="Reason" HeaderText="原因" HeaderStyle-Width="20%"></asp:BoundColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
