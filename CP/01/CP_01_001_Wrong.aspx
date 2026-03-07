<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_01_001_Wrong.aspx.vb" Inherits="WDAIIP.CP_01_001_Wrong" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>匯入 錯誤資料</title>
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
        <table id="Wrong_TABLE1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid1" runat="server" AutoGenerateColumns="false" Width="100%" CssClass="font" AllowPaging="true">
                        <AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
                        <HeaderStyle CssClass="head_navy" />
                        <Columns>
                            <asp:BoundColumn DataField="Index" HeaderText="第幾筆錯誤"></asp:BoundColumn>
                            <asp:BoundColumn DataField="OCID" HeaderText="班級代碼"></asp:BoundColumn>
                            <asp:BoundColumn DataField="ApplyDate" HeaderText="抽訪日期"></asp:BoundColumn>
                            <asp:BoundColumn DataField="VisitorName" HeaderText="訪查人員"></asp:BoundColumn>
                            <asp:BoundColumn DataField="Reason" HeaderText="原因"></asp:BoundColumn>
                        </Columns>
                        <PagerStyle Visible="false"></PagerStyle>
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
