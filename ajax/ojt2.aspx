<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ojt2.aspx.vb" Inherits="WDAIIP.ojt2" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:TextBox ID="TextBox1" runat="server" Width="88%" Rows="6" TextMode="MultiLine"></asp:TextBox>
            <br />
        </div>
        <div>
            <asp:Button ID="Button2" runat="server" Text="下載測試" />
        </div>
        <div>
            <asp:Label ID="labResult" runat="server" Text=""></asp:Label>
        </div>
    </form>
</body>
</html>
