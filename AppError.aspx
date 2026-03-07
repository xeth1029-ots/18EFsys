<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="AppError.aspx.vb" Inherits="WDAIIP.AppError" EnableViewState="False" EnableSessionState="True" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>AppError</title>
    <link rel="Shortcut Icon" href="./css/wdalogo.ico" type="image/x-icon" />
    <style id="antiClickjack" type="text/css">
        html { display: none; }
        body { display: none !important; background-color: #ffffff; }
    </style>
    <script type="text/javascript">
        //解決不支援 X-Frame-Options設定，需額外判斷
        //if (top != self) { top.location = self.location; }
        if (self === top || self == top) {
            var antiClickjack = document.getElementById("antiClickjack");
            if (antiClickjack) { antiClickjack.parentNode.removeChild(antiClickjack); }
            document.documentElement.style.display = 'block';
        }
        else { top.location = self.location; }
        if (parent.document.frames != undefined && parent.document.frames.length != 0) {
            top.location.replace(self.location);
        }
    </script>
</head>
<body>
    <div>
        系統發生錯誤，請聯絡系統管理員!!!
    </div>
    <asp:Panel ID="errMsg" runat="server">
        <asp:Label ID="labExMessage" runat="server" Text=""></asp:Label>
    </asp:Panel>
    <asp:Panel ID="errStackTrace" runat="server">
        <asp:Label ID="labStackTrace" runat="server" Text=""></asp:Label>
        <% %>
    </asp:Panel>
    <br />
    <a target="_self" href="Index">回首頁</a>
</body>
</html>
