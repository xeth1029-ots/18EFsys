<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ojt1.aspx.vb" Inherits="WDAIIP.ojt1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title></title>
    <script type="text/javascript" src="../Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript">
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <span id="spanTIMENM1">timeTick</span>
        </div>
        <div>
            <asp:Label ID="labTITLE1" runat="server" Text="報名資料查詢-每分鐘報名人數(count1)"></asp:Label>
        </div>
        <asp:GridView ID="GridView1" runat="server"></asp:GridView>
    </form>
    <script type="text/javascript">
        function generateRandomNumber() {
            // Math.random() 會產生一個介於 0 (包含) 到 1 (不包含) 之間的浮點數。
            // (5000 - (-5000)) = 10000 // Math.random() * 10000 會產生一個 0 到 10000 之間的數。 // - 5000 讓範圍從 -5000 開始，變成 -5000 到 5000。
            return 30000 + (Math.random() * 30000);
        }
        $(document).ready(function () {
            // 當 DOM 載入完成後，執行這段 jQuery 程式碼
            const timeTick5 = generateRandomNumber();
            var spTIMENM1 = "timeTick: " + Math.trunc(timeTick5/1000).toString() + " sec..";
            document.getElementById("spanTIMENM1").textContent = spTIMENM1;
            setTimeout(function () {
                location.reload();
            }, timeTick5); // 5000 毫秒 = 5 秒
        });
    </script>
</body>
</html>
