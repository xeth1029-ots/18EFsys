<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="wPopUp1.aspx.vb" Inherits="WDAIIP.wPopUp1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>可正常使用彈出PopUp 視窗</title>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <style type="text/css">
        .auto-style1 { font-size: large; display: inline-block; }
    </style>
    <script type="text/javascript">
        function startWinClose9Sec() {
            // 按下 start 後 id 為 timer 的 DIV 內容可以開始倒數到到 0。 
            var timer =  document.querySelector("#timer");
            var number = 9;
            setInterval(function () {
                number--;
                if (number <= 0) {
                    number = 0;
                    setTimeout("window.opener=null;window.close()", 20);
                }
                timer.innerText = number + 0
            }, 1000);
        }
        $(document).ready(function () {
            startWinClose9Sec();
            //setTimeout("window.opener=null;window.close()", 9000);
        })
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <div class="auto-style1"><strong>可正常使用彈出PopUp 視窗</strong></div>
        <br />
        <div class="auto-style1">
            <div class="auto-style1" id="timer">9</div>
            秒後關閉視窗
        </div>
    </form>
</body>
</html>
