<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_03_002_img.aspx.vb" Inherits="WDAIIP.SD_03_002_img" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD html 4.0 Transitional//EN">
<html>
<head>
    <title>學員資料維護</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1" />
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1" />
    <meta name="vs_defaultClientScript" content="JavaScript" />
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5" />
    <link rel="stylesheet" type="text/css" href="../../css/style.css" />
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        
        <table id="ImageShoeTable1" runat="server" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td align="center" class="whitecol">
                    <asp:Label ID="LabMsg1" runat="server" ForeColor="Red"></asp:Label>
                </td>
            </tr>
             <tr>
                <td align="center" class="whitecol">
                    <asp:Image ID="Image2" runat="server" Width="600px" Height="400px" />
                </td>
            </tr>
        </table>
          <table id="ButtonTable4" runat="server" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td align="center" class="whitecol">
                    <asp:Button ID="Button5" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="Hid_ERRMSG1" runat="server" />
    </form>   
</body>
</html>
