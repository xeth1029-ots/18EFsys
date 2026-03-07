<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_004_Classid.aspx.vb" Inherits="WDAIIP.TC_01_004_Classid" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>班別代碼</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function Search_click() {
            if (document.getElementById('txtSearch1') && document.getElementById('txtSearch2')) {
                if (event.keyCode == 13) {
                    document.getElementById('btnSearch').disabled = true;
                    document.getElementById('hbtnSearch').click();
                }
            }
        }

        function Search_click2() {
            document.getElementById('btnSearch').disabled = true;
            document.getElementById('hbtnSearch').click();
        }

        function returnValue() {
            if (isEmpty("classid")) {
                alert("請選擇班別代碼");
                return false;
            }
            var clsid = getValue("classid");
            //alert(clsid);
            if (clsid == "") {
                opener.document.form1.change.disabled = true;
                window.close();
            }
            else {
                opener.document.form1.clsid.value = clsid;
                opener.document.form1.change.disabled = false;
                window.close();
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <div id="divP2">
            <table class="font" width="100%" border="1" cellpadding="1" cellspacing="1">
                <tr>
                    <td class="bluecol" width="20%">訓練計畫： </td>
                    <td><asp:Label ID="ProecessType" runat="server" Width="100%" CssClass="font"></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol">關鍵字： </td>
                    <td class="whitecol">
                        <asp:TextBox ID="txtSearch1" runat="server" Width="40%"></asp:TextBox>
                        <asp:Button ID="btnSearch" runat="server" Text="搜尋" CssClass="asp_Export_M"></asp:Button>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">班別代碼： </td>
                    <td class="whitecol"><asp:TextBox ID="txtSearch2" runat="server" Width="30%"></asp:TextBox></td>
                </tr>
            </table>
        </div>
        <div id="divP1" align="center">
            <br /><table class="font" id="search_tbl" cellspacing="0" cellpadding="0" width="100%" border="1" runat="server"></table>
            <br /><input onclick="returnValue()" type="button" value="確定" class="asp_Export_M" />
            <input id="hbtnSearch" style="display: none" type="button" value="搜尋" name="hbtnSearch" runat="server" class="asp_button_M" />
        </div>
        <asp:HiddenField ID="Hid_PlanID" runat="server" />
        <asp:HiddenField ID="Hid_DistID" runat="server" />        
    </form>
</body>
</html>