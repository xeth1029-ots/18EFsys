<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_004_class.aspx.vb" Inherits="WDAIIP.TC_01_004_class" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>班別</title>
    <meta content="False" name="vs_showGrid">
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

        function returnValue(clsid, classid, cont, tmid, trainname) {
            opener.document.form1.clsid.value = clsid;
            opener.document.form1.TBclass_id.value = classid;
            opener.document.form1.TBclass.value = classid;
            opener.document.form1.ClassEngName.value = cont;
            opener.document.form1.ClassEng.value = cont;
            opener.document.form1.trainValue.value = tmid;
            opener.document.form1.TB_career_id.value = trainname;
            //window.close();
            self.close();
        }
    </script>
    <style type="text/css">
        .auto-style1 { color: Black; text-align: right; padding: 4px 6px; background-color: #f1f9fc; border-right: 3px solid #49cbef; height: 45px; }
        .auto-style2 { color: #333333; padding: 4px; height: 45px; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <asp:Panel ID="Panel2" Style="z-index: 102; left: 0px; top: 0px" runat="server" Width="100%">
            <table cellpadding="1" cellspacing="1" class="table_nw" width="100%">
                <tr>
                    <td class="bluecol" width="20%">訓練計畫：</td>
                    <td class="whitecol">
                        <asp:Label ID="ProecessType" runat="server" Width="100%" CssClass="font"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">關鍵字：</td>
                    <td class="whitecol">
                        <asp:TextBox ID="txtSearch1" runat="server" Width="40%"></asp:TextBox>
                        <asp:Button ID="btnSearch" runat="server" Text="搜尋" CssClass="asp_Export_M"></asp:Button>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">班別代碼：</td>
                    <td class="whitecol">
                        <asp:TextBox ID="txtSearch2" runat="server" Width="30%"></asp:TextBox>
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <div style="overflow-y: auto; height: 400px;">
            <asp:Panel ID="Panel" Style="z-index: 101; left: 0px; top: 200px" runat="server" Visible="true" Width="100%">
                <table class="font" id="search_tbl" cellspacing="0" cellpadding="0" width="100%" border="0" runat="server"></table>
                <input id="hbtnSearch" style="display: none" type="button" value="搜尋" name="hbtnSearch" runat="server" class="asp_Export_M">
            </asp:Panel>
        </div>
    </form>
</body>
</html>
