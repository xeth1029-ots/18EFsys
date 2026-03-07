<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_02_007_R.aspx.vb" Inherits="WDAIIP.CP_02_007_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>學員異動月報表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function choose_class() {
            var RID = document.form1.RIDValue.value;
            window.open('../CP_01_ch.aspx?RID=' + RID, '', 'width=550,height=520,location=0,status=0,menubar=0,scrollbars=1,resizable=0');
        }
        function print() {
            var msg = '';

            if (document.form1.syear.selectedIndex == 0) msg += '請選擇年度\n';
            if (document.form1.smonth.selectedIndex == 0) msg += '請選擇月份\n';
            //if (document.form1.DistID.selectedIndex == 0) msg += '請選擇轄區中心\n';
            if (document.form1.DistID.selectedIndex == 0) msg += '請選擇轄區分署\n';
            if (document.form1.TPlanID.selectedIndex == 0) msg += '請選擇訓練計畫\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練查核與績效管理&gt;&gt;公務統計報表 &gt;&gt;學員異動月報表</asp:Label>
                </td>
            </tr>
        </table>

        <table id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td class="bluecol_need" style="width: 20%">統計月份
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="syear" runat="server">
                    </asp:DropDownList>
                    <font face="新細明體">年</font>
                    <asp:DropDownList ID="smonth" runat="server">
                    </asp:DropDownList>
                    <font face="新細明體">月</font>
                </td>
            </tr>
            <tr>
                <%--<td class="bluecol_need">轄區中心</td>--%>
                <td class="bluecol_need">轄區分署</td>
                <td class="whitecol">
                    <asp:DropDownList ID="DistID" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">訓練計畫
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="TPlanID" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">匯出檔案格式</td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                        <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                        <asp:ListItem Value="ODS">ODS</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td colspan="2" align="center" class="whitecol">
                    <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                    <asp:Button ID="btnExport1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                </td>
            </tr>
        </table>
        <div align="center">
            <asp:Label ID="msg1" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
        </div>
        <br />
    </form>
</body>
</html>
