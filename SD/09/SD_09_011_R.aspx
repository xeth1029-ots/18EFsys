<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_09_011_R.aspx.vb" Inherits="WDAIIP.SD_09_011_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>職前訓練月報表</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1" />
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1" />
    <meta name="vs_defaultClientScript" content="JavaScript" />
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
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
        function print() {
            var msg = '';
            //if (document.form1.DistID.selectedIndex == 0) msg += '請選擇轄區中心\n';
            if (document.form1.DistID.selectedIndex == 0) msg += '請選擇轄區分署\n';
            if (document.form1.years.selectedIndex == 0) msg += '請選擇年度\n';
            if (document.form1.months.selectedIndex == 0) msg += '請選擇月份\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table id="Table1" cellspacing="1" cellpadding="1" width="600" border="0">
        <tr>
            <td>
                <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            首頁&gt;&gt;學員動態管理&gt;&gt;教務報表管理&gt;&gt; <font color="#990000">職前訓練月報表</font>
                        </td>
                    </tr>
                </table>
                <table class="table_sch" id="Table3" >
                    <tr>
                        <%--<td class="bluecol_need" width="100">轄區中心</td>--%>
                        <td class="bluecol_need" width="100">轄區分署</td>
                        <td class="whitecol" colspan="3">
                            <asp:DropDownList ID="DistID" runat="server">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol_need" width="100">月份</td>
                        <td class="whitecol">
                            <asp:DropDownList ID="years" runat="server">
                            </asp:DropDownList>
                            年
                            <asp:DropDownList ID="months" runat="server">
                            </asp:DropDownList>月
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <br />
    <div style="width:600" align="center">
        <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
    </div>
    </form>
</body>
</html>
