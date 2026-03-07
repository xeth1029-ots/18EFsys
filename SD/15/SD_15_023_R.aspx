<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_023_R.aspx.vb" Inherits="WDAIIP.SD_15_023_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>交叉分析統計表</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function PageLoad() {
            if (_isIE) { print(); window.close(); }
            else { window.print() }
        }
    </script>
</head>
<body onload="PageLoad();">
    <form id="form1" method="post" runat="server">
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0" onclick="window.close();">
            <tr>
                <td align="center">
                    <font face="標楷體" size="5">(<asp:Label ID="Label1" runat="server"></asp:Label>)年度&nbsp;
					<asp:Label ID="lb_Plan" runat="server"></asp:Label>&nbsp;交叉分析統計表</font>
                </td>
            </tr>
            <tr>
                <td align="right">
                    <font face="新細明體"><font face="標楷體">資料日期：</font>
                        <asp:Label ID="PrintDate" runat="server" Font-Names="標楷體"></asp:Label></font>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Table ID="DataTable1" runat="server" Width="100%" CellSpacing="0" CellPadding="2" Font-Names="標楷體" Font-Size="X-Small">
                    </asp:Table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
