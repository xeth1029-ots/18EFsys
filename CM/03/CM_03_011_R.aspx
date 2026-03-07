<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CM_03_011_R.aspx.vb" Inherits="WDAIIP.CM_03_011_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>交叉分析統計表</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <script type="text/javascript">
        function PageLoad() {
            document.body.innerHTML = document.body.innerHTML + "<IMG src='../../images/rptpic/temple/TIMS_1.jpg' style='z-index:-1;position:absolute;top:0px;left:0px;height:680px;width:640px;display:inline' />";
            document.all.factory.printing.header = "";
            document.all.factory.printing.footer = "";
            document.all.factory.printing.Print(true);
            window.close();
        }
    </script>
</head>
<body background="../../images/rptpic/temple/TIMS_1.jpg" onload="PageLoad();">
    <!-- MeadCo ScriptX -->
    <object id="factory" style="display: none" classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="../../scriptx/smsx.cab#Version=6,6,440,26"></object>
    <form id="form1" method="post" runat="server">
    <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
        <tr>
            <td align="center">
                (<asp:Label ID="Label1" runat="server"></asp:Label>)年度&nbsp;
                    <%--<asp:Label id="lb_Plan" runat="server"></asp:Label>--%>
                    交叉分析統計表
            </td>
        </tr>
        <tr>
            <td align="right">
                資料日期：<asp:Label ID="PrintDate" runat="server"></asp:Label>
            </td>
        </tr>
        <tr>
            <td>
                <asp:Table ID="DataTable1" runat="server" Width="100%" CellSpacing="0" CellPadding="2" Font-Size="X-Small"></asp:Table>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>