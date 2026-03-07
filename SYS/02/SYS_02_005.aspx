<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_02_005.aspx.vb" Inherits="WDAIIP.SYS_02_005" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>層級對調</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/common.js"></script>
</head>
<body>
    <form id="form1" method="post" runat="server">

        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
            <tr>
                <td>
                    <table id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
                        <tr>
                            <td>首頁&gt;&gt;系統管理&gt;&gt;帳號維護管理&gt;&gt;<font color="#990000">層級對調</font></td>
                        </tr>
                    </table>
                    <table id="Table3" class="table_sch">
                        <tr>
                            <td class="bluecol" width="100">帳號</td>
                            <td class="whitecol">
                                <asp:DropDownList ID="BefAcct" runat="server" AutoPostBack="True"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="100">新角色</td>
                            <td class="whitecol">
                                <asp:DropDownList ID="Role" runat="server" AutoPostBack="True"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol">交換對象</td>
                            <td class="whitecol">
                                <asp:DropDownList ID="NewAcct" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td colspan="2" class="whitecol" align="center">
                                <asp:Button ID="Button1" runat="server" Text="儲存" CssClass="asp_button_S"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>

    </form>
</body>
</html>
