<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_02_003.aspx.vb" Inherits="WDAIIP.SYS_02_003" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>授予相同權限</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/common.js"></script>
    <style type="text/css">
        .style1 { width: 72px; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="600">
            <tr>
                <td>
                    <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>首頁&gt;&gt;系統管理&gt;&gt;帳號維護管理&gt;&gt;<font color="#990000">授予相同權限</font></td>
                        </tr>
                    </table>
                    <table class="table_sch" id="Table3">
                        <tr>
                            <td width="100" class="bluecol">角色</td>
                            <td colspan="3" class="whitecol">
                                <asp:DropDownList ID="RoleName" runat="server" AutoPostBack="True"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol">帳號</td>
                            <td width="187" class="whitecol">&nbsp;&nbsp;&nbsp;&nbsp;<asp:DropDownList ID="BefAcct" runat="server" AutoPostBack="True"></asp:DropDownList><input id="RIDValue" type="hidden" name="RIDValue" runat="server"></td>
                            <td colspan="2" class="whitecol">&nbsp;&nbsp;&nbsp;&nbsp;<asp:RadioButtonList ID="PlanName" runat="server" AutoPostBack="True" CssClass="font" RepeatLayout="Flow"></asp:RadioButtonList></td>
                        </tr>
                        <tr>
                            <td class="bluecol" colspan="3">授予相同權限給</td>
                        </tr>
                        <tr>
                            <td class="bluecol">帳號
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:DropDownList ID="NewAcct" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td colspan="4" class="whitecol">
                                <p align="center">
                                    <asp:Button ID="Button1" runat="server" Text="儲存"></asp:Button>
                                </p>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
