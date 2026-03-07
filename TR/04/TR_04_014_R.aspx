<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_04_014_R.aspx.vb" Inherits="WDAIIP.TR_04_014_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>TR_04_014_R</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script src="../../js/common.js"></script>
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
        <tr>
            <td>
                <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            首頁&gt;&gt;訓練與就業需求管理&gt;&gt;就業追蹤報表&gt;&gt;<font color="#990000">受評單位各班別結訓學員就業狀況表</font>
                        </td>
                    </tr>
                </table>
                <table class="table_sch" id="Table2" cellspacing="1" cellpadding="1">
                    <tr>
                        <td class="bluecol" width="80">
                            結訓期間
                        </td>
                        <td class="whitecol">
                            <asp:DropDownList ID="SYear" runat="server">
                            </asp:DropDownList>
                            年
                            <asp:DropDownList ID="SMonth" runat="server">
                            </asp:DropDownList>
                            月～
                            <asp:DropDownList ID="FYear" runat="server">
                            </asp:DropDownList>
                            年
                            <asp:DropDownList ID="FMonth" runat="server">
                            </asp:DropDownList>
                            月
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            訓練計畫
                        </td>
                        <td class="whitecol" colspan="3">
                            <asp:CheckBox ID="SelectAllItem" runat="server" Text="全選"></asp:CheckBox><asp:CheckBoxList ID="TPlanID" runat="server" RepeatColumns="3" CssClass="font" RepeatDirection="Horizontal" CellSpacing="0" CellPadding="0">
                            </asp:CheckBoxList>
                        </td>
                    </tr>
                </table>
                <p align="center">
                    <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                </p>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
