<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_05_002_R.aspx.vb" Inherits="WDAIIP.TR_05_002_R" %>



<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>年度職業訓練行業別_性別分佈</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        function search() {
            var msg = '';
            if (document.form1.Syear1.selectedIndex == 0) msg += '請選擇起始年度\n';
            if (document.form1.Syear2.selectedIndex == 0) msg += '請選擇結束年度\n';
            if (parseInt(document.form1.Syear1.value) > parseInt(document.form1.Syear2.value)) msg += '請選擇起始年度不能大於結束年度\n';
            if (!isChecked(document.getElementsByName('PlanType'))) msg += '請選擇訓練性質\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練需求管理&gt;&gt;統計分析&gt;&gt;綜合動態報表</asp:Label>
                    <%--首頁&gt;&gt;訓練與就業需求管理&gt;&gt;統計分析&gt;&gt;<FONT color="#990000">年度職業訓練行業別_性別分佈</FONT>--%>
                </td>
            </tr>
        </table>

        <table id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td class="bluecol" width="20%">動態報表 </td>
                <td class="whitecol" width="80%">
                    <uc1:WUC2 runat="server" ID="WUC2" />
                </td>
            </tr>
            <tr>
                <td class="bluecol_need" width="20%">年度
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="Syear1" runat="server">
                    </asp:DropDownList>
                    ~
                    <asp:DropDownList ID="Syear2" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">訓練性質
                </td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="PlanType" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                        <asp:ListItem Value="">全部</asp:ListItem>
                        <asp:ListItem Value="1">自辦</asp:ListItem>
                        <asp:ListItem Value="2">委辦</asp:ListItem>
                        <asp:ListItem Value="3">合辦</asp:ListItem>
                        <asp:ListItem Value="4">補助</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td colspan="2" align="center">
                    <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                </td>
            </tr>
        </table>

    </form>
</body>
</html>
