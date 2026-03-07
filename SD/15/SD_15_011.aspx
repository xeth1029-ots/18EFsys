<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_011.aspx.vb" Inherits="WDAIIP.SD_15_011" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>開訓統計週報表</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;統計表&gt;&gt;綜合動態報表</asp:Label>
                    <%--首頁&gt;&gt;學員動態管理&gt;&gt;產學訓統計表&gt;&gt;<FONT color="#800000">開訓統計週報表</FONT>--%>                    
                </td>
            </tr>
        </table>
        <table class="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td class="bluecol" width="20%">動態報表 </td>
                <td class="whitecol" width="80%">
                    <uc1:WUC1 runat="server" ID="WUC1" />
                </td>
            </tr>
            <tr>
                <td class="bluecol">年度 </td>
                <td class="whitecol">
                    <asp:DropDownList ID="Years" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">轄區 </td>
                <td class="whitecol">
                    <%--<asp:RadioButtonList ID="DistID" runat="server" Width="440px" RepeatLayout="Flow" RepeatDirection="Horizontal">
                        <asp:ListItem Value="0" Selected="True">全部</asp:ListItem>
                        <asp:ListItem Value="001">北區</asp:ListItem>
                        <asp:ListItem Value="003">桃園</asp:ListItem>
                        <asp:ListItem Value="004">中區</asp:ListItem>
                        <asp:ListItem Value="005">台南</asp:ListItem>
                        <asp:ListItem Value="006">南區</asp:ListItem>
                    </asp:RadioButtonList>--%>
                    <asp:RadioButtonList ID="DistID" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">開訓日期 </td>
                <td class="whitecol" runat="server">
                    <%--onfocus="this.blur()"--%>
                    <asp:TextBox ID="start_date" runat="server" Columns="10" MaxLength="10"></asp:TextBox>
                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">
                    ～
				    <asp:TextBox ID="end_date" runat="server" Columns="10" MaxLength="10"></asp:TextBox>
                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">
                </td>
            </tr>
            <tr id="tr_AppStage_TP28" runat="server">
                <td class="bluecol">申請階段 </td>
                <td class="whitecol">
                    <asp:DropDownList ID="AppStage" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr id="trPlanKind" runat="server">
                <td class="bluecol" align="center">計畫範圍 </td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="OrgKind2" runat="server" Width="493px" RepeatLayout="Flow" RepeatDirection="Horizontal">
                        <asp:ListItem Value="G" Selected="True">產業人才投資計畫</asp:ListItem>
                        <asp:ListItem Value="W">提升勞工自主學習計畫</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr id="trPackageType" runat="server">
                <td class="bluecol" align="center">包班種類 </td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="PackageType" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                        <asp:ListItem Value="A" Selected="True">全部</asp:ListItem>
                        <%--<asp:ListItem Value="1">非包班</asp:ListItem>--%>
                        <asp:ListItem Value="2">企業包班</asp:ListItem>
                        <asp:ListItem Value="3">聯合企業包班</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td align="center" colspan="2">
                    <asp:Button ID="btnPrint1" runat="server" Text="列印" CssClass="asp_Export_M" />
                    <%--<input id="Print" value="列印" type="button" runat="server" class="asp_Export_M">--%>
                </td>
            </tr>

        </table>

    </form>
</body>
</html>
