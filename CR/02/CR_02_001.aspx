<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CR_02_001.aspx.vb" Inherits="WDAIIP.CR_02_001" %>

<%--<!DOCTYPE html>--%>
<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">

<%--<html xmlns="http://www.w3.org/1999/xhtml">--%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>核配額度上限控管彙整表</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;課程審查&gt;&gt;二階審查&gt;&gt;核配額度上限控管彙整表</asp:Label>
                </td>
            </tr>
        </table>
        <asp:Panel ID="PanelSch1" runat="server">
            <table width="100%" cellpadding="1" cellspacing="1" class="table_sch">
                <%--<tr>
                    <td class="bluecol_need" width="18%">年度</td>
                    <td class="whitecol" width="82%" colspan="3">
                        <asp:DropDownList ID="ddlYEARS_SCH" runat="server"></asp:DropDownList></td>
                </tr>--%>
                <tr>
                    <td class="bluecol_need">申請階段</td>
                    <td class="whitecol">
                        <asp:DropDownList ID="ddlAPPSTAGE_SCH" runat="server"></asp:DropDownList></td>
                </tr>
                <%--<tr>
                    <td class="bluecol">訓練機構 </td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="center" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox>
                        <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />&nbsp;
							<input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini" />
                        <span id="HistoryList2" style="position: absolute; display: none">
                            <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                        </span></td>
                </tr>--%>
                <tr id="TRPlanPoint28" runat="server">
                    <td class="bluecol">計畫 </td>
                    <td class="whitecol">
                        <asp:RadioButtonList ID="rblOrgKind2" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                            <asp:ListItem Value="G" Selected="True">產業人才投資計畫</asp:ListItem>
                            <asp:ListItem Value="W">提升勞工自主學習計畫</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <%--<tr>
                    <td class="bluecol">申請日期區間 </td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="APPLIEDDATE1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                        <img style="cursor: pointer" onclick="javascript:show_calendar('APPLIEDDATE1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                        ～<asp:TextBox ID="APPLIEDDATE2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                        <img style="cursor: pointer" onclick="javascript:show_calendar('APPLIEDDATE2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">開訓日期 </td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="STDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                        <img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                        ～<asp:TextBox ID="STDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                        <img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                    </td>
                </tr>--%>
                <tr>
                    <td class="bluecol">匯出檔案格式</td>
                    <td class="whitecol">
                        <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                            <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                            <asp:ListItem Value="ODS">ODS</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <%-- <tr>
                    <td class="bluecol">&nbsp;篩選範圍</td>
                    <td colspan="3" class="whitecol">
                        <asp:RadioButtonList ID="RBL_RANGE1_SCH" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                            <asp:ListItem Value="1" Selected="True">不區分</asp:ListItem>
                            <asp:ListItem Value="2">轄區單位</asp:ListItem>
                            <asp:ListItem Value="3">19大類主責課程</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">跨區/轄區提案</td>
                    <td class="whitecol" colspan="3">
                        <asp:RadioButtonList ID="RBL_CrossDist_SCH" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                            <asp:ListItem Value="D" Selected="True">不區分</asp:ListItem>
                            <asp:ListItem Value="C">跨區提案單位</asp:ListItem>
                            <asp:ListItem Value="J">轄區提案單位</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>--%>
                <%--<tr>
                    <td class="bluecol">匯出檔案格式</td>
                    <td class="whitecol" colspan="3">
                        <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                            <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                            <asp:ListItem Value="ODS">ODS</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>--%>
                <tr>
                    <td class="whitecol" align="center" colspan="2">
                        <%--<asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                        <asp:Button ID="BtnSearch" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>--%>
                        <asp:Button ID="BtnExport1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                        <%--<asp:Button ID="BtnExport2" runat="server" Text="匯出審查意見綜整表" CssClass="asp_Export_M"></asp:Button>--%>
                    </td>
                </tr>
            </table>
            <div align="center">
                <asp:Label ID="msg1" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
            </div>
        </asp:Panel>
    </form>
</body>
</html>
