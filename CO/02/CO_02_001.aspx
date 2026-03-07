<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CO_02_001.aspx.vb" Inherits="WDAIIP.CO_02_001" %>

<%--<!DOCTYPE html>--%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>審查計分綜合動態報表</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-confirm.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery.blockUI.js"></script>
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="../../js/TIMS.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <%--<script type="text/javascript">,$(document).ready(function () {,});,</script>,--%>
</head>
<body>
    <form id="form1" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;審查計分表&gt;&gt;統計表&gt;&gt;審查計分綜合動態報表</asp:Label>
                </td>
            </tr>
        </table>
        <div id="divSch1" runat="server">
            <table class="table_sch" cellpadding="1" cellspacing="1" width="100%" border="0">
                <tr>
                    <td class="bluecol" style="width: 20%">分署</td>
                    <td colspan="3" class="whitecol">
                        <asp:DropDownList ID="ddlDISTID_SCH" runat="server"></asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol" style="width: 20%">審查計分區間</td>
                    <td colspan="3" class="whitecol">
                        <asp:DropDownList ID="ddlSCORING_SCH" runat="server"></asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">訓練機構</td>
                    <td class="whitecol" style="width: 35%">
                        <asp:TextBox ID="ORGNAME_SCH" runat="server" MaxLength="50" Columns="60" Width="90%"></asp:TextBox>
                    </td>
                    <td class="bluecol" style="width: 20%">統一編號</td>
                    <td class="whitecol" style="width: 25%">
                        <asp:TextBox ID="COMIDNO_SCH" runat="server" MaxLength="15" Width="60%"></asp:TextBox>
                    </td>
                </tr>
                <tr id="TRPlanPoint28" runat="server">
                    <td class="bluecol_need">計畫
                    </td>
                    <td colspan="3" class="whitecol">
                        <asp:RadioButtonList ID="RBL_ORGPLANKIND_SCH" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow" AutoPostBack="True">
                            <asp:ListItem Value="G" Selected="True">產業人才投資計畫</asp:ListItem>
                            <asp:ListItem Value="W">提升勞工自主學習計畫</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <%--單位屬性,機構別--%>
                    <td class="bluecol">單位屬性</td>
                    <td colspan="3" class="whitecol">
                        <asp:DropDownList ID="DDL_TYPEID2_SCH" runat="server"></asp:DropDownList>
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
                </tr>
                <tr>
                    <td class="bluecol">單位負責人 </td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="MASTERNAME_SCH" runat="server" MaxLength="30" Columns="33" Width="33%"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol">匯出欄位
                    </td>
                    <td class="whitecol" colspan="3">
                        <asp:CheckBoxList ID="CBEXIT_SCH" runat="server" RepeatDirection="Horizontal" RepeatColumns="3" CssClass="whitecol">
                        </asp:CheckBoxList>
                        <asp:CheckBoxList ID="CBEXIT2_SCH" runat="server" RepeatDirection="Horizontal" RepeatColumns="1" CssClass="whitecol">
                        </asp:CheckBoxList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">匯出檔案格式</td>
                    <td class="whitecol" colspan="3">
                        <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                            <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                            <asp:ListItem Value="ODS">ODS</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td class="whitecol" colspan="4" align="center">
                        <asp:Button ID="BTN_EXP1" runat="server" Text="匯出審查計分綜合動態報表" CssClass="asp_Export_M" data-exp="Y"></asp:Button>
                    </td>
                </tr>
            </table>
            <div align="center">
                <asp:Label ID="labmsg1" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
            </div>
        </div>
    </form>
</body>
</html>
