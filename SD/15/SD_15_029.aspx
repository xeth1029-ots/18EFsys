<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_029.aspx.vb" Inherits="WDAIIP.SD_15_029" %>

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>核定課程統計表</title>
    <link rel="stylesheet" type="text/css" href="../../css/style.css" />
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function OpenOrg(vTPlanID) {
            var ddlDistID = document.getElementById('ddlDistID');
            if (!ddlDistID) return;
            if (ddlDistID.selectedIndex == 0) {
                alert('請先選擇轄區');
                return false;
            }
            wopen('../../common/MainOrg.aspx?DistID=' + ddlDistID.value + '&TPlanID=' + vTPlanID + '&YEARSTYPE=3', '', 600, 600, 'yes');
        }

        function BtnClear2Click() {
            var Objcenter = document.getElementById('center');
            var RIDValue = document.getElementById('RIDValue');
            var PlanID = document.getElementById('PlanID');
            if (Objcenter) Objcenter.value = "";
            if (RIDValue) RIDValue.value = "";
            if (PlanID) PlanID.value = "";
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;統計表&gt;&gt;核定課程統計表</asp:Label>
                </td>
            </tr>
        </table>
        <table width="100%" cellpadding="1" cellspacing="1" class="table_sch">
            <tr>
                <td class="bluecol_need" width="20%">訓練計畫</td>
                <td colspan="3" class="whitecol" width="80%">
                    <asp:DropDownList ID="TPlanlist1" runat="server"></asp:DropDownList>
                </td>
            </tr>
            <%-- <tr>
                <td class="bluecol">分署</td>
                <td class="whitecol">
                    <asp:DropDownList ID="ddlDistID" runat="server"></asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">機構</td>
                <td class="whitecol">
                    <asp:TextBox ID="center" runat="server" Width="70%" AutoCompleteType="Disabled" onfocus="this.blur()"></asp:TextBox>
                    <input id="Button1" type="button" value="..." name="Button1" runat="server" class="button_b_Mini" />
                    <input id="RIDValue" type="hidden" runat="server">
                    <input id="PlanID" type="hidden" runat="server">
                    <input id="BtnClear2" type="button" value="清空選擇" name="BtnClear2" runat="server" class="asp_Export_M" />
                </td>
            </tr>--%>
            <tr>
                <td class="bluecol" style="width: 20%">訓練機構</td>
                <td class="whitecol">
                    <asp:TextBox ID="center" runat="server" Width="70%" onfocus="this.blur()"></asp:TextBox>
                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                    <input id="BtnLevOrg1" type="button" value="..." name="BtnLevOrg1" runat="server" class="asp_button_Mini" />
                    <asp:Button Style="display: none" ID="BtnGETvalue2" runat="server"></asp:Button>
                    <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                        <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                    </span>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need" width="20%">目標計畫年度</td>
                <td colspan="3" class="whitecol" width="80%">
                    <asp:DropDownList ID="yearlist" runat="server"></asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">申請階段</td>
                <td class="whitecol">
                    <asp:DropDownList ID="ddlAPPSTAG1" runat="server"></asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">計畫範圍
                </td>
                <td colspan="3" class="whitecol">
                    <asp:RadioButtonList ID="OrgKind2" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                        <asp:ListItem Value="G" Selected="True">產業人才投資計畫</asp:ListItem>
                        <asp:ListItem Value="W">提升勞工自主學習計畫</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">課程狀態
                </td>
                <td colspan="3" class="whitecol">
                    <asp:RadioButtonList ID="RBL_CLASSSTATUS" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                        <asp:ListItem Value="1" Selected="True">已申請</asp:ListItem>
                        <asp:ListItem Value="2">已二階審查</asp:ListItem>
                        <asp:ListItem Value="3">已核定(班級審核)</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <%-- <tr>
                <td class="bluecol">匯出檔案格式</td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                        <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                        <asp:ListItem Value="ODS">ODS</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>--%>
        </table>
        <table width="100%">
            <tr>
                <td class="whitecol" align="center">
                    <asp:Button ID="bt_EXPORT" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                    <%--<asp:Button ID="bt_DOWNLOADFILE" runat="server" Text="檔案下載" CssClass="asp_Export_M"></asp:Button>--%>
                    <div align="center"></div>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                    </div>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="Hid_YEAR_ROC1" runat="server" />
        <asp:HiddenField ID="Hid_YEAR_ROC3" runat="server" />
        <%--<asp:HiddenField ID="Hid_DISTID" runat="server" />--%>
    </form>
</body>
</html>
