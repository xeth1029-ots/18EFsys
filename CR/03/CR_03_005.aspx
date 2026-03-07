<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CR_03_005.aspx.vb" Inherits="WDAIIP.CR_03_005" %>

<%--<!DOCTYPE html>--%>
<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">

<%--<html xmlns="http://www.w3.org/1999/xhtml">--%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>核定課程明細表</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="../../Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function SelectAll(obj, hidobj) {
            var num = getCheckBoxListValue(obj).length; //長度
            var myallcheck = document.getElementById(obj + '_' + 0); //第1個
            //alert(getCheckBoxListValue(obj));
            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0); //記憶
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
            else {
                //若有全選
                if (getCheckBoxListValue(obj).charAt(0) == '1') {
                    myallcheck.checked = false; //全選改為false
                    document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0); //記憶
                }
            }
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;課程審查&gt;&gt;【陳核版】課程核定報表&gt;&gt;核定課程明細表</asp:Label>
                </td>
            </tr>
        </table>
        <asp:Panel ID="PanelSch1" runat="server">
            <table width="100%" cellpadding="1" cellspacing="1" class="table_sch">
                <tr>
                    <td class="bluecol_need">轄區分署</td>
                    <td class="whitecol">
                        <asp:CheckBoxList ID="cblDistid" runat="server" RepeatDirection="Horizontal" RepeatColumns="3" CssClass="whitecol">
                        </asp:CheckBoxList>
                        <input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server">
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need" width="18%">申請階段</td>
                    <td class="whitecol" width="82%">
                        <asp:DropDownList ID="ddlAPPSTAGE_SCH" runat="server"></asp:DropDownList></td>
                </tr>
                <tr id="TRPlanPoint28" runat="server">
                    <td class="bluecol">計畫 </td>
                    <td class="whitecol">
                        <asp:RadioButtonList ID="rblOrgKind2" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                            <asp:ListItem Value="G" Selected="True">產業人才投資計畫</asp:ListItem>
                            <asp:ListItem Value="W">提升勞工自主學習計畫</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">匯出檔案格式</td>
                    <td class="whitecol">
                        <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                            <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                            <asp:ListItem Value="ODS">ODS</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <%-- <tr><td class="whitecol" colspan="2">※1..當【申請階段】選擇「3：政策性產業」，即不區分計畫查詢</td></tr>--%>
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
