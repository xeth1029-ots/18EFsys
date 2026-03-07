<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_001_add.aspx.vb" Inherits="WDAIIP.TC_01_001_add" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>計畫代碼設定</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <%--<script type="text/javascript" src="../../js/date-picker.js"></script>--%>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <%--
    <script language="javascript">
        function yearPlan(selectedPlanID) {
            var year = document.getElementById('yearlist_add');
            var parms = "[['year','" + year.value + "']]";      //透過 selectControl 傳遞給 SQLMap 的年度查詢條件, 格式請參考 selectControl 定義說明
            selectControl('ajaxTPlanList', 'planlist_add', 'PlanName', 'TPlanID', '請選擇', selectedPlanID, parms);
        }
    </script>
    --%>
    <script type="text/javascript">
        //決定date-picker元件使用的是西元年or民國年，by:20181018
        var bl_rocYear = "<%=ConfigurationManager.AppSettings("REPLACE2ROC_YEARS") %>";
        var scriptTag = document.createElement('script');
        var jsPath = (bl_rocYear == "Y" ? "../../js/date-picker2.js" : "../../js/date-picker.js");
        scriptTag.src = jsPath;
        document.head.appendChild(scriptTag);
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <table class="font" id="table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <%--
                    <table class="font" id="Table5" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td style="background-color:#f8f2d8">(<font color="#ff0000">*</font>為必填欄位)</td>
                        </tr>
                    </table>
                    --%>
                    <table class="table_sch" id="Table4" cellpadding="1" cellspacing="1">
                        <tr>
                            <td>
                                <table id="Table1" cellspacing="1" cellpadding="0" border="0" runat="server" width="100%">
                                    <tr>
                                        <td class="bluecol" width="20%">轄區</td>
                                        <td class="whitecol" width="80%" colspan="3">
                                            <asp:DropDownList ID="DistValue" runat="server"></asp:DropDownList></td>
                                    </tr>
                                    <tr>
                                        <td id="td1" class="bluecol_need" runat="server" width="20%">年度</td>
                                        <td class="whitecol" width="30%">
                                            <asp:DropDownList ID="yearlist_add" runat="server" AutoPostBack="True"></asp:DropDownList>
                                            <%--<asp:RequiredFieldValidator ID="Requiredfieldvalidator1" runat="server" ErrorMessage="請選擇年度" Display="None" ControlToValidate="yearlist_add"></asp:RequiredFieldValidator>--%>
                                        </td>
                                        <td id="td8" class="bluecol_need" runat="server" width="20%">序號</td>
                                        <td class="whitecol" width="30%">
                                            <asp:TextBox ID="seqno" runat="server" onfocus="this.blur()" MaxLength="3" Columns="5" Width="70%"></asp:TextBox><br />
                                            (系統自動產生)</td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need" width="20%">訓練計畫</td>
                                        <td colspan="3" class="whitecol" width="80%">
                                            <asp:DropDownList ID="planlist_add" runat="server" AutoPostBack="true"></asp:DropDownList>
                                            <%--<asp:RequiredFieldValidator ID="mustplan" runat="server" ErrorMessage="請選擇訓練計畫" Display="None" ControlToValidate="planlist_add"></asp:RequiredFieldValidator>--%>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td id="td3" runat="server" class="bluecol" width="20%">主辦單位</td>
                                        <td class="whitecol" width="30%">
                                            <asp:TextBox ID="main_center" runat="server" MaxLength="45" Columns="30" Width="80%"></asp:TextBox></td>
                                        <td id="td4" runat="server" class="bluecol" width="20%">協辦單位</td>
                                        <td class="whitecol" width="30%">
                                            <asp:TextBox ID="sub_center" runat="server" MaxLength="45" Columns="30" Width="80%"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td id="td5" runat="server" class="bluecol_need" width="20%">時效起、迄日</td>
                                        <td colspan="3" class="whitecol" width="80%">
                                            <asp:TextBox ID="start_date" runat="server" onfocus="this.blur()" Width="15%"></asp:TextBox>
                                            <span id="span1" runat="server">
                                                <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>～
                                            <asp:TextBox ID="end_date" runat="server" onfocus="this.blur()" Width="15%"></asp:TextBox>
                                            <span id="span2" runat="server">
                                                <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                            <%--
                                            <asp:CompareValidator ID="comdate" runat="server" ErrorMessage="迄日不得小於起日或日期格式不正確" Display="None" ControlToValidate="end_date" Type="Date" Operator="GreaterThan" ControlToCompare="start_date"></asp:CompareValidator>
                                            <asp:RequiredFieldValidator ID="mustsdate" runat="server" ErrorMessage="請輸入起日" Display="None" ControlToValidate="start_date"></asp:RequiredFieldValidator>
                                            <asp:RequiredFieldValidator ID="mustedate" runat="server" ErrorMessage="請輸入迄日" Display="None" ControlToValidate="end_date"></asp:RequiredFieldValidator>
                                            --%>
                                            <%--<asp:ValidationSummary ID="totalmsg" runat="server" ShowSummary="False" ShowMessageBox="True" DisplayMode="List"></asp:ValidationSummary>--%>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td id="Td6" runat="server" class="bluecol_need" width="20%">計畫種類</td>
                                        <td colspan="3" class="whitecol" width="80%">
                                            <asp:RadioButtonList ID="PlanKind" runat="server" AutoPostBack="True" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font"></asp:RadioButtonList>
                                            <%--<asp:RequiredFieldValidator ID="mustPlanKind" runat="server" ErrorMessage="請選擇計畫總類" Display="None" ControlToValidate="PlanKind"></asp:RequiredFieldValidator>--%>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td id="Td7" runat="server" class="bluecol_need" width="20%">彈性調整出缺勤</td>
                                        <td colspan="3" class="whitecol" width="80%">
                                            <asp:RadioButtonList ID="FlexTurnoutKind" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font"></asp:RadioButtonList>
                                            <%--<asp:RequiredFieldValidator ID="MustFlexTurnoutKind" runat="server" ErrorMessage="請選擇婚假是否列入缺曠課時數" Display="None" ControlToValidate="PlanKind"></asp:RequiredFieldValidator>--%>
                                        </td>
                                    </tr>
                                    <tr id="Tr18" runat="server">
                                        <td class="bluecol_need" width="20%">e網報名審核發送Email</td>
                                        <td colspan="3" class="whitecol" width="80%">
                                            <asp:RadioButtonList ID="R18" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="Y">是</asp:ListItem>
                                                <asp:ListItem Value="N">否</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">備註顯示</td>
                                        <td colspan="3" class="whitecol" width="80%">
                                            <asp:TextBox ID="SubTitle" runat="server" Columns="20" MaxLength="10" Width="50%"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">計畫說明</td>
                                        <td colspan="3" class="whitecol" width="80%">
                                            <asp:TextBox ID="PComment" runat="server" Columns="5" Width="50%" TextMode="MultiLine" Rows="6"></asp:TextBox></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="bt_addrow" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                                &nbsp;<asp:Button ID="Button1" runat="server" Text="回上一頁" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <%--<asp:HiddenField ID="hidPlanID" runat="server" />--%>
        <asp:HiddenField ID="Hid_rqeditid" runat="server" />
        <input id="hidPlanKind" type="hidden" runat="server" />
        <asp:HiddenField ID="Hid_FlexTurnoutKind_OPEN_CLOSE" runat="server" />
    </form>
</body>
</html>
