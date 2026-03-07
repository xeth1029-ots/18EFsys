<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_07_001_add.aspx.vb" Inherits="WDAIIP.CP_07_001_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>受訓期間學員滿意度</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
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
        //function chkinput() {
        //    var msg = '';
        //    //alert('test');
        //    //if (!confirm('確定儲存?')) { return false; }

        //    if (document.form1.Q1_1.selectedIndex == 0) {
        //        msg += '請填寫[一、課程安排] 1.您對整體課程安排是否滿意？\n';
        //    }

        //    if (document.form1.Q1_2.selectedIndex == 0) {
        //        msg += '請填寫[一、課程安排] 2.您對課程安排內容銜接情形是否滿意？\n';
        //    }
        //    if (document.form1.Q1_3.selectedIndex == 0) {
        //        msg += '請填寫[一、課程安排] 3.您對課程內容時數安排是否滿意？\n';
        //    }

        //    if (document.form1.Q2_1.selectedIndex == 0) {
        //        msg += '請填寫[二、師資與教學] 1.您對老師的教學方式是否滿意？\n';
        //    }
        //    if (document.form1.Q2_2.selectedIndex == 0) {
        //        msg += '請填寫[二、師資與教學] 2.您對老師的教學態度是否滿意？\n';
        //    }
        //    if (document.form1.Q2_3.selectedIndex == 0) {
        //        msg += '請填寫[二、師資與教學] 3.您對老師的專業知識是否滿意？\n';
        //    }

        //    if (document.form1.Q3_1.selectedIndex == 0) {
        //        msg += '請填寫[三、設備和教材] 1.您對上課期間，教材設備的充分利用情形是否滿意？\n';
        //    }
        //    if (document.form1.Q3_2.selectedIndex == 0) {
        //        msg += '請填寫[三、設備和教材] 2.您對教材的新穎是否滿意？\n';
        //    }
        //    if (document.form1.Q3_3.selectedIndex == 0) {
        //        msg += '請填寫[三、設備和教材] 3.您對教材內容的難易程度是否滿意？\n';
        //    }
        //    if (document.form1.Q4_1.selectedIndex == 0) {
        //        msg += '請填寫[四、行政措施] 1.您對導師關心學員學習狀況及解決問題的能力是否滿意？\n';
        //    }
        //    if (document.form1.Q4_2.selectedIndex == 0) {
        //        msg += '請填寫[四、行政措施] 2.您對求助及申訴管道是否滿意？\n';
        //    }
        //    if (document.form1.Q4_3.selectedIndex == 0) {
        //        msg += '請填寫[四、行政措施] 3.您對學習空間是否滿意？\n';
        //    }
        //    if (document.form1.Q4_4.selectedIndex == 0) {
        //        msg += '請填寫[四、行政措施] 4.您對上課地點場地及週遭環境清潔是否滿意？\n';
        //    }

        //    if (msg != "") {
        //        alert(msg);
        //        return false;
        //    }
        //}
        //function chkRadio1_1() {
        //    var radiolist = document.getElementsByName('Q1_1');
        //    if (!((radiolist[1].checked) || (radiolist[2].checked) || (radiolist[3].checked))) {
        //        return false;
        //    }
        //}
        //function ChgFillDate() {
        //    __doPostBack('LinkButton1', '');
        //}
    </script>

</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;訓練成效與滿意度&gt;&gt;受訓期間學員滿意度</asp:Label>
                </td>
            </tr>
        </table>
        <table class="table_nw" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td width="15%" class="bluecol">計畫名稱</td>
                <td class="whitecol" width="85%" colspan="3">
                    <asp:Label ID="lb_PlanName" runat="server"></asp:Label>
                </td>
            </tr>
            <tr>
                <td class="bluecol">訓練單位</td>
                <td class="whitecol" colspan="3">
                    <asp:Label ID="lb_OrgName" runat="server"></asp:Label>
                </td>
            </tr>
            <tr>
                <td class="bluecol">訓練班別</td>
                <td class="whitecol" colspan="3">
                    <asp:Label ID="lb_OCID" runat="server"></asp:Label>
                </td>
            </tr>
            <tr>
                <td class="bluecol">訓練期間</td>
                <td class="whitecol" colspan="3">
                    <asp:Label ID="lb_STDate" runat="server"></asp:Label>~
                <asp:Label ID="lb_FTDate" runat="server"></asp:Label>
                </td>
            </tr>
            <tr>
                <td class="bluecol">學員姓名</td>
                <td class="whitecol" colspan="3">
                    <asp:Label ID="lb_STDNAME" runat="server"></asp:Label>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">填表日期</td>
                <td class="whitecol" colspan="3">
                    <asp:TextBox ID="txFillDate" runat="server" Width="15%"></asp:TextBox>
                    <span id="span_FillDate" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= txFillDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                    </span>
                    <%--<asp:LinkButton ID="LinkButton1" runat="server"></asp:LinkButton>--%>
                </td>
            </tr>
            <tr id="tr_signer" runat="server">
                <td class="bluecol">抽訪人員姓名</td>
                <td class="whitecol">
                    <asp:TextBox ID="signer" runat="server" MaxLength="15" Width="30%"></asp:TextBox>
                    <%--<asp:Label ID="msg" runat="server" Width="168px" ForeColor="Red"></asp:Label>--%>
                </td>
            </tr>

        </table>
        <div align="center">
            <asp:Label ID="msg1" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
        </div>
        <%--<table class="font" width="100%">--%>
        <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0" id="tb_Ques1" runat="server">
            <tr>
                <td class="table_title">壹、參訓意見(請依滿意度擇一勾選)
                </td>
            </tr>
            <tr>
                <td align="center" class="table_title">一、課程安排
                </td>
            </tr>

            <tr>
                <td class="td_light">1.您對整體課程安排是否滿意？
                </td>
            </tr>
            <tr>
                <td class="whitecol">
                    <asp:RadioButtonList ID="Q1_1" runat="server" CssClass="font" RepeatDirection="Horizontal" Width="80%">
                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                        <asp:ListItem Value="2">滿意</asp:ListItem>
                        <asp:ListItem Value="3">尚可</asp:ListItem>
                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td></td>
            </tr>
            <tr>
                <td class="td_light">2.您對課程安排內容銜接情形是否滿意？
                </td>
            </tr>
            <tr>
                <td class="whitecol">
                    <asp:RadioButtonList ID="Q1_2" runat="server" CssClass="font" RepeatDirection="Horizontal" Width="80%">
                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                        <asp:ListItem Value="2">滿意</asp:ListItem>
                        <asp:ListItem Value="3">尚可</asp:ListItem>
                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td></td>
            </tr>
            <tr>
                <td class="td_light">3.您對課程內容時數安排是否滿意？
                </td>
            </tr>
            <tr>
                <td class="whitecol">
                    <asp:RadioButtonList ID="Q1_3" runat="server" CssClass="font" RepeatDirection="Horizontal" Width="80%">
                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                        <asp:ListItem Value="2">滿意</asp:ListItem>
                        <asp:ListItem Value="3">尚可</asp:ListItem>
                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td></td>
            </tr>
            <tr>
                <td align="center" class="table_title">二、師資、助教及教學
                </td>
            </tr>
            <tr>
                <td></td>
            </tr>
            <tr>
                <td class="td_light">1.您對老師的教學方式是否滿意？
                </td>
            </tr>
            <tr>
                <td class="whitecol">
                    <asp:RadioButtonList ID="Q2_1" runat="server" CssClass="font" RepeatDirection="Horizontal" Width="80%">
                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                        <asp:ListItem Value="2">滿意</asp:ListItem>
                        <asp:ListItem Value="3">尚可</asp:ListItem>
                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td></td>
            </tr>
            <tr>
                <td class="td_light">2.您對老師的教學態度是否滿意？
                </td>
            </tr>
            <tr>
                <td class="whitecol">
                    <asp:RadioButtonList ID="Q2_2" runat="server" CssClass="font" RepeatDirection="Horizontal" Width="80%">
                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                        <asp:ListItem Value="2">滿意</asp:ListItem>
                        <asp:ListItem Value="3">尚可</asp:ListItem>
                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td></td>
            </tr>
            <tr>
                <td class="td_light">3.您對老師的專業知識是否滿意？
                </td>
            </tr>
            <tr>
                <td class="whitecol">
                    <asp:RadioButtonList ID="Q2_3" runat="server" CssClass="font" RepeatDirection="Horizontal" Width="80%">
                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                        <asp:ListItem Value="2">滿意</asp:ListItem>
                        <asp:ListItem Value="3">尚可</asp:ListItem>
                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="td_light">4.您對助教的協助教學是否滿意? (無助教不須勾選)
                </td>
            </tr>
            <tr>
                <td class="whitecol">
                    <asp:RadioButtonList ID="Q2_4" runat="server" CssClass="font" RepeatDirection="Horizontal" Width="80%">
                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                        <asp:ListItem Value="2">滿意</asp:ListItem>
                        <asp:ListItem Value="3">尚可</asp:ListItem>
                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>

            <tr>
                <td></td>
            </tr>
            <tr>
                <td align="center" class="table_title">三、設備和教材
                </td>
            </tr>
            <tr>
                <td></td>
            </tr>
            <tr>
                <td class="td_light">1.您對上課期間，教材設備的充分利用情形是否滿意？
                </td>
            </tr>
            <tr>
                <td class="whitecol">
                    <asp:RadioButtonList ID="Q3_1" runat="server" CssClass="font" RepeatDirection="Horizontal" Width="80%">
                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                        <asp:ListItem Value="2">滿意</asp:ListItem>
                        <asp:ListItem Value="3">尚可</asp:ListItem>
                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td></td>
            </tr>
            <tr>
                <td class="td_light">2.您對教材的新穎是否滿意？
                </td>
            </tr>
            <tr>
                <td class="whitecol">
                    <asp:RadioButtonList ID="Q3_2" runat="server" CssClass="font" RepeatDirection="Horizontal" Width="80%">
                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                        <asp:ListItem Value="2">滿意</asp:ListItem>
                        <asp:ListItem Value="3">尚可</asp:ListItem>
                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td></td>
            </tr>
            <tr>
                <td class="td_light">3.您對教材內容的難易程度是否滿意？
                </td>
            </tr>
            <tr>
                <td class="whitecol">
                    <asp:RadioButtonList ID="Q3_3" runat="server" CssClass="font" RepeatDirection="Horizontal" Width="80%">
                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                        <asp:ListItem Value="2">滿意</asp:ListItem>
                        <asp:ListItem Value="3">尚可</asp:ListItem>
                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td></td>
            </tr>
            <tr>
                <td align="center" class="table_title">四、行政措施
                </td>
            </tr>
            <tr>
                <td></td>
            </tr>
            <tr>
                <td class="td_light">1.您對導師關心學員學習狀況及解決問題的能力是否滿意？
                </td>
            </tr>
            <tr>
                <td class="whitecol">
                    <asp:RadioButtonList ID="Q4_1" runat="server" CssClass="font" RepeatDirection="Horizontal" Width="80%">
                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                        <asp:ListItem Value="2">滿意</asp:ListItem>
                        <asp:ListItem Value="3">尚可</asp:ListItem>
                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td></td>
            </tr>
            <tr>
                <td class="td_light">2.您對求助及申訴管道是否滿意？
                </td>
            </tr>
            <tr>
                <td class="whitecol">
                    <asp:RadioButtonList ID="Q4_2" runat="server" CssClass="font" RepeatDirection="Horizontal" Width="80%">
                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                        <asp:ListItem Value="2">滿意</asp:ListItem>
                        <asp:ListItem Value="3">尚可</asp:ListItem>
                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td></td>
            </tr>
            <tr>
                <td class="td_light">3.您對學習空間是否滿意？
                </td>
            </tr>
            <tr>
                <td class="whitecol">
                    <asp:RadioButtonList ID="Q4_3" runat="server" CssClass="font" RepeatDirection="Horizontal" Width="80%">
                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                        <asp:ListItem Value="2">滿意</asp:ListItem>
                        <asp:ListItem Value="3">尚可</asp:ListItem>
                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td></td>
            </tr>
            <tr>
                <td class="td_light">4.您對上課地點場地及週遭環境清潔是否滿意？
                </td>
            </tr>
            <tr>
                <td class="whitecol">
                    <asp:RadioButtonList ID="Q4_4" runat="server" RepeatDirection="Horizontal" Width="80%">
                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                        <asp:ListItem Value="2">滿意</asp:ListItem>
                        <asp:ListItem Value="3">尚可</asp:ListItem>
                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td></td>
            </tr>
            <tr>
                <td class="table_title">貳、其它意見：</td>
            </tr>
            <tr>
                <td align="center" class="whitecol">
                    <br>
                    <asp:TextBox ID="txt_Suggestion" runat="server" TextMode="MultiLine" Width="80%" Rows="10"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td align="center" class="whitecol">
                    <asp:Button ID="bt_back" runat="server" Text="回上頁" CssClass="asp_button_S"></asp:Button>
                    <asp:Button ID="bt_save" runat="server" Text="儲存" CssClass="asp_button_S"></asp:Button>
                </td>
            </tr>
        </table>
        <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
        <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
        <input id="Hid_socid" type="hidden" runat="server" />
    </form>
</body>
</html>
