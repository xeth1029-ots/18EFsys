<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CM_03_007.aspx.vb" Inherits="WDAIIP.CM_03_007" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>主要特定對象統計表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }

        //檢查列印條件為
        function CheckPrint() {
            var STDate1 = document.getElementById('STDate1');
            var FTDate1 = document.getElementById('FTDate1');
            var STDate2 = document.getElementById('STDate2');
            var FTDate2 = document.getElementById('FTDate2');
            var msg = '';

            if (STDate1.value != '') {
                if (!checkDate(STDate1.value)) msg += '[開訓區間的起日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }
            if (FTDate1.value != '') {
                if (!checkDate(FTDate1.value)) msg += '[結訓區間的起日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }
            if (STDate2.value != '') {
                if (!checkDate(STDate2.value)) msg += '[開訓區間的迄日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }
            if (FTDate2.value != '') {
                if (!checkDate(FTDate2.value)) msg += '[結訓區間的迄日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }

            if (STDate2.value != '' && STDate1.value != '' && STDate2.value < STDate1.value)
            { msg += '[開訓區間的迄日]必需大於[開訓區間的起日]\n'; }
            if (FTDate2.value != '' && FTDate1.value != '' && FTDate2.value < FTDate1.value)
            { msg += '[結訓區間的迄日]必需大於[結訓區間的起日]\n'; }

            if (STDate1.value == '' && STDate2.value == '' && FTDate1.value == '' && FTDate2.value == '' && document.form1.Syear.selectedIndex == 0)
            { msg += '[年度]、[開訓區間]、[結訓區間],請擇一輸入查詢\n'; }

            var Identity1 = getCheckBoxListValue('Identity');
            var DistID1 = getCheckBoxListValue('DistID');
            var TPlanID1 = getCheckBoxListValue('TPlanID');

            if (parseInt(DistID1) == 0)
            { msg += '請選擇轄區\n'; }
            if (parseInt(TPlanID1) == 0)
            { msg += '請選擇計畫\n'; }

            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        //檢查統計項目若是身分別就把身分別隱藏					
        function ChangeMode() {
            var IdentityTR = document.getElementById('IdentityTR');
            //IdentityTR.style.display = 'inline';
            IdentityTR.style.display = '';
            if (getRadioValue(document.getElementsByName('rblMode1')) == '0') {
                IdentityTR.style.display = 'none';
            }
            //debugger;
            if (document.body) { window.scroll(0, document.body.scrollHeight); }
            if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        }

        //選擇全部
        function SelectAll(obj, hidobj) {
            var num = getCheckBoxListValue(obj).length;
            var myallcheck = document.getElementById(obj + '_' + 0);

            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
            else {
                for (var i = 1; i < num; i++) {
                    if ('0' == getCheckBoxListValue(obj).charAt(i)) {
                        document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(i);
                        var mycheck = document.getElementById(obj + '_' + i);
                        myallcheck.checked = mycheck.checked;
                        break;
                    }
                }
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
                    <%--  首頁&gt;&gt;訓練需求管理&gt;&gt;統計分析&gt;&gt;主要特定對象統計表--%>
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
                <td class="bluecol">年度 </td>
                <td class="whitecol">
                    <asp:DropDownList ID="Syear" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">開訓區間 </td>
                <td class="whitecol">
                    <asp:TextBox ID="STDate1" runat="server" Columns="10"></asp:TextBox>
                    <span runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    ~<asp:TextBox ID="STDate2" runat="server" Columns="10"></asp:TextBox>
                    <span runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                </td>
            </tr>
            <tr>
                <td class="bluecol">結訓區間 </td>
                <td class="whitecol">
                    <asp:TextBox ID="FTDate1" runat="server" Columns="10"></asp:TextBox>
                    <span runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    ~<asp:TextBox ID="FTDate2" runat="server" Columns="10"></asp:TextBox>
                    <span runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">查詢方式 </td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="rblSchType1" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                        <asp:ListItem Value="AND" Selected="True">AND(資料範圍同時包含開訓區間與結訓區間)</asp:ListItem>
                        <asp:ListItem Value="OR">OR(資料範圍分別符合開訓日期或結訓日期)</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">轄區 </td>
                <td class="whitecol">
                    <asp:CheckBoxList ID="DistID" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                    </asp:CheckBoxList>
                    <input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server">
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">訓練計畫 </td>
                <td class="whitecol">
                    <asp:CheckBoxList ID="TPlanID" runat="server" RepeatDirection="Horizontal" CssClass="font" CellSpacing="0" CellPadding="0" RepeatColumns="3">
                    </asp:CheckBoxList>
                    <input id="TPlanHidden" type="hidden" value="0" name="TPlanHidden" runat="server">
                </td>
            </tr>
            <tr>
                <td class="bluecol">預算來源 </td>
                <td class="whitecol">
                    <asp:CheckBoxList ID="BudgetList" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                    </asp:CheckBoxList>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">統計項目 </td>
                <td class="whitecol" style="height: 20px">
                    <asp:RadioButtonList ID="rblMode1" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                        <asp:ListItem Value="0" Selected="True">身分別</asp:ListItem>
                        <asp:ListItem Value="1">年齡</asp:ListItem>
                        <asp:ListItem Value="2">訓練職類</asp:ListItem>
                        <asp:ListItem Value="3">教育程度</asp:ListItem>
                        <asp:ListItem Value="4">性別</asp:ListItem>
                        <asp:ListItem Value="5">通俗職類</asp:ListItem>
                        <asp:ListItem Value="7">就職狀況</asp:ListItem>
                        <asp:ListItem Value="11">年齡2</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr id="IdentityTR" runat="server">
                <td class="bluecol">身分別 </td>
                <td class="whitecol">
                    <asp:CheckBoxList ID="Identity" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="4">
                    </asp:CheckBoxList>
                    <input id="Identity_List" type="hidden" value="0" name="Identity_List" runat="server" />
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

            <tr>
                <td colspan="2" align="center">
                    <asp:Button ID="Print" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                    &nbsp;<asp:Button ID="BtnExp" runat="server" CssClass="asp_Export_M" Text="匯出統計資料" />
                    &nbsp;<asp:Button ID="BtnExp2" runat="server" CssClass="asp_Export_M" Text="匯出班級明細資料" />
                    <%--<table class="table_sch" id="Table2" runat="server" cellspacing="1" cellpadding="1">
                    </table>
                    <p align="center">
                    </p>--%>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
