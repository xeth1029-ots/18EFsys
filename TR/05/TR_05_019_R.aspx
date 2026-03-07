<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_05_019_R.aspx.vb" Inherits="WDAIIP.TR_05_019_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>推動辦理身心障礙者職業訓練執行情形</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        //檢查列印條件為
        function CheckPrint() {
            var msg = '';

            if (document.form1.STDate1.value != '') {
                if (!checkDate(document.form1.STDate1.value)) msg += '[開訓區間的起日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }
            if (document.form1.STDate2.value != '') {
                if (!checkDate(document.form1.STDate2.value)) msg += '[開訓區間的迄日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }

            if (document.form1.FTDate1.value != '') {
                if (!checkDate(document.form1.FTDate1.value)) msg += '[結訓區間的起日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }
            if (document.form1.FTDate2.value != '') {
                if (!checkDate(document.form1.FTDate2.value)) msg += '[結訓區間的迄日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }

            if (msg == '') {
                if (document.form1.STDate2.value != '' && document.form1.STDate1.value != '' && document.form1.STDate2.value < document.form1.STDate1.value)
                { msg += '[開訓區間的迄日]必需大於[開訓區間的起日]\n'; }
            }
            if (msg == '') {
                if (document.form1.FTDate2.value != '' && document.form1.FTDate1.value != '' && document.form1.FTDate2.value < document.form1.FTDate1.value)
                { msg += '[結訓區間的迄日]必需大於[結訓區間的起日]\n'; }
            }

            if (document.form1.STDate1.value == ''
		    		&& document.form1.STDate2.value == ''
		    		&& document.form1.FTDate1.value == ''
		    		&& document.form1.FTDate2.value == ''
		    		&& document.form1.Syear.selectedIndex == 0)

            { msg += '[年度]、[開訓區間]、[結訓區間], 請擇一輸入查詢\n'; }

            //var Identity1=getCheckBoxListValue('Identity'); 
            var DistID1 = getCheckBoxListValue('DistID');
            var TPlanID1 = getCheckBoxListValue('TPlanID');

            if (parseInt(DistID1) == 0) { msg += '請選擇轄區\n'; }
            if (parseInt(TPlanID1) == 0) { msg += '請選擇計畫\n'; }

            if (msg != '') {
                alert(msg);
                return false;
            }
        }


        //選擇全部
        function SelectAll(obj, hidobj) {
            var num = getCheckBoxListValue(obj).length; //長度
            var myallcheck = document.getElementById(obj + '_' + 0); //第1個

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
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁>>訓練與需求管理>>統計分析>>推動辦理身心障礙者職業訓練執行情形</asp:Label>
                </td>
            </tr>
        </table>

        <table id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td class="bluecol" style="width: 20%">年度
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="Syear" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">開訓區間
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="STDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                    <img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top">~
                            <asp:TextBox ID="STDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                    <img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top">
                </td>
            </tr>
            <tr>
                <td class="bluecol">結訓區間
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="FTDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                    <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top">~
                            <asp:TextBox ID="FTDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                    <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top">
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">轄區
                </td>
                <td class="whitecol">
                    <asp:CheckBoxList ID="DistID" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="3">
                    </asp:CheckBoxList>
                    <input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server">
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">訓練計畫
                </td>
                <td class="whitecol">
                    <asp:CheckBoxList ID="TPlanID" runat="server" RepeatDirection="Horizontal" CssClass="font" CellSpacing="0" CellPadding="0" RepeatColumns="3">
                    </asp:CheckBoxList>
                    <input id="TPlanHidden" type="hidden" value="0" name="TPlanHidden" runat="server">
                </td>
            </tr>
            <tr>
                <td class="bluecol">計畫分類
                </td>
                <td class="whitecol">
                    <asp:RadioButtonList Style="z-index: 0" ID="rblPlanType" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                        <asp:ListItem Value="1" Selected="True">自辦</asp:ListItem>
                        <asp:ListItem Value="2">委辦</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>

            <%--<tr>
                <td class="bluecol">訓練類別
                </td>
                <td class="whitecol">
                    <asp:RadioButtonList Style="z-index: 0" ID="rblPropertyID" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                        <asp:ListItem Value="0" Selected="True">職前</asp:ListItem>
                        <asp:ListItem Value="1">在職</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>--%>
            <tr>
                <td class="bluecol">匯出種類
                </td>
                <td class="whitecol">
                    <asp:RadioButtonList Style="z-index: 0" ID="rblType1" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                        <asp:ListItem Value="1" Selected="True">融合式訓練辦理情形</asp:ListItem>
                        <asp:ListItem Value="2">融合式訓練職類統計</asp:ListItem>
                        <asp:ListItem Value="3">(專班)辦理情形</asp:ListItem>
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
            <tr>
                <td align="center" colspan="2" class="whitecol">
                    <asp:Button ID="Export1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                </td>
            </tr>

            <%--<tr><td><table class="table_sch" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server"></table></td></tr>--%>
        </table>
    </form>
</body>
</html>
