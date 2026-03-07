<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_05_018_R.aspx.vb" Inherits="WDAIIP.TR_05_018_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>四大、六大、十大產業課程統計表</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../style.css" type="text/css" rel="stylesheet">
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
            var num = getCheckBoxListValue(obj).length;//長度
            var myallcheck = document.getElementById(obj + '_' + 0);//第1個

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
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="600" border="0">
            <tr>
                <td>
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁>>訓練與需求管理>>統計分析>>四大、六大、十大產業課程統計表</asp:Label></td>
                        </tr>
                    </table>
                    <table class="font" id="Table2" style="width: 696px; height: 285px" cellspacing="1" cellpadding="1"
                        width="696" border="0" runat="server">
                        <tr>
                            <td class="CM_TD1" width="130">&nbsp;&nbsp;&nbsp; 年度</td>
                            <td class="CM_TD2">
                                <asp:DropDownList ID="Syear" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="CM_TD1" style="height: 28px" width="80"><font face="新細明體">&nbsp;&nbsp;&nbsp; 
										開訓區間</font></td>
                            <td class="CM_TD2" style="height: 28px">
                                <asp:TextBox ID="STDate1" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');"
                                    alt="" src="../../images/show-calendar.gif" align="top">~<asp:TextBox ID="STDate2" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');"
                                        alt="" src="../../images/show-calendar.gif" align="top"></td>
                        </tr>
                        <tr>
                            <td class="CM_TD1"><font face="新細明體">&nbsp;&nbsp;&nbsp; 結訓區間</font></td>
                            <td class="CM_TD2">
                                <asp:TextBox ID="FTDate1" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');"
                                    alt="" src="../../images/show-calendar.gif" align="top">~<asp:TextBox ID="FTDate2" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');"
                                        alt="" src="../../images/show-calendar.gif" align="top"></td>
                        </tr>
                        <tr>
                            <td class="CM_TD1"><font face="新細明體">&nbsp;&nbsp;&nbsp; 轄區 <font color="#ff0000">*</font></font></td>
                            <td class="CM_TD2">
                                <asp:CheckBoxList ID="DistID" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font"></asp:CheckBoxList><input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="CM_TD1"><font face="新細明體">&nbsp;&nbsp;&nbsp; 訓練計畫 <font color="#ff0000">*</font></font></td>
                            <td class="CM_TD2">
                                <asp:CheckBoxList ID="TPlanID" runat="server" RepeatDirection="Horizontal" CssClass="font" CellSpacing="0"
                                    CellPadding="0" RepeatColumns="3">
                                </asp:CheckBoxList><input id="TPlanHidden" type="hidden" value="0" name="TPlanHidden" runat="server">
                            </td>
                        </tr>
                        <tr id="KID_6_TR" runat="server">
                            <td class="CM_TD1">&nbsp;&nbsp;&nbsp;六大新興產業</td>
                            <td class="CM_TD2">
                                <asp:CheckBoxList ID="KID_6" runat="server" RepeatColumns="5" RepeatDirection="Horizontal" Font-Size="X-Small"></asp:CheckBoxList>
                                <input id="KID_6_hid" value="0" type="hidden" name="HID_DepID_6" runat="server" size="1">
                            </td>
                        </tr>
                        <tr id="KID_10_TR" runat="server">
                            <td class="CM_TD1">&nbsp;&nbsp;&nbsp;十大重點服務業</td>
                            <td class="CM_TD2">
                                <asp:CheckBoxList ID="KID_10" runat="server" RepeatColumns="4" RepeatDirection="Horizontal" Font-Size="X-Small"></asp:CheckBoxList>
                                <input id="KID_10_hid" value="0" type="hidden" name="HID_DepID_6" runat="server" size="1">
                            </td>
                        </tr>
                        <tr id="KID_4_TR" runat="server">
                            <td class="CM_TD1">&nbsp;&nbsp; 四大新興智慧型產業</td>
                            <td class="CM_TD2">
                                <asp:CheckBoxList ID="KID_4" runat="server" RepeatColumns="5" RepeatDirection="Horizontal" Font-Size="X-Small"></asp:CheckBoxList>
                                <input id="KID_4_hid" value="0" type="hidden" name="HID_DepID_6" runat="server" size="1">
                            </td>
                        </tr>
                        <tr>
                            <td class="CM_TD1">&nbsp;&nbsp;&nbsp; 預算來源
                            </td>
                            <td class="CM_TD2">
                                <asp:CheckBoxList ID="BudgetList" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal"
                                    CssClass="font">
                                </asp:CheckBoxList><input id="hidBudgetList" value="0" type="hidden" name="hidBudgetList" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="2"><font face="新細明體">
                                <asp:Button ID="Print" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>&nbsp;
										<asp:Button ID="Export1" runat="server" Text="匯出明細" CssClass="asp_Export_M"></asp:Button></font></td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
