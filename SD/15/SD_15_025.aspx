<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_025.aspx.vb" Inherits="WDAIIP.SD_15_025" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>5＋N產業課程統計</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <%--<script type="text/javascript" language="javascript" src="../../js/date-picker2.js"></script>--%>
    <script type="text/javascript" language="javascript" src="../../js/date-picker2.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
       <%-- var bl_rocYear = "<%=ConfigurationManager.AppSettings("REPLACE2ROC_YEARS") %>";
        var scriptTag = document.createElement('script');
        var jsPath = (bl_rocYear == "Y" ? "../../js/date-picker2.js" : "../../js/date-picker.js");
        scriptTag.src = jsPath;
        document.head.appendChild(scriptTag);--%>

        //檢查列印條件為
        function CheckPrint() {
            var msg = '';
            if (document.form1.STDate1.value != '') {
                if (bl_rocYear == "Y") {
                    if (!checkRocDate(document.form1.STDate1.value)) msg += '[開訓區間的起日]不是正確的日期格式,請輸入正確的日期格式,YYY/MM/DD!!\n';
                }
                else {
                    if (!checkDate(document.form1.STDate1.value)) msg += '[開訓區間的起日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
                }
            }
            if (document.form1.STDate2.value != '') {
                if (bl_rocYear == "Y") {
                    if (!checkRocDate(document.form1.STDate2.value)) msg += '[開訓區間的迄日]不是正確的日期格式,請輸入正確的日期格式,YYY/MM/DD!!\n';
                }
                else {
                    if (!checkDate(document.form1.STDate2.value)) msg += '[開訓區間的迄日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
                }
            }
            if (document.form1.FTDate1.value != '') {
                if (bl_rocYear == "Y") {
                    if (!checkRocDate(document.form1.FTDate1.value)) msg += '[結訓區間的起日]不是正確的日期格式,請輸入正確的日期格式,YYY/MM/DD!!\n';
                }
                else {
                    if (!checkDate(document.form1.FTDate1.value)) msg += '[結訓區間的起日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
                }
            }
            if (document.form1.FTDate2.value != '') {
                if (bl_rocYear == "Y") {
                    if (!checkRocDate(document.form1.FTDate2.value)) msg += '[結訓區間的迄日]不是正確的日期格式,請輸入正確的日期格式,YYY/MM/DD!!\n';
                }
                else {
                    if (!checkDate(document.form1.FTDate2.value)) msg += '[結訓區間的迄日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
                }
            }
            if (msg == '') {
                if (document.form1.STDate2.value != '' && document.form1.STDate1.value != '' && document.form1.STDate2.value < document.form1.STDate1.value) {
                    msg += '[開訓區間的迄日]必需大於[開訓區間的起日]\n';
                }
                if (document.form1.FTDate2.value != '' && document.form1.FTDate1.value != '' && document.form1.FTDate2.value < document.form1.FTDate1.value) {
                    msg += '[結訓區間的迄日]必需大於[結訓區間的起日]\n';
                }
            }
            if (document.form1.STDate1.value == '' && document.form1.STDate2.value == '' && document.form1.FTDate1.value == '' && document.form1.FTDate2.value == '' && document.form1.Syear.selectedIndex == 0) {
                msg += '[年度]、[開訓區間]、[結訓區間], 請擇一輸入查詢\n';
            }
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
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練需求管理&gt;&gt;統計分析&gt;&gt;5＋N產業課程統計</asp:Label>
                </td>
            </tr>
        </table>
        <table id="myTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <table id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td class="bluecol" width="20%">年度</td>
                            <td class="whitecol" width="80%">
                                <asp:DropDownList ID="Syear" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol">開訓區間</td>
                            <td class="whitecol">
                                <asp:TextBox ID="STDate1" runat="server" Columns="10" Width="16%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top">
                                ~
                                <asp:TextBox ID="STDate2" runat="server" Columns="10" Width="16%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">結訓區間</td>
                            <td class="whitecol">
                                <asp:TextBox ID="FTDate1" runat="server" Columns="10" Width="16%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top">
                                ~
                                <asp:TextBox ID="FTDate2" runat="server" Columns="10" Width="16%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">轄區</td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="DistID" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="3"></asp:CheckBoxList>
                                <input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">訓練計畫</td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="TPlanID" runat="server" RepeatDirection="Horizontal" CssClass="font" CellSpacing="0" CellPadding="0" RepeatColumns="3"></asp:CheckBoxList>
                                <input id="TPlanHidden" type="hidden" value="0" name="TPlanHidden" runat="server">
                            </td>
                        </tr>
                        <tr id="trKID20" runat="server">
                            <td class="bluecol" width="20%">&nbsp; 政府政策性產業 </td>
                            <td class="whitecol" width="80%">
                                <table border="0" cellspacing="1" cellpadding="1" width="100%">
                                    <tr>
                                        <td class="bluecol" width="20%">「5+2」產業創新計畫</td>
                                        <td class="whitecol">
                                            <asp:CheckBoxList ID="CBLKID20_1" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">台灣AI行動計畫</td>
                                        <td class="whitecol">
                                            <asp:CheckBoxList ID="CBLKID20_2" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">數位國家創新經濟<br />
                                            發展方案</td>
                                        <td class="whitecol">
                                            <asp:CheckBoxList ID="CBLKID20_3" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">國家資通安全發展方案</td>
                                        <td class="whitecol">
                                            <asp:CheckBoxList ID="CBLKID20_4" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">前瞻基礎建設計畫</td>
                                        <td class="whitecol">
                                            <asp:CheckBoxList ID="CBLKID20_5" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">新南向政策</td>
                                        <td class="whitecol" width="80%">
                                            <asp:CheckBoxList ID="CBLKID20_6" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">進階政策性產業類別</td>
                                        <td class="whitecol">
                                            <asp:CheckBoxList ID="CBLKID22" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <%-- <tr>
                            <td class="bluecol">5＋N產業</td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="KID_5" runat="server" RepeatColumns="5" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                <input id="KID_5_hid" value="0" type="hidden" name="HID_DepID_6" runat="server" size="1">
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
                        <tr>
                            <td align="center" colspan="2">
                                <asp:Button ID="Print" runat="server" Text="列印" CssClass="asp_Export_M" Visible="false"></asp:Button>
                                <asp:Button ID="Export1" runat="server" Text="匯出明細" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
