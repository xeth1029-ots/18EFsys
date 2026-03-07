<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_04_004_R.aspx.vb" Inherits="WDAIIP.SD_04_004_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>列印課程表</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var cst_inline1 = "";  //'inline';
        function GETvalue() {
            document.getElementById('Button3').click();
        }

        function choose_class() {
            var RID = document.form1.RIDValue.value;
            openClass('../02/SD_02_ch.aspx?RID=' + RID + '&BtnName=Button2');
            //setRadioValue(document.form1.RadioButtonList1,2);
            //document.form1.start_date.value=document.form1.SDate_month.value;
            //document.form1.end_date.value=document.form1.FDate_month.value;
            //window.open('../02/SD_02_ch.aspx?special=1','','width=550,height=250,location=0,status=0,menubar=0,scrollbars=1,resizable=0');
        }

        function ReportPrint() {
            var msg = '';
            if (getRadioValue(document.form1.RadioButtonList1) != 3) {
                if (document.form1.OCIDValue1.value == '') msg += '請選擇班級職類\n';
            }
            if (!isChecked(document.form1.RadioButtonList1))
                msg += '請選擇範圍\n';
            else {
                switch (parseInt(getRadioValue(document.getElementsByName('RadioButtonList1')))) {
                    case 0:
                        if (document.getElementById('start_date').value == '') msg += '請輸入起始日期\n';
                        else if (!checkDate(document.getElementById('start_date').value)) msg += '起始日期不是正確的時間格式\n';
                        if (document.getElementById('end_date').value == '') msg += '請輸入結束日期\n';
                        else if (!checkDate(document.getElementById('end_date').value)) msg += '結束日期不是正確的時間格式\n';
                        break;
                    case 1:
                        if (document.getElementById('Years').selectedIndex == 0) msg += '請選擇年度\n';
                        break;
                    case 2:
                        if (document.getElementById('Years').selectedIndex == 0) msg += '請選擇年度\n';
                        if (document.getElementById('Months').selectedIndex == 0) msg += '請選擇月份\n';
                        break;
                    case 3:
                        if (document.getElementById('DecDay').value == '') msg += '請輸入結束日期\n';
                        else if (!checkDate(document.getElementById('DecDay').value)) msg += '結束日期不是正確的時間格式\n';
                        break;
                    case 4:
                        if (document.getElementById('s_date').value == '') msg += '請輸入起始日期\n';
                        else if (!checkDate(document.getElementById('s_date').value)) msg += '起始日期不是正確的時間格式\n';
                        if (document.getElementById('e_date').value == '') msg += '請輸入結束日期\n';
                        else if (!checkDate(document.getElementById('e_date').value)) msg += '結束日期不是正確的時間格式\n';
                        break;
                }
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function ChangeMode() {
            document.getElementById('TR1').style.display = 'none';
            document.getElementById('TR2').style.display = 'none';
            document.getElementById('TR3').style.display = 'none';
            document.getElementById('TR4').style.display = 'none';
            document.getElementById('TR5').style.display = 'none';
            document.getElementById('TR6').style.display = 'none';
            switch (getRadioValue(document.getElementsByName('RadioButtonList1'))) {
                case '0':
                    document.getElementById('TR1').style.display = cst_inline1;  //'inline';
                    break;
                case '1':
                    document.getElementById('TR2').style.display = cst_inline1;  //'inline';
                    break;
                case '2':
                    document.getElementById('TR2').style.display = cst_inline1;  //'inline';
                    document.getElementById('TR3').style.display = cst_inline1;  //'inline';
                    document.getElementById('TR5').style.display = cst_inline1;  //'inline';
                    break;
                case '3':
                    document.getElementById('TR4').style.display = cst_inline1;  //'inline';
                    break;
                case '4':
                    document.getElementById('TR6').style.display = cst_inline1;  //'inline';
                    break;
            }
        }

        function CleanClass() {
            document.getElementById('TMID1').value = '';
            document.getElementById('TMIDValue1').value = '';
            document.getElementById('OCID1').value = '';
            document.getElementById('OCIDValue1').value = '';
            document.getElementById('start_date').value = '';
            document.getElementById('end_date').value = '';
            document.getElementById('s_date').value = '';
            document.getElementById('e_date').value = '';
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;課程管理&gt;&gt;列印課程表</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" class="table_nw" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <table id="Table3" class="table_nw" width="100%" cellpadding="1" cellspacing="1">
                        <tr>
                            <td class="bluecol" width="20%">訓練機構</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="center" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="Button8" value="..." type="button" name="Button8" runat="server" class="button_b_Mini" />
                                <input style="width: 6%;" id="OrgidValue" size="5" type="hidden" name="OrgidValue" runat="server" />
                                <asp:Button Style="display: none" ID="Button3" runat="server" Text="Button3"></asp:Button>
                                <span style="position: absolute; display: none" id="HistoryList2" onclick="GETvalue()"><asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="20%">班別/職類</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" value="..." type="button" class="button_b_Mini" />
                                <input id="OCIDValue1" type="hidden" name="Hidden2" runat="server" />
                                <input id="TMIDValue1" type="hidden" name="Hidden1" runat="server" />
                                <asp:Button Style="display: none" ID="Button2" runat="server" Text="查詢班級資料(隱藏)"></asp:Button>
                                <input onclick="CleanClass();" value="清除" type="button" />
                                <span style="position: absolute; display: none; left: 50%" id="HistoryList"><asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="20%">範圍</td>
                            <td class="whitecol" width="80%">
                                <asp:RadioButtonList ID="RadioButtonList1" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="0">全期</asp:ListItem>
                                    <asp:ListItem Value="1">年</asp:ListItem>
                                    <asp:ListItem Value="2" Selected="True">月</asp:ListItem>
                                    <asp:ListItem Value="4">週</asp:ListItem>
                                    <asp:ListItem Value="3">日</asp:ListItem>
                                </asp:RadioButtonList>
                                (以「日」為範圍的時候可以不用選擇班級)
                            </td>
                        </tr>
                        <tr id="TR1" runat="server">
                            <td class="bluecol" width="20%">日期</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="start_date" runat="server" Width="15%"></asp:TextBox>
                                <span id="span1" runat="server"><img style="cursor: pointer" onclick="PublicCalendar('start_date');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30" /></span> ～
                                <asp:TextBox ID="end_date" runat="server" Width="15%"></asp:TextBox>
                                <span id="span2" runat="server"><img style="cursor: pointer" onclick="PublicCalendar('end_date');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30" /></span>
                            </td>
                        </tr>
                        <tr id="TR2" runat="server">
                            <td class="bluecol" width="20%">列印年度</td>
                            <td class="whitecol" width="80%"><asp:DropDownList ID="Years" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr id="TR3" runat="server">
                            <td class="bluecol" width="20%">列印月份</td>
                            <td class="whitecol" width="80%"><asp:DropDownList ID="Months" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr id="TR4" runat="server">
                            <td class="bluecol" width="20%">列印日期</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="DecDay" runat="server" Width="15%"></asp:TextBox>
                                <span id="span3" runat="server"><img style="cursor: pointer" onclick="PublicCalendar('DecDay');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30" /></span>
                            </td>
                        </tr>
                        <tr id="TR5" runat="server">
                            <td class="bluecol" width="20%">列印方式</td>
                            <td class="whitecol" width="80%">
                                <asp:RadioButtonList ID="Monthlist" runat="server" Width="200px" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="1" Selected="True">不分頁</asp:ListItem>
                                    <asp:ListItem Value="0">分二頁</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="TR6" runat="server">
                            <td class="bluecol" width="20%">日期</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="s_date" runat="server" Width="15%"></asp:TextBox>
                                <span id="span4" runat="server"><img style="cursor: pointer" onclick="PublicCalendar('s_date');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30" /></span> ～
                                <asp:TextBox ID="e_date" runat="server" Width="15%"></asp:TextBox>
                                <span id="span5" runat="server"><img style="cursor: pointer" onclick="PublicCalendar('e_date');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30" /></span>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td class="whitecol" align="center">
                                <div align="center"><asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button></div>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>