<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CM_03_013.aspx.vb" Inherits="WDAIIP.CM_03_013" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>開訓、離退訓、結訓人數統計表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        //檢查列印條件為
        function CheckPrint() {
            var msg = '';
            if (document.form1.FTDate1.value != '') {
                if (!checkDate(document.form1.FTDate1.value)) {
                    msg += '[結訓區間 起日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
                }
            }

            if (document.form1.FTDate2.value != '') {
                if (!checkDate(document.form1.FTDate2.value)) {
                    msg += '[結訓區間 迄日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
                }
            }
            if (msg == ''
		     		&& document.form1.FTDate2.value != ''
		     		&& document.form1.FTDate1.value != ''
		     		&& document.form1.FTDate2.value < document.form1.FTDate1.value) {
                msg += '[結訓區間的迄日]必需大於[結訓區間的起日]\n';
            }

            var DistID1 = getCheckBoxListValue('DistID');
            var TPlanID1 = getCheckBoxListValue('TPlanID');

            if (parseInt(DistID1) == 0) {
                msg += '請選擇轄區\n';
            }
            if (parseInt(TPlanID1) == 0) {
                msg += '請選擇計畫\n';
            }

            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        //選擇全部 與未選擇全部
        function SelectAll(obj, hidobj) {
            var num = getCheckBoxListValue(obj).length;
            var myallcheck = document.getElementById(obj + '_' + 0); //第0個選項

            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
            else {
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    if ('0' == getCheckBoxListValue(obj).charAt(i)) {
                        document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(i);
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
        <p>
            <table id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
                <tr>
                    <td>
                        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <tr>
                                <td>
                                    <asp:Label ID="TitleLab1" runat="server"></asp:Label><asp:Label ID="TitleLab2" runat="server">首頁>>訓練與需求管理>>統計分析>>開訓、離退訓、結訓人數統計表</asp:Label>
                                </td>
                            </tr>
                        </table>
                        <table class="table_sch" id="Table2" runat="server" cellspacing="1" cellpadding="1">
                            <tr>
                                <td class="bluecol_need" width="80">年度
                                </td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="Syear" runat="server">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">結訓區間
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="FTDate1" runat="server" MaxLength="10" Columns="10"></asp:TextBox>
                                    <span runat="server">
                                        <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span> ~
                                <asp:TextBox ID="FTDate2" runat="server" MaxLength="10" Columns="10"></asp:TextBox>
                                    <span runat="server">
                                        <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">轄區
                                </td>
                                <td class="whitecol">
                                    <asp:CheckBoxList ID="DistID" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                                    </asp:CheckBoxList>
                                    <input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server">
                                </td>
                            </tr>
                        </table>
                        <p align="center">
                            <asp:Button ID="Print" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                        </p>
                    </td>
                </tr>
            </table>
        </p>
    </form>
</body>
</html>
