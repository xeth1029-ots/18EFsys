<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_09_012_R.aspx.vb" Inherits="WDAIIP.SD_09_012_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_09_012_R</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1" />
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1" />
    <meta name="vs_defaultClientScript" content="JavaScript" />
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
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
        }

        function print() {
            var msg = '';
            if (document.form1.start_date.value == '') msg += '請選擇退訓日期起日!!\n';
            if (document.form1.end_date.value == '') msg += '請選擇退訓日期迄日!!\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function check_date() {
            if (!checkDate(form1.start_date.value) || !checkDate(form1.end_date.value)) {
                document.form1.start_date.value = '';
                document.form1.end_date.value = '';
                alert('請輸入正確的日期格式,YYYY/MM/DD!!\n');
            }

        }		
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table id="Table1" cellspacing="1" cellpadding="1" width="740" border="0">
        <tr>
            <td>
                <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            首頁&gt;&gt;學員動態管理&gt;&gt;教務報表管理&gt;&gt; <font color="#990000">學員退訓賠償未結案件統計表</font>
                        </td>
                    </tr>
                </table>
                <table class="table_sch" id="Table3" cellspacing="1" cellpadding="1">
                    <tr>
                        <td class="bluecol_need" width="100">退訓日期</td>
                        <td class="whitecol">
                            <asp:TextBox ID="start_date" runat="server" Width="100px"></asp:TextBox>
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');"
                                alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                            ～
                            <asp:TextBox ID="end_date" runat="server" Width="100px"></asp:TextBox>
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');"
                                alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol" width="100">訓練機構</td>
                        <td class="whitecol">
                            <asp:CheckBoxList ID="DistrictList" runat="server" CssClass="font" RepeatDirection="Horizontal"
                                Width="512px" Height="11px">
                            </asp:CheckBoxList>
                            <input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server" />
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol" width="100">訓練計畫</td>
                        <td class="whitecol">
                            <asp:CheckBoxList ID="PlanList" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3">
                            </asp:CheckBoxList>
                            <input id="TPlanHidden" type="hidden" value="0" name="TPlanHidden" runat="server" class="button_b_Mini" />
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2" class="whitecol">
                            <font class="font">說明：<font color="red"> ＊學員退訓賠償從2008年度開始</font></font>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <br />
    <div style="width:600" align="center">
        <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
    </div>
    </form>
</body>
</html>
