<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_05_012_R.aspx.vb" Inherits="WDAIIP.TR_05_012_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>原住民訓練人數統計表</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
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


        function search() {

            var msg = '';
            if (isEmpty(document.form1.start_date) && isEmpty(document.form1.end_date)) msg += '請選擇日期範圍!\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>

                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">
										首頁&gt;&gt;訓練與就業需求管理&gt;&gt;統計分析&gt;&gt;<FONT color="#990000">原住民訓練人數統計表</FONT>
                                </asp:Label>
                            </td>
                        </tr>
                    </table>

                    <table class="table_sch" id="Table2" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol_need" width="100">開訓期間
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="start_date" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer"
                                    onclick="javascript:show_calendar('start_date','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif"
                                    align="top" width="24" height="24">~<asp:TextBox ID="end_date" runat="server" Columns="10"></asp:TextBox><img
                                        style="cursor: pointer" onclick="javascript:show_calendar('end_date','','','CY/MM/DD');"
                                        alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練計畫
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="TPlanID" runat="server" CssClass="font" RepeatDirection="Horizontal"
                                    RepeatColumns="3">
                                </asp:CheckBoxList>
                                <input id="TPlanHidden" type="hidden" value="0" runat="server" name="TPlanHidden">
                            </td>
                        </tr>
                    </table>
                    <p align="center">
                        <asp:Button ID="Button1" runat="server" Text="列印"
                            CssClass="asp_Export_M"></asp:Button>
                    </p>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
