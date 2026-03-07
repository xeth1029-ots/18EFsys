<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_03_013.aspx.vb" Inherits="WDAIIP.SYS_03_013" %>

<!DOCTYPE html PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>學員資料整批匯出</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
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
                    var mycheck = document.getElementById(obj + '_' + i);
                    if ('0' == getCheckBoxListValue(obj).charAt(i)) {
                        document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(i);
                        myallcheck.checked = mycheck.checked;
                        break;
                    }
                }
            }
        }

        function search() {
            var msg = '';

            if (document.form1.Stdate1.value != '') {
                if (!checkDate(document.form1.Stdate1.value))
                    msg += '[開訓期間起日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }

            if (document.form1.Stdate2.value != '') {
                if (!checkDate(document.form1.Stdate2.value))
                    msg += '[開訓期間迄日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }

            if (document.form1.Stdate2.value != '' && document.form1.Stdate1.value != '' && document.form1.Stdate2.value < document.form1.Stdate1.value) {
                msg += '[開訓期間迄日]必需大於[開訓期間起日]\n';
            }

            //'限定輸入取消 by AMU 20091006
            //if (document.form1.Stdate1.value == '' && document.form1.Stdate2.value == '') { 
            //	msg += '請輸入[開訓期間]的起迄日期\n';
            //}　

            if (document.form1.Ftdate1.value != '') {
                if (!checkDate(document.form1.Ftdate1.value))
                    msg += '[結訓期間起日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }

            if (document.form1.Ftdate2.value != '') {
                if (!checkDate(document.form1.Ftdate2.value))
                    msg += '[結訓期間迄日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }

            if (document.form1.Ftdate2.value != '' && document.form1.Ftdate1.value != '' && document.form1.Ftdate2.value < document.form1.Ftdate1.value) {
                msg += '[結訓期間迄日]必需大於[結訓期間起日]\n';
            }

            if (msg != '') {
                alert(msg);
                return false;
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;功能權限管理&gt;&gt;學員資料整批匯出</asp:Label>
                </td>
            </tr>
        </table>
    <table id="table1" cellspacing="1" cellpadding="1" width="100%" border="0">
        <tr>
            <td>
                <%--<table class="font" id="table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            <asp:Label ID="TitleLab1" runat="server"></asp:Label><asp:Label ID="TitleLab2" runat="server">
										首頁&gt;&gt;系統管理&gt;&gt;功能權限管理&gt;&gt;<font style="COLOR: brown">學員資料整批匯出</font></asp:Label>
                        </td>
                    </tr>
                </table>--%>
                <table class="table_nw" id="table3" cellspacing="1" cellpadding="1" width="100%">
                    <tr>
                        <td class="bluecol" style="width:20%">
                            年度
                        </td>
                        <td class="whitecol">
                            <asp:DropDownList ID="yearlist" runat="server">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            轄區
                        </td>
                        <td class="whitecol" colspan="4">
                            <asp:CheckBoxList ID="DistrictList" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                            </asp:CheckBoxList>
                            <input id="DistHidden" type="hidden" size="2" value="0" name="DistHidden" runat="server">
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            訓練計畫
                        </td>
                        <td class="whitecol">
                            <asp:CheckBoxList ID="PlanList" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3" CellPadding="0" CellSpacing="0">
                            </asp:CheckBoxList>
                            <input id="TPlanHidden" type="hidden" size="2" value="0" name="TPlanHidden" runat="server">
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            開訓期間
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="Stdate1" runat="server" Columns="10" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('Stdate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />～
                            <asp:TextBox ID="Stdate2" runat="server" Columns="10" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('Stdate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            結訓期間
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="Ftdate1" runat="server" Columns="10" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('Ftdate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />～
                            <asp:TextBox ID="Ftdate2" runat="server" Columns="10" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('Ftdate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                        </td>
                    </tr>
                     
                    <tr>
                        <td class="whitecol" colspan="2">
                            <p align="center">
                                <asp:Button ID="Button1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                            </p>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
