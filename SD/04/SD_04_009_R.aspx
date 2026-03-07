<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_04_009_R.aspx.vb" Inherits="WDAIIP.SD_04_009_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>教師時數統計總表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        function ReportPrint() {
            var msg = ''
            //if (document.form1.Years.value == '') msg += '請選擇年度!!\n';
            //if (document.form1.Months.value == '') msg += '請選擇月份!!\n';
            if (parseInt(getCheckBoxListValue('InTeach')) == 0 && parseInt(getCheckBoxListValue('OutTeach')) == 0) { msg += '請選擇老師\n'; }
            if (document.form1.Mode.value == '') msg += '請選擇列印格式!!\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
            return true;
        }

        function GetAllTeach(Mode, flag) {
            var obj = (Mode == 1) ? 'InTeach' : 'OutTeach';
            var objcount = getCheckBoxListValue(obj).length;
            for (var i = 0; i < objcount; i++) {
                document.getElementById(obj + '_' + i).checked = flag;
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;教師資料管理&gt;&gt;教師時數統計總表</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <%--<table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            首頁&gt;&gt;學員動態管理&gt;&gt;課程管理&gt;&gt;教師時數統計總表
                        </td>
                    </tr>
                </table>--%>
                    <table class="table_nw" id="Table3" width="100%" cellpadding="1" cellspacing="1">
                        <tr>
                            <td class="bluecol_need" style="width: 20%">老師姓名
                            </td>
                            <td class="whitecol">
                                <table class="font" id="Table5" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td>內聘
                                        <asp:CheckBox ID="InSelectAll" runat="server" Text="全選"></asp:CheckBox>
                                            <asp:Button Style="z-index: 0" ID="Button3" runat="server" Text="勾選可列印最大筆數" CssClass="asp_button_M"></asp:Button>
                                            <asp:Label Style="z-index: 0" ID="Label3" runat="server" CssClass="font" ForeColor="Red">自動由勾選處再勾取</asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:CheckBoxList ID="InTeach" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="6">
                                            </asp:CheckBoxList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>外聘
                                        <asp:CheckBox ID="OutSelectAll" runat="server" Text="全選"></asp:CheckBox>
                                            <asp:Button Style="z-index: 0" ID="Button1" runat="server" Text="勾選可列印最大筆數" CssClass="asp_button_M"></asp:Button>
                                            <asp:Label Style="z-index: 0" ID="Label2" runat="server" CssClass="font" ForeColor="Red">自動由勾選處再勾取</asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:CheckBoxList ID="OutTeach" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="6">
                                            </asp:CheckBoxList>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">列印依據
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="Printtype" runat="server" RepeatDirection="Horizontal" AutoPostBack="True" CssClass="font">
                                    <asp:ListItem Value="0" Selected="True">年月</asp:ListItem>
                                    <asp:ListItem Value="1">全期</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="Months_TR" runat="server">
                            <td class="bluecol_need">統計月份
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="Years" runat="server">
                                </asp:DropDownList>
                                年
                            <asp:DropDownList ID="Months" runat="server">
                            </asp:DropDownList>
                                月
                            </td>
                        </tr>
                        <tr id="Allday_TR" runat="server">
                            <td class="bluecol">排課期間
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="s_date1" runat="server" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="PublicCalendar('s_date1');" alt="" align="top" src="../../images/show-calendar.gif">
                                ～
                            <asp:TextBox ID="s_date2" runat="server" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="PublicCalendar('s_date2');" alt="" align="top" src="../../images/show-calendar.gif">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">列印格式
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="Mode" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="1" Selected="True">依班別統計</asp:ListItem>
                                    <asp:ListItem Value="2">依課程統計</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">教師名字
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="tTeachName" runat="server" MaxLength="10" Width="20%"></asp:TextBox>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="Button2" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <asp:Label Style="z-index: 0" ID="Label1" runat="server" CssClass="font" ForeColor="Red">老師數量選擇，請不要超過30筆，避免資料遺失(有資安問題)</asp:Label>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
