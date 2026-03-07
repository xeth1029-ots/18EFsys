<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CM_03_010.aspx.vb" Inherits="WDAIIP.CM_03_010" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>縣市政府各類身分別人數統計</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        //列印
        function chkPrt() {
            var msg = "";
            var msg1 = "";
            var msg2 = "";
            var ddlYear = document.getElementById('ddlYear');
            var txtSTSDate = document.getElementById('txtSTSDate');
            var txtSTEDate = document.getElementById('txtSTEDate');
            var txtETSDate = document.getElementById('txtETSDate');
            var txtETEDate = document.getElementById('txtETEDate');
            var cblCity = document.getElementById('cblCity');

            if (ddlYear.value == '' && txtSTSDate.value == '' && txtSTEDate.value == '' && txtETSDate.value == '' & txtETEDate.value == '') {
                msg1 += '年度、開訓區間及結訓區間必選其一為查詢條件!\n';

            } else {
                //判斷開訓期間
                if (txtSTSDate.value != '') {
                    if (!checkDate(txtSTSDate.value)) {
                        msg1 += '開訓區間(起)日期格式錯誤!\n';
                    }
                }

                if (txtSTEDate.value != '') {
                    if (!checkDate(txtSTEDate.value)) {
                        msg1 += '開訓區間(迄)日期格式錯誤!\n';
                    }
                }

                if (txtSTSDate.value != '' && txtSTEDate.value != '' && msg1 == "") {
                    if (getDiffDay(getAdDate(txtSTSDate.value), getAdDate(txtSTEDate.value)) < 0) {
                        msg1 += '開訓區間日期(起)日不得大於(迄)日!\n';
                    }
                }

                //判斷結訓期間
                if (txtETSDate.value != '') {
                    if (!checkDate(txtETSDate.value)) {
                        msg2 += '結訓區間(起)日期格式錯誤!\n';
                    }
                }

                if (txtETEDate.value != '') {
                    if (!checkDate(txtETEDate.value)) {
                        msg2 += '結訓區間(迄)日期格式錯誤!\n';
                    }
                }

                if (txtETSDate.value != '' && txtETEDate.value != '' && msg2 == "") {
                    if (getDiffDay(getAdDate(txtETSDate.value), getAdDate(txtETEDate.value)) < 0) {
                        msg2 += '結訓區間日期(起)日不得大於(迄)日!\n';
                    }
                }
            }

            msg = msg1 + msg2;

            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>首頁&gt;&gt;訓練與需求管理&gt;&gt;統計分析&gt;&gt;縣市政府各類身分別人數統計
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" width="100">年度
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="ddlYear" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="80">開訓區間
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="txtSTSDate" runat="server" Columns="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('txtSTSDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                ~
                            <asp:TextBox ID="txtSTEDate" runat="server" Columns="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('txtSTEDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">結訓區間
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="txtETSDate" runat="server" Columns="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('txtETSDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                ~
                            <asp:TextBox ID="txtETEDate" runat="server" Columns="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('txtETEDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">縣市別
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="rblCity" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="8" runat="server">
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="2" class="whitecol">
                                <asp:Button ID="btnPrt" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
