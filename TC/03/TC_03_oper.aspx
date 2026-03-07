<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_03_oper.aspx.vb" Inherits="WDAIIP.TC_03_oper" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>時間換算</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <%--<script type="text/javascript" src="../../js/OpenWin/openwin.js"></script>--%>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function chkDataFormat() {
            var msg = '';
            var RadioButtonList1 = document.form1.RadioButtonList1;
            if (!isChecked(RadioButtonList1)) {
                msg += '請先選擇要換算的總類! \n';
            }
            if (document.form1.start_date.value == '') {
                msg += '請輸入起始日期! \n';
            }
            if (document.form1.oneday.value == '') {
                msg += '請輸入一天時數! \n';
            }
            if (document.form1.oneday.value != '' && !isUnsignedInt(document.form1.oneday.value)) {
                msg += '一天時數必須為正整數! \n';
            }
            if (getRadioValue(RadioButtonList1) == 1 && document.form1.hours.value == '') {
                msg += '請輸入時數! \n';
            }
            if (getRadioValue(RadioButtonList1) == 2 && document.form1.end_date.value == '') {
                msg += '請輸入結訓日期! \n';
            }
            if (msg != '') {
                //opener.alert(msg);
                $('#lblMsg').text(msg);
                return false;
            }
            else {
                $('#lblMsg').text("");
                return true;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="Table1" cellspacing="1" width="100%" border="0">
            <tr>
                <td>
                    <table id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td align="center" colspan="2" class="table_title">此試算僅適用六、日不上課者</td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">換算種類<font color="#ff0000">*</font></td>
                            <td class="whitecol" width="80%">
                                <asp:RadioButtonList ID="RadioButtonList1" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                    <asp:ListItem Value="1">結訓日</asp:ListItem>
                                    <asp:ListItem Value="2">時數</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">一天時數<font color="#ff0000">*</font></td>
                            <td class="whitecol">
                                <asp:TextBox ID="oneday" runat="server" Width="20%" MaxLength="3"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">開訓日<font color="#ff0000">*</font></td>
                            <td class="whitecol">
                                <asp:TextBox ID="start_date" runat="server" Width="15%" MaxLength="10"></asp:TextBox>
                                <span id="span1" runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">結訓日</td>
                            <td class="whitecol">
                                <asp:TextBox ID="end_date" runat="server" Width="15%" MaxLength="10"></asp:TextBox>
                                <span id="span2" runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">時數</td>
                            <td class="whitecol">
                                <asp:TextBox ID="hours" runat="server" Width="20%" MaxLength="4"></asp:TextBox></td>
                        </tr>
                    </table>
                    <div align="center" class="whitecol">
                        <asp:Button ID="Button1" runat="server" Text="換算" CssClass="asp_button_M"></asp:Button>
                        <input type="button" value="關閉視窗" onclick="javascript: window.close();" class="asp_button_M" />
                        <br />
                        <br />
                        <asp:Label ID="lblMsg" runat="server" Text="" ForeColor="Red"></asp:Label>
                    </div>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
