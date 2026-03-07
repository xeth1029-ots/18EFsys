<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_09_004_R.aspx.vb" Inherits="WDAIIP.SD_09_004_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>列印點名表</title>
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
        function search() {
            var msg = ''
            if (document.form1.OCIDValue1.value == '') msg += '必須選擇班別職類\n';
            if (document.form1.start_date.value == '') msg += '請選擇點名日期起\n';
            if (document.form1.start_date.value != '' && !checkDate(document.form1.start_date.value)) msg += '點名日期格式不正確\n';
            if (!isChecked(document.form1.RadioButtonList1)) msg += '請選擇列印方式\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
            return true;
        }

        function choose_class() {
            openClass('../02/SD_02_ch.aspx');
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務報表管理&gt;&gt;列印點名表</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tbody>
                <tr>
                    <td>

                        <table id="Table3" class="table_sch" cellspacing="1" cellpadding="1">
                            <tbody>
                                <tr>
                                    <td class="bluecol_need" width="20%">班別/職類
                                    </td>
                                    <td class="whitecol">
                                        <asp:Button ID="Button2" Style="display: none" runat="server" Text="Button2"></asp:Button>
                                        <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="25%"></asp:TextBox>
                                        <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                        <input type="button" value="..." onclick="choose_class()" class="button_b_Mini" />
                                        <input id="OCIDValue1" type="hidden" name="Hidden2" runat="server" />
                                        <input id="TMIDValue1" type="hidden" name="Hidden1" runat="server" />
                                        <span id="HistoryList" style="display: none; left: 270px; position: absolute">
                                            <asp:Table ID="HistoryTable" runat="server" Width="100%">
                                            </asp:Table>
                                        </span>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">點名日期
                                    </td>
                                    <td class="whitecol">
                                        <span id="span01" runat="server">
                                            <asp:TextBox ID="start_date" runat="server" Width="15%"></asp:TextBox>
                                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />～
                                        <asp:TextBox ID="end_date" runat="server" Width="15%"></asp:TextBox>
                                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                                        </span>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">列印方式
                                    </td>
                                    <td class="whitecol">
                                        <asp:RadioButtonList ID="RadioButtonList1" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                            <asp:ListItem Value="1" Selected="True">週</asp:ListItem>
                                            <asp:ListItem Value="2">天</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </td>
                                </tr>
                            </tbody>
                        </table>
                    </td>
                </tr>
            </tbody>
        </table>
        <br />
        <div style="width: 100%" align="center" class="whitecol">
            <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
        </div>
    </form>
</body>
</html>
