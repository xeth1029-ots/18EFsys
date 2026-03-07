<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_007_add.aspx.vb" Inherits="WDAIIP.SD_05_007_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_05_007_add</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script src="../../js/common.js"></script>
    <script language="javascript">
        function GETvalue() {
            document.getElementById('Button5').click();
        }

        function choose_class() {
            document.form1.Hidden1.value = 1;
            var RID = document.form1.RIDValue.value;
            openClass('../02/SD_02_ch.aspx?special=2&&RID=' + RID);
        }
        function chkdata() {
            var msg = '';
            if (document.form1.SanDate.value == '') msg += '請輸入申請日\n';

            var mytable = document.getElementById('Table4');
            for (var i = 1; i < mytable.rows.length; i++) {
                for (var j = 2; j < mytable.rows(i).cells.length; j++) {
                    var mytext = mytable.rows(i).cells(j).children(0);
                    if (mytext.value != '' && !isUnsignedInt(mytext.value)) {
                        msg += '獎懲必須輸入數字(第' + i + '行)\n';
                        break;
                    }
                }
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
    <table id="Table1" cellspacing="1" cellpadding="1" width="740" border="0">
        <tr>
            <td>
                <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">學員獎懲作業</font>
                        </td>
                    </tr>
                </table>
                <table class="table_sch" id="Table3" cellpadding="1" cellspacing="1">
                    <tr>
                        <td width="100" class="bluecol">
                            訓練機構
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="410px"></asp:TextBox>
                            <input id="Button3" type="button" value="..." name="Button3" runat="server">
                            <input id="RIDValue" type="hidden" name="Hidden3" runat="server" size="1">
                            <input id="Hidden1" type="hidden" name="Hidden1" runat="server" size="1">
                            <asp:Button ID="Button5" Style="display: none" runat="server" Text="Button5"></asp:Button>
                            <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                <asp:Table ID="HistoryRID" runat="server" Width="310px">
                                </asp:Table>
                            </span>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol_need">
                            職類/班別
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="210px"></asp:TextBox>
                            <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="200px"></asp:TextBox>
                            <input id="Button2" type="button" value="..." name="Button2" runat="server">
                            <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" size="1">
                            <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" size="1">
                            <span id="HistoryList" style="position: absolute; display: none; left: 270px">
                                <asp:Table ID="HistoryTable" runat="server" Width="310">
                                </asp:Table>
                            </span>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol_need">
                            申請日
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="SanDate" runat="server" Width="75px"></asp:TextBox>
                            <img id="IMG1" style="cursor: pointer" onclick="javascript:show_calendar('SanDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30">
                        </td>
                    </tr>
                </table>
                <asp:Table ID="Table4" runat="server" Width="100%" CssClass="font" CellSpacing="0" CellPadding="1">
                </asp:Table>
                <p align="center">
                    <asp:Button ID="Button1" runat="server" Text="儲存" CssClass="asp_button_S"></asp:Button>&nbsp;<input id="Button4" type="button" value="回上一頁" name="Button4" runat="server" class="asp_button_S"></p>
                <p align="center">
                    <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label></p>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
