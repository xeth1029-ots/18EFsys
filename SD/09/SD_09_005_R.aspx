<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_09_005_R.aspx.vb" Inherits="TIMS.SD_09_005_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_02_001_R</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        function choose_class() {
            window.open('../02/SD_02_ch.aspx', '', 'width=550,height=520,location=0,status=0,menubar=0,scrollbars=1,resizable=0');
        }
        function print() {
            var msg = '';
            if (document.form1.OCIDValue1.value == '') msg += '請選擇班級職類\n';
            if (document.form1.years.selectedIndex == 0) msg += '請選擇年度\n';
            if (document.form1.months.selectedIndex == 0) msg += '請選擇月份\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table id="Table1" cellspacing="1" cellpadding="1" width="600" border="0">
        <tr>
            <td>
                <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            首頁&gt;&gt;學員動態管理&gt;&gt;教務報表管理&gt;&gt; <font color="#990000">(超時)鍾點費印領清冊</font>
                        </td>
                    </tr>
                </table>
                <table class="table_sch" id="Table3">
                    <tr>
                        <td class="bluecol_need" width="100">
                            班別/職類
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="210px"></asp:TextBox>
                            <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="200px"></asp:TextBox>
                            <input onclick="choose_class()" type="button" value="..." class="button_b_Mini" />
                            <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" size="1" />
                            <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" size="1" />
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol_need" width="100">
                            月份
                        </td>
                        <td class="whitecol">
                            <asp:DropDownList ID="years" runat="server">
                            </asp:DropDownList>
                            年
                            <asp:DropDownList ID="months" runat="server">
                            </asp:DropDownList>
                            月
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <br />
    <div style="width: 600" align="center">
        <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_button_S"></asp:Button>
    </div>
    </form>
</body>
</html>
