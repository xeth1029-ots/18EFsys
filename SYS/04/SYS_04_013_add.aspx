<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_04_013_add.aspx.vb" Inherits="WDAIIP.SYS_04_013_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>個人行事曆(維護)</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
		function checkSave1() {
			var msg = '';
			if (document.form1.subject.value == '') {
				msg += "必須填寫 主旨\n";
			}

			if (document.form1.OSDate.value != '') {
				if (!checkDate(document.form1.OSDate.value)) msg += '[日期區間的起日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
			}
			else { msg += "必須填寫 日期區間的起日\n"; }

			if (document.form1.OFDate.value != '') {
				if (!checkDate(document.form1.OFDate.value)) msg += '[日期區間的迄日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
			}
			else { msg += "必須填寫 日期區間的迄日\n"; }

			if (msg == '') {
				if (document.form1.OSDate.value != ''
					&& document.form1.OFDate.value != ''
					&& document.form1.OFDate.value < document.form1.OSDate.value)
				{ msg += '[日期區間的迄日]必需大於[日期區間的起日]\n'; }
			}

			if (document.form1.txtcontext.value == '') {
				msg += "必須填寫 內容\n";
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
        <table class="font" width="740">
            <tr>
                <td class="font">首頁&gt;&gt;系統管理&gt;&gt;系統參數管理&gt;&gt;<font color="#990000">個人行事曆(維護)</font>
                </td>
            </tr>
        </table>
        <table class="table_nw" width="740" cellpadding="1" cellspacing="1">
            <tr>
                <td class="bluecol_need" align="left">主旨
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="subject" runat="server" Width="350px" MaxLength="100"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need" align="left" width="100">日期區間
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="OSDate" runat="server" onfocus="this.blur()" Width="80px"></asp:TextBox>
                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= OSDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top">～
				<asp:TextBox ID="OFDate" runat="server" onfocus="this.blur()" Width="80px"></asp:TextBox>
                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= OFDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top">
                </td>
            </tr>
            <tr>
                <td class="bluecol_need" align="left">內容
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="txtcontext" runat="server" Width="450px" Style="z-index: 0" MaxLength="1000"></asp:TextBox>
                </td>
            </tr>
        </table>
        <table width="740">
            <tr>
                <td class="whitecol" colspan="2" align="center">
                    <asp:Button ID="btnSave1" runat="server" Text="儲存" CssClass="asp_button_S"></asp:Button>
                    <font face="新細明體">&nbsp;</font>
                    <input id="back" type="button" value="回上一頁" name="back" runat="server" class="asp_button_S">
                </td>
            </tr>
        </table>
        <input type="hidden" id="hidcalID" runat="server" size="1">
    </form>
</body>
</html>
