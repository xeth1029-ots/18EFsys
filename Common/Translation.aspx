<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Translation.aspx.vb" Inherits="WDAIIP.Translation" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>中翻英</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link rel="stylesheet" type="text/css" href="../../css/style.css" />
    <script type="text/javascript" src="../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        //判斷查詢
        function chkSch() {
            var msg = '';
            var txtName1 = document.getElementById('txtName1');
            var txtName2 = document.getElementById('txtName2');
            var txtEng1 = document.getElementById('txtEng1');
            var txtEng2 = document.getElementById('txtEng2');
            if (isBlank(txtName1)) {
                txtEng1.value = '';
                msg += '請輸入中文姓名第一欄!\n';
            }
            if (isBlank(txtName2)) {
                txtEng2.value = '';
                msg += '請輸入中文姓名第二欄!\n';
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        //回傳英文值
        function sentValue() {
            var hidSN = document.getElementById('hidSN');
            var hidRtnID1 = document.getElementById('hidRtnID1');
            var hidRtnID2 = document.getElementById('hidRtnID2');
            var txtEng1 = document.getElementById('txtEng1');
            var txtEng2 = document.getElementById('txtEng2');
            if (hidSN.value == 'stud') {
                window.opener.document.getElementById(hidRtnID1.value).value = txtEng1.value;
                window.opener.document.getElementById(hidRtnID2.value).value = txtEng2.value;
            } else {
                if (txtEng1.value != '' && txtEng2.value != '') {
                    window.opener.document.getElementById(hidRtnID1.value).value = txtEng1.value + ',' + txtEng2.value;
                } else if (txtEng1.value != '') {
                    window.opener.document.getElementById(hidRtnID1.value).value = txtEng1.value;
                } else if (txtEng2.value != '') {
                    window.opener.document.getElementById(hidRtnID1.value).value = txtEng2.value;
                } else {
                    window.opener.document.getElementById(hidRtnID1.value).value = '';
                }
            }
            window.close();
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <asp:Panel ID="panSch" runat="server">
            <input id="hidSN" type="hidden" runat="server">
            <input id="hidName" type="hidden" runat="server">
            <input id="hidRtnID1" type="hidden" runat="server">
            <input id="hidRtnID2" type="hidden" runat="server">
            <input id="hidName1" type="hidden" runat="server">
            <input id="hidName2" type="hidden" runat="server">
            <br style="line-height: 5px">
            <table class="table_nw" cellspacing="1" cellpadding="1" width="100%" align="center" border="0">
                <tr>
                    <td width="20%" class="bluecol">中文姓名</td>
                    <td class="whitecol" width="80%">
                        <asp:TextBox ID="txtName1" runat="server" CssClass="ipt" MaxLength="3" Width="15%"></asp:TextBox>
                        <asp:TextBox ID="txtName2" runat="server" CssClass="ipt" MaxLength="5" Width="15%"></asp:TextBox>
                        <asp:Button ID="btnSch" runat="server" Text="翻譯" CssClass="asp_button_M"></asp:Button>
                    </td>
                </tr>
                <tr>
                    <td width="20%" class="bluecol">英文姓名</td>
                    <td class="whitecol">
                        <asp:TextBox ID="txtEng1" runat="server" CssClass="ipt" Width="15%"></asp:TextBox>
                        <asp:TextBox ID="txtEng2" runat="server" CssClass="ipt" Width="15%"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td align="center" colspan="2" class="whitecol">
                        <asp:Button ID="btnAdd" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="btnSent" runat="server" Text="確定" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="btnLev" runat="server" Text="離開" CssClass="asp_button_M"></asp:Button>
                    </td>
                </tr>
            </table>
            <table style="font-size: 9pt" cellspacing="1" cellpadding="1" width="100%" align="center" border="0">
                <tr>
                    <td style="color: red">
                        1.若姓名中,有特殊字或難字時,可能會造成無法英譯之問題,建議<br>&nbsp;&nbsp;&nbsp;將特殊字或難字以同音代替,以利正常翻譯.<br>
                        2.或至<a onclick="window.open('http://www.boca.gov.tw/sp.asp?xdURL=E2C/c2102-5.asp&amp;CtNodeID=58&amp;mp=1','','');window.close();" href="#">外交部領事事務局</a>查詢.
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="panAdd" Visible="False" runat="server">
            <table class="table_nw" cellspacing="1" cellpadding="1" width="100%" align="center" border="0">
                <tr>
                    <td class="bluecol" width="20%">英文:</td>
                    <td class="whitecol"><asp:TextBox ID="txtAEng" runat="server"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol">中文:</td>
                    <td class="whitecol"><asp:TextBox ID="txtAWord" runat="server"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="whitecol">
                        <asp:Button ID="btnUpdate" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="btnBack" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                    </td>
                </tr>
            </table>
        </asp:Panel>
    </form>
</body>
</html>