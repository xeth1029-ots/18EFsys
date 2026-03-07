<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="EXAM_02_002.aspx.vb" Inherits="WDAIIP.EXAM_02_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>題組題庫維護</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script src="../../js/common.js"></script>
    <script src="../../js/TIMS.js"></script>
    <script language="javascript">
        function check_select() {
            var msg = '';
            if (document.getElementById('ddl_qtype').value == '0') {
                msg = '請選擇題目類型\n';
            }
            if ((!document.getElementById('ddl_select2').disabled) && (document.getElementById('ddl_select2').value == '0')) {
                msg += '請選擇選項數目\n';
            }
            if (msg != '') {
                window.alert(msg);
                return false;
            }
        }

        function check_save() {
            var i = 0;
            var sum = 0;
            var msg = '';
            var tmpfiletypepos = 0;
            var tmpfiletype = '';

            if (document.getElementById('ddl_etid').value == '0' || document.getElementById('ddl_etid').value == '') {
                msg += '請選擇題組類別\n';
            }
            if (document.getElementById('txt_question').value == '') {
                msg += '請填入題目名稱\n';
            }
            if (document.getElementById('ddl_qtype').value == '0') {
                msg += '請選擇題目類型\n';
            }
            if ((document.getElementById('ddl_qtype').value == '4') && (document.getElementById('txt_ans4').value == '')) {
                msg += '請填寫【解答】內容\n';
            }
            if (document.getElementById('ddl_qtype').value == '2') {
                if (document.getElementById('ddl_select2').value == '2') {
                    for (i = 1; i < 3; i++) {
                        if (document.getElementById('txt_ans2_' + i).value == '') {
                            msg += '請填寫【項目' + i + '】內容\n';
                        }
                        if (document.getElementById('chk_ans2_' + i).checked) {
                            sum = sum + 1;
                        }
                    }
                    if (sum == 0) {
                        msg += '請選擇正確答案\n';
                    }
                    else if (sum > 1) {
                        msg += '正確答案只有一個，請重新選擇\n';
                    }
                }
                if (document.getElementById('ddl_select2').value == '3') {
                    for (i = 1; i < 4; i++) {
                        if (document.getElementById('txt_ans2_' + i).value == '') {
                            msg += '請填寫【項目' + i + '】內容\n';
                        }
                        if (document.getElementById('chk_ans2_' + i).checked) {
                            sum = sum + 1;
                        }
                    }
                    if (sum == 0) {
                        msg += '請選擇正確答案\n';
                    }
                    else if (sum > 1) {
                        msg += '正確答案只有一個，請重新選擇\n';
                    }
                }
                if (document.getElementById('ddl_select2').value == '4') {
                    for (i = 1; i < 5; i++) {
                        if (document.getElementById('txt_ans2_' + i).value == '') {
                            msg += '請填寫【項目' + i + '】內容\n';
                        }
                        if (document.getElementById('chk_ans2_' + i).checked) {
                            sum = sum + 1;
                        }
                    }
                    if (sum == 0) {
                        msg += '請選擇正確答案\n';
                    }
                    else if (sum > 1) {
                        msg += '正確答案只有一個，請重新選擇\n';
                    }
                }
                if (document.getElementById('ddl_select2').value == '5') {
                    for (i = 1; i < 6; i++) {
                        if (document.getElementById('txt_ans2_' + i).value == '') {
                            msg += '請填寫【項目' + i + '】內容\n';
                        }
                        if (document.getElementById('chk_ans2_' + i).checked) {
                            sum = sum + 1;
                        }
                    }
                    if (sum == 0) {
                        msg += '請選擇正確答案\n';
                    }
                    else if (sum > 1) {
                        msg += '正確答案只有一個，請重新選擇\n';
                    }
                }
            }
            if (document.getElementById('ddl_qtype').value == '3') {
                if (document.getElementById('ddl_select2').value == '2') {
                    for (i = 1; i < 3; i++) {
                        if (document.getElementById('txt_ans2_' + i).value == '') {
                            msg += '請填寫【項目' + i + '】內容\n';
                        }
                        if (document.getElementById('chk_ans2_' + i).checked) {
                            sum = sum + 1;
                        }
                    }
                    if (sum == 0) {
                        msg += '請選擇正確答案\n';
                    }
                }
                if (document.getElementById('ddl_select2').value == '3') {
                    for (i = 1; i < 4; i++) {
                        if (document.getElementById('txt_ans2_' + i).value == '') {
                            msg += '請填寫【項目' + i + '】內容\n';
                        }
                        if (document.getElementById('chk_ans2_' + i).checked) {
                            sum = sum + 1;
                        }
                    }
                    if (sum == 0) {
                        msg += '請選擇正確答案\n';
                    }
                }
                if (document.getElementById('ddl_select2').value == '4') {
                    for (i = 1; i < 5; i++) {
                        if (document.getElementById('txt_ans2_' + i).value == '') {
                            msg += '請填寫【項目' + i + '】內容\n';
                        }
                        if (document.getElementById('chk_ans2_' + i).checked) {
                            sum = sum + 1;
                        }
                    }
                    if (sum == 0) {
                        msg += '請選擇正確答案\n';
                    }
                }
                if (document.getElementById('ddl_select2').value == '5') {
                    for (i = 1; i < 6; i++) {
                        if (document.getElementById('txt_ans2_' + i).value == '') {
                            msg += '請填寫【項目' + i + '】內容\n';
                        }
                        if (document.getElementById('chk_ans2_' + i).checked) {
                            sum = sum + 1;
                        }
                    }
                    if (sum == 0) {
                        msg += '請選擇正確答案\n';
                    }
                }
            }
            if (msg != '') {
                window.alert(msg);
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
        <tr>
            <td>
                <%--<table class="font" id="tab_title" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                            <asp:Label ID="TitleLab2" runat="server">
									首頁&gt;&gt;招生甄試成績管理&gt;&gt;甄試題組題庫設定&gt;&gt;<font color="#990000">題組題庫維護</font>
                            </asp:Label>
                        </td>
                    </tr>
                </table>--%>
                <table class="table_sch" id="table_Sch" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                    <tr>
                        <td class="bluecol" style="width:20%">
                            題組類別
                        </td>
                        <td class="whitecol">
                            <asp:DropDownList ID="ddl_etid" runat="server" AutoPostBack="True">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            題組子類別
                        </td>
                        <td class="whitecol">
                            <asp:DropDownList ID="ddl_cETID" runat="server">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            題目名稱
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="txt_question" runat="server" Width="40%"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            題目類型
                        </td>
                        <td class="whitecol">
                            <asp:DropDownList ID="ddl_qtype" runat="server" AutoPostBack="true">
                                <asp:ListItem Value="0" Selected="True">--題目類型--</asp:ListItem>
                                <asp:ListItem Value="1">是非題</asp:ListItem>
                                <asp:ListItem Value="2">選擇題</asp:ListItem>
                                <asp:ListItem Value="3">複選題</asp:ListItem>
                                <asp:ListItem Value="4">問答題</asp:ListItem>
                            </asp:DropDownList>
                            &nbsp;&nbsp;&nbsp;
                            <asp:DropDownList ID="ddl_select2" runat="server" Enabled="False">
                                <asp:ListItem Value="0">--選項數目--</asp:ListItem>
                                <asp:ListItem Value="2">2</asp:ListItem>
                                <asp:ListItem Value="3">3</asp:ListItem>
                                <asp:ListItem Value="4">4</asp:ListItem>
                                <asp:ListItem Value="5">5</asp:ListItem>
                            </asp:DropDownList>
                            &nbsp;
                            <asp:Button ID="btn_check" runat="server" Text="確定" CssClass="asp_button_M"></asp:Button>&nbsp;
                            <asp:Button ID="btn_rst" runat="server" Text="重選" Visible="False" CssClass="asp_button_M"></asp:Button>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                           不啟用
                        </td>
                        <td  class="whitecol">
                            <asp:CheckBox Style="z-index: 0" ID="chkStopUse" runat="server" Text="不啟用"></asp:CheckBox>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <asp:Panel ID="tab_select1" Visible="False" runat="server" Height="10px">
                                <table id="Table6" class="font" border="0" cellspacing="1" cellpadding="1" width="100%"
                                    runat="server">
                                    <tr>
                                        <td width="15%" align="center" class="bluecol">
                                            <asp:Label ID="Label5" runat="server">解答</asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td width="15%" align="center" class="whitecol">
                                            <asp:RadioButtonList ID="rdo_ans1" runat="server" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="Y" Selected="True">○</asp:ListItem>
                                                <asp:ListItem Value="N">╳</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                </table>
                            </asp:Panel>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <asp:Panel ID="tab_select2" Visible="False" runat="server">
                                <table id="Table2" class="font" border="0" cellspacing="1" cellpadding="1" width="100%"
                                    runat="server">
                                    <tr>
                                        <td width="7%" class="bluecol">
                                            <asp:Label ID="Label2" runat="server">選號</asp:Label>
                                        </td>
                                        <td width="86%" class="bluecol">
                                            <asp:Label ID="Label7" runat="server">選項</asp:Label>
                                        </td>
                                        <td width="7%" class="bluecol">
                                            <asp:Label ID="Label8" runat="server">解答</asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="height: 10px" width="7%" align="center" class="whitecol">
                                            1
                                        </td>
                                        <td style="height: 10px" width="86%" class="whitecol">
                                            <asp:TextBox ID="txt_ans2_1" runat="server" Width="100%"></asp:TextBox>
                                        </td>
                                        <td style="height: 10px" width="7%" align="center" class="whitecol">
                                            <asp:CheckBox ID="chk_ans2_1" runat="server"></asp:CheckBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="height: 10px" width="7%" align="center" class="whitecol">
                                            2
                                        </td>
                                        <td style="height: 10px" width="86%" class="whitecol">
                                            <asp:TextBox ID="txt_ans2_2" runat="server" Width="100%"></asp:TextBox>
                                        </td>
                                        <td width="7%" align="center" class="whitecol">
                                            <asp:CheckBox ID="chk_ans2_2" runat="server"></asp:CheckBox>
                                        </td>
                                    </tr>
                                </table>
                            </asp:Panel>
                            <asp:Panel ID="tab_select3" Visible="False" runat="server">
                                <table id="Table3" class="font" border="0" cellspacing="1" cellpadding="1" width="100%"
                                    runat="server">
                                    <tr class="whitecol">
                                        <td width="7%" align="center">
                                            3
                                        </td>
                                        <td width="86%">
                                            <asp:TextBox ID="txt_ans2_3" runat="server" Width="100%"></asp:TextBox>
                                        </td>
                                        <td width="7%" align="center">
                                            <asp:CheckBox ID="chk_ans2_3" runat="server"></asp:CheckBox>
                                        </td>
                                    </tr>
                                </table>
                            </asp:Panel>
                            <asp:Panel ID="tab_select4" Visible="False" runat="server">
                                <table id="Table1" class="font" border="0" cellspacing="1" cellpadding="1" width="100%"
                                    runat="server">
                                    <tr class="whitecol">
                                        <td width="7%" align="center">
                                            4
                                        </td>
                                        <td width="86%">
                                            <asp:TextBox ID="txt_ans2_4" runat="server" Width="100%"></asp:TextBox>
                                        </td>
                                        <td width="7%" align="center">
                                            <asp:CheckBox ID="chk_ans2_4" runat="server"></asp:CheckBox>
                                        </td>
                                    </tr>
                                </table>
                            </asp:Panel>
                            <asp:Panel ID="tab_select5" Visible="False" runat="server">
                                <table id="Table4" class="font" border="0" cellspacing="1" cellpadding="1" width="100%"
                                    runat="server">
                                    <tr class="whitecol">
                                        <td width="7%" align="center">
                                            5
                                        </td>
                                        <td width="86%">
                                            <asp:TextBox ID="txt_ans2_5" runat="server" Width="100%"></asp:TextBox>
                                        </td>
                                        <td width="7%" align="center">
                                            <asp:CheckBox ID="chk_ans2_5" runat="server"></asp:CheckBox>
                                        </td>
                                    </tr>
                                </table>
                            </asp:Panel>
                            <asp:Panel ID="tab_select6" Visible="False" runat="server" Height="10px">
                                <table id="Table5" class="font" border="0" cellspacing="1" cellpadding="1" width="100%"
                                    runat="server">
                                    <tr>
                                        <td width="100%" class="bluecol">
                                            <asp:Label ID="Label1" runat="server">解答</asp:Label>
                                        </td>
                                    </tr>
                                    <tr class="whitecol">
                                        <td align="center">
                                            <asp:TextBox ID="txt_ans4" runat="server" Width="100%"></asp:TextBox>
                                        </td>
                                    </tr>
                                </table>
                            </asp:Panel>
                        </td>
                    </tr>
                    <tr>
                        <td align="center" colspan="2" class="whitecol">
                            <asp:Button ID="btn_save" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>&nbsp;
                            <asp:Button ID="btn_exit" runat="server" Text="離開" CssClass="asp_button_M"></asp:Button>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
