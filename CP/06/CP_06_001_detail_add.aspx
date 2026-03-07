<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_06_001_detail_add.aspx.vb" Inherits="WDAIIP.CP_06_001_detail_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>CP_06_001_detail_add</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script src="../../js/common.js"></script>
    <script>
        //改變答案設定-選項種類
        function ChangeAnsType() {
            if (getRadioValue(document.form1.Path) == '2') {
                if (getRadioValue(document.form1.AnsType) == '01') {
                    document.getElementById('Table_Multiline').style.display = 'none';
                    document.getElementById('Table_AnsChooseN').style.display = 'none';
                    document.getElementById('Table_AnsList').style.display = 'inline';
                }

                if (getRadioValue(document.form1.AnsType) == '02') {
                    document.getElementById('Table_Multiline').style.display = 'none';
                    document.getElementById('Table_AnsChooseN').style.display = 'inline';
                    document.getElementById('Table_AnsList').style.display = 'inline';
                }

                if (getRadioValue(document.form1.AnsType) == '03') {
                    document.getElementById('Table_Multiline').style.display = 'inline';
                    document.getElementById('Table_AnsChooseN').style.display = 'inline';
                    document.getElementById('Table_AnsList').style.display = 'none';
                }
            }
        }


        //改變答案設定-問答題是否多行
        function ChangeMultiline() {
            if (getRadioValue(document.form1.Multiline) != '') {
                if (getRadioValue(document.form1.Multiline) == 'N') {
                    document.getElementById('MultilineTable').style.display = 'none';
                }
                else {
                    document.getElementById('MultilineTable').style.display = 'inline';
                }
            }
        }

        function ChkData() {
            var msg = '';
            var a = '';
            var AnsListNum = new Array();

            if (document.getElementById('Heading').value == '') msg += '請輸入題目文字\n';
            if (document.getElementById('Seq').value == '') msg += '請輸入排序\n';
            else if (!isUnsignedInt(document.form1.Seq.value)) msg += '排序不是正確的數字\n';

            //選項次-子項後要做的檢查			
            if (getRadioValue(document.form1.Path) == '2') {
                if (document.form1.PathOGHID_List.selectedIndex == 0) msg += '請選擇項次下拉選項或新增父項\n';
                if (getRadioValue(document.form1.AnsType) == '') msg += '請選擇答案設定-選項種類\n';

                if (getRadioValue(document.form1.AnsType) == '03') {
                    if (getRadioValue(document.form1.Multiline) == '') msg += '請選擇問答題是否多行\n';
                    if (getRadioValue(document.form1.Multiline) == 'Y') {
                        if (document.getElementById('Rows').value == '') msg += '請輸入行數\n';
                        else if (!isUnsignedInt(document.form1.Rows.value)) msg += '行數必須為整數\n';
                        else if (document.getElementById('Rows').value == '1') msg += '行數數字請大於1\n';
                    }
                }

                if (getRadioValue(document.form1.AnsType) == '02' || getRadioValue(document.form1.AnsType) == '03') {
                    if (!isUnsignedInt(document.form1.AnsChooseN.value)) msg += '複選答案數量限制(問答長度)必須為整數\n';
                    if (document.form1.AnsChooseN.value == '0') msg += '複選答案數量限制(問答長度)不能為0\n';
                }

                if (getRadioValue(document.form1.AnsType) == '01' || getRadioValue(document.form1.AnsType) == '02') {
                    if (document.form1.AnsList.value == '') msg += '請輸入答案列示\n';
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
    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
        <tr>
            <td colspan="2">
                <font color="red">題目設定</font>
            </td>
        </tr>
        <tr>
            <td style="width: 100px; height: 50px" align="left" width="35" bgcolor="#cc6666">
                <font color="#ffffff">&nbsp;&nbsp;&nbsp; 項次<font color="#ffff66"><strong>*</strong></font></font>
            </td>
            <td style="height: 46px" bgcolor="#ffecec">
                <asp:RadioButtonList ID="Path" runat="server" AutoPostBack="True" RepeatDirection="Horizontal" CellSpacing="0" CellPadding="0" CssClass="font" Width="20%">
                    <asp:ListItem Value="1" Selected="True">父項 </asp:ListItem>
                    <asp:ListItem Value="2">子項</asp:ListItem>
                </asp:RadioButtonList>
                <table class="font" id="PathOGHID_Table" style="border-collapse: collapse" bordercolor="darkseagreen" cellspacing="0" cellpadding="0" width="100%" border="0" runat="server">
                    <tr>
                        <td>
                            <asp:DropDownList ID="PathOGHID_List" runat="server">
                            </asp:DropDownList>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td style="width: 100px; height: 24px" align="left" width="35" bgcolor="#cc6666">
                <font color="#ffffff">&nbsp;&nbsp;&nbsp; 題目文字<font color="#ffff66"><strong>*</strong></font></font>
            </td>
            <td style="height: 24px" bgcolor="#ffecec">
                <asp:TextBox ID="Heading" runat="server" Width="416px"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td style="width: 100px" align="left" width="35" bgcolor="#cc6666">
                <font color="#ffffff">&nbsp;&nbsp;&nbsp; 排序<font color="#ffff66"><strong>*</strong></font></font>
            </td>
            <td bgcolor="#ffecec">
                <asp:TextBox ID="Seq" runat="server"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td style="width: 100px" align="left" width="35" bgcolor="#cc6666">
                <font color="#ffffff">&nbsp;&nbsp;&nbsp; 是否啟用</font>
            </td>
            <td bgcolor="#ffecec">
                <asp:CheckBox ID="Useing" runat="server" Checked="True"></asp:CheckBox>
            </td>
        </tr>
    </table>
    <asp:Panel ID="PanelAns" runat="server" Width="100%">
        <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td colspan="2">
                    <font color="red">答案設定</font>
                </td>
            </tr>
            <tr>
                <td style="width: 98px" align="left" width="98" bgcolor="#cc6666">
                    <font color="#ffffff">&nbsp;&nbsp;&nbsp; 選項種類<font color="#ffff66"><strong>*</strong></font></font>
                </td>
                <td bgcolor="#ffecec">
                    <asp:RadioButtonList ID="AnsType" runat="server" CellSpacing="0" CellPadding="0" CssClass="font" Width="60%" RepeatLayout="Flow">
                        <asp:ListItem Value="01">選擇(是非，多選1)</asp:ListItem>
                        <asp:ListItem Value="02">複選(多選n)</asp:ListItem>
                        <asp:ListItem Value="03">問答(答案長度限制)</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
        </table>
        <table class="font" id="Table_Multiline" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td style="width: 98px; height: 24px" align="left" width="98" bgcolor="#cc6666">
                    <font color="#ffffff">&nbsp;&nbsp;&nbsp; 問答題是<br>
                        &nbsp;&nbsp;&nbsp; 否多行<font color="#ffff66"><strong>*</strong></font></font>
                </td>
                <td style="height: 24px" bgcolor="#ffecec">
                    <asp:RadioButtonList ID="Multiline" runat="server" RepeatDirection="Horizontal" CellSpacing="0" CellPadding="0" CssClass="font" Width="20%">
                        <asp:ListItem Value="N">否</asp:ListItem>
                        <asp:ListItem Value="Y">是</asp:ListItem>
                    </asp:RadioButtonList>
                    <table class="font" id="MultilineTable" style="border-collapse: collapse" bordercolor="darkseagreen" cellspacing="0" cellpadding="0" width="100%" border="1" runat="server">
                        <tr>
                            <td>
                                <asp:TextBox ID="Rows" runat="server"></asp:TextBox>行
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <table class="font" id="Table_AnsChooseN" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td style="width: 98px" align="left" width="98" bgcolor="#cc6666">
                    <font color="#ffffff">&nbsp;&nbsp;&nbsp; 複選答案數量限<br>
                        &nbsp;&nbsp;&nbsp; 制(問答長度)<font color="#ffff66"><strong>*</strong></font></font>
                </td>
                <td bgcolor="#ffecec">
                    <asp:TextBox ID="AnsChooseN" runat="server"></asp:TextBox>
                </td>
            </tr>
        </table>
        <table class="font" id="Table_AnsList" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td colspan="2">
                    <font color="red">答案群</font>
                </td>
            </tr>
            <tr>
                <td bgcolor="#cc6666" colspan="2">
                    <font color="#ffffff">答案列示 (答案請用半型逗點分隔)<font color="#ffff66"><strong>*</strong></font></font>
                </td>
            </tr>
            <tr>
                <td colspan="2">
                    <asp:TextBox ID="AnsList" runat="server" Width="360px"></asp:TextBox>
                </td>
            </tr>
        </table>
    </asp:Panel>
    <table class="font" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
        <tr>
            <td align="center">
                <asp:Button ID="Save" runat="server" Text="儲存"></asp:Button>&nbsp;
                <asp:Button ID="return_btn" runat="server" Text="回上一頁"></asp:Button>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
