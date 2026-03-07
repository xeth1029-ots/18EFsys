<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_03_004_R.aspx.vb" Inherits="TIMS.SD_03_004_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_03_004_R</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">

        function change() {
            var j = 4;
            if (document.form1.RadioButton3.checked) {
                for (var i = 0; i < j; i++) {
                    document.form1.elements['Sort1:' + i].disabled = false;
                }
            }
            else {
                for (var i = 0; i < j; i++) {
                    document.form1.elements['Sort1:' + i].disabled = true;
                }
            }
        }

        function print() {
            var msg = '';
            if (!document.form1.RadioButton1.checked && !document.form1.RadioButton2.checked && !document.form1.RadioButton3.checked) msg += '請選擇列印種類';
            if (document.form1.RadioButton3.checked && getCheckBoxListValue('Sort1').toString(10) == 0) msg += '請勾選列印格式\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
</head>
<body onload="change()" ms_positioning="FlowLayout">
    <form id="form1" method="post" runat="server">
    <font face="新細明體">
        <table id="Table1" cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label><asp:Label ID="TitleLab2" runat="server">
										首頁&gt;&gt;學員動態管理&gt;&gt;招生報名&gt;&gt;<font color="#990000">報名人數分析統計表</font>
                                </asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="table_nw" id="Table3" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td width="100" class="bluecol">
                                職類
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" Width="210px" onfocus="this.blur()"></asp:TextBox>
                                <input id="TMIDValue1" style="width: 40px; height: 22px" type="hidden" size="1" name="TMIDValue1" runat="server" />
                            </td>
                            <td width="100" class="bluecol">
                                班別
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="OCID1" runat="server" Width="210px" onfocus="this.blur()"></asp:TextBox>
                                <input id="OCIDValue1" style="width: 24px; height: 22px" type="hidden" size="1" name="OCIDValue1" runat="server" />
                                <input onclick="javascript:wopen('../02/SD_02_ch.aspx','課程',520,520,1)" type="button" value="..." name="Submit" class="button_b_Mini" />
                                <span id="HistoryList" style="display: none; position: absolute">
                                    <asp:Table ID="HistoryTable" runat="server" Width="230px">
                                    </asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td width="100" class="bluecol_need">
                                列印種類
                            </td>
                            <td class="whitecol" colspan="3">
                                <table class="font" id="Table4" cellspacing="1" cellpadding="1" width="100%">
                                    <tr>
                                        <td width="100">
                                            <asp:RadioButton ID="RadioButton1" runat="server" GroupName="R1" Text="依照身分別"></asp:RadioButton>
                                        </td>
                                        <td>
                                            &nbsp;
                                        </td>
                                    </tr>
                                    <tr>
                                        <td width="100">
                                            <asp:RadioButton ID="RadioButton2" runat="server" GroupName="R1" Text="依照縣市"></asp:RadioButton>
                                        </td>
                                        <td>
                                            &nbsp;
                                        </td>
                                    </tr>
                                    <tr>
                                        <td width="100">
                                            <asp:RadioButton ID="RadioButton3" runat="server" GroupName="R1" Text="其他"></asp:RadioButton>
                                        </td>
                                        <td>
                                            <asp:CheckBoxList ID="Sort1" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                                                <asp:ListItem Value="性別">性別</asp:ListItem>
                                                <asp:ListItem Value="報名管道">報名管道</asp:ListItem>
                                                <asp:ListItem Value="年齡">年齡</asp:ListItem>
                                                <asp:ListItem Value="學歷">學歷</asp:ListItem>
                                            </asp:CheckBoxList>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="4" class="whitecol">
                                <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_button_S"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </font>
    </form>
</body>
</html>
