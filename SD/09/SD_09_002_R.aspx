<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_09_002_R.aspx.vb" Inherits="WDAIIP.SD_09_002_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <%--<title>列印屆退官兵訓練人員名冊</title>--%>
    <title>列印志願役官兵訓練人員名冊</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function search(args) {
            var msg = ''
            //if (document.form1.OCIDValue1.value=='') msg+='必須選擇班別職類\n';
            if (!isChecked(document.form1.RadioButtonList1)) msg += '請選擇學員狀態';

            if (msg != '') {
                alert(msg);
                return false;
            }
            else {
                if (getRadioValue(document.form1.RadioButtonList1) == 0) {
                    window.open('../../SQControl.aspx?' + args + '&OCID=' + document.getElementById('OCIDValue1').value + '&CJOB_UNKEY=' + document.getElementById('cjobValue').value, 'print', 'toolbar=0,location=0,status=0,menubar=0,resizable=1');
                }
                else {
                    window.open('../../SQControl.aspx?' + args + '&OCID=' + document.getElementById('OCIDValue1').value + '&CJOB_UNKEY=' + document.getElementById('cjobValue').value + '&StudStatus=' + getRadioValue(document.form1.RadioButtonList1), 'print', 'toolbar=0,location=0,status=0,menubar=0,resizable=1');
                }
            }
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務報表管理&gt;&gt;列印屆退軍官訓練人員名冊</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>

                    <table class="table_sch" id="Table3" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" style="width: 20%">職類/班別
                            </td>
                            <td class="whitecol">
                                <asp:Button ID="Button2" Style="display: none" runat="server" Text="Button2"></asp:Button>
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="25%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="button_b_Mini" />
                                <input id="OCIDValue1" type="hidden" name="Hidden2" runat="server" />
                                <input id="TMIDValue1" type="hidden" name="Hidden1" runat="server" />
                                <span id="HistoryList" style="display: none; left: 270px; position: absolute">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label>
                            </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="30%"></asp:TextBox>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="button_b_Mini" />
                                <input id="cjobValue" type="hidden" name="cjobValue" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">學員狀況
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="RadioButtonList1" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="0">不區分</asp:ListItem>
                                    <asp:ListItem Value="1">在訓</asp:ListItem>
                                    <asp:ListItem Value="2">離訓</asp:ListItem>
                                    <asp:ListItem Value="3">退訓</asp:ListItem>
                                    <asp:ListItem Value="4">續訓</asp:ListItem>
                                    <asp:ListItem Value="5">結訓</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>

        <div style="width: 100%" align="center" class="whitecol">
            <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
        </div>
    </form>
</body>
</html>
