<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_002_R.aspx.vb" Inherits="WDAIIP.SD_05_002_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>學員出缺勤明細表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        function GETvalue() {
            document.getElementById('Button3').click();
        }

        function checkRBL1VAL() {
            //var RBL1_1 = document.getElementById('RadioButtonList1_0');
            //var RBL1_2 = document.getElementById('RadioButtonList1_1');
            //var RBL2_2 = document.getElementById('RadioButtonList2');
            //var fg1 = (RBL2_2 && RBL1_1 && RBL1_1.checked);
            //$("#RadioButtonList1_0").is(":checked")
            if ($("#RadioButtonList1_0").is(":checked")) {
                $("#RadioButtonList2").show();
            }
            else {
                $("#RadioButtonList2").hide();
            }

        }
        function print() {
            var msg = ''
            if (document.form1.start_date.value == '' && document.form1.end_date.value == '') msg += '請輸入時間區間\n';
            if (!isChecked(document.form1.RadioButtonList1)) msg += '請選擇列印方式\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
            return true;
        }

        function choose_class() {
            //onclick="javascript:openClass('../02/SD_02_ch.aspx?special=5&RID='+document.form1.RIDValue.value);"
            //special=5 提供開訓、結訓日期欄位special=5&DateS=start_date&DateF=end_date
            openClass('../02/SD_02_ch.aspx?special=5&DateS=start_date&DateF=end_date&RID=' + document.form1.RIDValue.value);
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;學員出缺勤明細表</asp:Label>
                </td>
            </tr>
        </table>
        <input id="Hiditem1" type="hidden" name="Hiditem1" runat="server">
        <input id="Hiditem2" type="hidden" name="Hiditem2" runat="server">
        <input id="Hiditem3" type="hidden" name="Hiditem3" runat="server">
        <input id="Hiditem4" type="hidden" name="Hiditem4" runat="server">
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <%--<table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">學員出缺勤明細表</font> </td>
					</tr>
				</table>--%>
                    <table class="table_sch" id="Table3" cellpadding="1" cellspacing="1">
                        <tr>
                            <td class="bluecol" width="20%">訓練機構</td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="Button2" type="button" value="..." name="Button2" runat="server">
                                <input id="RIDValue" type="hidden" name="Hidden3" runat="server">
                                <asp:Button ID="Button3" Style="display: none" runat="server" Text="Button3"></asp:Button>
                                <span id="HistoryList2" style="display: none; position: absolute" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%">
                                    </asp:Table>
                                </span></td>
                        </tr>
                        <tr>
                            <td class="bluecol">職類/班別</td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="...">
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                <span id="HistoryList" style="display: none; left: 270px; position: absolute">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%">
                                    </asp:Table>
                                </span></td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">時間區間</td>
                            <td class="whitecol">
                                <span id="span01" runat="server">
                                    <asp:TextBox ID="start_date" runat="server" Width="15%" onfocus="this.blur()"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">～
							    <asp:TextBox ID="end_date" runat="server" Width="15%" onfocus="this.blur()"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                </span>
                                <asp:Button ID="btnEnterDate" runat="server" Text="帶入班級時間" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">通俗職類</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="25%"></asp:TextBox>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server">
                                <input id="cjobValue" type="hidden" name="cjobValue" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">列印格式 </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="RadioButtonList1" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="1">出缺勤明細表</asp:ListItem>
                                    <asp:ListItem Value="2">請假、缺曠課累計時數統計表</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">列印版型</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="RadioButtonList2" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="1">直版</asp:ListItem>
                                    <asp:ListItem Value="2">橫版</asp:ListItem>
                                </asp:RadioButtonList>
                                <%--<asp:TextBox ID="prtPageSize" runat="server" Width="6%" MaxLength="3">26/42</asp:TextBox>--%></td>
                        </tr>
                    </table>
                    <p align="center" class="whitecol">
                        <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                    </p>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
