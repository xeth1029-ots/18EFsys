<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_04_013_R.aspx.vb" Inherits="WDAIIP.SD_04_013_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>助教授課時數統計報表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        function GETvalue() {
            document.getElementById('Button3').click();
        }
        function visble() {
            document.getElementById('Years').style.visibility = 'hidden'
            document.getElementById('Months').style.visibility = 'hidden'
        }
        function choose_class() {
            var RIDValue = document.getElementById('RIDValue');
            openClass('../02/SD_02_ch.aspx?RID=' + RIDValue.value, 'Class');
        }

        function ReportPrint() {
            var Printtype = document.getElementsByName('Printtype');
            var Years = document.getElementById('Years');
            var Months = document.getElementById('Months');
            var s_date = document.getElementById('s_date');
            var e_date = document.getElementById('e_date');

            var msg = '';
            if (getRadioValue(Printtype) == '0') {
                if (Years.value == '') msg += '請選擇【年度】!!\n';
                if (Months.value == '') msg += '請選擇【月份】!!\n';
            }
            else {
                if (s_date.value == '') msg += '請選擇【全期起日】!!\n';
                if (e_date.value == '') msg += '請選擇【全期迄日】!!\n';
            }

            if (msg != '') {
                alert(msg);
                return false;
            }
            return true;
        }

        function ChangeMode() {
            switch (getRadioValue(document.getElementsByName('Printtype'))) {
                case '0':
                    document.getElementById('Months_TR').style.display = '';
                    document.getElementById('Allday_TR').style.display = 'none';
                    break;
                case '1':
                    document.getElementById('Months_TR').style.display = 'none';
                    document.getElementById('Allday_TR').style.display = '';
                    break;
            }
        }

    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;教師資料管理&gt;&gt;第二教師授課時數統計報表</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <%--<table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>
							<asp:Label ID="TitleLab1" runat="server"></asp:Label>
							<asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;課程管理&gt;&gt;<FONT color="#990000">第二教師授課時數統計報表
										</FONT></asp:Label>
						</td>
					</tr>
				</table>--%>
                    <table class="table_nw" id="Table3" width="100%" cellpadding="1" cellspacing="1">
                        <tr>
                            <td class="bluecol" style="width: 20%">訓練機構
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" onfocus="this.blur()" runat="server" Width="40%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                                <asp:Button ID="Button3" Style="display: none" runat="server" Text="Button3"></asp:Button>
                                <span id="HistoryList2" style="display: none; position: absolute" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">職類/班別
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="25%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="button_b_Mini">
                                <input id="OCIDValue1" type="hidden" name="Hidden2" runat="server">
                                <input id="TMIDValue1" type="hidden" name="Hidden1" runat="server">
                                <span id="HistoryList" style="display: none; left: 270px; position: absolute">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label>
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="30%"></asp:TextBox>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="button_b_Mini">
                                <input id="cjobValue" type="hidden" name="cjobValue" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">列印依據
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="Printtype" runat="server" AutoPostBack="True" RepeatDirection="Horizontal" CssClass="font">
                                    <asp:ListItem Value="0" Selected="True">月</asp:ListItem>
                                    <asp:ListItem Value="1">全期</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="Months_TR" runat="server">
                            <td class="bluecol_need">依月份
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="Years" runat="server">
                                </asp:DropDownList>
                                年
							<asp:DropDownList ID="Months" runat="server">
                            </asp:DropDownList>
                            </td>
                        </tr>
                        <tr id="Allday_TR" runat="server">
                            <td class="bluecol_need">依全期
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="s_date" runat="server" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="PublicCalendar('s_date');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                ～
							<asp:TextBox ID="e_date" runat="server" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="PublicCalendar('e_date');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td class="whitecol">
                                <p align="center">
                                    <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                                </p>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
