<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_04_008.aspx.vb" Inherits="WDAIIP.CP_04_008" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>CP_04_008</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <style type="text/css">
        .class_link A { color: #000000; }
            .class_link A:link { color: #0000ff; }
            .class_link A:hover { color: #0000ff; }
        A:visited { color: #0000ff; }
        A:active { color: #0000ff; }
    </style>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript">
        //選擇全部
        function SelectAll(obj, hidobj) {
            var num = getCheckBoxListValue(obj).length;
            var myallcheck = document.getElementById(obj + '_' + 0);

            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
        }

        //檢查日期格式
        function check_date() {
            var msg = '';

            if (!IsDate(form1.SSTDate.value)) {
                document.form1.SSTDate.value = '';
                msg += '開訓起日請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }

            if (!IsDate(form1.ESTDate.value)) {
                document.form1.ESTDate.value = '';
                msg += '開訓迄日請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }

            if (!IsDate(form1.SFTDate.value)) {
                document.form1.SFTDate.value = '';
                msg += '結訓起日請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }

            if (!IsDate(form1.EFTDate.value)) {
                document.form1.EFTDate.value = '';
                msg += '結訓迄日請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }

            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function IsDate(MyDate) {
            if (MyDate != '') {
                if (!checkDate(MyDate))
                    return false;
            }
            return true;
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <table class="font" width="100%">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label><asp:Label ID="TitleLab2" runat="server">
							<FONT face="新細明體">首頁&gt;&gt;訓練查核與績效管理&gt;&gt;訓練資料查詢&gt;&gt;<FONT class="font" color="#800000" size="2">招生中班別查詢</FONT></FONT>
                                </asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch" id="Table1" cellspacing="1" cellpadding="1">
                        <tr>
                            <td style="width: 10%; height: 14px" class="bluecol_need">年度
                            </td>
                            <td style="height: 14px" class="whitecol">
                                <asp:DropDownList ID="yearlist" runat="server">
                                </asp:DropDownList>
                                <asp:RequiredFieldValidator ID="MustYear" runat="server" ErrorMessage="請選擇年度" Display="Dynamic" ControlToValidate="yearlist" CssClass="font"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 10%" class="bluecol">轄區
                            </td>
                            <td class="whitecol">
                                <table id="Table2" style="width: 100%; height: 52px" cellspacing="1" cellpadding="1" width="536" border="0">
                                </table>
                                <asp:CheckBoxList ID="DistrictList" runat="server" CssClass="font" Height="11px" Width="512px" RepeatDirection="Horizontal">
                                </asp:CheckBoxList>
                                <input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 10%" class="bluecol">縣市
                            </td>
                            <td class="whitecol">
                                <table id="Table15" style="width: 100%; height: 52px" cellspacing="1" cellpadding="1" width="536" border="0">
                                </table>
                                <asp:CheckBoxList ID="CityList" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="8">
                                </asp:CheckBoxList>
                                <input id="CityHidden" type="hidden" value="0" name="CityHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 10%" class="bluecol">訓練計畫
                            </td>
                            <td class="whitecol">
                                <table id="Table3" style="width: 100%; height: 52px" cellspacing="1" cellpadding="1" width="536" border="0">
                                </table>
                                <asp:CheckBoxList ID="PlanList" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3">
                                </asp:CheckBoxList>
                                <input id="TPlanHidden" type="hidden" value="0" name="TPlanHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 10%; height: 26px" class="bluecol">開訓日期
                            </td>
                            <td style="height: 26px" class="whitecol">
                                <table id="Table4" style="width: 100%; height: 52px" cellspacing="1" cellpadding="1" width="536" border="0">
                                </table>
                                <asp:TextBox ID="SSTDate" runat="server" Columns="10" MaxLength="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= SSTDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">&nbsp;~
							<asp:TextBox ID="ESTDate" runat="server" Columns="10" MaxLength="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= ESTDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 10%; height: 26px" class="bluecol">結訓日期
                            </td>
                            <td style="height: 26px" class="whitecol">
                                <table id="Table7" style="width: 100%; height: 52px" cellspacing="1" cellpadding="1" width="536" border="0">
                                </table>
                                <asp:TextBox ID="SFTDate" runat="server" Columns="10" MaxLength="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= SFTDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">&nbsp;~
							<asp:TextBox ID="EFTDate" runat="server" Columns="10" MaxLength="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= EFTDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 10%" class="bluecol">開班狀態
                            </td>
                            <td class="whitecol">
                                <table id="Table5" style="width: 100%; height: 52px" cellspacing="1" cellpadding="1" width="536" border="0">
                                </table>
                                <asp:CheckBoxList ID="NotOpenStaus" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="開班" Selected="True">開班</asp:ListItem>
                                    <asp:ListItem Value="不開班">不開班</asp:ListItem>
                                </asp:CheckBoxList>
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 10%" class="bluecol">機構名稱
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="OrgName" runat="server"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 10%" class="bluecol">班級名稱
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="ClassCName" runat="server"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 10%" class="bluecol">職類
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="TMID" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <%--
				            <TR>
					            <TD style="WIDTH: 10%" bgColor="#cc6666"><FONT class="font" face="新細明體" color="#ffffff" size="2">預算來源</FONT></TD>
					            <TD bgColor="#ffecec">
						            <TABLE id="Table8" style="WIDTH: 100%; HEIGHT: 52px" cellSpacing="1" cellPadding="1" width="536"
							            border="0">
							            <TR>
							            </TR>
						            </TABLE>
						            <asp:checkboxlist id="BudgetList" runat="server" CssClass="font" Height="11px" Width="170px" RepeatDirection="Horizontal"></asp:checkboxlist></TD>
				            </TR>
                        --%>
                    </table>
                </td>
            </tr>
        </table>
        <table class="font" id="Table6" cellspacing="0" cellpadding="0" width="740" border="0">
            <tr align="center">
                <td>
                    <font face="新細明體">
                        <asp:Button ID="bt_search" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>
                        <asp:Button ID="bt_reset" runat="server" Text="重新設定" CssClass="asp_button_M"></asp:Button></font>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
