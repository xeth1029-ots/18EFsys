<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_04_003.aspx.vb" Inherits="WDAIIP.CP_04_003" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>開班資料</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
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
            var SSTDate = document.getElementById("SSTDate");
            var ESTDate = document.getElementById("ESTDate");
            var SFTDate = document.getElementById("SFTDate");
            var EFTDate = document.getElementById("EFTDate");
            if (!IsDate(SSTDate.value)) {
                SSTDate.focus();
                msg += '開訓起日請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }
            if (!IsDate(ESTDate.value)) {
                ESTDate.focus();
                msg += '開訓迄日請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }
            if (!IsDate(SFTDate.value)) {
                SFTDate.focus();
                msg += '結訓起日請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }
            if (!IsDate(EFTDate.value)) {
                EFTDate.focus();
                msg += '結訓迄日請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function IsDate(MyDate) {
            if (MyDate != '') {
                if (!checkDate(MyDate)) { return false; }
            }
            return true;
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務報表管理&gt;&gt;開班資料</asp:Label>
                </td>
            </tr>
        </table>
        <table cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch" id="Table1" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol_need" width="20%">年度</td>
                            <td class="whitecol">
                                <asp:DropDownList ID="yearlist" runat="server"></asp:DropDownList>
                                <asp:RequiredFieldValidator ID="MustYear" runat="server" ErrorMessage="請選擇年度" Display="Dynamic" ControlToValidate="yearlist" CssClass="font"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">轄區</td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="DistrictList" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="4"></asp:CheckBoxList>
                                <input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">縣市</td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="CityList" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="8"></asp:CheckBoxList>
                                <input id="CityHidden" type="hidden" value="0" name="CityHidden" runat="server" />
                            </td>
                        </tr>
                        <tr id="TPlan_item_TR" runat="server">
                            <td class="bluecol">訓練計畫</td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="PlanList" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3"></asp:CheckBoxList>
                                <input id="TPlanHidden" type="hidden" value="0" name="TPlanHidden" runat="server" />
                            </td>
                        </tr>
                        <%--
                        <tr id="TPlanID0_TR" runat="server">
						    <td class="bluecol" width="100">訓練計畫(職前)</td>
						    <td class="whitecol">
							    <asp:CheckBoxList ID="chkTPlanID0" runat="server" CellSpacing="0" CellPadding="0" RepeatColumns="3" RepeatDirection="Horizontal" CssClass="font"></asp:CheckBoxList>
							    <input id="TPlanID0HID" value="0" type="hidden" name="TPlanID0HID" runat="server">
						    </td>
					    </tr>
					    <tr id="TPlanID1_TR" runat="server">
						    <td class="bluecol">訓練計畫(在職)</td>
						    <td class="whitecol">
							    <asp:CheckBoxList ID="chkTPlanID1" runat="server" CellSpacing="0" CellPadding="0" RepeatColumns="3" RepeatDirection="Horizontal" CssClass="font"></asp:CheckBoxList>
							    <input id="TPlanID1HID" value="0" type="hidden" name="TPlanID1HID" runat="server">
						    </td>
					    </tr>
					    <tr id="TPlanIDX_TR" runat="server">
						    <td class="bluecol">訓練計畫(其他)</td>
						    <td class="whitecol">
							    <asp:CheckBoxList ID="chkTPlanIDX" runat="server" CellSpacing="0" CellPadding="0" RepeatColumns="3" RepeatDirection="Horizontal" CssClass="font"></asp:CheckBoxList>
							    <input id="TPlanIDXHID" value="0" type="hidden" name="TPlanIDXHID" runat="server">
						    </td>
					    </tr>
                        --%>
                        <tr>
                            <td class="bluecol">開訓日期</td>
                            <td class="whitecol">
                                <span id="span01" runat="server">
                                    <asp:TextBox ID="SSTDate" runat="server" Columns="10" MaxLength="10" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= SSTDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">&nbsp;~
                                    <asp:TextBox ID="ESTDate" runat="server" Columns="10" MaxLength="10" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= ESTDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">結訓日期</td>
                            <td class="whitecol">
                                <span id="span02" runat="server">
                                    <asp:TextBox ID="SFTDate" runat="server" Columns="10" MaxLength="10" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= SFTDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">&nbsp;~
                                    <asp:TextBox ID="EFTDate" runat="server" Columns="10" MaxLength="10" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= EFTDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">開班狀態</td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="NotOpenStaus" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="開班" Selected="True">開班</asp:ListItem>
                                    <asp:ListItem Value="不開班">不開班</asp:ListItem>
                                </asp:CheckBoxList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">機構名稱</td>
                            <td class="whitecol">
                                <asp:TextBox ID="OrgName" runat="server" Width="20%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">班級名稱</td>
                            <td class="whitecol">
                                <asp:TextBox ID="ClassCName" runat="server" Width="20%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">職類</td>
                            <td class="whitecol">
                                <asp:DropDownList ID="TMID" runat="server"></asp:DropDownList></td>
                        </tr>
                        <%--<tr>
                            <td class="bluecol">預算來源</td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="BudgetList" runat="server" CssClass="font" RepeatDirection="Horizontal"></asp:CheckBoxList></td>
                        </tr>--%>
                        <tr>
                            <td class="bluecol">匯出檔案格式</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                                    <asp:ListItem Value="ODS">ODS</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>

                    </table>
                </td>
            </tr>
        </table>
        <table class="font" id="Table6" cellspacing="0" cellpadding="0" border="0" width="100%">
            <tr align="center">
                <td class="whitecol">
                    <font>
                        <asp:Button ID="bt_search" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="bt_reset" runat="server" Text="重新設定" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="bt_export" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                    </font>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="HIDOrgID" runat="server" />
    </form>
</body>
</html>
