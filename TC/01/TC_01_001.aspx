<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_001.aspx.vb" Inherits="WDAIIP.TC_01_001" %>

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>計畫代碼設定</title>
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

        function but_edit(id) {
            location.href = 'TC_01_001_add.aspx?ID=' + getParamValue('ID') + '&editid=' + id;
        }

        /*
        function chkSubmit() {
            var yearlist = document.getElementById('yearlist');
            var planlist = document.getElementById('planlist');
            var msg = '';
            if (getSelectValue(yearlist) == '') msg += '請選擇年度。\n';
            if (getSelectValue(planlist) == '') msg += '請選擇計劃。\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }
        */

        /*
		function yearPlan(selectedPlanID) {
		var year = document.getElementById('yearlist');
		var parms = "[['year','" + year.value + "']]";   //透過selectControl傳遞給SQLMap的年度查詢條件,格式請參考selectControl定義說明
		selectControl('ajaxTPlanList', 'planlist', 'PlanName', 'TPlanID', '請選擇', selectedPlanID, parms);
		}
		*/
    </script>
    <style type="text/css">
        .auto-style1 { color: #FF0000; text-align: right; padding: 4px 6px; background-color: #f1f9fc; border-right: 3px solid #49cbef; height: 40px; }
        .auto-style2 { color: #333333; padding: 4px; height: 40px; }
        .auto-style3 { color: Black; text-align: right; padding: 4px 6px; background-color: #f1f9fc; border-right: 3px solid #49cbef; height: 64px; }
        .auto-style4 { color: #333333; padding: 4px; height: 64px; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;基本資料設定&gt;&gt;計畫代碼設定</asp:Label>
                </td>
            </tr>
        </table>
        <br>
        <table width="100%" cellpadding="1" cellspacing="1" class="table_sch">
            <tr id="trDistid" runat="server">
                <%--<td class="bluecol_need" width="20%">轄區中心</td>--%>
                <td class="bluecol_need" width="20%">轄區分署</td>
                <td class="whitecol" width="80%">
                    <asp:CheckBoxList ID="DistID" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="4"></asp:CheckBoxList>
                    <input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server">
                </td>
            </tr>
            <tr style="display: none">
                <td class="bluecol" width="20%">其他選項</td>
                <td class="whitecol" width="80%">
                    <asp:CheckBox Style="z-index: 0" ID="cbk1" runat="server" Text="含不啟用的計畫"></asp:CheckBox>&nbsp;
                    <asp:Button Style="z-index: 0" ID="bt_search2" runat="server" Text="查詢" CssClass="asp_button_M" Visible="False"></asp:Button>
                </td>
            </tr>
            <tr>
                <td id="td6" runat="server" class="bluecol_need" width="20%">年度</td>
                <td class="whitecol" width="80%">
                    <asp:DropDownList ID="yearlist" AutoPostBack="True" runat="server"></asp:DropDownList>
                    <%--<asp:RequiredFieldValidator ID="MustYear" runat="server" ControlToValidate="yearlist" Display="None" ErrorMessage="請輸入年度"></asp:RequiredFieldValidator>--%>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need" width="20%">訓練計畫</td>
                <td class="whitecol" width="80%">
                    <asp:DropDownList ID="planlist" runat="server"></asp:DropDownList></td>
            </tr>
        </table>
        <table width="100%">
            <tr>
                <td class="whitecol" align="center" colspan="2">
                    <asp:Button ID="bt_search" runat="server" Text="查詢" CssClass="asp_button_M" AuthType="QRY"></asp:Button>
                    <asp:Button ID="bt_add" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_M" AuthType="ADD"></asp:Button>
                    <asp:Button ID="btn_ImpYears" runat="server" Text="匯入年度計畫" CssClass="asp_Export_M" AuthType="IMPYEAR"></asp:Button>
                    <%--<input id="Button3" type="button" value="匯入年度計畫" name="Button3" runat="server" class="asp_button_M" />--%>
                </td>
            </tr>
        </table>
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td align="center">
                    <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" AllowPaging="True" Width="100%" AutoGenerateColumns="False" CellPadding="8">
                        <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn HeaderText="序號" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="8%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="PlanName" HeaderText="計畫代碼">
                                <HeaderStyle HorizontalAlign="center" Width="40%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="SDate" HeaderText="時效起日" DataFormatString="{0:d}">
                                <HeaderStyle HorizontalAlign="center" Width="15%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="EDate" HeaderText="時效迄日" DataFormatString="{0:d}">
                                <HeaderStyle HorizontalAlign="center" Width="15%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="功能" ItemStyle-HorizontalAlign="Center">
                                <HeaderStyle HorizontalAlign="center" Width="22%"></HeaderStyle>
                                <ItemTemplate>
                                    <asp:Button ID="Button1" runat="server" Text="修改" CommandName="edit" CssClass="asp_button_M" AuthType="UPD" />
                                    <asp:Button ID="Button2" runat="server" Text="刪除" CommandName="del" CssClass="asp_button_M" AuthType="DEL" />
                                    <asp:Button ID="btnView" runat="server" Text="查詢賦予帳號" CommandName="view" CssClass="asp_button_M" AuthType="QRY" />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                </td>
            </tr>
        </table>
        <table id="Table2" cellspacing="1" cellpadding="8" width="100%" border="0" runat="server">
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CellPadding="8">
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <HeaderStyle CssClass="head_navy" Width="33%" />
                        <ItemStyle CssClass="font" HorizontalAlign="Center" />
                    </asp:DataGrid>
                </td>
            </tr>
        </table>
        <br>
        <asp:Panel ID="table_pnl" runat="server" Visible="False">
            <table id="search_tbl" class="font" border="1" cellspacing="0" cellpadding="0" width="100%" runat="server"></table>
        </asp:Panel>
        <%--<asp:ValidationSummary ID="Total" runat="server" ShowSummary="False" ShowMessageBox="True" DisplayMode="List"></asp:ValidationSummary>--%>
    </form>
</body>
</html>
