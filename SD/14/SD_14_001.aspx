<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_14_001.aspx.vb" Inherits="WDAIIP.SD_14_001" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訓練單位基本資料表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function clearSelectValue() {
            var SelectValue = document.getElementById('SelectValue');
            SelectValue.value = '';
            var MyTable = document.getElementById('DataGrid1');
            if (MyTable) {
                for (i = 1; i < MyTable.rows.length; i++) {
                    MyTable.rows[i].cells[0].children[0].checked = false;
                    //SelectItem(false, MyTable.rows[i].cells[0].children[0].value);
                }
            }
            return true;
        }

        function SelectAll(Flag) {
            var MyTable = document.getElementById('DataGrid1');
            if (MyTable) {
                for (i = 1; i < MyTable.rows.length; i++) {
                    //debugger;
                    MyTable.rows[i].cells[0].children[0].checked = Flag;
                    //SelectItem(Flag, MyTable.rows[i].cells[0].children[0].value);
                }
            }
        }

        function SelectItem(Flag, MyValue) {
            var SelectValue = document.getElementById('SelectValue');
            //SelectValue
            if (Flag) {
                //insert
                if (SelectValue.value != '') { SelectValue.value += ','; }
                SelectValue.value += "'" + MyValue + "'";
            }
            else {
                //delete
                if (SelectValue.value.indexOf(',' + "'" + MyValue + "'" + ',') != -1) { SelectValue.value = SelectValue.value.replace(',' + "'" + MyValue + "'" + ',', ',') }
                if (SelectValue.value.indexOf(',' + "'" + MyValue + "'") != -1) { SelectValue.value = SelectValue.value.replace(',' + "'" + MyValue + "'", '') }
                if (SelectValue.value.indexOf("'" + MyValue + "'" + ',') != -1) { SelectValue.value = SelectValue.value.replace("'" + MyValue + "'" + ',', '') }
                if (SelectValue.value.indexOf("'" + MyValue + "'") != -1) { SelectValue.value = SelectValue.value.replace("'" + MyValue + "'", '') }
                if (SelectValue.value.indexOf(',,') != -1) { SelectValue.value = SelectValue.value.replace(',,', ',') }
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;表單列印&gt;&gt;訓練單位基本資料表</asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch" id="Table2">
                        <tr>
                            <td class="bluecol" style="width: 20%">機構名稱</td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="OrgName" runat="server" MaxLength="40" Columns="40" Width="80%"></asp:TextBox></td>
                            <td class="bluecol" style="width: 20%">統一編號</td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="ComIDNO" runat="server" MaxLength="10" Columns="10" Width="50%"></asp:TextBox></td>
                        </tr>
                        <tr id="tr_rbl_AppliedResult54" runat="server">
                            <td class="bluecol">課程審核狀態</td>
                            <td class="whitecol" colspan="3">
                                <%-- 01:不區分,02:審核通過,03:審核中,04:不通過--%>
                                <asp:RadioButtonList ID="rbl_AppliedResult54" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="01" Selected="True">不區分</asp:ListItem>
                                    <asp:ListItem Value="02">審核通過</asp:ListItem>
                                    <asp:ListItem Value="03">審核中</asp:ListItem>
                                    <asp:ListItem Value="04">不通過</asp:ListItem>
                                </asp:RadioButtonList></td>
                        </tr>
                        <tr id="TRPlanPoint28" runat="server">
                            <td class="bluecol">計畫</td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="PlanPoint" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="1" Selected="True">產業人才投資計畫</asp:ListItem>
                                    <asp:ListItem Value="2">提升勞工自主學習計畫</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="tr_AppStage_TP28" runat="server">
                            <td class="bluecol">申請階段</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="AppStage" runat="server"></asp:DropDownList></td>
                        </tr>
                    </table>
                    <div align="center" class="whitecol">
                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" MaxLength="2" Width="6%">10</asp:TextBox>
                        <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                    </div>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                    </div>
                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AllowPaging="True" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#f5f5f5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <HeaderTemplate>
                                                <input onclick="SelectAll(this.checked)" type="checkbox">
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <input id="RSID" type="checkbox" runat="server" />
                                                <input id="RID" type="hidden" runat="server" name="RID" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="DistName" HeaderText="轄區">
                                            <HeaderStyle Width="25%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構">
                                            <HeaderStyle Width="40%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ComIDNO" HeaderText="統一編號">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="PlanMaster" HeaderText="計畫主持人">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="PlanMasterPhone" HeaderText="主持人電話">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="Button4" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button></td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <input id="ROC_Years" type="hidden" runat="server" />
        <input id="SelectValue" type="hidden" runat="server" />
        <input id="KindValue" type="hidden" name="KindValue" runat="server" />
        <input id="hid_planid" type="hidden" name="hid_planid" runat="server" />
    </form>
</body>
</html>
