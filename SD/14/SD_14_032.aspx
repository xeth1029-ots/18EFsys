<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_14_032.aspx.vb" Inherits="WDAIIP.SD_14_032" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>收據</title>
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
        function GETvalue() {
            document.getElementById('Button4').click();
        }

        function CheckSearch() {
            var OCIDValue1 = document.getElementById('OCIDValue1');
            if (OCIDValue1.value == '') {
                alert('請選擇班級')
                return false;
            }
        }

        function SelectItem(Flag, MyValue) {
            var SOCIDValue = document.getElementById('SOCIDValue');
            if (Flag) {
                if (SOCIDValue.value != '') { SOCIDValue.value += ','; }
                SOCIDValue.value += MyValue;
            }
            else {
                if (SOCIDValue.value.indexOf(',' + MyValue) != -1)
                    SOCIDValue.value = SOCIDValue.value.replace(',' + MyValue, '')
                else if (SOCIDValue.value.indexOf(MyValue + ',') != -1)
                    SOCIDValue.value = SOCIDValue.value.replace(MyValue + ',', '')
                else if (SOCIDValue.value.indexOf(MyValue) != -1)
                    SOCIDValue.value = SOCIDValue.value.replace(MyValue, '')
            }
        }

        function SelectAll(Flag) {
            var MyTable = document.getElementById('DataGrid1');
            for (i = 1; i < MyTable.rows.length; i++) {
                MyTable.rows[i].cells[0].children[0].checked = Flag;
                SelectItem(Flag, MyTable.rows[i].cells[0].children[0].value);
            }
        }

        function choose_class() {
            var RIDValue = document.getElementById('RIDValue');
            document.getElementById('OCID1').value = '';
            document.getElementById('TMID1').value = '';
            document.getElementById('OCIDValue1').value = '';
            document.getElementById('TMIDValue1').value = '';
            openClass('../02/SD_02_ch.aspx?&RID=' + RIDValue.value);
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;表單列印&gt;&gt;收據</asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch">
                        <tr>
                            <td class="bluecol" style="width: 20%">訓練機構</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="center" runat="server" Width="55%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="Hidden2" runat="server">
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                                <asp:Button ID="Button4" Style="display: none" runat="server" Text="Button4"></asp:Button>
                                <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">職類/班別</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="25%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" Width="30%" onfocus="this.blur()"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="button_b_Mini">
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                <span id="HistoryList" style="position: absolute; display: none; left: 28%">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr id="TRPlanPoint28" runat="server">
                            <td class="bluecol">計畫</td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="PlanPoint" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" CssClass="font">
                                    <asp:ListItem Value="1" Selected="True">產業人才投資計畫</asp:ListItem>
                                    <asp:ListItem Value="2">提升勞工自主學習計畫</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                    <div align="center" class="whitecol">
                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                        <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                    </div>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                    </div>
                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#f5f5f5" />
                                    <HeaderStyle CssClass="head_navy" HorizontalAlign="Center" />
                                    <ItemStyle HorizontalAlign="Center" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="8%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <HeaderTemplate>
                                                <input onclick="SelectAll(this.checked);" type="checkbox">
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <input id="SOCID" type="checkbox" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="STUDID2" HeaderText="學號">
                                            <HeaderStyle Width="12%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="NAME" HeaderText="姓名">
                                            <HeaderStyle Width="12%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="IDNO_MK" HeaderText="身分證號碼">
                                            <HeaderStyle Width="15%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="SEX_N" HeaderText="性別">
                                            <HeaderStyle Width="8%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="STUDSTATUS_N" HeaderText="狀態">
                                            <HeaderStyle Width="12%"></HeaderStyle>
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
                                <asp:Button ID="BtnPrint1" runat="server" Text="列印" CssClass="asp_Export_M" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <input id="Hid_OrgKind2" type="hidden" name="Hid_OrgKind2" runat="server" />
        <input id="Hid_MSD" type="hidden" name="Hid_MSD" runat="server" />
        <input id="Years" type="hidden" name="Years" runat="server" />
        <input id="SOCIDValue" type="hidden" name="SOCIDValue" runat="server" />
        <%--<input id="KindValue" type="hidden" name="KindValue" runat="server">--%>
    </form>
</body>
</html>
