<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_14_004.aspx.vb" Inherits="WDAIIP.SD_14_004" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>師資基本資料表</title>
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
        function SelectItem(Flag, MyValue) {
            var TechIDValue = document.getElementById('TechIDValue');
            if (Flag && MyValue != '') {
                if (TechIDValue.value != '') TechIDValue.value += ',';
                TechIDValue.value += MyValue;
            }
            else {
                if (TechIDValue.value.indexOf(',' + MyValue) != -1)
                    TechIDValue.value = TechIDValue.value.replace(',' + MyValue, '');
                else if (TechIDValue.value.indexOf(MyValue + ',') != -1)
                    TechIDValue.value = TechIDValue.value.replace(MyValue + ',', '');
                else if (TechIDValue.value.indexOf(MyValue) != -1)
                    TechIDValue.value = TechIDValue.value.replace(MyValue, '');
            }
        }

        function CheckPrint() {
            var TechIDValue = document.getElementById('TechIDValue');
            var Years = document.getElementById('Years');
            if (TechIDValue.value == '') {
                alert('請選擇要列印的師資');
                return false;
            }
            else {
                openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=BussinessTrain&filename=SD_14_004&path=TIMS&TechID=' + TechIDValue.value + '&Years=' + Years.value, '', '');
            }
        }

        function SelectAll(Flag) {
            var MyTable = document.getElementById('DataGrid1');
            for (i = 1; i < MyTable.rows.length; i++) {
                MyTable.rows[i].cells[0].children[0].checked = Flag;
                SelectItem(Flag, MyTable.rows[i].cells[0].children[0].value);
            }
        }
    </script>
    <style type="text/css">
        .auto-style1 { color: Black; text-align: right; padding: 4px 6px; background-color: #f1f9fc; border-right: 3px solid #49cbef; height: 45px; }
        .auto-style2 { color: #333333; padding: 4px; height: 45px; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <%--<asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;表單列印&gt;&gt;師資資料表</asp:Label>--%>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;表單列印&gt;&gt;師資基本資料表</asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table id="Table2" class="table_sch">
                        <tr>
                            <td class="bluecol" width="20%">訓練機構</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" style="width: 33px; height: 22px">
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">身分證號碼</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="IDNO" runat="server" Width="22%" MaxLength="15"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol" style="width: 20%">講師代碼</td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="TeacherID" runat="server" Width="66%" MaxLength="15"></asp:TextBox></td>
                            <td class="bluecol" style="width: 20%">講師姓名</td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="TeachCName" runat="server" Width="66%" MaxLength="33"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">內外聘</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="KindEngage" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="不區分" Selected="True">不區分</asp:ListItem>
                                    <asp:ListItem Value="1">內聘</asp:ListItem>
                                    <asp:ListItem Value="2">外聘</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                            <td class="bluecol">排課使用</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="WorkStatus" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="不區分" Selected="True">不區分</asp:ListItem>
                                    <asp:ListItem Value="1">是</asp:ListItem>
                                    <asp:ListItem Value="2">否</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="TRPlanPoint28" runat="server">
                            <td class="auto-style1">計畫</td>
                            <td class="auto-style2" colspan="3">
                                <asp:RadioButtonList ID="PlanPoint" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="1" Selected="True">產業人才投資計畫</asp:ListItem>
                                    <asp:ListItem Value="2">提升勞工自主學習計畫</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                    <div align="center" class="whitecol">
                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" MaxLength="2" Width="6%">10</asp:TextBox>
                        <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                    </div>

                    <div align="center" class="whitecol">
                    </div>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                    </div>
                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" class="font" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AutoGenerateColumns="False" AllowPaging="True" CssClass="font" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#f5f5f5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <HeaderTemplate>
                                                <input onclick="SelectAll(this.checked);" type="checkbox">
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <input id="TechID" type="checkbox" runat="server">
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="TeacherID" HeaderText="講師代碼">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TeachCName" HeaderText="講師姓名">
                                            <HeaderStyle Width="15%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="IDNO" HeaderText="身分證號碼">
                                            <HeaderStyle Width="15%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="KindEngage_N" HeaderText="內外聘">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="WorkStatus_N" HeaderText="排課使用">
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
                                <%--<input id="Button3" type="button" value="列印" name="Button3" runat="server" class="asp_button_S" onclick="return Button3_onclick()">--%>
                                <asp:Button ID="BtnPrint1" runat="server" Text="列印" CssClass="asp_Export_M" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <input id="TechIDValue" type="hidden" runat="server" />
        <input id="Years" type="hidden" name="Years" runat="server" />
    </form>
</body>
</html>
