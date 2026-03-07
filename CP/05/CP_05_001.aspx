<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_05_001.aspx.vb" Inherits="WDAIIP.CP_05_001" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>CP_05_001</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript">
        function SelectAll(flag, num) {
            for (i = 0; i < num; i++) {
                var mycheck = document.getElementById('CTID_' + i);
                mycheck.checked = flag;
            }
        }

        function check_data() {
            var msg = '';
            if (parseInt(getCheckBoxListValue('CTID')) == 0) msg += '請選擇縣市\n';
            if (document.form1.start_date.value != '' && !checkDate(document.form1.start_date.value)) msg += '開結訓日起日輸入格式不正確\n';
            if (document.form1.end_date.value != '' && !checkDate(document.form1.end_date.value)) msg += '開結訓日迄日輸入格式不正確\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">

        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <font class="font" size="2">首頁&gt;&gt;訓練查核與績效管理&gt;&gt;專案計畫查核&gt;&gt;</font><font class="font" color="#800000" size="2">新興資軟查核</font>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="SearchTable" runat="server" cellpadding="1" cellspacing="1" width="100%">
                        <tr>
                            <td align="center">
                                <table class="table_sch" id="Table2" cellpadding="1" cellspacing="1">
                                    <tr>
                                        <td width="100" class="bluecol">縣市
                                        </td>
                                        <td class="td_light">
                                            <asp:CheckBox ID="CheckBox1" runat="server" Text="全部"></asp:CheckBox><asp:CheckBoxList ID="CTID" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" RepeatColumns="6">
                                            </asp:CheckBoxList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">機構名稱
                                        </td>
                                        <td class="td_light">
                                            <asp:TextBox ID="SearchOrgName" runat="server"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">開結訓日
                                        </td>
                                        <td class="td_light">
                                            <asp:TextBox ID="start_date" runat="server" Width="100px"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">～
                                            <asp:TextBox ID="end_date" runat="server" Width="100px"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                        </td>
                                    </tr>
                                </table>
                                <p align="center">
                                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue" DESIGNTIMEDRAGDROP="30">顯示列數</asp:Label>
                                    <asp:TextBox ID="TxtPageSize" runat="server" Width="23px" MaxLength="2">10</asp:TextBox>
                                    <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>
                                </p>
                                <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                    <tr>
                                        <td>
                                            <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" AllowPaging="True">
                                                <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                <Columns>
                                                    <asp:BoundColumn HeaderText="序號">
                                                        <HeaderStyle Width="25px"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn HeaderText="管控單位">
                                                        <HeaderStyle Width="150px"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="OrgName" HeaderText="機構名稱"></asp:BoundColumn>
                                                    <asp:TemplateColumn HeaderText="功能">
                                                        <HeaderStyle Width="200px"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Button ID="Button2" runat="server" Text="查詢班級" CommandName="ShowClass"></asp:Button>
                                                            <asp:Button ID="Button3" runat="server" Text="列印開課一覽表" CommandName="PrintClass" CssClass="asp_Export_M"></asp:Button>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
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
                                </table>
                                <asp:Label ID="msg1" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="DetailTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:Label ID="OrgName" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td align="right">點選班級名稱可以查詢或新增實地訪查紀錄表
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <table id="DataGridTable2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                    <tr>
                                        <td>
                                            <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" AllowPaging="True">
                                                <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                <Columns>
                                                    <asp:BoundColumn HeaderText="序號">
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:TemplateColumn HeaderText="班級名稱">
                                                        <ItemTemplate>
                                                            <asp:LinkButton ID="LinkButton1" runat="server" ForeColor="Blue" CommandName="record">LinkButton</asp:LinkButton>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:BoundColumn HeaderText="開訓日期">
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:TemplateColumn HeaderText="功能">
                                                        <ItemTemplate>
                                                            <asp:Button ID="Button5" runat="server" Text="課程表" CommandName="Course"></asp:Button>
                                                            <asp:Button ID="Button6" runat="server" Text="學員名冊" CommandName="StudentInfo"></asp:Button>
                                                            <asp:Button ID="Button7" runat="server" Text="出缺勤" CommandName="TurnOut"></asp:Button><br>
                                                            <asp:Button ID="Button8" runat="server" Text="生活津貼印領清冊" Width="118px" CommandName="Subsidy"></asp:Button>
                                                            <asp:Button ID="Button9" runat="server" Text="生活津貼統計明細表" Width="134px" CommandName="Subsidy2"></asp:Button>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                </Columns>
                                                <PagerStyle Visible="False"></PagerStyle>
                                            </asp:DataGrid>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="center">
                                            <uc1:PageControler ID="PageControler2" runat="server"></uc1:PageControler>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <asp:Label ID="msg2" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <asp:Button ID="Button4" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
