<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_015.aspx.vb" Inherits="WDAIIP.SD_15_015" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>師資授課時數統計表</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
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
    <script type="text/javascript">
        function CheckSearch1() {
            var STDate1 = document.getElementById('STDate1').value;
            var STDate2 = document.getElementById('STDate2').value;
            var FTDate1 = document.getElementById('FTDate1').value;
            var FTDate2 = document.getElementById('FTDate2').value;

            var msg = '';
            if (document.getElementById('yearlist').selectedIndex == 0) msg += '請選擇年度\n';
            if (document.getElementById('DistID').selectedIndex == 0) msg += '請選擇轄區\n';

            if (!checkDate(STDate1) && STDate1 != '') msg += '開訓起始日期必須為正確日期格式\n';
            if (!checkDate(STDate2) && STDate2 != '') msg += '開訓結束日期必須為正確日期格式\n';

            if (!checkDate(FTDate1) && FTDate1 != '') msg += '結訓起始日期必須為正確日期格式\n';
            if (!checkDate(FTDate2) && FTDate2 != '') msg += '結訓結束日期必須為正確日期格式\n';

            if (msg != '') {
                alert(msg);
                return false;
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;統計表&gt;&gt;師資授課時數統計表</asp:Label>
                </td>
            </tr>
        </table>
        <table class="table_sch" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td class="bluecol" width="20%">年度
                </td>
                <td class="whitecol" colspan="3">
                    <asp:DropDownList ID="yearlist" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">轄區
                </td>
                <td class="whitecol" colspan="3">
                    <asp:DropDownList ID="DistID" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">講師姓名
                </td>
                <td class="whitecol" colspan="3">
                    <asp:TextBox Style="z-index: 0" ID="txtName" runat="server" MaxLength="200" Columns="50" Width="20%"></asp:TextBox>
                    (可用半型逗點分隔,不可有空隔或特殊符號)
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">身分證字號
                </td>
                <td class="whitecol" colspan="3">
                    <asp:TextBox Style="z-index: 0" ID="txtIDNO" runat="server" MaxLength="200" Columns="50" Width="20%"></asp:TextBox>
                    (可用半型逗點分隔,不可有空隔或特殊符號)
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">開訓期間
                </td>
                <td class="whitecol" colspan="3">
                    <span id="span01" runat="server">
                        <asp:TextBox ID="STDate1" runat="server" Width="15%"></asp:TextBox>&nbsp;
                    <img style="cursor: pointer" alt="#" onclick="Javascript:show_calendar('<%= STDate1.ClientId %>','','','CY/MM/DD');" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                        &nbsp;~&nbsp;
                    <asp:TextBox ID="STDate2" runat="server" Width="15%"></asp:TextBox>&nbsp;
                    <img style="cursor: pointer" alt="#" onclick="Javascript:show_calendar('<%= STDate2.ClientId %>','','','CY/MM/DD');" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                    </span>
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">結訓期間
                </td>
                <td class="whitecol" colspan="3">
                    <span id="span02" runat="server">
                        <asp:TextBox ID="FTDate1" runat="server" Width="15%" Style="z-index: 0"></asp:TextBox>&nbsp;
                    <img style="cursor: pointer" alt="#" onclick="Javascript:show_calendar('<%= FTDate1.ClientId %>','','','CY/MM/DD');" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                        &nbsp;~&nbsp;
                    <asp:TextBox ID="FTDate2" runat="server" Width="15%"></asp:TextBox>&nbsp;
                    <img style="cursor: pointer" alt="#" onclick="Javascript:show_calendar('<%= FTDate2.ClientId %>','','','CY/MM/DD');" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                    </span>
                </td>
            </tr>
            <tr>
                <td class="bluecol">課程審核狀況</td>
                <td class="whitecol" colspan="3">
                    <asp:RadioButtonList ID="audit" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                        <asp:ListItem Value="" Selected="True">不拘</asp:ListItem>
                        <asp:ListItem Value="N">審核中</asp:ListItem>
                        <asp:ListItem Value="Y">已通過</asp:ListItem>
                    </asp:RadioButtonList>
                    <%--<asp:RadioButtonList ID="audit" runat="server" CssClass="font" RepeatDirection="horizontal" RepeatLayout="flow" Visible="false">
                    <asp:ListItem Value="A" Selected="true">不區分</asp:ListItem>
                    <asp:ListItem Value="Y">己審核</asp:ListItem>
                    <asp:ListItem Value="N">審核中</asp:ListItem>
                    </asp:RadioButtonList>--%>
                </td>
            </tr>
            <tr id="tr_AppStage_TP28" runat="server">
                <td class="bluecol">申請階段 </td>
                <td class="whitecol" colspan="3">
                    <asp:DropDownList ID="AppStage" runat="server"></asp:DropDownList></td>
            </tr>
            <tr>
                <td class="bluecol">匯出檔案格式</td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                        <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                        <asp:ListItem Value="ODS">ODS</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td align="center" colspan="4">(依目前登入計畫搜尋)</td>
            </tr>
            <tr>
                <td align="center" colspan="4" class="whitecol">
                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                    <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                    <asp:Button ID="btnSearch1" runat="server" Text="查詢" Style="z-index: 0" CssClass="asp_button_M"></asp:Button>&nbsp;
                    <asp:Button ID="btnPrint1" runat="server" Text="列印" Style="z-index: 0" CssClass="asp_Export_M"></asp:Button>&nbsp;
                    <asp:Button ID="btnExport1" runat="server" Text="匯出明細表" Style="z-index: 0" CssClass="asp_Export_M"></asp:Button>&nbsp;
                    <asp:Button ID="btnExport2" runat="server" Text="匯出統計表" Style="z-index: 0" CssClass="asp_Export_M"></asp:Button>
                </td>
            </tr>
            <tr>
                <td align="center" colspan="4">
                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                </td>
            </tr>
        </table>
        <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                        <AlternatingItemStyle BackColor="#f5f5f5" />
                        <HeaderStyle CssClass="head_navy" />
                        <Columns>
                            <asp:BoundColumn HeaderText="序號">
                                <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center" Wrap="False"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="TeachCName" HeaderText="師資姓名"></asp:BoundColumn>
                            <asp:BoundColumn DataField="IDNO_MK" HeaderText="身分證字號"></asp:BoundColumn>
                            <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構"></asp:BoundColumn>
                            <asp:BoundColumn DataField="CLASSNAME2" HeaderText="班別名稱"></asp:BoundColumn>
                            <asp:BoundColumn DataField="SFTDate" HeaderText="訓練期間"></asp:BoundColumn>
                            <asp:BoundColumn DataField="WEEKDAY1" HeaderText="星期"></asp:BoundColumn>
                            <asp:BoundColumn DataField="PName" HeaderText="上課時間"></asp:BoundColumn>
                            <asp:BoundColumn DataField="PHour" HeaderText="授課時數"></asp:BoundColumn>
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
    </form>
</body>
</html>
