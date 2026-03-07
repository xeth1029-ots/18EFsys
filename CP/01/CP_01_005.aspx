<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_01_005.aspx.vb" Inherits="WDAIIP.CP_01_005" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>實地訪查稽核表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="../../js/TIMS.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function HiddenTable(TableName, State) {
            if (document.getElementById(TableName)) {
                if (document.getElementById(State).value == '1') {
                    document.getElementById(TableName).style.display = 'none';
                    document.getElementById(State).value = '0';
                }
                else {
                    document.getElementById(TableName).style.display = '';
                    document.getElementById(State).value = '1';
                }
            }
            return false;
        }

        function search() {
            var msg = '';
            if (!CheckMyDate(document.getElementById('SDate').value)) msg += '起始日期不是正確的時間格式\n';
            if (!CheckMyDate(document.getElementById('EDate').value)) msg += '結束日期不是正確的時間格式\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function CheckMyDate(MyDate) {
            if (MyDate != '') {
                return checkDate(MyDate)
            }
            return true;
        }

        function check_data() {
            var msg = '';
            if (document.getElementById('TraceDate').value == '') msg += '請輸入追蹤日期\n';
            else if (!checkDate(document.getElementById('TraceDate').value)) msg += '追蹤日期不是正確時間格式\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function ChangeMode() {
            var num = getRadioValue(document.getElementsByName('ShowMode'));
            if (num == '1') {
                document.getElementById('ShowModeTable1').style.display = '';
                document.getElementById('ShowModeTable2').style.display = 'none';
            }
            else {
                document.getElementById('ShowModeTable1').style.display = 'none';
                document.getElementById('ShowModeTable2').style.display = '';
            }
        }
    </script>
</head>
<body onload="FrameLoad();">
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;查核績效管理&gt;&gt;實地訪查稽核表</asp:Label>
                </td>
            </tr>
        </table>

        <table class="font" id="FrameTable1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table id="Page1" runat="server" width="100%" class="font" cellspacing="1" cellpadding="1">
                        <tr>
                            <td align="center">
                                <table class="table_sch" id="SearchTable" runat="server" cellspacing="1" cellpadding="1">
                                    <tr>
                                        <td class="bluecol" width="20%">年度</td>
                                        <td class="whitecol">
                                            <asp:DropDownList ID="ddlYears" runat="server" AutoPostBack="True"></asp:DropDownList></td>
                                    </tr>
                                    <tr id="TR1" runat="server">
                                        <td class="bluecol">轄區</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="rblDistID" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="3"></asp:RadioButtonList></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">訓練計畫</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="rblTPlanID" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="3"></asp:RadioButtonList></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">機構名稱</td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="SOrgName" runat="server" Width="40%"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">只顯示有<br>
                                            查核資料</td>
                                        <td class="whitecol">
                                            <asp:CheckBox ID="chkVisirOnly" runat="server" Text="是"></asp:CheckBox></td>
                                    </tr>
                                </table>
                                <div align="center" class="whitecol">
                                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue" DESIGNTIMEDRAGDROP="30">顯示列數</asp:Label>
                                    <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                    <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                    <input id="TableState1" type="hidden" value="1" name="TableState1" runat="server">
                                </div>
                                <table class="font" id="DataGridTable1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                    <tr>
                                        <td align="right">
                                            <asp:LinkButton ID="LinkButton3" runat="server" ForeColor="Blue">關閉/展開搜尋條件</asp:LinkButton></td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Width="100%" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                                                <AlternatingItemStyle BackColor="WhiteSmoke"></AlternatingItemStyle>
                                                <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                                                <Columns>
                                                    <asp:BoundColumn DataField="DistName" HeaderText="轄區">
                                                        <HeaderStyle Width="11%"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="PlanName" HeaderText="訓練計畫">
                                                        <HeaderStyle Width="11%"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn HeaderText="訪查設定">
                                                        <HeaderStyle Width="11%"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:TemplateColumn HeaderText="訓練機構">
                                                        <HeaderStyle Width="12%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:LinkButton ID="LinkButton1" runat="server" ForeColor="Blue"></asp:LinkButton>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:BoundColumn DataField="ClassCount" HeaderText="開班數">
                                                        <HeaderStyle Width="11%"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="FinCount" HeaderText="結訓班數">
                                                        <HeaderStyle Width="11%"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn HeaderText="預定次數"></asp:BoundColumn>
                                                    <asp:BoundColumn DataField="VisitCount" HeaderText="實際次數">
                                                        <HeaderStyle Width="11%"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn HeaderText="燈號">
                                                        <HeaderStyle Width="11%"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                </Columns>
                                                <PagerStyle Visible="False"></PagerStyle>
                                            </asp:DataGrid>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="center" class="whitecol">
                                            <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                            <br>
                                            <asp:Button ID="Button9" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                                        </td>
                                    </tr>
                                </table>
                                <asp:Label ID="msg1" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="Page2" cellspacing="0" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td align="center">
                                <table class="table_sch" id="FilterTable" runat="server">
                                    <tr>
                                        <td class="bluecol" width="20%">訪查日期</td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="SDate" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                            <img style="cursor: pointer" onclick="javascript:show_calendar('SDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                            ～
                                                <asp:TextBox ID="EDate" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                            <img style="cursor: pointer" onclick="javascript:show_calendar('EDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">查核結果</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="VisitResult" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                                                <asp:ListItem Value="不區分" Selected="True">不區分</asp:ListItem>
                                                <asp:ListItem Value="3">綠燈</asp:ListItem>
                                                <asp:ListItem Value="2">黃燈</asp:ListItem>
                                                <asp:ListItem Value="1">紅燈</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">結案狀況</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="IsClear" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                                                <asp:ListItem Value="不區分" Selected="True">不區分</asp:ListItem>
                                                <asp:ListItem Value="1">已結案</asp:ListItem>
                                                <asp:ListItem Value="0">未結案</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                </table>
                                <div align="center" class="whitecol">
                                    <asp:Label ID="labPageSize2" runat="server" ForeColor="SlateBlue" DESIGNTIMEDRAGDROP="30">顯示列數</asp:Label>
                                    <asp:TextBox ID="TxtPageSize2" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                    <asp:Button ID="Button2" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                    <input id="TableState2" type="hidden" value="0" name="TableState2" runat="server">
                                    <table class="table_nw" id="Table5" cellspacing="1" cellpadding="1" width="100%" border="0">
                                        <tr>
                                            <td class="bluecol" width="20%">訓練計畫：<asp:Label ID="PlanName" runat="server"></asp:Label></td>
                                            <td class="bluecol">訓練機構：<asp:Label ID="OrgName" runat="server"></asp:Label></td>
                                            <td class="bluecol">
                                                <asp:LinkButton ID="LinkButton2" runat="server" ForeColor="Blue">關閉/展開搜尋條件</asp:LinkButton></td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">訪查類型：<asp:Label ID="CheckMode" runat="server"></asp:Label></td>
                                            <td class="bluecol">預定次數：<asp:Label ID="DeCount" runat="server"></asp:Label></td>
                                            <td class="bluecol">實際次數：<asp:Label ID="RelCount" runat="server"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol" colspan="3">顯示模式：
                                                    <asp:RadioButtonList ID="ShowMode" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                                                        <asp:ListItem Value="1" Selected="True">已查核的班級</asp:ListItem>
                                                        <asp:ListItem Value="2">尚未查核的班級</asp:ListItem>
                                                    </asp:RadioButtonList>
                                            </td>
                                        </tr>
                                    </table>
                                </div>
                                <table class="font" id="ShowModeTable1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                    <tr>
                                        <td>
                                            <table class="font" id="DataGridTable2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                                <tr>
                                                    <td>
                                                        <asp:DataGrid ID="DataGrid2" runat="server" CssClass="font" Width="100%" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                                                            <AlternatingItemStyle BackColor="WhiteSmoke"></AlternatingItemStyle>
                                                            <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                                                            <Columns>
                                                                <asp:BoundColumn DataField="ApplyDate" HeaderText="訪查日期" DataFormatString="{0:d}">
                                                                    <HeaderStyle Width="14%" />
                                                                </asp:BoundColumn>
                                                                <asp:BoundColumn DataField="ClassCName" HeaderText="班別名稱">
                                                                    <HeaderStyle Width="14%" />
                                                                </asp:BoundColumn>
                                                                <asp:BoundColumn DataField="FTDate" HeaderText="結訓日期" DataFormatString="{0:d}">
                                                                    <HeaderStyle Width="14%" />
                                                                </asp:BoundColumn>
                                                                <asp:BoundColumn HeaderText="訪查結果">
                                                                    <HeaderStyle Width="16%" />
                                                                </asp:BoundColumn>
                                                                <asp:BoundColumn DataField="WrongCount" HeaderText="缺失項目個數">
                                                                    <HeaderStyle Width="14%" />
                                                                </asp:BoundColumn>
                                                                <asp:BoundColumn HeaderText="結案狀況">
                                                                    <HeaderStyle Width="14%" />
                                                                </asp:BoundColumn>
                                                                <asp:TemplateColumn HeaderText="功能">
                                                                    <HeaderStyle Width="14%" />
                                                                    <ItemStyle HorizontalAlign="Center" />
                                                                    <ItemTemplate>
                                                                        <asp:Button ID="Button3" runat="server" Text="查看" CommandName="view" CssClass="asp_button_M"></asp:Button>
                                                                        <asp:Button ID="Button4" runat="server" Text="結案" CommandName="clear" CssClass="asp_button_M"></asp:Button>
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
                                                <tr>
                                                    <td align="center"></td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="center">
                                            <asp:Label ID="msg2" runat="server" ForeColor="Red"></asp:Label></td>
                                    </tr>
                                </table>
                                <table class="font" id="ShowModeTable2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                    <tr>
                                        <td>
                                            <table class="font" id="DataGridTable3" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                                <tr>
                                                    <td>
                                                        <asp:DataGrid ID="DataGrid3" runat="server" CssClass="font" Width="100%" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                                                            <AlternatingItemStyle BackColor="WhiteSmoke"></AlternatingItemStyle>
                                                            <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                                                            <Columns>
                                                                <asp:BoundColumn DataField="CLASSCNAME2" HeaderText="班級名稱">
                                                                    <HeaderStyle Width="80%" />
                                                                </asp:BoundColumn>
                                                                <asp:BoundColumn DataField="STDate" HeaderText="開訓日期" DataFormatString="{0:d}">
                                                                    <HeaderStyle Width="10%"></HeaderStyle>
                                                                </asp:BoundColumn>
                                                                <asp:BoundColumn DataField="FTDate" HeaderText="結訓日期" DataFormatString="{0:d}">
                                                                    <HeaderStyle Width="10%"></HeaderStyle>
                                                                </asp:BoundColumn>
                                                            </Columns>
                                                            <PagerStyle Visible="False"></PagerStyle>
                                                        </asp:DataGrid>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td align="center">
                                                        <uc1:PageControler ID="PageControler3" runat="server"></uc1:PageControler>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td align="center"></td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="center">
                                            <asp:Label ID="msg3" runat="server" ForeColor="Red"></asp:Label></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="Button5" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button></td>
                        </tr>
                    </table>
                    <table class="font" id="Page3" cellspacing="0" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <table class="font" id="Table4" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td>
                                            <table class="font" id="Table6" cellspacing="1" cellpadding="1" width="100%" border="0">
                                                <tr>
                                                    <td colspan="2" class="bluecol" width="50%">訓練機構：<asp:Label ID="OrgName1" runat="server"></asp:Label></td>
                                                    <td colspan="2" class="bluecol" width="50%">班別名稱：<asp:Label ID="ClassCName1" runat="server"></asp:Label></td>
                                                </tr>
                                                <tr>
                                                    <td class="bluecol" width="20%">訪查日期：<asp:Label ID="ApplyDate1" runat="server"></asp:Label></td>
                                                    <td class="bluecol" width="30%">追蹤日期：<asp:TextBox ID="TraceDate" runat="server" Columns="10" Width="60%" CssClass="whitecol"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('TraceDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></td>
                                                    <td class="bluecol" width="30%">查核結果：<asp:Label ID="VisitResult1" runat="server"></asp:Label></td>
                                                    <td class="bluecol" width="20%">結案狀況：<asp:Label ID="IsClear1" runat="server"></asp:Label></td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <table class="table_sch" id="Table7" runat="server">
                                                <tr>
                                                    <td class="bluecol">訪查項目</td>
                                                    <td class="bluecol">書面項目<br>
                                                        是否具備</td>
                                                    <td class="bluecol">處裡情形</td>
                                                    <td class="bluecol">追蹤狀況</td>
                                                    <td class="bluecol">追蹤備註</td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(書)訓練日誌</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Data1" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="DataCopy1" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Data1Trace" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Data1TNote" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(書)學員名冊</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Data2" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="DataCopy2" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Data2Trace" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Data2TNote" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(書)學員簽到(退)表</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Data3" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="DataCopy3" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Data3Trace" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Data3TNote" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(書)課程表</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Data4" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="DataCopy4" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Data4Trace" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Data4TNote" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(書)請假單</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Data5" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="DataCopy5" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Data5Trace" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Data5TNote" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(書)退訓申請單</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Data6" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="DataCopy6" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Data6Trace" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Data6TNote" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(書)勞保投保資料</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Data7" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="DataCopy7" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Data7Trace" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Data7TNote" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)有無週(月)課程表?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item1_1" runat="server"></asp:Label></td>
                                                    <td rowspan="2" class="whitecol">
                                                        <asp:Label ID="Item1Pros" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item1_1Trace" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item1_1TNote" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)是否依課程表授課?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item1_2" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item1_2Trace" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item1_2TNote" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)教學(訓練)日誌是否確實填寫?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item2_1" runat="server"></asp:Label></td>
                                                    <td rowspan="2" class="whitecol">
                                                        <asp:Label ID="Item2Pros" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item2_1Trace" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item2_1TNote" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)有否按時呈主管核閱?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item2_2" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item2_2Trace" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item2_2TNote" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)教師(職業訓練師)與助教姓名?是否與計畫相符?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item3_1" runat="server"></asp:Label></td>
                                                    <td rowspan="2" class="whitecol">
                                                        <asp:Label ID="Item3Pros" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item3_1Trace" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item3_1TNote" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)學員學習情況是否良好?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item3_2" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item3_2Trace" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item3_2TNote" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)教學環境(教室或訓練工場)是否整齊?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item4_1" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item4Pros" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item4_1Trace" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item4_1TNote" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)有無具體輔導活動或其他事項?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item5_1" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item5Pros" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item5_1Trace" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item5_1TNote" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)職業訓練生活津貼是否依規定申請並發放?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item6_1" runat="server"></asp:Label></td>
                                                    <td class="whitecol"></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item6_1Trace" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item6_1TNote" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <table class="table_sch" id="Table7_97" runat="server">
                                                <tr>
                                                    <td class="bluecol">訪查項目</td>
                                                    <td class="bluecol">書面項目<br>
                                                        是否具備</td>
                                                    <td class="bluecol">處裡情形</td>
                                                    <td class="bluecol">追蹤狀況</td>
                                                    <td class="bluecol">追蹤備註</td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(書)教學訓練日誌</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Data1_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="DataCopy1_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Data1Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Data1TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(書)學員簽到(退)表</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Data3_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="DataCopy3_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Data3Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Data3TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(書)請假單</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Data5_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="DataCopy5_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Data5Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Data5TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(書)退訓/提前就業申請</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Data6_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="DataCopy6_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Data6Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Data6TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr id="Data7_TR" runat="server">
                                                    <td class="whitecol">(書)勞保加/退保明細表</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Data7_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="DataCopy7_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Data7Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Data7TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr id="Data9_TR" runat="server">
                                                    <td class="whitecol">(書)學員書籍(講義)、材料領用表</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Data9_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="DataCopy9_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Data9Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Data9TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(書)職訓生活津貼補助印領清冊</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Data10_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="DataCopy10_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Data10Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Data10TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr id="Data11_TR" runat="server">
                                                    <td class="whitecol">(書)學員服務手冊或權利義務公告相關文件</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Data11_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="DataCopy11_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Data11Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Data11TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)有無週(月)課程表?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item1_1_97" runat="server"></asp:Label></td>
                                                    <td rowspan="8" class="whitecol">
                                                        <asp:Label ID="Item1Pros_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item1_1Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item1_1TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)是否依課程表授課?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item1_2_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item1_2Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item1_2TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)教師與助教是否與計畫相符?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item3_1_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item3_1Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item3_1TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)指導輔導員是否在旁協助?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item14_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item14Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item14TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)是否協助報到學員登錄帳號?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item15_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item15Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item15TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)是否依規定隨到隨訓?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item16_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item16Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item16TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)數位學習中心研習的學員是否與系統顯示之學員相符?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item17_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item17Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item17TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)學員是否皆持學習卷參加線上研習?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item18_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item18Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item18TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)有無書籍(講義)領用表?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item19_97" runat="server"></asp:Label></td>
                                                    <td rowspan="4" class="whitecol">
                                                        <asp:Label ID="Item2Pros_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item19Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item19TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)有無材料領用表?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item20_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item20Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item20TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)訓練設施設備是否依契約提供學員使用?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item21_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item21Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item21TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)課程連線是否正常?(學習券線上研習)</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item22_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item22Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item22TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)教學(訓練)日誌是否確實填寫?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item2_1_97" runat="server"></asp:Label></td>
                                                    <td rowspan="7" class="whitecol">
                                                        <asp:Label ID="Item3Pros_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item2_1Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item2_1TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)有否按時呈主管核閱?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item2_2_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item2_2Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item2_2TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)學員生活、就業輔導與管理機制是否依契約規範辦理?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item23_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item23Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item23TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)是否依契約規範提供學員問題反應申訴管道?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item24_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item24Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item24TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)是否為參訓學員辦理勞工保險加退保?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item25_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item25Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item25TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)是否依契約規範公告學員權益義務或製參訓學員服務手冊?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item26_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item26Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item26TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)對己結訓學員是否依規定發給研習證書?(學習券線上研習)</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item27_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item27Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item27TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr id="Item28_TR" runat="server">
                                                    <td class="whitecol">(查)有無自費參訓學員?幾人?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item28_97" runat="server"></asp:Label></td>
                                                    <td rowspan="4" class="whitecol">
                                                        <asp:Label ID="Item4Pros_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item28Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item28TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr id="Item28_2_TR" runat="server">
                                                    <td class="whitecol">(查)訓練費用是否繳交主辦單位?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item28_2_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item28_2Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item28_2TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">(查)職業訓練生活津貼是否依規定申請並核發?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item6_1_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item6_1Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item6_1TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr id="Item29_TR" runat="server">
                                                    <td class="whitecol">(查)培訓單位無巧立名目強制收取費用?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item29_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item29Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item29TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                                <tr id="Item30_TR" runat="server">
                                                    <td class="whitecol">(查)職業訓練機構是否依規定懸掛設立許可證書?</td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item30_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:Label ID="Item5Pros_97" runat="server"></asp:Label></td>
                                                    <td class="whitecol">
                                                        <asp:DropDownList ID="Item30Trace_97" runat="server">
                                                            <asp:ListItem Value="1">待追蹤</asp:ListItem>
                                                            <asp:ListItem Value="2">已改善</asp:ListItem>
                                                            <asp:ListItem Value="3">未改善</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td class="whitecol" align="center">
                                                        <asp:TextBox ID="Item30TNote_97" runat="server" Columns="15" TextMode="MultiLine" Width="80%"></asp:TextBox></td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="center" class="whitecol">
                                            <asp:Button ID="Button7" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                                            <asp:Button ID="Button6" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>

    </form>
</body>
</html>
