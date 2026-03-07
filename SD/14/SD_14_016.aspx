<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_14_016.aspx.vb" Inherits="WDAIIP.SD_14_016" %>
 

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>參訓學員補助費簽領清冊</title>
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
    <script type="text/javascript" language="javascript">
        function printkind() {
            document.getElementById('print_orderyby').value = 'd.IDNO'
            if (getValue("print_type") == '2') {
                document.getElementById('print_orderyby').value = 'c.StudentID'
            }
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
                                <asp:Label ID="TitleLab2" runat="server">參訓學員補助費簽領清冊</asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch">
                        <tr>
                            <td class="bluecol" style="width: 20%">訓練機構</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="center" runat="server" Width="50%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="Hidden2" runat="server">
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                                <span id="HistoryList2" style="position: absolute; display: none"><asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">開訓期間</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="STDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <span runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span> ~
                                <asp:TextBox ID="STDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <span runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">結訓期間</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="FTDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <span runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span> ~
                                <asp:TextBox ID="FTDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <span runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">列印排序方式</td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="print_type" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="1">依【身分證字號】排序列印</asp:ListItem>
                                    <asp:ListItem Value="2" Selected="True">依【學號】排序列印</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                    <div align="center" class="whitecol">
                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                        <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                    </div>
                    <div align="center"><asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></div>
                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AllowPaging="True" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <AlternatingItemStyle BackColor="#f5f5f5" />
                                    <Columns>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構" HeaderStyle-Width="35%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassCName" HeaderText="班別名稱" HeaderStyle-Width="35%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDate" HeaderText="開訓日期" DataFormatString="{0:d}" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="FTDate" HeaderText="結訓日期" DataFormatString="{0:d}" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderStyle-Width="10%">
                                            <ItemStyle HorizontalAlign="Center"/>
                                            <ItemTemplate>
                                                <asp:Label ID="labMsg1" runat="server"></asp:Label>
                                                <asp:Button ID="BtnPrint" class="asp_Export_M" runat="server" Text="列印" CommandName="print" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center"><uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler></td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <%--
	    <input id="Years" type="hidden" name="Years" runat="server">
	    <input id="SelectValue" type="hidden" name="SelectValue" runat="server">
        <input id="PlanID" type="hidden" name="PlanID" runat="server">
        --%>
        <input id="print_orderyby" type="hidden" name="print_orderyby" runat="server">
    </form>
</body>
</html>
