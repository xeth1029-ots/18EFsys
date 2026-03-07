<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_14_006.aspx.vb" Inherits="WDAIIP.SD_14_006" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訓練計畫場地資料表</title>
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
        function CheckSearch() {
            var STDate1 = document.getElementById('STDate1').value;
            var STDate2 = document.getElementById('STDate2').value;
            var FTDate1 = document.getElementById('FTDate1').value;
            var FTDate2 = document.getElementById('FTDate2').value;
            var msg = '';
            if (!checkDate(STDate1) && STDate1 != '') msg += '開訓起始日期必須為正確日期格式\n';
            if (!checkDate(STDate2) && STDate2 != '') msg += '開訓結束日期必須為正確日期格式\n';
            if (!checkDate(FTDate1) && FTDate1 != '') msg += '結訓起始日期必須為正確日期格式\n';
            if (!checkDate(FTDate2) && FTDate2 != '') msg += '結訓結束日期必須為正確日期格式\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function SelectAll(Flag) {
            var MyTable = document.getElementById('DataGrid1');
            for (var i = 1; i < MyTable.rows.length; i++) {
                var chkSeqNo = MyTable.rows[i].cells[0].children[0];
                chkSeqNo.checked = Flag;
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
                                <%--<asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;產學訓表單列印&gt;&gt;訓練計畫場地資料表</asp:Label>--%>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;表單列印&gt;&gt;訓練計畫場地資料表</asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="bluecol" width="20%">訓練機構</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" runat="server">
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">班級狀態</td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="Radio1" runat="server" CssClass="font" AutoPostBack="True" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="0">未轉班</asp:ListItem>
                                    <asp:ListItem Value="1">已轉班</asp:ListItem>
                                    <asp:ListItem Value="2">變更待審</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="OpenTR" runat="server">
                            <td class="bluecol">開訓期間</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="STDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span> ~
                                <asp:TextBox ID="STDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr id="CloseTR" runat="server">
                            <td class="bluecol">結訓期間
                            </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="FTDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>～
                                <asp:TextBox ID="FTDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                    </table>
                    <div align="center" class="whitecol">
                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                        <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                    </div>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></div>
                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#f5f5f5" />
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <HeaderTemplate>
                                                <input onclick="SelectAll(this.checked);" type="checkbox">
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <input id="chkSeqNo" type="checkbox" runat="server">
                                                <input id="HidRadio1" type="hidden" runat="server">
                                                <input id="hidPlanID" type="hidden" runat="server">
                                                <input id="hidComIDNO" type="hidden" runat="server">
                                                <input id="hidSeqNo" type="hidden" runat="server">
                                                <input id="hidSubSeqNO" type="hidden" runat="server">
                                                <input id="hidCDateValue" type="hidden" runat="server">
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構">
                                            <HeaderStyle Width="30%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassCName" HeaderText="班級名稱">
                                            <HeaderStyle Width="30%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDate" HeaderText="開訓日期" DataFormatString="{0:d}">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="FTDate" HeaderText="結訓日期" DataFormatString="{0:d}">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn>
                                            <HeaderTemplate>申請變更時間</HeaderTemplate>
                                            <HeaderStyle Width="15%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="labModifydate" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <%--<asp:BoundColumn HeaderText="申請變更時間"></asp:BoundColumn>--%>
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
                                <%--<input id="print" type="button" value="列印" runat="server" class="asp_button_S">--%>
                                <asp:Button ID="BtnPrint" runat="server" Text="列印" CssClass="asp_Export_M" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <input id="ROC_Years" type="hidden" runat="server" />
        <input id="Years2" type="hidden" name="Years2" runat="server">
    </form>
</body>
</html>
