<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_14_010.aspx.vb" Inherits="WDAIIP.SD_14_010" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>計畫變更申請書</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function ClearData() {
            document.getElementById('TMID1').value = '';
            document.getElementById('OCID1').value = '';
            document.getElementById('TMIDValue1').value = '';
            document.getElementById('OCIDValue1').value = '';
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
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;表單列印&gt;&gt;訓練計畫變更表</asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch">
                        <tr>
                            <td class="bluecol" width="20%">訓練機構 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" runat="server">
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr id="ClassTR" runat="server">
                            <td class="bluecol">職類/班別 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input type="button" value="..." class="button_b_Mini" onclick="choose_class();" />
                                <input type="button" value="清除" class="asp_button_M" onclick="ClearData();" />
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                <span id="HistoryList" style="position: absolute; display: none; left: 28%">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" style="width: 20%">班級名稱 </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="ClassName" runat="server" Width="60%"></asp:TextBox></td>
                            <td class="bluecol" style="width: 20%">期別 </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="CyclType" runat="server" Columns="3" Width="30%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">變更項目 </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="ChgItem" runat="server"></asp:DropDownList>
                                <input id="chgState" type="hidden" value="0" name="chgState" runat="server" />
                            </td>

                            <td class="bluecol">審核狀態</td>
                            <td class="whitecol">
                                <asp:DropDownList ID="CheckMode" runat="server">
                                    <asp:ListItem Value="">==請選擇==</asp:ListItem>
                                    <asp:ListItem Value="1">審核不通過</asp:ListItem>
                                    <asp:ListItem Value="2">審核中</asp:ListItem>
                                    <asp:ListItem Value="0">審核完成</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr id="TRPlanPoint28" runat="server">
                            <td class="bluecol">計畫 </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="PlanPoint" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
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
                        <asp:Button ID="Button4" runat="server" Text="空白變更申請書" CssClass="asp_Export_M"></asp:Button>
                    </div>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                    </div>
                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <div id="divTip" runat="server" visible="false" style="text-align: right;"><font color="#0066FF">※只有「審核中」的資料才提供列印</font></div>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#f5f5f5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構" HeaderStyle-Width="25%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassName" HeaderText="班別名稱" HeaderStyle-Width="25%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="CDate" HeaderText="申請日期" HeaderStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="變更項目" HeaderStyle-Width="15%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="ReviseStatus" HeaderText="審核狀態" HeaderStyle-Width="10%" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="列印" HeaderStyle-Width="15%">
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <input id="Button3" type="button" value="計畫變更表" runat="server" class="asp_Export_M">
                                                <input id="bt_print" type="button" value="變更後課程表" runat="server" class="asp_Export_M">
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
                </td>
            </tr>
        </table>
        <input id="ROC_Years" type="hidden" runat="server" />
        <input id="KindValue" type="hidden" runat="server" />
        <input id="Kind2Value" type="hidden" runat="server" />
    </form>
</body>
</html>
