<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_14_027.aspx.vb" Inherits="WDAIIP.SD_14_027" %>

<%--<!DOCTYPE html>--%>
<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>學員簽到(退)及教學日誌</title>
    <link rel="stylesheet" type="text/css" href="../../css/style.css" />
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
        function ClearData() {
            document.getElementById('TMID1').value = '';
            document.getElementById('OCID1').value = '';
            document.getElementById('TMIDValue1').value = '';
            document.getElementById('OCIDValue1').value = '';
        }

        function choose_class() {
            ClearData();
            var RIDValue = document.getElementById('RIDValue');
            openClass('../02/SD_02_ch.aspx?&RID=' + RIDValue.value);
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;表單列印&gt;&gt;學員簽到(退)及教學日誌</asp:Label>
                </td>
            </tr>
        </table>
        <table class="table_sch">
            <tr>
                <td class="bluecol" style="width: 20%">訓練機構 </td>
                <td class="whitecol" colspan="3">
                    <asp:TextBox ID="center" runat="server" Width="60%"></asp:TextBox>
                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                    <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                    <span id="HistoryList2" style="position: absolute; display: none">
                        <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                    </span>
                </td>
            </tr>
            <tr id="ClassTR" runat="server">
                <td class="bluecol_need">職類/班別 </td>
                <td class="whitecol" colspan="3">
                    <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                    <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                    <input onclick="choose_class()" type="button" value="..." class="button_b_Mini">
                    <input id="Button4" type="button" value="清除" name="Button4" runat="server" class="asp_button_M">
                    <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                    <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                    <span id="HistoryList" style="position: absolute; display: none; left: 270px">
                        <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                    </span>
                </td>
            </tr>
            <%-- <tr id="TRPlanPoint28" runat="server">
                <td class="bluecol">計畫 </td>
                <td class="whitecol" colspan="3">
                    <asp:RadioButtonList ID="PlanPoint" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                        <asp:ListItem Value="1" Selected="True">產業人才投資計畫</asp:ListItem>
                        <asp:ListItem Value="2">提升勞工自主學習計畫</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>--%>
            <tr>
                <td class="bluecol">學員名單列印範圍 </td>
                <td class="whitecol" colspan="3">
                    <asp:RadioButtonList ID="rbl_StudlistRange" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                        <asp:ListItem Value="1">錄訓作業正取名單</asp:ListItem>
                        <asp:ListItem Value="2" Selected="True">完成報到學員名單</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">上課日期 </td>
                <td class="whitecol" colspan="3">
                    <asp:TextBox ID="STRAINDATE" runat="server" Columns="10" Width="15%"></asp:TextBox>
                    <span runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('STRAINDATE','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                </td>
            </tr>
        </table>
        <div align="center" class="whitecol">
            <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
            <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
            <asp:Button ID="btnSearch1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
        </div>
        <div align="center">
            <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
        </div>
        <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td>
                    <%--顯示欄位：訓練機構、班級名稱、開訓日期、結訓日期、功能--%>
                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                        <AlternatingItemStyle BackColor="#f5f5f5" />
                        <HeaderStyle CssClass="head_navy" HorizontalAlign="Center" />
                        <Columns>
                            <asp:BoundColumn HeaderText="序號" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                            <%--<asp:BoundColumn DataField="OrgName" HeaderText="訓練機構"></asp:BoundColumn>
                            <asp:BoundColumn DataField="ClassCName" HeaderText="班級名稱"></asp:BoundColumn>
                            <asp:BoundColumn DataField="STDate" HeaderText="開訓日期"></asp:BoundColumn>
                            <asp:BoundColumn DataField="FTDate" HeaderText="結訓日期"></asp:BoundColumn>--%>
                            <asp:BoundColumn DataField="STRAINDATE" HeaderText="上課日期"></asp:BoundColumn>
                            <asp:BoundColumn DataField="TPX" HeaderText="授課時段"></asp:BoundColumn>
                            <asp:BoundColumn DataField="PONCLASS1" HeaderText="授課時間"></asp:BoundColumn>
                            <asp:BoundColumn DataField="TPERIOD28_N" HeaderText="課程進度/內容"></asp:BoundColumn>
                            <asp:TemplateColumn>
                                <HeaderTemplate>功能</HeaderTemplate>
                                <ItemStyle HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <asp:Button ID="btnPrint1" runat="server" Text="列印" CssClass="asp_Export_M" CommandName="Print1"></asp:Button>
                                    <asp:HiddenField ID="OCID" runat="server" />
                                    <asp:HiddenField ID="PCSValue" runat="server" />
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
        <asp:HiddenField ID="Hid_OCID1" runat="server" />
        <asp:HiddenField ID="hidPCSValue" runat="server" />
        <asp:HiddenField ID="hidSTRAINDATE" runat="server" />
        <%-- <asp:HiddenField ID="hidPCSValue" runat="server" />
        <asp:HiddenField ID="hidOCIDValue" runat="server" />--%>
    </form>
</body>
</html>
