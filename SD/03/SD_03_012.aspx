<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_03_012.aspx.vb" Inherits="WDAIIP.SD_03_012" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <%--報名暨報到名單確認--%>
    <title>參訓學員名單確認</title>
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
        function GETvalue() { document.getElementById('Button3').click(); }

        function SetOneOCID() { document.getElementById('Button4').click(); }

        function choose_class() {
            var Button4 = document.getElementById('Button4');
            var OCID1 = document.getElementById('OCID1');
            if (OCID1.value == '') { Button4.click(); }
            var RIDValue = document.getElementById('RIDValue');
            openClass('../02/SD_02_ch.aspx?RID=' + RIDValue.value);
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
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;學員資料管理&gt;&gt;參訓學員名單確認</asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch" id="table_sch" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td class="bluecol_need" style="width: 20%">訓練機構 </td>
                            <td class="whitecol" style="width: 80%">
                                <asp:TextBox ID="center" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox>
                                <input id="Button8" type="button" value="..." runat="server" class="asp_button_Mini">
                                <asp:Button ID="Button4" Style="display: none" runat="server"></asp:Button>
                                <asp:Button ID="Button3" Style="display: none" runat="server"></asp:Button>
                                <%--
                                <asp:Button ID="Button4" runat="server"></asp:Button>
                                <asp:Button ID="Button3" runat="server"></asp:Button>
                                --%>
                                <input id="RIDValue" type="hidden" runat="server" />
                                <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td width="20%" class="bluecol">職類/班別</td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="asp_button_Mini" />
                                <input id="OCIDValue1" type="hidden" runat="server" />
                                <input id="TMIDValue1" type="hidden" runat="server" />
                                <span id="HistoryList" style="position: absolute; display: none; left: 28%">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">開訓期間 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="STDATE1" runat="server" Columns="10" MaxLength="10" Width="20%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('STDATE1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                ~
                                <asp:TextBox ID="STDATE2" runat="server" Columns="10" MaxLength="10" Width="20%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('STDATE2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">結訓期間 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="FTDATE1" runat="server" Columns="10" MaxLength="10" Width="20%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('FTDATE1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                ~
                                <asp:TextBox ID="FTDATE2" runat="server" Columns="10" MaxLength="10" Width="20%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('FTDATE2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">確認日期 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="CONFIRDATE1" runat="server" Columns="10" MaxLength="10" Width="20%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('CONFIRDATE1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                ~
                                <asp:TextBox ID="CONFIRDATE2" runat="server" Columns="10" MaxLength="10" Width="20%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('CONFIRDATE2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            </td>
                        </tr>
                        <tr id="TRPlanPoint28" runat="server">
                            <td class="bluecol">計畫 </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="PlanPoint" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="0" Selected="True">不區分</asp:ListItem>
                                    <asp:ListItem Value="1">產業人才投資計畫</asp:ListItem>
                                    <asp:ListItem Value="2">提升勞工自主學習計畫</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="tr_ddl_INQUIRY_S" runat="server">
                            <td class="bluecol_need">查詢原因</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="ddl_INQUIRY_Sch" runat="server"></asp:DropDownList>
                            </td>
                        </tr>
                    </table>
                    <table width="100%" class="font">
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                <asp:Button ID="BtnSearch1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="BtnInsert1" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Label ID="Msg1" runat="server" ForeColor="Red"></asp:Label></td>
                        </tr>
                    </table>
                    <table id="DataGridTable1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server" class="font">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#f5f5f5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構">
                                            <HeaderStyle Width="15%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassCName2" HeaderText="班級名稱">
                                            <HeaderStyle Width="15%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDate" HeaderText="開訓日期">
                                            <HeaderStyle Width="15%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="FTDate" HeaderText="結訓日期">
                                            <HeaderStyle Width="15%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="15%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <HeaderTemplate>確認日期</HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:Label ID="LabCONFIRDATE" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn>
                                            <HeaderTemplate>確認者</HeaderTemplate>
                                            <HeaderStyle Width="15%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:Label ID="LabCONFIRNAME" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn>
                                            <HeaderTemplate>報名名單</HeaderTemplate>
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" Font-Size="Small" />
                                            <ItemTemplate>
                                                <asp:HiddenField ID="Hid_CFGUID" runat="server" />
                                                <asp:HiddenField ID="Hid_OCID" runat="server" />
                                                <asp:HiddenField ID="Hid_CFSEQNO" runat="server" />
                                                <asp:LinkButton ID="BtnEDIT1" runat="server" Text="檢視" CommandName="BtnEDIT1" CssClass="linkbutton"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn>
                                            <HeaderTemplate>報到名單</HeaderTemplate>
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" Font-Size="Small" />
                                            <ItemTemplate>
                                                <asp:LinkButton ID="BtnEDIT2" runat="server" Text="檢視" CommandName="BtnEDIT2" CssClass="linkbutton"></asp:LinkButton>
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
    </form>
</body>
</html>
