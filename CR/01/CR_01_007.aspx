<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CR_01_007.aspx.vb" Inherits="WDAIIP.CR_01_007" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>審查幕僚意見開關機制(產業人才專用)</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker2.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td class="font">
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;課程審查&gt;&gt;一階審查&gt;&gt;審查幕僚意見開關機制</asp:Label>
                </td>
            </tr>
        </table>
        <asp:Panel ID="panel_sch" runat="server" Width="100%">
            <table id="tb_table_sch" runat="server" class="table_sch" cellspacing="1" cellpadding="1" width="100%">
                <tr>
                    <td class="bluecol_need" runat="server" width="20%">年度</td>
                    <td class="whitecol" width="80%">
                        <asp:DropDownList ID="ddlYearlist_sch" runat="server"></asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">申請階段</td>
                    <td class="whitecol">
                        <asp:DropDownList ID="ddlAppStage_sch" runat="server"></asp:DropDownList></td>
                </tr>
                <tr>
                    <td class="whitecol" align="center" colspan="2">
                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>&nbsp;
                        <asp:Button ID="bt_search" runat="server" Text="查詢" CssClass="asp_button_M" AuthType="QRY"></asp:Button>
                        <asp:Button ID="bt_addnew" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_M" AuthType="ADD"></asp:Button>
                    </td>
                </tr>
                <tr>
                    <td align="center" colspan="2">
                        <asp:Label ID="msg1" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="center" colspan="2">
                        <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td>
                                    <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Width="100%" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                                        <Columns>
                                            <asp:BoundColumn HeaderText="序號">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" Width="6%"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="YEARS_ROC" HeaderText="年度">
                                                <ItemStyle HorizontalAlign="Center" Width="12%"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="APPSTAGE_N" HeaderText="申請階段">
                                                <ItemStyle HorizontalAlign="Center" Width="12%"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="SFOPENDATE_ROC" HeaderText="受理期間設定">
                                                <ItemStyle HorizontalAlign="Center" Width="20%"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="功能">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" Width="15%"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:LinkButton ID="BTNUPDATE" runat="server" Text="修改" CommandName="BTNUPDATE" CssClass="linkbutton"></asp:LinkButton>
                                                    <asp:LinkButton ID="BTNDELETE" runat="server" Text="刪除" CommandName="BTNDELETE" CssClass="linkbutton"></asp:LinkButton>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                        <PagerStyle Visible="False"></PagerStyle>
                                    </asp:DataGrid>
                                    <div align="center">
                                        <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                    </div>
                                </td>
                            </tr>
                        </table>

                    </td>
                </tr>
            </table>
        </asp:Panel>

        <asp:Panel ID="panel_EDIT1" runat="server" Width="100%" Visible="False">
            <table id="tb_EDIT1" runat="server" class="table_nw" cellspacing="1" cellpadding="1" width="100%" border="0">
                <tr>
                    <td class="bluecol_need" runat="server" width="20%">年度</td>
                    <td class="whitecol" width="80%">
                        <asp:DropDownList ID="ddlYEARS" runat="server"></asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need">申請階段</td>
                    <td class="whitecol"><asp:DropDownList ID="ddlAppStage" runat="server"></asp:DropDownList></td>
                </tr>
                <tr>
                    <td class="bluecol_need">「審查幕僚意見」開放增修<br />起始日期</td>
                    <td class="whitecol">
                        <asp:TextBox ID="TB_SOPENDATE" runat="server" onfocus="this.blur()" Columns="20"></asp:TextBox>
                        <img id="date1" style="cursor: pointer" onclick="javascript:show_calendar('TB_SOPENDATE','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30">
                        <asp:DropDownList ID="TB_SOPENDATE_HR" runat="server"></asp:DropDownList>時：
                    <asp:DropDownList ID="TB_SOPENDATE_MM" runat="server"></asp:DropDownList>分
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need">「審查幕僚意見」開放增修<br />結束日期</td>
                    <td class="whitecol">
                        <asp:TextBox ID="TB_FOPENDATE" runat="server" onfocus="this.blur()" Columns="20"></asp:TextBox>
                        <img id="date2" style="cursor: pointer" onclick="javascript:show_calendar('TB_FOPENDATE','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30">
                        <asp:DropDownList ID="TB_FOPENDATE_HR" runat="server"></asp:DropDownList>時：
                    <asp:DropDownList ID="TB_FOPENDATE_MM" runat="server"></asp:DropDownList>分
                    </td>
                </tr>
                <tr>
                    <td align="center" colspan="4" class="whitecol">
                        <asp:Button ID="BtnSAVEDATA1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="BtnBack1" runat="server" Text="回查詢頁面" CssClass="asp_button_M"></asp:Button>
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:HiddenField ID="Hid_PSSID" runat="server" />
    </form>
</body>
</html>
