<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CR_01_006.aspx.vb" Inherits="WDAIIP.CR_01_006" %>

<%--<!DOCTYPE html>--%>
<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">

<%--<html xmlns="http://www.w3.org/1999/xhtml">--%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>年度主責分署設定</title>
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
</head>
<body>
    <form id="form1" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;課程審查&gt;&gt;一階審查&gt;&gt;年度主責分署設定</asp:Label>
                </td>
            </tr>
        </table>
        <asp:Panel ID="panelSch" runat="server">
            <table width="100%" cellpadding="1" cellspacing="1" class="table_sch">
                <tr>
                    <td class="bluecol_need" width="18%">年度</td>
                    <td class="whitecol" width="82%">
                        <asp:DropDownList ID="ddlYEARS_SCH" runat="server"></asp:DropDownList></td>
                </tr>
                <tr>
                    <td class="bluecol">申請階段</td>
                    <td class="whitecol">
                        <asp:DropDownList ID="ddlAPPSTAGE_SCH" runat="server"></asp:DropDownList></td>
                </tr>
                <tr>
                    <td class="whitecol" align="center" colspan="2">
                        <asp:Button ID="BtnSearch" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="BtnAddNew1" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>
                    </td>
                </tr>
            </table>
            <div align="center">
                <asp:Label ID="msg1" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
            </div>
            <table id="tbDataGrid1" runat="server" cellspacing="1" cellpadding="1" width="100%" border="0">
                <tr>
                    <td align="center">
                        <asp:DataGrid ID="DataGrid1" runat="server" AllowPaging="false" Width="100%" AutoGenerateColumns="False" CellPadding="8">
                            <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                            <Columns>
                                <asp:BoundColumn DataField="YEARS_ROC" HeaderText="年度" ItemStyle-HorizontalAlign="Center"><HeaderStyle HorizontalAlign="center" Width="20%"></HeaderStyle></asp:BoundColumn>
                                <asp:BoundColumn DataField="APPSTAGE_N" HeaderText="申請階段" ItemStyle-HorizontalAlign="Center"><HeaderStyle HorizontalAlign="center" Width="20%"></HeaderStyle></asp:BoundColumn>
                                <asp:TemplateColumn HeaderText="主責分署" ItemStyle-HorizontalAlign="Center">
                                    <HeaderStyle HorizontalAlign="center" Width="60%"></HeaderStyle>
                                    <ItemStyle CssClass="whitecol" />
                                    <ItemTemplate>
                                        <asp:HiddenField ID="Hid_YEARS" runat="server" />
                                        <asp:HiddenField ID="Hid_APPSTAGE" runat="server" />
                                        <asp:HiddenField ID="Hid_DISTID" runat="server" />
                                        <asp:RadioButtonList ID="rbl_DISTNM" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow"></asp:RadioButtonList>
                                        <%--<asp:TextBox ID="txtCLASSQUOTA" runat="server" MaxLength="10"></asp:TextBox>--%>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                            </Columns>
                            <PagerStyle Visible="False"></PagerStyle>
                        </asp:DataGrid>
                        <%--<uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>--%>
                    </td>
                </tr>
                <%--<tr>
                    <td class="whitecol" align="center">
                        <asp:Button ID="BtnSaveData2" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                    </td>
                </tr>--%>
            </table>
        </asp:Panel>
        <asp:Panel ID="PanelEdit1" runat="server" Visible="false">
            <table id="tbPanelEdit1" runat="server" class="table_sch" border="0" cellspacing="1" cellpadding="1" width="100%">
                <tr>
                    <td class="bluecol_need" width="18%">年度</td>
                    <td class="whitecol" width="82%">
                        <asp:DropDownList ID="ddlYEARS" runat="server"></asp:DropDownList></td>
                </tr>
                <tr>
                    <td class="bluecol_need">申請階段</td>
                    <td class="whitecol">
                        <asp:DropDownList ID="ddlAPPSTAGE" runat="server"></asp:DropDownList></td>
                </tr>
                <tr>
                    <td class="bluecol_need">主責分署</td>
                    <td class="whitecol">
                        <asp:RadioButtonList ID="rblDISTMAIN" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow"></asp:RadioButtonList></td>
                </tr>
                <tr>
                    <td class="whitecol" align="center" colspan="2">
                        <asp:Button ID="BtnSaveData1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="BtnBack1" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                    </td>
                </tr>
            </table>
        </asp:Panel>
    </form>
</body>
</html>
