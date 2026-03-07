<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_06_012.aspx.vb" Inherits="WDAIIP.SYS_06_012" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD html 4.0 Transitional//EN">
<html>
<head>
    <title>資料交換平台</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1" />
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1" />
    <meta name="vs_defaultClientScript" content="JavaScript" />
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <%--<script type="text/javascript" src="../../js/date-picker2.js"></script>--%>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="../../js/TIMS.js"></script>
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
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;資料交換平台</asp:Label>
                </td>
            </tr>
        </table>
        <div id="divSch1" runat="server">
            <table class="table_sch" cellpadding="1" cellspacing="1" width="100%" border="0">
                <tr>
                    <td class="bluecol" style="width: 20%">掛載系統名稱
                    </td>
                    <td class="whitecol">
                        <asp:TextBox ID="sMSNAME" runat="server" MaxLength="100" Width="70%"></asp:TextBox>
                    </td>
                    <td class="bluecol" style="width: 20%">掛載系統IP
                    </td>
                    <td class="whitecol">
                        <asp:TextBox ID="sMSIP4" runat="server" MaxLength="100" Width="70%"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol" style="width: 20%">掛載系統PORT
                    </td>
                    <td class="whitecol">
                        <asp:TextBox ID="sMSPORT" runat="server" MaxLength="100" Width="70%"></asp:TextBox>
                    </td>
                    <td class="bluecol" style="width: 20%">掛載模組
                    </td>
                    <td class="whitecol">
                        <asp:TextBox ID="sMSMODULE" runat="server" MaxLength="100" Width="70%"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol" style="width: 20%">通訊協定
                    </td>
                    <td class="whitecol" colspan="3">
                        <asp:DropDownList ID="sddlMSPROTOCOL" runat="server" AppendDataBoundItems="True">
                        </asp:DropDownList>
                    </td>
                </tr>

                <tr>
                    <td class="whitecol" colspan="4" align="center">
                        <asp:Button ID="btnQuery" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>
                        <asp:Button ID="BtnReset2" runat="server" Text="重設查詢" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="btnAddNew1" runat="server" Text="新增" CssClass="asp_button_S"></asp:Button>
                        <%--<asp:Button ID="BtnBack2" runat="server" Text="回上頁" CssClass="asp_button_M"></asp:Button>--%>
                        <%--<asp:Button ID="btnImp1" runat="server" Text="匯入總場次" CssClass="asp_button_S"></asp:Button>--%>
                        <%--<asp:Button ID="btnExp1" runat="server" Text="匯出場次代碼" CssClass="asp_Export_M"></asp:Button>--%>
                    </td>
                </tr>
            </table>
            <table id="table_sch_show" cellspacing="1" cellpadding="1" width="100%" border="0">
                <tr>
                    <td align="center">
                        <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td align="center">
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AllowSorting="True" PagerStyle-HorizontalAlign="Left"
                                        PagerStyle-Mode="NumericPages" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                                        <Columns>
                                            <asp:BoundColumn HeaderText="編號" HeaderStyle-Width="5%">
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>

                                            <asp:BoundColumn DataField="MSNAME" HeaderText="掛載系統名稱" HeaderStyle-Width="8%">
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="MSIP4" HeaderText="掛載系統IP" HeaderStyle-Width="8%">
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="MSPORT" HeaderText="掛載系統PORT" HeaderStyle-Width="8%">
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="MSMODULE" HeaderText="掛載模組" HeaderStyle-Width="8%">
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <%--<asp:BoundColumn DataField="MSPROTOCOL" HeaderText="通訊協定" HeaderStyle-Width="8%">
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>--%>
                                            <asp:TemplateColumn HeaderText="通訊協定">
                                                <ItemStyle HorizontalAlign="Center" Width="8%" />
                                                <ItemTemplate>
                                                    <asp:Label ID="labMSPROTOCOL" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="功能" HeaderStyle-Width="12%">
                                                <ItemStyle HorizontalAlign="Center" Font-Size="Small" />
                                                <ItemTemplate>
                                                    <asp:LinkButton ID="lbtEdit" runat="server" Text="編輯" CommandName="btnEdit" CssClass="linkbutton"></asp:LinkButton>
                                                    <asp:LinkButton ID="lbtDel1" runat="server" Text="刪除" CommandName="btnDel" CssClass="linkbutton"></asp:LinkButton>
                                                    <asp:LinkButton ID="lbtTest1" runat="server" Text="測試" CommandName="btnTest" CssClass="linkbutton"></asp:LinkButton>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>

                                        </Columns>
                                        <PagerStyle Visible="false"></PagerStyle>
                                    </asp:DataGrid>
                                </td>
                            </tr>

                        </table>
                    </td>
                </tr>
                <tr>
                    <td align="center">
                        <asp:Label ID="msg1" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                    </td>
                </tr>
            </table>
        </div>

        <div id="divEdt1" runat="server">
            <table class="table_sch" cellpadding="1" cellspacing="1" width="100%" border="0">
                <tr>
                    <td class="bluecol_need" style="width: 20%">掛載系統名稱
                    </td>
                    <td class="whitecol">
                        <asp:TextBox ID="tMSNAME" runat="server" MaxLength="100" Width="70%"></asp:TextBox>
                    </td>
                    <td class="bluecol_need" style="width: 20%">掛載系統IP
                    </td>
                    <td class="whitecol">
                        <asp:TextBox ID="tMSIP4" runat="server" MaxLength="100" Width="70%"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need" style="width: 20%">掛載系統PORT
                    </td>
                    <td class="whitecol">
                        <asp:TextBox ID="tMSPORT" runat="server" MaxLength="100" Width="70%"></asp:TextBox>
                    </td>
                    <td class="bluecol_need" style="width: 20%">掛載模組
                    </td>
                    <td class="whitecol">
                        <asp:TextBox ID="tMSMODULE" runat="server" MaxLength="100" Width="70%"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need" style="width: 20%">通訊協定
                    </td>
                    <td class="whitecol" colspan="3">
                        <asp:DropDownList ID="ddlMSPROTOCOL" runat="server" AppendDataBoundItems="True">
                        </asp:DropDownList>
                    </td>
                </tr>

            </table>
            <table width="100%">
                <tr>
                    <td class="whitecol">
                        <div align="center">
                            <%--<asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                            <asp:TextBox ID="TxtPageSize" runat="server" MaxLength="2" Width="55px">10</asp:TextBox>--%>
                            <asp:Button ID="BtnBack1" runat="server" Text="回上頁" CssClass="asp_button_M"></asp:Button>
                            <asp:Button ID="BtnSaveData1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                        </div>
                    </td>
                </tr>
            </table>
        </div>

        <asp:HiddenField ID="HID_WSENO" runat="server" />

        <%-- <asp:HiddenField ID="HID_WSENO" runat="server" />
        <asp:HiddenField ID="Hid_DISTID" runat="server" />
        <asp:HiddenField ID="Hid_TPLANID" runat="server" />
        <asp:HiddenField ID="Hid_HALFYEAR" runat="server" />
        <asp:HiddenField ID="Hid_PTYID" runat="server" />--%>
    </form>
</body>
</html>
