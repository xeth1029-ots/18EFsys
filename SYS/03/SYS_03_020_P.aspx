<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_03_020_P.aspx.vb" Inherits="WDAIIP.SYS_03_020_P" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>功能頁面設定-關聯報表</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;功能權限管理&gt;&gt;功能頁面設定</asp:Label>
                </td>
            </tr>
        </table>

        <table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr>
                <td>

                    <table id="tbSch" runat="server" width="100%" cellspacing="1" cellpadding="1">
                        <tr>
                            <td>
                                <table width="100%" class="table_nw" cellpadding="1" cellspacing="1">
                                    <tr>
                                        <td class="bluecol" width="20%">功能ID</td>
                                        <td class="whitecol" width="80%">
                                            <asp:Label ID="ssLabFUNID" runat="server"></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">功能名稱</td>
                                        <td class="whitecol" width="80%">
                                            <asp:Label ID="ssLabFUNNAME" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">功能位置</td>
                                        <td class="whitecol" width="80%">
                                            <asp:Label ID="ssLabSPAGE" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">報表名稱代號</td>
                                        <td class="whitecol" width="80%">
                                            <asp:TextBox ID="sRPTNAME" runat="server" MaxLength="100" Width="30%"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td align="center" colspan="2" class="whitecol">
                                            <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue" DESIGNTIMEDRAGDROP="30">顯示列數</asp:Label>
                                            <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                            <asp:Button ID="btnSearch1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                            &nbsp;<asp:Button ID="btnAdd1" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                                            &nbsp;<asp:Button ID="btnBack2" runat="server" Text="回上頁" CssClass="asp_button_M"></asp:Button>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <div align="center">
                                    <asp:Label ID="labmsg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                                </div>
                                <table id="tbList" runat="server" cellspacing="0" bordercolordark="#ffffff" cellpadding="0" width="100%" align="left" bordercolorlight="#666666" border="0">
                                    <tr>
                                        <td>
                                            <div align="center">
                                                <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Width="100%" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <Columns>
                                                        <asp:BoundColumn HeaderText="序號">
                                                            <HeaderStyle HorizontalAlign="Center" Width="6%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>

                                                        <asp:BoundColumn DataField="RPTNAME" HeaderText="報表名稱代號">
                                                            <HeaderStyle HorizontalAlign="Center" Width="16%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>

                                                        <asp:TemplateColumn HeaderText="功能" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="8%">
                                                            <ItemStyle HorizontalAlign="Center" Font-Size="Small" />
                                                            <ItemTemplate>
                                                                <asp:LinkButton ID="lbtUpdate" Text="修改" runat="server" CssClass="linkbutton" CommandName="UPD"></asp:LinkButton>&nbsp;&nbsp;
                                                                <%--<asp:LinkButton ID="lbtDelete" Text="刪除" runat="server" CssClass="linkbutton" CommandName="DEL"></asp:LinkButton>&nbsp;&nbsp;--%>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                    </Columns>
                                                    <PagerStyle Visible="False"></PagerStyle>
                                                </asp:DataGrid>
                                                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                            </div>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table id="tbEdit" runat="server" class="table_nw" width="100%" border="0" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" width="20%">功能ID</td>
                            <td class="whitecol" width="80%">
                                <asp:Label ID="LabFUNID" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">功能名稱</td>
                            <td class="whitecol" width="80%">
                                <asp:Label ID="LabFUNNAME" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">功能位置</td>
                            <td class="whitecol" width="80%">
                                <asp:Label ID="LabSPAGE" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">報表名稱代號</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="txRPTNAME" runat="server" MaxLength="100" Width="30%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td colspan="2" class="whitecol" align="center">
                                <asp:Button ID="btnSave1" Text="儲存" runat="server" CssClass="asp_Export_M"></asp:Button>&nbsp;
                                <asp:Button ID="btnBack1" runat="server" Text="回上頁" CssClass="asp_button_M"></asp:Button>&nbsp;
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="Hid_FRSEQ" runat="server" />
        <asp:HiddenField ID="Hid_FUNID" runat="server" />
    </form>
</body>
</html>
