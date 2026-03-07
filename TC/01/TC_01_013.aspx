<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_013.aspx.vb" Inherits="WDAIIP.TC_01_013" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>計畫場地設定</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
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
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;基本資料設定&gt;&gt;計畫場地設定</asp:Label>
                </td>
            </tr>
        </table>
        <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td class="bluecol" width="20%">訓練機構</td>
                <td class="whitecol">
                    <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                    <input type="button" value="..." id="Org" name="Button5" runat="server" class="button_b_Mini" />
                    <input id="RIDValue" type="hidden" runat="server" name="RIDValue" />&nbsp;
                    <input id="orgid_value" type="hidden" runat="server" name="orgid_value" />&nbsp;<%--<br>--%>
                    <span id="HistoryList2" style="position: absolute; display: none">
                        <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                    </span>
                </td>
            </tr>
            <tr>
                <td class="bluecol">場地代號</td>
                <td class="whitecol">
                    <asp:TextBox ID="PlaceNo" runat="server" Width="20%" MaxLength="30"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="bluecol">場地名稱</td>
                <td class="whitecol">
                    <asp:TextBox ID="Place" runat="server" Width="25%" MaxLength="200"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="bluecol">地址關鍵字</td>
                <td class="whitecol">
                    <asp:TextBox ID="txtAddress" runat="server" Width="60%" MaxLength="200"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="bluecol">含已停用資料</td>
                <td class="whitecol">
                    <asp:CheckBox ID="chkboxDelData" runat="server" Text="含已停用資料" /></td>
            </tr>
        </table>
        <table width="100%">
            <tr>
                <td class="whitecol">
                    <div align="center">
                        <asp:Label ID="labPageSize" runat="server" DESIGNTIMEDRAGDROP="30" ForeColor="SlateBlue">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                        <asp:Button ID="bt_search" Text="查詢" runat="server" CssClass="asp_button_M"></asp:Button>&nbsp;
                        <asp:Button ID="bt_add" Text="新增" runat="server" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                    </div>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
                    </div>
                </td>
            </tr>
        </table>
        <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td align="center">
                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                        <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn HeaderText="序號">
                                <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="PlaceID" HeaderText="埸地編號">
                                <HeaderStyle HorizontalAlign="Center" Width="11%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="PlaceName" HeaderText="場地名稱">
                                <HeaderStyle HorizontalAlign="Center" Width="11%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Left"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Address" HeaderText="場地地址">
                                <HeaderStyle HorizontalAlign="Center" Width="14%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Left"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="FactMode" HeaderText="場地類型">
                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Left"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="PlacePic1" HeaderText="上傳圖檔1">
                                <HeaderStyle HorizontalAlign="Center" Width="16%"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="PlacePic2" HeaderText="上傳圖檔2">
                                <HeaderStyle HorizontalAlign="Center" Width="16%"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="功能">
                                <HeaderStyle HorizontalAlign="Center" Width="17%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                                <ItemTemplate>
                                    <asp:LinkButton ID="lbtView" runat="server" Text="檢視" CommandName="view" CssClass="linkbutton"></asp:LinkButton>
                                    <asp:LinkButton ID="lbtEdit" runat="server" Text="修改" CommandName="edit" CssClass="linkbutton"></asp:LinkButton>
                                    <asp:LinkButton ID="lbtReturn" runat="server" Text="啟用" CommandName="return" CssClass="linkbutton"></asp:LinkButton>
                                    <asp:LinkButton ID="lbtStop" runat="server" Text="停用" CommandName="stop" CssClass="linkbutton"></asp:LinkButton>
                                    <asp:LinkButton ID="lbtDel" runat="server" Text="刪除" CommandName="del" CssClass="linkbutton"></asp:LinkButton>
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
        <%--<asp:HiddenField ID="Hid_ORGID" runat="server" />--%>
    </form>
</body>
</html>
