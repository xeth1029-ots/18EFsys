<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_014.aspx.vb" Inherits="WDAIIP.TC_01_014" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>開班計畫表資料維護作業(產業人才專用)</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function Search() {
            if (document.form1.yearlist.value == '') {
                alert('請輸入年度');
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td class="font">
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;開班計畫表資料維護作業</asp:Label>
                </td>
            </tr>
        </table>
        <table class="table_nw" width="100%">
            <tr id="tr_center" runat="server">
                <td id="Td3" runat="server" class="bluecol_need" width="20%">訓練機構 </td>
                <td colspan="3" class="whitecol">
                    <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                    <input id="Org" type="button" value="..." name="Org" runat="server" class="button_b_Mini">
                    <span id="HistoryList2" style="position: absolute; display: none">
                        <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                    </span>
                </td>
            </tr>
            <tr>
                <td id="Td1" runat="server" class="bluecol_need">年度 </td>
                <td colspan="3" class="whitecol">
                    <asp:DropDownList ID="yearlist" runat="server"></asp:DropDownList></td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">班級名稱 </td>
                <td class="whitecol" width="30%">
                    <asp:TextBox ID="ClassName" runat="server" Columns="30" MaxLength="30" Width="80%"></asp:TextBox></td>
                <td class="bluecol" width="20%">期別 </td>
                <td class="whitecol" width="30%">
                    <asp:TextBox ID="CyclType" runat="server" Columns="5" MaxLength="3" Width="30%"></asp:TextBox></td>
            </tr>
            <tr>
                <td id="Td2" runat="server" class="bluecol_need">資料類型 </td>
                <td colspan="3" class="whitecol">
                    <asp:RadioButtonList ID="IsApprPaper" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                        <asp:ListItem Value="Y" Selected="True">正式</asp:ListItem>
                        <asp:ListItem Value="N">草稿</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td colspan="4" class="whitecol" width="100%">
                    <div align="center">
                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue" DESIGNTIMEDRAGDROP="30">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" MaxLength="2" Width="6%">10</asp:TextBox>
                        <asp:Button ID="bt_search" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                    </div>
                    <div align="center">
                        <br />
                        <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label><br />
                    </div>
                </td>
            </tr>
        </table>
        <asp:Panel ID="Panel" runat="server" Width="100%" Visible="False">
            <asp:DataGrid ID="DG_Org" runat="server" CssClass="font" Width="100%" Visible="False" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                <Columns>
                    <asp:BoundColumn HeaderText="編號">
                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center" Width="6%"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="OrgName" HeaderText="機構名稱">
                        <ItemStyle HorizontalAlign="Center" Width="20%"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ClassName" HeaderText="班別名稱">
                        <ItemStyle HorizontalAlign="Center" Width="22%"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ClassCateText" HeaderText="課程類別">
                        <ItemStyle HorizontalAlign="Center" Width="20%"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:TemplateColumn HeaderText="審核狀態">
                        <ItemStyle HorizontalAlign="Center" Width="11%"></ItemStyle>
                        <ItemTemplate>
                            <asp:Label ID="LSecResult" runat="server"></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateColumn>
                    <asp:TemplateColumn HeaderText="功能">
                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center" Width="21%" Font-Size="Small"></ItemStyle>
                        <ItemTemplate>
                            <asp:HiddenField ID="HidPCS" runat="server" />
                            <asp:HiddenField ID="HidClassCate" runat="server" />
                            <asp:HiddenField ID="HidClassID" runat="server" />
                            <asp:LinkButton ID="add_but" runat="server" Text="新增" CommandName="add" CssClass="linkbutton"></asp:LinkButton>
                            <asp:LinkButton ID="edit_but" runat="server" Text="修改" CommandName="edit" CssClass="linkbutton"></asp:LinkButton>
                            <asp:LinkButton ID="del_but" runat="server" Text="刪除" CommandName="del" CssClass="linkbutton"></asp:LinkButton>
                            <asp:LinkButton ID="view_but" runat="server" Text="查詢" CommandName="view" CssClass="linkbutton"></asp:LinkButton>
                        </ItemTemplate>
                    </asp:TemplateColumn>
                </Columns>
                <PagerStyle Visible="False"></PagerStyle>
            </asp:DataGrid>
            <div align="center">
                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
            </div>
        </asp:Panel>
    </form>
</body>
</html>
