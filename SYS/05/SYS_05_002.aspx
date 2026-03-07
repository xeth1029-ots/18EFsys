<%@ Page Language="vb" AutoEventWireup="true" CodeBehind="SYS_05_002.aspx.vb" Inherits="WDAIIP.SYS_05_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>上稿維護-公告維護</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
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
        //檢查日期格式
        function check_date() {
            if (!checkDate(form1.SPostDate.value) || !checkDate(form1.EPostDate.value)) {
                document.form1.SPostDate.value = '';
                document.form1.EPostDate.value = '';
                alert('請輸入正確的日期格式,YYYY/MM/DD!!\n');
            }
        }
    </script>
    <style type="text/css">
        A:link { color: #000000; text-decoration: none; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;公告維護</asp:Label>
                </td>
            </tr>
        </table>
        <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td class="bluecol" width="20%">項目 </td>
                <td class="whitecol">
                    <asp:DropDownList ID="ItemList" runat="server">
                        <asp:ListItem Value="1">News</asp:ListItem>
                        <asp:ListItem Value="2">新功能</asp:ListItem>
                        <asp:ListItem Value="3">文件下載</asp:ListItem>
                        <asp:ListItem Value="4">影音教學</asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">發布日期 </td>
                <td class="whitecol">
                    <font face="新細明體">
                        <asp:TextBox ID="SPostDate" runat="server" MaxLength="10" ToolTip="日期格式:99/01/31" Columns="13" Width="15%"></asp:TextBox>
                        <span id="span_SPostDate" runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= SPostDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">&nbsp;~
                        </span>
                        <asp:TextBox ID="EPostDate" runat="server" MaxLength="10" ToolTip="日期格式:99/01/31" Columns="13" Width="15%"></asp:TextBox>
                        <span id="span_EPostDate" runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= EPostDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                        </span>
                    </font>
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">優先排序 </td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="RBL_ORDERBY" runat="server" RepeatDirection="Horizontal">
                        <asp:ListItem Value="M" Selected="True">異動日期</asp:ListItem>
                        <asp:ListItem Value="P">發布日期</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">永遠顯示 </td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="RBL_SHOWTYPE" runat="server" RepeatDirection="Horizontal">
                        <asp:ListItem Value="" Selected="True">(不區分)</asp:ListItem>
                        <asp:ListItem Value="Y">是</asp:ListItem>
                        <asp:ListItem Value="N">否</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td align="center" colspan="2" class="whitecol">
                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue" DESIGNTIMEDRAGDROP="30">顯示列數</asp:Label>
                    <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                    <asp:Button ID="bt_search" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                    &nbsp;<asp:Button ID="bt_add" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                    &nbsp;<asp:Button ID="reset" runat="server" Text="重新設定" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
            <tr>
                <td colspan="2" align="center">
                    <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                </td>
            </tr>
            <tr>
                <td colspan="2" align="center">

                    <table id="Table5" cellspacing="0" cellpadding="0" width="100%" align="left" border="0">
                        <tbody>
                            <tr>
                                <td>
                                    <div align="center">
                                        <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Width="100%" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                                            <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                            <Columns>
                                                <asp:BoundColumn HeaderText="序號">
                                                    <HeaderStyle HorizontalAlign="Center" Width="5%" CssClass="head_navy"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:TemplateColumn HeaderText="項目">
                                                    <HeaderStyle HorizontalAlign="Center" Width="10%" CssClass="head_navy"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:Label ID="LabType2" runat="server"></asp:Label>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:TemplateColumn HeaderText="發布日期">
                                                    <HeaderStyle HorizontalAlign="Center" Width="10%" CssClass="head_navy"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:Label ID="Label4" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.PostDate", "{0:d}") %>'>
                                                        </asp:Label>
                                                    </ItemTemplate>
                                                    <EditItemTemplate>
                                                        &nbsp;
                                                    </EditItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:TemplateColumn HeaderText="發布主題">
                                                    <HeaderStyle HorizontalAlign="Center" Width="40%" CssClass="head_navy"></HeaderStyle>
                                                    <ItemTemplate>
                                                        <asp:Label ID="Label2" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.Subject2") %>'>
                                                        </asp:Label>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:TemplateColumn HeaderText="永遠顯示">
                                                    <HeaderStyle HorizontalAlign="Center" Width="5%" CssClass="head_navy"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:Label ID="LabisShow" runat="server"></asp:Label>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:BoundColumn DataField="Name" HeaderText="異動者">
                                                    <HeaderStyle HorizontalAlign="Center" Width="10%" CssClass="head_navy"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="ModifyDate" HeaderText="異動時間">
                                                    <HeaderStyle HorizontalAlign="Center" Width="10%" CssClass="head_navy"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:TemplateColumn HeaderText="功能" ItemStyle-HorizontalAlign="Center" HeaderStyle-CssClass="head_navy" HeaderStyle-Width="10%">
                                                    <ItemStyle Font-Size="Small" />
                                                    <ItemTemplate>
                                                        <asp:LinkButton ID="lbtUpdate" Text="修改" runat="server" CssClass="linkbutton" CommandName="UCmd1"></asp:LinkButton>
                                                        <asp:LinkButton ID="lbtDelete" Text="刪除" runat="server" CssClass="linkbutton" CommandName="DCmd1"></asp:LinkButton>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                            </Columns>
                                            <PagerStyle Visible="False"></PagerStyle>
                                        </asp:DataGrid>
                                        <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                    </div>
                                </td>
                            </tr>

                        </tbody>
                    </table>

                </td>
            </tr>

        </table>

    </form>
</body>
</html>
