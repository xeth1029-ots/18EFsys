<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="EXAM_01_001.aspx.vb" Inherits="WDAIIP.EXAM_01_001" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>題組類別維護</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script src="../../js/common.js"></script>
    <script src="../../js/TIMS.js"></script>
    <script language="javascript">

    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;甄試管理&gt;&gt;甄試題組類別設定</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <%--<table class="font" id="tab_title" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                            <asp:Label ID="TitleLab2" runat="server">
										首頁&gt;&gt;招生甄試成績管理&gt;&gt;甄試題組類別設定&gt;&gt;<font color="#990000">題組類別維護</font>
                            </asp:Label>
                        </td>
                    </tr>
                </table>--%>
                    <asp:Panel ID="Panel_Sch" runat="server" Visible="True">
                        <table id="table_Sch" class="table_sch" runat="server" cellspacing="1" cellpadding="1">
                            <tr>
                                <%--<td class="bluecol" style="width:20%">轄區中心</td>--%>
                                <td class="bluecol" style="width: 20%">轄區分署</td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="ddl_DistID" runat="server">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">類別名稱
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_pName" runat="server" Width="30%"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">子類別名稱
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_cName" runat="server" Width="30%"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">匯入題組類別
                                </td>
                                <td class="whitecol">
                                    <p>
                                        <input id="File1" size="50" type="file" name="File1" runat="server" cssclass="asp_button_M" accept=".xls,.ods" />
                                        (必須為ods或xls格式)<br />
                                        <%--
                                        --%>
                                        <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="../../Doc/ExamType_Input1.zip"
                                            CssClass="font" ForeColor="#8080FF">下載整批上載格式檔</asp:HyperLink>
                                        <asp:Button Style="z-index: 0" ID="BtnImport1" runat="server" Text="匯入題組類別" CssClass="asp_button_M"></asp:Button>
                                    </p>
                                </td>
                            </tr>
                        </table>
                        <table width="100%">
                            <tr>
                                <td align="center" class="whitecol">
                                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                    <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                    <asp:Button ID="btn_Sch" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>&nbsp;
                                <asp:Button ID="btn_Add" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>&nbsp;
                                <div align="center">
                                    <asp:Label ID="msg" runat="server" Visible="False" CssClass="font" ForeColor="Red">查無資料!!</asp:Label>
                                    </div>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                    <table cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td align="center" runat="server">
                                <asp:Panel ID="Panel_View" runat="server" Visible="False">
                                    <asp:DataGrid ID="dg_Sch" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                                        <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                        <Columns>
                                            <asp:BoundColumn HeaderText="序號">
                                                <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="pName" HeaderText="類別名稱">
                                                <HeaderStyle HorizontalAlign="Center" Width="19%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="cName" HeaderText="子類別名稱">
                                                <HeaderStyle HorizontalAlign="Center" Width="19%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                            </asp:BoundColumn>
                                            <%--<asp:BoundColumn DataField="DistName" HeaderText="轄區中心">--%>
                                            <asp:BoundColumn DataField="DistName" HeaderText="轄區分署">
                                                <HeaderStyle HorizontalAlign="Center" Width="19%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="Avail" HeaderText="啟用">
                                                <HeaderStyle HorizontalAlign="Center" Width="19%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="功能">
                                                <HeaderStyle HorizontalAlign="Center" Width="19%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Button ID="btn_edit" runat="server" Text="修改" CommandName="edit" CssClass="asp_button_M"></asp:Button>
                                                    <asp:Button ID="btn_del" runat="server" Text="刪除" CommandName="del" CssClass="asp_button_M"></asp:Button>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                        <PagerStyle Visible="False"></PagerStyle>
                                    </asp:DataGrid>
                                    <div align="center" runat="server">
                                        <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                    </div>
                                </asp:Panel>
                                <asp:Panel ID="Panel_Add" runat="server" Visible="False">
                                    <table class="table_sch" runat="server" cellspacing="1" cellpadding="1">
                                        <tr>
                                            <%--<td class="bluecol" style="width:20%">轄區中心</td>--%>
                                            <td class="bluecol" style="width: 20%">轄區分署</td>
                                            <td class="whitecol">
                                                <asp:Label ID="txt_distid" runat="server"></asp:Label>
                                                <input id="hDistIDVal" type="hidden" runat="server">
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">上層類別
                                            </td>
                                            <td class="whitecol">
                                                <asp:DropDownList ID="ddlParent" runat="server"></asp:DropDownList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">子類別名稱
                                            </td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="txt_dname" runat="server" Width="30%" MaxLength="25"></asp:TextBox>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">是否啟用
                                            </td>
                                            <td class="whitecol">
                                                <asp:RadioButtonList ID="rbl_avail" runat="server" Width="10%" RepeatDirection="Horizontal">
                                                    <asp:ListItem Value="1" Selected="True">是</asp:ListItem>
                                                    <asp:ListItem Value="0">否</asp:ListItem>
                                                </asp:RadioButtonList>
                                            </td>
                                        </tr>
                                    </table>
                                    <table width="100%">
                                        <tr>
                                            <td align="center" class="whitecol">
                                                <asp:Button ID="btn_save" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                                                <asp:Button ID="btn_exit" runat="server" Text="離開" CssClass="asp_button_M"></asp:Button>
                                            </td>
                                        </tr>
                                    </table>
                                </asp:Panel>
                                <asp:Panel ID="Panel_edit" runat="server" Visible="False">
                                    <table id="Table1" class="table_sch" runat="server" cellspacing="1" cellpadding="1" width="100%">
                                        <tr>
                                            <%--<td class="bluecol" style="width:20%">轄區中心</td>--%>
                                            <td class="bluecol" style="width: 20%">轄區分署</td>
                                            <td class="whitecol">
                                                <asp:DropDownList ID="ddl_editDistID" runat="server">
                                                </asp:DropDownList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">上層類別
                                            </td>
                                            <td class="whitecol">
                                                <asp:DropDownList ID="ddleditParent" runat="server"></asp:DropDownList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">子類別名稱
                                            </td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="txt_editdname" runat="server" Width="30%" MaxLength="25"></asp:TextBox>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">是否啟用
                                            </td>
                                            <td class="whitecol">
                                                <asp:RadioButtonList ID="rbl_eavail" runat="server" Width="10%" RepeatDirection="Horizontal">
                                                    <asp:ListItem Value="1" Selected="True">是</asp:ListItem>
                                                    <asp:ListItem Value="0">否</asp:ListItem>
                                                </asp:RadioButtonList>
                                            </td>
                                        </tr>
                                    </table>
                                    <table width="100%">
                                        <tr>
                                            <td align="center" class="whitecol">
                                                <asp:Button ID="btn_editsave" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                                                <asp:Button ID="btn_editexit" runat="server" Text="離開" CssClass="asp_button_M"></asp:Button>
                                            </td>
                                        </tr>
                                    </table>
                                </asp:Panel>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
