<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="OB_01_005.aspx.vb" Inherits="WDAIIP.OB_01_005" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>評選項目資料查詢</title>
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
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <table class="font" id="tab_title" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server">
										<FONT face="新細明體">首頁&gt;&gt;委外訓練管理&gt;&gt;</FONT>
                                </asp:Label><asp:Label ID="TitleLab2" runat="server">
										<font color="#990000">評選項目資料查詢</font>
                                </asp:Label>
                            </td>
                        </tr>
                    </table>
                    <asp:Panel ID="Panel_Sch" runat="server" Visible="True">
                        <table id="table_Sch" class="table_sch" runat="server" cellspacing="1" cellpadding="1">
                            <tr>
                                <%--<td class="bluecol" width="100">轄區中心</td>--%>
                                <td class="bluecol" width="100">轄區分署</td>
                                <td bgcolor="#ecf7ff" class="whitecol">
                                    <asp:DropDownList ID="ddl_DistID" runat="server">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">評選項目名稱
                                </td>
                                <td bgcolor="#ebf8ff" class="whitecol">
                                    <asp:TextBox ID="txt_ORName" runat="server" Width="200px"></asp:TextBox>
                                </td>
                            </tr>
                        </table>
                        <p align="center">
                            <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                            <asp:TextBox ID="TxtPageSize" runat="server" Width="23px" MaxLength="2">10</asp:TextBox>
                            <asp:Button ID="btn_Sch" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>&nbsp;
                        <asp:Button ID="btn_Add" runat="server" Text="新增" CssClass="asp_button_S"></asp:Button>
                        </p>
                        <p align="center">
                            <asp:Label ID="msg" runat="server" Visible="False" ForeColor="Red" CssClass="font">查無資料!!</asp:Label>
                        </p>
                    </asp:Panel>
                    <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td id="Td1" align="center" runat="server">
                                <asp:Panel ID="Panel_View" runat="server" Visible="False">
                                    <asp:DataGrid ID="dg_Sch" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False"
                                        AllowPaging="True">
                                        <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                        <Columns>
                                            <asp:BoundColumn HeaderText="序號">
                                                <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="ORName" HeaderText="評選項目名稱">
                                                <HeaderStyle HorizontalAlign="Center" Width="50%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="left"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="num" HeaderText="細項數目">
                                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="ORAvail" HeaderText="啟用">
                                                <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="功能">
                                                <HeaderStyle HorizontalAlign="Center" Width="26%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Button ID="btn_view" runat="server" Text="細項" CommandName="view"></asp:Button>
                                                    <asp:Button ID="btn_edit" runat="server" Text="修改" CommandName="edit"></asp:Button>
                                                    <%--
														<asp:Button id="btn_del" runat="server" Text="刪除" CommandName="del"></asp:Button>
                                                    --%>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                        <PagerStyle Visible="False"></PagerStyle>
                                    </asp:DataGrid>
                                    <div id="Div1" align="center" runat="server">
                                        <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                    </div>
                                </asp:Panel>
                                <asp:Panel ID="Panel_View2" runat="server" Visible="False">
                                    <asp:DataGrid ID="dg_Sch2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False"
                                        AllowPaging="True">
                                        <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                        <Columns>
                                            <asp:BoundColumn HeaderText="序號">
                                                <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="ORName" HeaderText="評選項目名稱">
                                                <HeaderStyle HorizontalAlign="Center" Width="68%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="left"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="ORAvail" HeaderText="啟用">
                                                <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="功能">
                                                <HeaderStyle HorizontalAlign="Center" Width="18%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Button ID="btn_edit2" runat="server" Text="修改" CommandName="edit"></asp:Button>
                                                    <%--
														<asp:Button id="btn_del2" runat="server" Text="刪除" CommandName="del"></asp:Button>
                                                    --%>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                        <PagerStyle Visible="False"></PagerStyle>
                                    </asp:DataGrid>
                                    <div align="center">
                                        <asp:Label ID="msg2" runat="server" Visible="False" ForeColor="Red" CssClass="font">查無資料!!</asp:Label>
                                    </div>
                                    <asp:Button ID="btn_Add2" runat="server" Text="新增" CssClass="asp_button_S"></asp:Button>
                                    <font face="新細明體">&nbsp;</font>
                                    <asp:Button ID="btn_back2" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                                </asp:Panel>
                                <asp:Panel ID="Panel_Add_Edit" runat="server" Visible="False">
                                    <table id="Table2" class="table_sch" runat="server" cellspacing="1" cellpadding="1">
                                        <tr>
                                            <td width="100" class="bluecol_need">評選項目名稱
                                            </td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="txt_Name" runat="server" Width="100%"></asp:TextBox>
                                            </td>
                                        </tr>
                                        <tr>
                                            <%--<td width="100" class="bluecol">轄區中心</td>--%>
                                            <td width="100" class="bluecol">轄區分署</td>
                                            <td class="whitecol">
                                                <asp:Label ID="txt_DistID" runat="server" Font-Size="10"></asp:Label>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td style="height: 13px" width="100" class="bluecol">是否啟用
                                            </td>
                                            <td style="height: 13px" class="whitecol">
                                                <asp:CheckBox ID="ckb_ORAvail" runat="server" Width="104px" Checked="True"></asp:CheckBox>
                                            </td>
                                        </tr>
                                    </table>
                                    <p align="center">
                                        <asp:Button ID="btn_save" runat="server" Text="儲存" CssClass="asp_button_S"></asp:Button>&nbsp;
                                    <asp:Button ID="btn_lev" runat="server" Text="離開" CssClass="asp_button_S"></asp:Button>
                                    </p>
                                </asp:Panel>
                                <asp:Panel ID="Panel_Add_Edit2" runat="server" Visible="False">
                                    <table id="Table3" class="table_sch" runat="server">
                                        <tr>
                                            <td class="bluecol_need" width="100">評選細項名稱
                                            </td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="txt_Name2" runat="server" Width="100%"></asp:TextBox>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">是否啟用
                                            </td>
                                            <td style="height: 20px" class="whitecol">
                                                <asp:CheckBox ID="ckb_ORAvail2" runat="server" Width="104px" Checked="True"></asp:CheckBox>
                                            </td>
                                        </tr>
                                    </table>
                                    <p align="center">
                                        <asp:Button ID="btn_save2" runat="server" Text="儲存" CssClass="asp_button_S"></asp:Button>&nbsp;
                                    <asp:Button ID="btn_lev2" runat="server" Text="離開" CssClass="asp_button_S"></asp:Button>
                                    </p>
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
