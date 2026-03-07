<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="EXAM_02_001.aspx.vb" Inherits="WDAIIP.EXAM_02_001" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>題組題庫維護</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1" />
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1" />
    <meta name="vs_defaultClientScript" content="JavaScript" />
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5" />
    <link rel="stylesheet" type="text/css" href="../../css/style.css" />
    <script language="javascript" type="text/javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script language="javascript" type="text/javascript" src="../../js/common.js"></script>
    <script language="javascript" type="text/javascript" src="../../js/TIMS.js"></script>
    <script language="javascript" type="text/javascript"></script>
    <style type="text/css">
        .auto-style1 { color: Black; line-height: 26px; background-color: #e9f1fe; padding: 4px; height: 24px; }
        .auto-style2 { color: #333333; padding: 4px; height: 24px; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;甄試管理&gt;&gt;甄試題組題庫設定</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable1" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>

                    <asp:Panel ID="Panel_Sch" runat="server" Visible="True">
                        <table id="table_Sch" class="table_sch" runat="server" cellspacing="1" cellpadding="1">
                            <tr>

                                <td class="bluecol" style="width: 20%">轄區分署</td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="ddl_DistID" runat="server" AutoPostBack="True"></asp:DropDownList></td>
                            </tr>
                            <tr>
                                <td class="bluecol">題組類別</td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="ddl_etid" runat="server" AutoPostBack="True"></asp:DropDownList></td>
                            </tr>
                            <tr>
                                <td class="bluecol">題組子類別</td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="ddl_cETID" runat="server"></asp:DropDownList></td>
                            </tr>
                            <tr>
                                <td class="bluecol">題目類型</td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="ddl_qtype" runat="server">
                                        <asp:ListItem Value="0" Selected="True">--請選擇--</asp:ListItem>
                                        <asp:ListItem Value="1">是非題</asp:ListItem>
                                        <asp:ListItem Value="2">選擇題</asp:ListItem>
                                        <asp:ListItem Value="3">複選題</asp:ListItem>
                                        <asp:ListItem Value="4">問答題</asp:ListItem>
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">匯入題庫</td>
                                <td class="whitecol">
                                    <p>
                                        <input id="File1" size="50" type="file" name="File1" runat="server" accept=".xls,.ods" />(必須為ods或xls格式)<br>
                                        <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="../../Doc/question_input1.zip" CssClass="font" ForeColor="#8080FF">下載整批上載格式檔</asp:HyperLink>
                                        <asp:Button Style="z-index: 0" ID="BtnImport1" runat="server" Text="匯入題庫" CssClass="asp_button_M"></asp:Button>
                                    </p>
                                </td>
                            </tr>
                        </table>
                        <table width="100%">
                            <tr>
                                <td align="center" class="whitecol">
                                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                    <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                    <asp:Button ID="btn_Sch" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="btn_Add" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>
                                    <div align="center">
                                        <asp:Label ID="msg" runat="server" Visible="False" CssClass="font" ForeColor="Red">查無資料!!</asp:Label>
                                    </div>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                    <table id="Table1" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tbody>
                            <tr>
                                <td id="Td1" align="center" runat="server">
                                    <asp:Panel ID="Panel_View" runat="server" Visible="False">
                                        <asp:DataGrid ID="dg_Sch" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                                            <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                            <Columns>
                                                <asp:BoundColumn HeaderText="序號">
                                                    <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="pName" HeaderText="題組類別">
                                                    <HeaderStyle HorizontalAlign="Center" Width="25%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="cName" HeaderText="題組子類別">
                                                    <HeaderStyle HorizontalAlign="Center" Width="25%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="qtype" HeaderText="題目類型">
                                                    <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="total" HeaderText="題數">
                                                    <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:TemplateColumn HeaderText="功能">
                                                    <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:Button ID="btn_view" runat="server" Text="檢視" CommandName="view" CssClass="asp_button_M"></asp:Button>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                            </Columns>
                                            <PagerStyle Visible="False"></PagerStyle>
                                        </asp:DataGrid>
                                        <div id="Div1" align="center" runat="server">
                                            <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                        </div>
                                    </asp:Panel>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:Panel ID="Panel_View2" runat="server" Visible="False">
                                        <table id="tab_title2" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                                            <tr>
                                                <td width="25%">
                                                    <asp:Label ID="lbl_title1" runat="server" Font-Size="Small" Font-Bold="True">題組類別：</asp:Label>
                                                    <asp:Label ID="lbl_petidname" runat="server" Font-Size="Small" Font-Bold="True"></asp:Label>
                                                </td>
                                                <td width="25%">
                                                    <asp:Label ID="Label3" runat="server" Font-Size="Small" Font-Bold="True">題組子類別：</asp:Label>
                                                    <asp:Label ID="lbl_cetidname" runat="server" Font-Size="Small" Font-Bold="True"></asp:Label>
                                                </td>
                                                <td width="35%">
                                                    <asp:Label ID="lbl_title2" runat="server" Font-Size="Small" Font-Bold="True">題目類型：</asp:Label>
                                                    <asp:Label ID="lbl_qtype" runat="server" Font-Size="Small" Font-Bold="True"></asp:Label>
                                                </td>
                                                <td width="15%" align="right">
                                                    <asp:Label ID="lbl_title3" runat="server" Font-Size="Small" Font-Bold="True">題數：</asp:Label>
                                                    <asp:Label ID="lbl_math" runat="server" Font-Size="Small" Font-Bold="True"></asp:Label>
                                                </td>
                                            </tr>
                                        </table>
                                        <br />
                                        <font class="font">題目搜尋：</font>
                                        <div class="whitecol">
                                            <asp:TextBox ID="txtQuestion" runat="server" Width="20%"></asp:TextBox>
                                            <asp:Button ID="btnSch" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                        </div>
                                        <asp:DataGrid ID="dg_Sch2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                                            <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                            <Columns>
                                                <asp:BoundColumn HeaderText="序號">
                                                    <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="question" HeaderText="題目">
                                                    <HeaderStyle HorizontalAlign="Center" Width="40%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="left"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="total" HeaderText="選項數目">
                                                    <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="stopuse" HeaderText="不啟用">
                                                    <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:TemplateColumn HeaderText="功能">
                                                    <HeaderStyle HorizontalAlign="Center" Width="25%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:Button ID="btn_view2" runat="server" Text="檢視" CommandName="view" CssClass="asp_button_M"></asp:Button>
                                                        <asp:Button ID="btn_edit2" runat="server" Text="修改" CommandName="edit" CssClass="asp_button_M"></asp:Button>
                                                        <asp:Button ID="btn_del2" runat="server" Text="刪除" CommandName="del" CssClass="asp_button_M"></asp:Button>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                            </Columns>
                                            <PagerStyle Visible="False"></PagerStyle>
                                        </asp:DataGrid>
                                        <div id="Div2" align="center" runat="server">
                                            <uc1:PageControler ID="PageControler2" runat="server"></uc1:PageControler>
                                            <br />
                                            <asp:Label ID="msg2" runat="server" Visible="False" CssClass="font" ForeColor="Red">查無資料!!</asp:Label>
                                        </div>
                                        <div>&nbsp;</div>
                                        <div align="center" class="whitecol">
                                            <asp:Button ID="btn_back" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                                        </div>
                                    </asp:Panel>
                                    <asp:Panel ID="Panel_show" runat="server" Visible="False">
                                        <table id="Table2" class="table_sch" cellspacing="1" cellpadding="1" runat="server">
                                            <tr>
                                                <td class="bluecol" style="width: 20%">題組類別
                                                </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="lbl_vpetidname" runat="server"></asp:Label>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td class="bluecol">題組子類別
                                                </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="lbl_vcetidname" runat="server"></asp:Label>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td class="bluecol">題目名稱
                                                </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="lbl_vquestion" runat="server"></asp:Label>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td class="bluecol">題目類型
                                                </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="lbl_vqtype" runat="server"></asp:Label>
                                                </td>
                                            </tr>
                                        </table>
                                    </asp:Panel>
                                    <asp:Panel ID="tab_select1" runat="server" Visible="False">
                                        <table id="Table3" class="table_sch" cellspacing="1" cellpadding="1" runat="server">
                                            <tr>
                                                <td width="15%" align="center" class="bluecol">
                                                    <asp:Label ID="Label5" runat="server">解答</asp:Label>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td width="15%" align="center" class="whitecol">
                                                    <asp:Label ID="lbl_ans1" runat="server"></asp:Label>
                                                </td>
                                            </tr>
                                        </table>
                                    </asp:Panel>
                                    <asp:Panel ID="tab_select2" Visible="False" runat="server">
                                        <table id="Table4" class="table_sch" cellspacing="1" cellpadding="1" runat="server">
                                            <tr>
                                                <td width="7%" class="bluecol">
                                                    <asp:Label ID="Label2" runat="server">
															選號</asp:Label>
                                                </td>
                                                <td width="86%" class="bluecol">
                                                    <asp:Label ID="Label7" runat="server">
															選項</asp:Label>
                                                </td>
                                                <td width="7%" class="bluecol">
                                                    <asp:Label ID="Label8" runat="server">
															解答</asp:Label>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td width="7%" align="center" class="auto-style1">1
                                                </td>
                                                <td width="86%" class="auto-style2">
                                                    <asp:Label ID="lbl_ans2_1" runat="server" Visible="False"></asp:Label>
                                                    <asp:LinkButton ID="lkb_ans2_1" runat="server" Visible="False" ForeColor="Blue"></asp:LinkButton>
                                                </td>
                                                <td width="7%" align="center" class="auto-style2">
                                                    <asp:Label ID="lbl_chk2_1" runat="server"></asp:Label>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="height: 25px" width="7%" align="center" class="td_light">2
                                                </td>
                                                <td style="height: 25px" width="86%" class="whitecol">
                                                    <asp:Label ID="lbl_ans2_2" runat="server" Visible="False"></asp:Label>
                                                    <asp:LinkButton ID="lkb_ans2_2" runat="server" Visible="False" ForeColor="Blue"></asp:LinkButton>
                                                </td>
                                                <td width="7%" align="center" class="whitecol">
                                                    <asp:Label ID="lbl_chk2_2" runat="server"></asp:Label>
                                                </td>
                                            </tr>
                                        </table>
                                    </asp:Panel>
                                    <asp:Panel ID="tab_select3" Visible="False" runat="server">
                                        <table id="Table5" class="table_sch" cellspacing="1" cellpadding="1" runat="server">
                                            <tr>
                                                <td width="7%" align="center" class="td_light">3
                                                </td>
                                                <td width="86%" class="whitecol">
                                                    <asp:Label ID="lbl_ans2_3" runat="server" Visible="False"></asp:Label>
                                                    <asp:LinkButton ID="lkb_ans2_3" runat="server" Visible="False" ForeColor="Blue"></asp:LinkButton>
                                                </td>
                                                <td width="7%" align="center" class="whitecol">
                                                    <asp:Label ID="lbl_chk2_3" runat="server"></asp:Label>
                                                </td>
                                            </tr>
                                        </table>
                                    </asp:Panel>
                                    <asp:Panel ID="tab_select4" Visible="False" runat="server">
                                        <table id="Table6" class="table_sch" cellspacing="1" cellpadding="1" runat="server">
                                            <tr>
                                                <td style="height: 25px" width="7%" align="center" class="td_light">4
                                                </td>
                                                <td style="height: 25px" width="86%" class="whitecol">
                                                    <asp:Label ID="lbl_ans2_4" runat="server" Visible="False"></asp:Label>
                                                    <asp:LinkButton ID="lkb_ans2_4" runat="server" Visible="False" ForeColor="Blue"></asp:LinkButton>
                                                </td>
                                                <td width="7%" align="center" class="whitecol">
                                                    <asp:Label ID="lbl_chk2_4" runat="server"></asp:Label>
                                                </td>
                                            </tr>
                                        </table>
                                    </asp:Panel>
                                    <asp:Panel ID="tab_select5" Visible="False" runat="server">
                                        <table id="Table7" class="table_sch" cellspacing="1" cellpadding="1" runat="server">
                                            <tr>
                                                <td style="height: 25px" width="7%" align="center" class="td_light">5
                                                </td>
                                                <td style="height: 25px" width="86%" class="whitecol">
                                                    <asp:Label ID="lbl_ans2_5" runat="server" Visible="False"></asp:Label>
                                                    <asp:LinkButton ID="lkb_ans2_5" runat="server" Visible="False" ForeColor="Blue"></asp:LinkButton>
                                                </td>
                                                <td width="7%" align="center" class="whitecol">
                                                    <asp:Label ID="lbl_chk2_5" runat="server"></asp:Label>
                                                </td>
                                            </tr>
                                        </table>
                                    </asp:Panel>
                                    <asp:Panel ID="tab_select6" Visible="False" runat="server" Height="10px">
                                        <table id="Table8" class="table_sch" cellspacing="1" cellpadding="1" runat="server">
                                            <tr>
                                                <td width="100%" class="bluecol">
                                                    <asp:Label ID="Label1" runat="server">解答</asp:Label>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td class="whitecol">
                                                    <asp:Label ID="lbl_ans4" runat="server"></asp:Label>
                                                </td>
                                            </tr>
                                        </table>
                                    </asp:Panel>
                                    <div>
                                        &nbsp;
                                    </div>
                                    <div align="center" class="whitecol">
                                        <asp:Button ID="btn_back2" runat="server" Visible="False" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
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
