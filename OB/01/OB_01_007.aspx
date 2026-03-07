 

<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="OB_01_007.aspx.vb" Inherits="WDAIIP.OB_01_007" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>工作小組評選結果資料</title>
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
        /*
        function choose_stn(){
        document.getElementById('txt_name').value='';
        openClass('OB_01_ch.aspx?sort=0');
        }
		
        function check_select(){
        if(document.getElementById('txt_name').value==''){
        window.alert("請選擇【標案名稱】");
        return false; 	
        }
        }
        */
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
        <tr>
            <td>
                <table class="font" id="tab_title" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            <asp:Label ID="TitleLab1" runat="server">
										<FONT face="新細明體">首頁&gt;&gt;委外訓練管理&gt;&gt;</FONT>
                            </asp:Label><asp:Label ID="TitleLab2" runat="server">
										<font color="#990000">工作小組評選結果查詢</font>
                            </asp:Label>
                        </td>
                    </tr>
                </table>
                <asp:Panel ID="Panel_Sch" runat="server">
                    <table id="Table11" class="table_sch" cellspacing="1" cellpadding="1">
                        <tr>
                            <td width="15%" class="bluecol">
                                年度
                            </td>
                            <td style="height: 19px" class="whitecol">
                                <asp:DropDownList ID="ddl_years" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td style="height: 18px" width="10%" class="bluecol">
                                訓練計畫
                            </td>
                            <td style="height: 18px" class="whitecol">
                                <asp:DropDownList ID="ddl_TPlanID" runat="server">
                                </asp:DropDownList>
                                <asp:TextBox ID="PlanName" runat="server" MaxLength="20"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                標案名稱
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="txt_TenderCName" runat="server" MaxLength="20" Width="250px"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                主辦單位
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="txt_Sponsor" runat="server" MaxLength="20" Width="400px"></asp:TextBox>
                            </td>
                        </tr>
                    </table>
                    <p align="center">
                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" MaxLength="2" Width="23px">10</asp:TextBox>&nbsp;
                        <asp:Button ID="btn_Sch" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>
                        <%--
						    <asp:button id="btn_Add" runat="server" Text="新增"></asp:button>
                        --%>
                    </p>
                    <p align="center">
                        <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font" Visible="False">查無資料!!</asp:Label>
                    </p>
                </asp:Panel>
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
                            <asp:BoundColumn DataField="tsn" Visible="False"></asp:BoundColumn>
                            <asp:BoundColumn DataField="years" HeaderText="年度別">
                                <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="PlanName" HeaderText="訓練計畫">
                                <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="TenderCName" HeaderText="標案名稱">
                                <HeaderStyle HorizontalAlign="Center" Width="30%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="left"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="TenderSDate" HeaderText="投標日期">
                                <HeaderStyle HorizontalAlign="Center" Width="12%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="功能">
                                <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Button ID="btn_add" runat="server" Text="新增" CommandName="add"></asp:Button>
                                    <asp:Button ID="btn_edit" runat="server" Text="檢視" CommandName="edit"></asp:Button>
                                    <asp:Button ID="btn_del" runat="server" Text="刪除" CommandName="del"></asp:Button>
                                    <asp:Button ID="btn_prt" runat="server" Text="列印空白表格" CommandName="prt"></asp:Button>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                    <div id="Div1" align="center" runat="server">
                        <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                    </div>
                </asp:Panel>
                <asp:Panel ID="Panel_edit" runat="server" Visible="False">
                    <asp:DataGrid ID="dg_edit" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False"
                        AllowPaging="True">
                        <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn HeaderText="序號">
                                <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn Visible="False" DataField="msn"></asp:BoundColumn>
                            <asp:BoundColumn DataField="memname" HeaderText="工作小組">
                                <HeaderStyle HorizontalAlign="Center" Width="69%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="left"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="功能">
                                <HeaderStyle HorizontalAlign="Center" Width="24%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center" Wrap="False"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Button ID="btn_add2" runat="server" Text="新增" CommandName="add"></asp:Button>
                                    <asp:Button ID="btn_edit2" runat="server" Text="修改" CommandName="edit"></asp:Button>
                                    <asp:Button ID="btn_del2" runat="server" Text="刪除" CommandName="del"></asp:Button>
                                    <asp:Button ID="btn_prt2" runat="server" Text="列印" CommandName="prt"></asp:Button>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                    <div align="center">
                        <asp:Button ID="btn_back" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button></div>
                </asp:Panel>
                <asp:Panel ID="Panel_Add_Edit" runat="server" Visible="False">
                    <table id="TableLay2" class="table_sch" border="0" cellspacing="1" cellpadding="1"
                        width="100%">
                        <tr>
                            <td width="15%" class="bluecol">
                                標案名稱
                            </td>
                            <td style="height: 19px" class="whitecol">
                                <asp:TextBox ID="txt_Name" runat="server" Width="100px" onfocus="this.blur()"></asp:TextBox><input
                                    id="txt_tsn" type="hidden" name="txt_tsn" runat="server">
                                <%--
										<asp:Button id="btn_choose" runat="server" Text="..."></asp:Button>
										<asp:button id="btn_select" runat="server" Text="確定"></asp:button>
										<asp:Button id="btn_clear" runat="server" Text="重選" Visible="False"></asp:Button>
                                --%>
                            </td>
                        </tr>
                        <tr>
                            <td style="height: 18px" width="10%" class="bluecol">
                                工作小組
                            </td>
                            <td style="height: 18px" class="whitecol">
                                <asp:DropDownList ID="ddl_Member" runat="server">
                                    <asp:ListItem Value="---">---</asp:ListItem>
                                </asp:DropDownList>
                                <input id="txt_msn" type="hidden" name="txt_tsn" runat="server">
                            </td>
                        </tr>
                    </table>
                    <input id="txt_num" type="hidden" name="txt_tsn" runat="server">
                    <input id="txt_ORSN" type="hidden" name="txt_tsn" runat="server">
                    <br />
                </asp:Panel>
                <asp:Panel ID="Panel_Item" runat="server" Visible="False">
                    <asp:DataGrid ID="dg_item" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False"
                        AllowPaging="True">
                        <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn DataField="id" Visible="False"></asp:BoundColumn>
                            <asp:BoundColumn DataField="memname" HeaderText="工作小組" Visible="False">
                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="orgname" HeaderText="投標廠商">
                                <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="評選項目1" Visible="False">
                                <HeaderStyle HorizontalAlign="left"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:TextBox ID="txt_item1" runat="server" Width="100%"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="評選項目2" Visible="False">
                                <HeaderStyle HorizontalAlign="left"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:TextBox ID="txt_item2" runat="server" Width="100%"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="評選項目3" Visible="False">
                                <HeaderStyle HorizontalAlign="left"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:TextBox ID="txt_item3" runat="server" Width="100%"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="評選項目4" Visible="False">
                                <HeaderStyle HorizontalAlign="left"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:TextBox ID="txt_item4" runat="server" Width="100%"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="評選項目5" Visible="False">
                                <HeaderStyle HorizontalAlign="left"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:TextBox ID="txt_item5" runat="server" Width="100%"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="評選項目6" Visible="False">
                                <HeaderStyle HorizontalAlign="left"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:TextBox ID="txt_item6" runat="server" Width="100%"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="評選項目7" Visible="False">
                                <HeaderStyle HorizontalAlign="left"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:TextBox ID="txt_item7" runat="server" Width="100%"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="評選項目8" Visible="False">
                                <HeaderStyle HorizontalAlign="left"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:TextBox ID="txt_item8" runat="server" Width="100%"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="評選項目9" Visible="False">
                                <HeaderStyle HorizontalAlign="left"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:TextBox ID="txt_item9" runat="server" Width="100%"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="評選項目10" Visible="False">
                                <HeaderStyle HorizontalAlign="left"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:TextBox ID="txt_item10" runat="server" Width="100%"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                </asp:Panel>
                <div align="center">
                    <asp:Button ID="btn_save" runat="server" Text="儲存" Visible="False" CssClass="asp_button_S">
                    </asp:Button>&nbsp;
                    <asp:Button ID="btn_lev" runat="server" Text="離開" Visible="False" CssClass="asp_button_S">
                    </asp:Button></div>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
