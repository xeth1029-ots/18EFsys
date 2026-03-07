<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="OB_01_006.aspx.vb" Inherits="WDAIIP.OB_01_006" %>

 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>標案評選項目查詢</title>
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
        function check_select() {
            var msg = '';
            if (document.getElementById('ddl_years2').value == '') {
                msg += '請選擇【年度】內容\n';
            }
            if (document.getElementById('ddl_TPlanID2').value == '') {
                msg += '請選擇【訓練計畫】內容\n';
            }
            if (msg != '') {
                window.alert(msg);
                return false;
            }

        }
        function check_score() {
            var msg = '';
            if (document.getElementById('txt_score').value == '0') {
                msg += '【評選項目配分】不可為0\n';
            }
            else {
                if (!isUnsignedInt(document.getElementById('txt_score').value)) {
                    msg += '【評選項目配分】填寫內容有誤\n';
                }
            }
            if (msg != '') {
                window.alert(msg);
                return false;
            }
        }
		
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
										<font color="#990000">標案評選項目查詢</font>
                            </asp:Label>
                        </td>
                    </tr>
                </table>
                <asp:Panel ID="Panel_Sch" runat="server" Visible="True">
                    <table id="TableLay2" class="table_sch" cellspacing="1" cellpadding="1">
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
                    </p>
                    <p align="center">
                        <asp:Label ID="msg" runat="server" Visible="False" ForeColor="Red" CssClass="font">查無資料!!</asp:Label>
                    </p>
                </asp:Panel>
                <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0"
                    runat="server">
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
                                        <asp:BoundColumn DataField="tsn" Visible="False"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="tisn" Visible="False"></asp:BoundColumn>
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
                                            <HeaderStyle HorizontalAlign="Center" Width="24%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Button ID="btn_Add" runat="server" Text="新增" CommandName="add"></asp:Button>
                                                <asp:Button ID="btn_edit" runat="server" Text="修改" CommandName="edit"></asp:Button>
                                                <asp:Button ID="btn_del" runat="server" Text="刪除" CommandName="del"></asp:Button>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                                <div align="center" runat="server">
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                </div>
                            </asp:Panel>
                            <asp:Panel ID="Panel_Add_Edit" Visible="False" runat="server">
                                <table id="Table2" class="table_sch" runat="server">
                                    <tr>
                                        <td class="bluecol_need" rowspan="2" width="15%">
                                            標案內容
                                        </td>
                                        <td class="bluecol" width="22%" align="center">
                                            年度
                                        </td>
                                        <td class="bluecol" width="53%" align="center">
                                            訓練計畫
                                        </td>
                                        <td style="display: none" class="bluecol" width="10%" align="center">
                                            功能
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="center" class="whitecol">
                                            <asp:Label ID="labYears" runat="server"></asp:Label>
                                        </td>
                                        <td align="center" class="whitecol">
                                            <asp:Label ID="labPlanName" runat="server"></asp:Label>
                                        </td>
                                        <td style="display: none" align="center" class="whitecol">
                                            <asp:Button ID="btn_select" runat="server" Text="確定" CssClass="asp_button_S"></asp:Button><%--
													<asp:button id="btn_clear" runat="server" Visible="False" Text="重選"></asp:button>
                                            --%>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td width="15%" class="bluecol_need">
                                            標案名稱
                                        </td>
                                        <td style="height: 13px" colspan="3" align="center" class="whitecol">
                                            <asp:Label ID="LabTenderCName" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                </table>
                                <p></p>
                            </asp:Panel>
                            <asp:Panel ID="Panel_ListBox" Visible="False" runat="server">
                                <table id="Table3" class="font" border="0" cellspacing="1" cellpadding="1" width="100%"
                                    runat="server">
                                    <tr>
                                        <td colspan="3" align="center" class="bluecol">
                                            
                                                <asp:Label ID="LabAction" runat="server"></asp:Label>設定評選大項&nbsp;&nbsp;第
                                            <asp:Label ID="lbl_num" runat="server">1</asp:Label>項
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="height: 14px" colspan="3" align="center" class="whitecol">
                                            <asp:DropDownList ID="ddl_ORName" runat="server" Width="100%" AutoPostBack="True">
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="center">
                                            可選用子項目
                                        </td>
                                        <td>
                                        </td>
                                        <td align="center">
                                            選擇後子項目
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="center">
                                            <asp:ListBox ID="lbx_Source" runat="server" Width="270px" Height="150px"></asp:ListBox>
                                        </td>
                                        <td>
                                            <p>
                                                <asp:ImageButton ID="img_Add" runat="server" ImageUrl="../../images/right2.gif" BackColor="Transparent">
                                                </asp:ImageButton>
                                                <p>
                                                    <asp:ImageButton ID="img_AddAll" runat="server" ImageUrl="../../images/right.gif"
                                                        BackColor="Transparent"></asp:ImageButton>
                                                    <p>
                                                        <asp:ImageButton ID="img_Remove" runat="server" ImageUrl="../../images/left2.gif"
                                                            BackColor="Transparent"></asp:ImageButton>
                                                        <p>
                                                            <asp:ImageButton ID="img_RemoveAll" runat="server" ImageUrl="../../images/left.gif"
                                                                BackColor="Transparent"></asp:ImageButton></p>
                                        </td>
                                        <td align="center">
                                            <asp:ListBox ID="lbx_Get" runat="server" Width="270px" Height="150px"></asp:ListBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td colspan="3" align="center">
                                            <font color="red">*</font>
                                            <asp:Label ID="Label1" runat="server">評選項目配分:</asp:Label>
                                            <asp:TextBox ID="txt_score" runat="server" Width="50px"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td colspan="3" align="center">
                                            <asp:Button ID="btn_ORItem" runat="server" Text="評選項目確定" 
                                                CssClass="asp_button_M"></asp:Button>
                                        </td>
                                    </tr>
                                </table>
                                <asp:DataGrid ID="dg_ORItem" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False"
                                    AllowPaging="True">
                                    <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn HeaderText="順序">
                                            <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="id" Visible="False"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="data" HeaderText="評選項目">
                                            <HeaderStyle HorizontalAlign="Center" Width="65%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="left"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="score" HeaderText="配分">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <HeaderStyle HorizontalAlign="Center" Width="18%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Button ID="btn_edit2" runat="server" Text="修改" CommandName="edit"></asp:Button>
                                                <asp:Button ID="btn_del2" runat="server" Text="刪除" CommandName="del"></asp:Button>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                                <div align="center">
                                    <asp:Button ID="btn_save" runat="server" Text="儲存" CssClass="asp_button_S"></asp:Button>&nbsp;
                                    <asp:Button ID="btn_lev" runat="server" Text="離開" CssClass="asp_button_S"></asp:Button></div>
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
