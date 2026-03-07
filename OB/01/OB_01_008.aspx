<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="OB_01_008.aspx.vb" Inherits="WDAIIP.OB_01_008" %>

 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>評審委員評選分數新增查詢</title>
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
        function choose_stn() {
            document.getElementById('txt_name').value = '';
            openClass('OB_01_ch.aspx?sort=1');
        }
        function sum(num) {
            var mytable = document.getElementById('dg_item');
            document.getElementById('txt_score').value = 0;
            for (var table_num = 0; table_num < parseInt(num) + 1; table_num++) {
                var txt2 = mytable.rows(table_num).cells(2).innerHTML;
                var txt = mytable.rows(table_num).cells(3).children(0);
                if (table_num > 0) {
                    if (txt.value != '') {
                        if (!isUnsignedInt(txt.value)) {
                            window.alert('輸入格式有誤，請輸入數字');
                            txt.select();
                            txt.focus();
                        }
                        else {
                            if (parseInt(txt2, 10) < parseInt(txt.value, 10) && txt2 != '配分') {
                                window.alert('輸入分數大於配分，請輸入數字');
                                txt.select();
                                txt.focus();
                            }
                            else {
                                document.getElementById('txt_score').value = parseInt(document.getElementById('txt_score').value, 10) + parseInt(txt.value, 10);
                            }
                        }
                    }
                }
            }
        }
        function check_save() {
            var msg = '';
            if (document.getElementById('txt_judgenumber').value == '') {
                msg = '請填寫【評選編號】內容\n';
            }
            if (document.getElementById('txt_judgedate').value == '') {
                msg += '請選擇【評選日期】內容\n';
            }
            if (document.getElementById('ddl_tcsn').value == '') {
                msg += '請選擇【申請單位】內容\n';
            }
            var j = 0;
            var mytable = document.getElementById('dg_item');
            for (var i = 1; i < mytable.rows.length; i++) {
                var txt = mytable.rows(i).cells(3).children(0);
                if (txt.value == '') {
                    j = j + 1;
                }
            }
            if (j != '0') {
                msg += '【評選分數】內容尚有填寫\n';
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
                            <asp:Label ID="TitleLab1" runat="server"> <FONT face="新細明體">首頁&gt;&gt;委外訓練管理&gt;&gt;</FONT></asp:Label><asp:Label ID="TitleLab2" runat="server"> <font color="#990000">評審委員評選分數新增查詢</font> </asp:Label>
                        </td>
                    </tr>
                </table>
                <asp:Panel ID="Panel_Sch" runat="server">
                    <%--
							<TR>
								<TD width="15%" bgColor="#2aafc0"><FONT face="新細明體" color="#ffffff">&nbsp;&nbsp; 標案名稱</FONT></TD>
								<TD class="SD_TD2" style="HEIGHT: 19px" colSpan="3">
									<asp:textbox id="txt_Name" runat="server" Width="100px" onfocus="this.blur()"></asp:textbox><INPUT id="txt_tsn" type="hidden" name="txt_tsn" runat="server">
									<asp:Button id="btn_choose" runat="server" Text="..."></asp:Button></TD>
							</TR>
							<TR>
								<TD bgColor="#2aafc0"><FONT face="新細明體" color="#ffffff">&nbsp;&nbsp; 主辦單位</FONT>
								</TD>
								<TD class="SD_TD2" colSpan="3">
									<asp:textbox id="txt_Sponsor" runat="server" Width="400px" MaxLength="20"></asp:textbox></TD>
							</TR>
                    --%>
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
                    </p>
                    <p align="center">
                        <asp:Label ID="msg" runat="server" ForeColor="Red" Visible="False" CssClass="font">查無資料!!</asp:Label>
                    </p>
                </asp:Panel>
                <asp:Panel ID="Panel_View" runat="server" Visible="False">
                    <asp:DataGrid ID="dg_Sch" runat="server" Width="100%" CssClass="font" AllowPaging="True"
                        AutoGenerateColumns="False">
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
                                <HeaderStyle HorizontalAlign="Center" Width="35%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="left"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="num" HeaderText="評選數">
                                <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="功能">
                                <HeaderStyle HorizontalAlign="Center" Width="24%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Button ID="btn_add" runat="server" Text="新增" CommandName="add"></asp:Button>
                                    <asp:Button ID="btn_edit" runat="server" Text="檢視" CommandName="edit"></asp:Button>
                                    <asp:Button ID="btn_del" runat="server" Text="刪除" CommandName="del"></asp:Button>
                                    <asp:Button ID="btn_prt" runat="server" Text="列印空白表" CommandName="prt"></asp:Button>
                                    <asp:Label ID="view_msg" runat="server" Visible="False">評選項目未建置</asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                    <div id="Div1" align="center" runat="server">
                        <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                    </div>
                </asp:Panel>
                <asp:Panel ID="Panel_Edit" runat="server" Visible="False">
                    <asp:DataGrid ID="dg_edit" runat="server" Width="100%" CssClass="font" AllowPaging="True"
                        AutoGenerateColumns="False">
                        <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn HeaderText="序號">
                                <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn Visible="False" DataField="otssn"></asp:BoundColumn>
                            <asp:BoundColumn DataField="judgenumber" HeaderText="評選編號">
                                <HeaderStyle HorizontalAlign="Center" Width="30%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="orgname" HeaderText="申請單位">
                                <HeaderStyle HorizontalAlign="Center" Width="47%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="left"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="功能">
                                <HeaderStyle HorizontalAlign="Center" Width="16%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Button ID="btn_edit2" runat="server" Text="修改" CommandName="edit"></asp:Button>
                                    <asp:Button ID="btn_del2" runat="server" Text="刪除" CommandName="del"></asp:Button>
                                    <asp:Button ID="btn_prt2" runat="server" Text="列印評選表" CommandName="prt"></asp:Button>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                    <div align="center">
                        <asp:Button ID="btn_back" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button></div>
                </asp:Panel>
                <asp:Panel ID="Panel_Add_Edit" runat="server" Visible="False">
                    <table id="Table2" class="table_sch">
                        <tr>
                            <td height="28" width="15%" class="bluecol_need">
                                標案名稱
                            </td>
                            <td class="whitecol">
                                <asp:Label ID="lbl_tsn" runat="server"></asp:Label><input id="hid_tsn" type="hidden"
                                    runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td width="15%" class="bluecol_need">
                                評選編號
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="txt_judgenumber" runat="server"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td width="15%" class="bluecol_need">
                                評選日期
                            </td>
                            <td style="height: 20px" class="whitecol" colspan="3">
                                <asp:TextBox ID="txt_judgedate" runat="server" onfocus="this.blur()" Width="80px"></asp:TextBox><img
                                    style="cursor: pointer" onclick="javascript:show_calendar('<%= txt_judgedate.ClientId %>','','','CY/MM/DD');"
                                    alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">
                            </td>
                        </tr>
                        <tr>
                            <td width="15%" class="bluecol_need">
                                申請單位
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="ddl_tcsn" runat="server">
                                </asp:DropDownList>
                                <input id="hid_tisn" type="hidden" runat="server">
                            </td>
                        </tr>
                    </table>
                    <asp:DataGrid ID="dg_item" runat="server" Width="100%" CssClass="font" AllowPaging="True"
                        AutoGenerateColumns="False">
                        <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn HeaderText="序號">
                                <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn Visible="False" DataField="id"></asp:BoundColumn>
                            <asp:BoundColumn DataField="data" HeaderText="評選項目">
                                <HeaderStyle HorizontalAlign="Center" Width="50%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Left"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="score" HeaderText="配分">
                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="小計">
                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:TextBox ID="txt_dscore" runat="server" Width="100%"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="備註">
                                <HeaderStyle HorizontalAlign="Center" Width="23%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:TextBox ID="txt_opinion" runat="server" Width="100%"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                    <div align="center">
                        <asp:Label ID="Label1" runat="server">統計：</asp:Label>
                        <asp:TextBox ID="txt_score" runat="server" Width="30px" onfocus="this.blur()">0</asp:TextBox>
                        <asp:Label ID="Label2" runat="server">分</asp:Label></div>
                    <div align="left">
                        <strong><font size="3">評審意見：</font></strong></div>
                    <asp:TextBox ID="txt_commnet" runat="server" Width="100%" Height="100px"></asp:TextBox>
                    <div align="center">
                        <asp:Button ID="btn_save" runat="server" Text="儲存" CssClass="asp_button_S"></asp:Button>&nbsp;
                        <asp:Button ID="btn_lev" runat="server" Text="離開" CssClass="asp_button_S"></asp:Button></div>
                </asp:Panel>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
