<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_06_003.aspx.vb" Inherits="WDAIIP.SYS_06_003" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>郵件發送</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;郵件發送</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>

                    <asp:Panel ID="panelSearch" runat="server">
                        <table id="Table3" class="table_sch" border="0" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol" style="width: 20%">發送範圍</td>
                                <td class="whitecol">
                                    <asp:RadioButtonList ID="rblPlanID_s" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                        <asp:ListItem Value="0" Selected="True">全系統</asp:ListItem>
                                        <asp:ListItem Value="1">指定計畫</asp:ListItem>
                                    </asp:RadioButtonList>
                                    <asp:DropDownList Style="z-index: 0" ID="ddlplanlist_s" runat="server" CssClass="font">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">發送對象</td>
                                <td class="whitecol">
                                    <asp:CheckBoxList Style="z-index: 0" ID="cblobjecttype_s" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow"
                                        CssClass="font">
                                    </asp:CheckBoxList>
                                </td>
                            </tr>
                             <tr>
                                <td class="bluecol">測試Email信箱</td>
                                <td class="whitecol"><asp:TextBox ID="txtTestEmail1" runat="server" MaxLength="600" Width="88%"></asp:TextBox></td>
                            </tr>

                            <tr>
                                <td colspan="2" align="center" class="whitecol">
                                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                    <asp:TextBox ID="TxtPageSize" runat="server" MaxLength="2" Width="6%">10</asp:TextBox>
                                    <asp:Button ID="btnQuery" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>&nbsp;
                                    <asp:Button ID="btnAddnew" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>&nbsp;
                                    <asp:Button ID="btnMailTest1" runat="server" Text="ws郵件測試" CssClass="asp_button_M"></asp:Button>&nbsp;
                                    <asp:Button ID="btnMailTest2" runat="server" Text="sp郵件測試" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="btnMailInfo1" runat="server" Text="檢視寄送資訊" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2" align="center">
                                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                                </td>
                            </tr>
                        </table>
                        <table id="DataGridTable1" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                            <tr>
                                <td align="center">
                                    <asp:DataGrid Style="z-index: 0" ID="DataGrid1" runat="server" CssClass="font" Width="100%" AutoGenerateColumns="False" AllowPaging="True">
                                        <AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
                                        <ItemStyle BackColor="White"></ItemStyle>
                                        <HeaderStyle CssClass="head_navy" />
                                        <Columns>
                                            <asp:BoundColumn HeaderText="序號">
                                                <HeaderStyle Width="5%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="Subject" HeaderText="主題">
                                                <HeaderStyle Width="40%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="SendDate" HeaderText="發送日期">
                                                <HeaderStyle Width="10%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="PlanName" HeaderText="發送範圍">
                                                <HeaderStyle Width="10%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="發送對象">
                                                <HeaderStyle Width="10%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="lbobjectType" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:BoundColumn DataField="SendState" HeaderText="發送狀況">
                                                <HeaderStyle Width="10%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="功能">
                                                <HeaderStyle Width="15%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:LinkButton ID="btnEdit1" runat="server" Text="修改" CommandName="Edit1" CssClass="linkbutton"></asp:LinkButton>
                                                    <asp:LinkButton ID="btnDelete1" runat="server" Text="刪除" CommandName="Delete1" CssClass="linkbutton"></asp:LinkButton>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                        <PagerStyle Visible="False"></PagerStyle>
                                    </asp:DataGrid>
                                </td>
                            </tr>
                            <tr>
                                <td style="height: 31px" align="center">
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>

                    <asp:Panel ID="PanelEdit1" runat="server">
                        <table id="Table4" class="table_sch" border="0" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol" style="width: 20%">主題<font color="red">*</font>
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox Style="z-index: 0" ID="txtSubject" runat="server" MaxLength="100" Columns="50" Width="50%"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">發送日期<font color="red">*</font>
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="SendDate" MaxLength="10" Width="15%" onfocus="this.blur()" runat="server"></asp:TextBox>
                                    <span runat="server">
                                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= SendDate.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif">
                                    </span>
                                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                            </tr>
                            <tr>
                                <td class="bluecol">發送範圍<font color="red">*</font>
                                </td>
                                <td class="whitecol">
                                    <asp:RadioButtonList Style="z-index: 0" ID="rblPlanID" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                        <asp:ListItem Value="0">全系統</asp:ListItem>
                                        <asp:ListItem Value="1">指定計畫</asp:ListItem>
                                    </asp:RadioButtonList>
                                    <asp:DropDownList Style="z-index: 0" ID="ddlplanlist" runat="server">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">發送對象<font color="red">*</font>
                                </td>
                                <td class="whitecol">
                                    <asp:CheckBoxList Style="z-index: 0" ID="cblobjecttype" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" CssClass="font">
                                    </asp:CheckBoxList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">郵件內容<font color="red">*</font>
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox Style="z-index: 0" ID="txtcontents" runat="server" Width="50%" Rows="10" TextMode="MultiLine"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">狀態
                                </td>
                                <td class="whitecol">
                                    <asp:Label ID="labIsApprPaper" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">寄送報告EMAIL
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txtAcctEmail" runat="server" MaxLength="100" Width="50%"></asp:TextBox>
                                </td>
                            </tr>


                            <tr>
                                <td class="bluecol">收件人數
                                </td>
                                <td class="whitecol">
                                    <asp:Button ID="BtnQuery1" runat="server" Text="收件人數查詢" CssClass="asp_button_M" />
                                </td>
                            </tr>

                            <tr>
                                <td class="whitecol" colspan="2" align="center">
                                    <asp:Button ID="btnSave1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>&nbsp;
                                <asp:Button ID="btnSave2" runat="server" Text="草稿儲存" CssClass="asp_button_M"></asp:Button>&nbsp;
                                <asp:Button ID="btnBack1" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                </td>
            </tr>
            <%--
				<TR>
					<TD>
					</TD>
				</TR>
            --%>
        </table>
        <asp:HiddenField ID="Hid_MSID" runat="server" />
        <asp:HiddenField ID="Hid_IsApprPaper" runat="server" />

        <%--<input id="hidmsid" type="hidden" name="hidmsid" runat="server">--%>
    </form>
</body>
</html>
