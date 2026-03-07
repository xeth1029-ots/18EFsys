<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_05_004.aspx.vb" Inherits="WDAIIP.SYS_05_004" %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>產投線上報名停止設定</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
</head>
<body>
    <form id="form1" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;上稿維護&gt;&gt;產投線上報名停止設定</asp:Label>
                </td>
            </tr>
        </table>
        <br style="line-height: 1px" />
        <table id="FrameTable3" width="100%" border="0" cellpadding="0" cellspacing="0">
            <%-- <tr> <td></td> </tr>--%>
            <tr>
                <td>
                    <table id="tbSch" runat="server" width="100%" cellspacing="1" cellpadding="1">
                        <tr>
                            <td>
                                <table width="100%" class="table_nw" cellpadding="1" cellspacing="1">
                                    <tr>
                                        <td class="bluecol" width="20%">查詢種類 </td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="RBL_QTYPE" runat="server" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="H3" Selected="True">產投</asp:ListItem>
                                                <asp:ListItem Value="H4">非產投</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" style="width: 20%">停止開始日期</td>
                                        <td class="whitecol" runat="server">
                                            <asp:TextBox ID="StopSDate1" runat="server" Columns="10" MaxLength="10" Width="15%"></asp:TextBox>&nbsp;<img style="cursor: pointer" onclick="javascript:show_calendar('<%= StopSDate1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />&nbsp; ~<asp:TextBox ID="StopSDate2" runat="server" Columns="10" MaxLength="10" Width="15%"></asp:TextBox>&nbsp;<img style="cursor: pointer" onclick="javascript:show_calendar('<%= StopSDate2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />&nbsp;
                                        </td>
                                    </tr>
                                    <tr>
                                        <td width="15%" class="bluecol">停止結束日期</td>
                                        <td class="whitecol" runat="server">
                                            <asp:TextBox ID="StopEDate1" runat="server" Columns="10" MaxLength="10" Width="15%"></asp:TextBox>&nbsp;<img style="cursor: pointer" onclick="javascript:show_calendar('<%= StopEDate1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />&nbsp; ~<asp:TextBox ID="StopEDate2" runat="server" Columns="10" MaxLength="10" Width="15%"></asp:TextBox>&nbsp;<img style="cursor: pointer" onclick="javascript:show_calendar('<%= StopEDate2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />&nbsp;
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="center" colspan="2" class="whitecol">
                                            <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue" DESIGNTIMEDRAGDROP="30">顯示列數</asp:Label>
                                            <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                            <asp:Button ID="btnSearch1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                            &nbsp;<asp:Button ID="btnAdd1" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <div align="center">
                                    <asp:Label ID="labmsg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                                </div>
                                <table id="tbList" runat="server" cellspacing="0" bordercolordark="#ffffff" cellpadding="0" width="100%" align="left" bordercolorlight="#666666" border="0">
                                    <tr>
                                        <td>
                                            <div align="center">
                                                <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Width="100%" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <Columns>
                                                        <asp:BoundColumn HeaderText="序號">
                                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:TemplateColumn HeaderText="停止開始日期">
                                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="StopSDate" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.StopSDate") %>'>
                                                                </asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="停止結束日期">
                                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="StopEDate" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.StopEDate") %>'>
                                                                </asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="發布主題">
                                                            <HeaderStyle HorizontalAlign="Center" Width="45%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="labSubject" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.Subject") %>'>
                                                                </asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:BoundColumn DataField="ModifyName" HeaderText="異動者">
                                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="PostDate" HeaderText="發佈日期">
                                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:TemplateColumn HeaderText="功能" ItemStyle-HorizontalAlign="Center">
                                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" Font-Size="Small" />
                                                            <ItemTemplate>
                                                                <asp:LinkButton ID="lbtUpdate" Text="修改" runat="server" CssClass="linkbutton" CommandName="UPD"></asp:LinkButton>
                                                                <asp:LinkButton ID="lbtDelete" Text="刪除" runat="server" CssClass="linkbutton" CommandName="DEL"></asp:LinkButton>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                    </Columns>
                                                    <PagerStyle Visible="False"></PagerStyle>
                                                </asp:DataGrid>
                                                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                            </div>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table id="tbEdit" runat="server" class="table_sch" width="100%" border="0" cellspacing="1" cellpadding="1">
                          <tr>
                            <td class="bluecol_need">查詢種類
                            </td>
                            <td class="whitecol" runat="server">
                                <asp:Label ID="labQTYPE_N" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" style="width: 20%">停止日期區間
                            </td>
                            <td class="whitecol" runat="server">
                                <asp:TextBox ID="StopSDate" runat="server" Columns="10" MaxLength="10" Width="15%"></asp:TextBox>&nbsp;<img style="cursor: pointer" onclick="javascript:show_calendar('<%= StopSDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />&nbsp;
							<asp:DropDownList ID="HR1" runat="server">
                            </asp:DropDownList>
                                <asp:DropDownList ID="MM1" runat="server">
                                </asp:DropDownList>
                                ~<asp:TextBox ID="StopEDate" runat="server" Columns="10" MaxLength="10" Width="15%"></asp:TextBox>&nbsp;<img style="cursor: pointer" onclick="javascript:show_calendar('<%= StopEDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />&nbsp;
							<asp:DropDownList ID="HR2" runat="server">
                            </asp:DropDownList>
                                <asp:DropDownList ID="MM2" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">發布日期
                            </td>
                            <td class="whitecol" runat="server">
                                <asp:TextBox ID="PostDate" runat="server" Columns="10" MaxLength="10" Width="15%"></asp:TextBox>&nbsp;<img style="cursor: pointer" onclick="javascript:show_calendar('<%= PostDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />&nbsp;
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">發布主題
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="txtSubject" runat="server" Columns="58" Rows="8" TextMode="MultiLine" Width="40%"></asp:TextBox>
                            </td>
                        </tr>
                        <%-- <tr>
                            <td class="bluecol">上傳檔案
                            </td>
                            <td class="whitecol">
                                <input id="File1" type="file" size="60" name="File1" runat="server" accept=".xls,.pdf" />
                                <asp:Button ID="BtnUpload" runat="server" CausesValidation="False" CssClass="asp_button_S" Text="匯入"></asp:Button>
                            </td>
                        </tr>--%>
                        <tr>
                            <td class="whitecol" colspan="2" align="center">
                                <asp:Label ID="lab_UpMsg1" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2" class="whitecol" align="center">
                                <asp:Button ID="btnSave1" Text="儲存" runat="server" CssClass="asp_button_M"></asp:Button>
                                &nbsp;<asp:Button ID="btnBack1" runat="server" Text="回上頁" CssClass="asp_button_M"></asp:Button>
                                &nbsp;
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="HidQTYPE" runat="server" />
        <asp:HiddenField ID="HidHN3ID" runat="server" />
        <asp:HiddenField ID="HidHN4ID" runat="server" />
    </form>
</body>
</html>
