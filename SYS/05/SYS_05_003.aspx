<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_05_003.aspx.vb" Inherits="WDAIIP.SYS_05_003" %>

<%--ValidateRequest="false"--%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>產投報名網站公告維護</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script src="../../js/common.js" type="text/javascript"></script>
    <%--<script src="../../ckeditor/ckeditor.js" type="text/javascript"></script> css:ckeditor--%>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
</head>
<body>
    <form id="form1" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;上稿維護&gt;&gt;產投報名網站公告維護</asp:Label>
                </td>
            </tr>
        </table>
        <br style="line-height: 1px" />
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr>
                <td></td>
            </tr>
            <tr>
                <td>
                    <%--<table id="tbTitle1" runat="server" class="font" width="100%" border="0">
					<tr>
						<td>
							<asp:Label ID="TitleLab1" runat="server"></asp:Label>
							<asp:Label ID="TitleLab2" runat="server">
				            首頁&gt;&gt;系統管理&gt;&gt;上稿維護&gt;&gt;<font color="#990000">產投報名網站公告維護</font>
							</asp:Label>
						</td>
					</tr>
				</table>--%>
                    <table id="tbSch" runat="server" width="100%" cellspacing="1" cellpadding="1">
                        <tr>
                            <td>
                                <table width="100%" class="table_nw" cellpadding="1" cellspacing="1">
                                    <tr>
                                        <td class="bluecol" style="width: 20%">頁籤項目
                                        </td>
                                        <td class="whitecol">
                                            <asp:DropDownList ID="ddlQTabNum" runat="server">
                                                <asp:ListItem Value="6">報名資料維護</asp:ListItem>
                                                <asp:ListItem Value="1">開班資料查詢</asp:ListItem>
                                                <asp:ListItem Value="2">線上報名</asp:ListItem>
                                                <asp:ListItem Value="3">線上報名查詢</asp:ListItem>
                                                <asp:ListItem Value="4">補助金申請查詢</asp:ListItem>
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">內部序號
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="txtQSeqno" runat="server" Columns="10" MaxLength="10"></asp:TextBox>
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
                                                        <asp:TemplateColumn HeaderText="頁籤">
                                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="labTABNUM2" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.TABNUM2") %>'>
                                                                </asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="內部序號">
                                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="labSeqno" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.Seqno") %>'>
                                                                </asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="最新">
                                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="labShowNews" runat="server">
                                                                </asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="發布主題">
                                                            <HeaderStyle HorizontalAlign="Center" Width="40%"></HeaderStyle>
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
                            <td class="bluecol" style="width: 20%">頁籤項目
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="TabNum" runat="server">
                                    <asp:ListItem Value="6">報名資料維護</asp:ListItem>
                                    <asp:ListItem Value="1">開班資料查詢</asp:ListItem>
                                    <asp:ListItem Value="2">線上報名</asp:ListItem>
                                    <asp:ListItem Value="3">線上報名查詢</asp:ListItem>
                                    <asp:ListItem Value="4">補助金申請查詢</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">內部序號
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="Seqno" runat="server" Columns="10" MaxLength="10" Width="20%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" bgcolor="#96b5e3">發布日期
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="PostDate" runat="server" Columns="10" MaxLength="10" Width="15%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('<%= PostDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">&nbsp;
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">發布主題
                            </td>
                            <td class="whitecol">
                                <%--CssClass="ckeditor"--%>
                                <asp:TextBox ID="Subject" runat="server" Columns="58" Rows="8" TextMode="MultiLine" Width="40%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">顯示最新消息圖示
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="ShowNews" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="Y">是</asp:ListItem>
                                    <asp:ListItem Value="N" Selected="True">否</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2" class="whitecol" align="center">
                                <asp:Button ID="btnSave1" Text="儲存" runat="server" CssClass="asp_button_M"></asp:Button>
                                &nbsp;<asp:Button ID="btnBack1" runat="server" Text="回上頁" CssClass="asp_button_M"></asp:Button>
                                &nbsp;<asp:Button ID="btnPreview" runat="server" Text="預覽" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="HidHN2ID" runat="server" />
    </form>
</body>
</html>
