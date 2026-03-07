<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_04_013.aspx.vb" Inherits="WDAIIP.SYS_04_013" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>個人行事曆</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;系統參數管理&gt;&gt;<font color="#990000">個人行事曆</font></asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table3" class="table_sch" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td class="bluecol" width="20%">主旨
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="schSubject" runat="server" MaxLength="20" Columns="40" Width="50%"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="bluecol">開始日期區間
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="schOSDate1" runat="server" Width="100px" onfocus="this.blur()"></asp:TextBox>
                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= schOSDate1.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">
                    ～
					<asp:TextBox ID="schOSDate2" runat="server" Width="100px" onfocus="this.blur()"></asp:TextBox>
                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= schOSDate2.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">
                </td>
            </tr>
        </table>
        <table width="100%">
            <tr>
                <td class="whitecol" colspan="2" align="center">
                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue" CssClass="font">顯示列數</asp:Label><asp:TextBox ID="TxtPageSize" runat="server" MaxLength="2" Width="23px">10</asp:TextBox>&nbsp;
							<asp:Button ID="btnSearch" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>&nbsp;
							<asp:Button ID="btnInsert" runat="server" Text="新增" CssClass="asp_button_S"></asp:Button>
                </td>
            </tr>
            <tr>
                <td class="whitecol" colspan="2" align="center">&nbsp;<asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
                </td>
            </tr>
        </table>
        <table id="DataGridTable" class="font" border="0" cellspacing="1" cellpadding="1" runat="server" width="100%">
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid1" CssClass="font" AllowPaging="True" AutoGenerateColumns="False" runat="server" Width="100%">
                        <%--<AlternatingItemStyle BackColor="White"></AlternatingItemStyle>--%>
                        <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                        <%--<ItemStyle BackColor="#FFECEC"></ItemStyle>--%>
                        <%--<HeaderStyle  ForeColor="White" BackColor="#CC6666"></HeaderStyle>--%>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn HeaderText="序號">
                                <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="主旨">
                                <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="alsubject" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="日期開始">
                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="alOSDate" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="日期結束">
                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="alOFDate" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="功能">
                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:LinkButton ID="BtnEdit" runat="server" Text="修改" CommandName="Edit" CssClass="linkbutton"></asp:LinkButton>
                                    <asp:LinkButton ID="BtnDel" runat="server" Text="刪除" CommandName="Del" CssClass="linkbutton"></asp:LinkButton>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                </td>
            </tr>
        </table>
        <%--<table id="Table1" class="font" border="0" cellspacing="1" cellpadding="1" width="740">
		<tr>
			<td>
				 
			
			</td>
		</tr>
	</table>--%>
    </form>
</body>
</html>
