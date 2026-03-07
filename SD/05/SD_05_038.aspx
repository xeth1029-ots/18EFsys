<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_038.aspx.vb" Inherits="WDAIIP.SD_05_038" %>
 

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>送訓官兵名冊</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
	<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
	<meta name="vs_defaultClientScript" content="JavaScript">
	<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	<link rel="stylesheet" type="text/css" href="../../css/style.css">
	<script type="text/javascript" src="../../js/date-picker.js"></script>
	<script type="text/javascript" src="../../js/openwin/openwin.js"></script>
	<script type="text/javascript" src="../../js/common.js"></script>
</head>
<body>
    <form id="form1" runat="server">
        <br style="line-height: 1px" />
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr>
                <td>
                    <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;送訓官兵名冊</asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table id="tbSch" runat="server" width="100%" cellspacing="1" cellpadding="1">
                        <tr>
                            <td>
                                <table width="100%" class="table_nw" cellpadding="1" cellspacing="1">
                                    <tr>
                                        <td class="bluecol" width="20%">姓名</td>
                                        <td class="whitecol" width="80%"><asp:TextBox ID="sCNAME" runat="server" MaxLength="12" Width="20%"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">身分證號碼</td>
                                        <td class="whitecol" width="80%"><asp:TextBox ID="sIDNO" runat="server" MaxLength="10" Width="20%"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">出生年月日</td>
                                        <td class="whitecol" width="80%">
                                            <asp:TextBox ID="sBIRTHDAY1" runat="server" Columns="10" MaxLength="10" Width="15%"></asp:TextBox>
                                            <span id="span1" runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('<%= sBIRTHDAY1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span> &nbsp;~
                                            <asp:TextBox ID="sBIRTHDAY2" runat="server" Columns="10" MaxLength="10" Width="15%"></asp:TextBox>
                                            <span id="span2" runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('<%= sBIRTHDAY2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">預定退伍日</td>
                                        <td class="whitecol" width="80%">
                                            <asp:TextBox ID="sPREEXDATE1" runat="server" Columns="10" MaxLength="10" Width="15%"></asp:TextBox>
                                            <span id="span3" runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('<%= sPREEXDATE1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span> &nbsp;~
                                            <asp:TextBox ID="sPREEXDATE2" runat="server" Columns="10" MaxLength="10" Width="15%"></asp:TextBox>
                                            <span id="span4" runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('<%= sPREEXDATE2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">薦訓至分署</td>
                                        <td class="whitecol" width="80%"><asp:DropDownList ID="sRECOMMDISTID" runat="server"></asp:DropDownList></td>
                                    </tr>
                                    <tr>
                                        <td align="center" colspan="2" class="whitecol">
                                            <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue" DESIGNTIMEDRAGDROP="30">顯示列數</asp:Label>
                                            <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                            <asp:Button ID="btnSearch1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <div align="center"><asp:Label ID="labmsg" runat="server" CssClass="font" ForeColor="Red"></asp:Label></div>
                                <table id="tbList" runat="server" cellspacing="0" bordercolordark="#ffffff" cellpadding="0" width="100%" align="left" bordercolorlight="#666666" border="0">
                                    <tr>
                                        <td>
                                            <div align="center">
                                                <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Width="100%" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <Columns>
                                                        <asp:BoundColumn HeaderText="序號">
                                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="NAME" HeaderText="姓名">
                                                            <HeaderStyle HorizontalAlign="Center" Width="16%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="IDNO" HeaderText="身分證字號">
                                                            <HeaderStyle HorizontalAlign="Center" Width="18%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="DISTNAME" HeaderText="薦訓至分署">
                                                            <HeaderStyle HorizontalAlign="Center" Width="30%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:TemplateColumn HeaderText="出生年月日">
                                                            <HeaderStyle HorizontalAlign="Center" Width="14%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="BIRTHDAY" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.BIRTHDAY") %>'>
                                                                </asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="預定退伍日">
                                                            <HeaderStyle HorizontalAlign="Center" Width="14%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="PREEXDATE" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.PREEXDATE") %>'>
                                                                </asp:Label>
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
        </table>
    </form>
</body>
</html>