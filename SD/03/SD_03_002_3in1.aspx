<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_03_002_3in1.aspx.vb" Inherits="WDAIIP.SD_03_002_3in1" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_03_002_3in1</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript">
        function checkRadio(objName, num) {
            var mytable = document.getElementById(objName);
            for (var i = 1; i < mytable.rows.length; i++) {
                var myradio = mytable.rows[i].cells[0].children[0];
                if (num == i)
                    myradio.checked = true;
                else
                    myradio.checked = false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <%--<table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;報到&gt;&gt;<FONT color="#990000">學員資料維護-學習券查詢</FONT></asp:Label>
                            </td>
                        </tr>
                    </table>--%>
                    <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="bluecol" width="20%">發卷單位<font color="#ff0000">*</font></td>
                            <td class="whitecol" width="80%">
                                <asp:DropDownList ID="center" runat="server" AutoPostBack="True"></asp:DropDownList>
                                <asp:DropDownList ID="station" runat="server" AutoPostBack="True"></asp:DropDownList>
                                <asp:DropDownList ID="tai" runat="server"></asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">身分證號碼</td>
                            <td class="whitecol">
                                <asp:TextBox ID="IDNO" runat="server" Width="40%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">開券日期</td>
                            <td class="whitecol">
                                <asp:TextBox ID="start_date" runat="server" Width="20%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: hand" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span> ～
                                <asp:TextBox ID="end_date" runat="server" Width="20%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: hand" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="4" class="whitecol">
                                <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="Button2" runat="server" Text="回學員資料維護" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="Datagrid1" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" AllowPaging="True" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <%-- <ItemStyle BackColor="#ECF7FF"></ItemStyle>--%>
                                    <HeaderStyle ForeColor="White" CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="2%" />
                                            <ItemTemplate>
                                                <input id="Radio1" type="radio" value='<%# DataBinder.Eval(Container.DataItem,"IDNO")%>' name="IDNO" runat="server">
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="name" HeaderText="姓名">
                                            <HeaderStyle Width="14%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="IDNO" HeaderText="身分證號碼">
                                            <HeaderStyle Width="14%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TICKET_NO" HeaderText="學習券編號">
                                            <HeaderStyle Width="14%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="APPLY_DATE" HeaderText="發卷日期" DataFormatString="{0:d}">
                                            <HeaderStyle Width="14%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Share_Name" HeaderText="適用對象">
                                            <HeaderStyle Width="14%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TICKET_TYPE" HeaderText="參加學習單元">
                                            <HeaderStyle Width="14%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Station_Name" HeaderText="就服中心">
                                            <HeaderStyle Width="14%" />
                                        </asp:BoundColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <uc1:DataGridPage ID="DataGridPage1" runat="server"></uc1:DataGridPage></td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="Button3" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button></td>
                        </tr>
                    </table>
                    <div style="width: 100%" align="center">
                        <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></div>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
