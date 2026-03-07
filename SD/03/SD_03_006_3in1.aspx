<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_03_006_3in1.aspx.vb" Inherits="TIMS.SD_03_006_3in1" %>

<%@ Register TagPrefix="uc1" TagName="DataGridPage" Src="../../DataGridPage.ascx" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_03_006_3in1</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script>
        function checkRadio(objName, num) {
            var mytable = document.getElementById(objName);

            for (var i = 1; i < mytable.rows.length; i++) {
                var myradio = mytable.rows(i).cells(0).children(0);
                if (num == i)
                    myradio.checked = true;
                else
                    myradio.checked = false;
            }
        }
    </script>
</head>
<body ms_positioning="FlowLayout">
    <form id="form1" method="post" runat="server">
    <font face="新細明體">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td align="center">
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">
										首頁&gt;&gt;學員動態管理&gt;&gt;報到&gt;&gt;<FONT color="#990000">結訓學員資料維護-學習券查詢</FONT>
                                </asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch" id="Table2">
                        <tr>
                            <td class="bluecol_need" width="100">
                                發卷單位
                            </td>
                            <td style="height: 16px" colspan="3" class="whitecol">
                                <asp:DropDownList ID="center" runat="server" AutoPostBack="True">
                                </asp:DropDownList>
                                <asp:DropDownList ID="station" runat="server" AutoPostBack="True">
                                </asp:DropDownList>
                                <asp:DropDownList ID="tai" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                身分證號碼
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="IDNO" runat="server" MaxLength="20"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                開券日期
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="start_date" runat="server" Width="100px"></asp:TextBox><img style="cursor: hand" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="24" height="24"><font face="新細明體">～</font>
                                <asp:TextBox ID="end_date" runat="server" Width="100px"></asp:TextBox><img style="cursor: hand" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="24" height="24">
                            </td>
                        </tr>
                    </table>
                    <p align="center">
                        <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>
                        <asp:Button ID="Button2" runat="server" Width="109px" Text="回學員資料維護" CssClass="asp_button_M"></asp:Button>
                    </p>
                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="Datagrid1" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" AllowPaging="True">
                                    <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <ItemTemplate>
                                                <input id="Radio1" type="radio" value='<%# DataBinder.Eval(Container.DataItem,"IDNO")%>' name="IDNO" runat="server">
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="name" HeaderText="姓名"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="IDNO" HeaderText="身分證號碼"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="TICKET_NO" HeaderText="學習券編號"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="APPLY_DATE" HeaderText="發卷日期" DataFormatString="{0:d}"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="Share_Name" HeaderText="適用對象"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="TICKET_TYPE" HeaderText="參加學習單元"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="Station_Name" HeaderText="就服中心"></asp:BoundColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <uc1:DataGridPage ID="DataGridPage1" runat="server"></uc1:DataGridPage>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <asp:Button ID="Button3" runat="server" Text="新增" CssClass="asp_button_S"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                </td>
            </tr>
        </table>
    </font>
    </form>
</body>
</html>
