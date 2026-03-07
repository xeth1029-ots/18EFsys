<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="OB_01_ch.aspx.vb" Inherits="WDAIIP.OB_01_ch" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>標案名稱查詢</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script src="../../js/common.js"></script>
    <script src="../../js/TIMS.js"></script>
    <script language="javascript">
        function returnNum() {
            window.opener.form1.TMID1.value = document.form1.class1.value;
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%"
        border="0">
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
                <asp:Panel ID="Panel_Sch" runat="server" Visible="True">
                    <table class="font" id="TableLay2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td width="15%" bgcolor="#2aafc0">
                                <font face="新細明體" color="#ffffff">&nbsp;&nbsp; 年度</font>
                            </td>
                            <td class="SD_TD2" style="height: 19px" colspan="3">
                                <asp:DropDownList ID="ddl_years" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td style="height: 18px" width="10%" bgcolor="#2aafc0">
                                <font face="新細明體" color="#ffffff">&nbsp;&nbsp; 訓練計畫</font>
                            </td>
                            <td class="SD_TD2" style="height: 18px" colspan="3">
                                <asp:DropDownList ID="ddl_TPlanID" runat="server">
                                </asp:DropDownList>
                                <asp:TextBox ID="PlanName" runat="server" MaxLength="20"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="4">
                                &nbsp;
                                <asp:Button ID="btn_Sch" runat="server" Text="查詢"></asp:Button>
                                <div align="center">
                                    <asp:Label ID="msg" runat="server" Visible="False" ForeColor="Red" CssClass="font">查無資料!!</asp:Label></div>
                            </td>
                        </tr>
                    </table>
                </asp:Panel>
                <asp:Panel ID="Panel_View" runat="server" Visible="False">
                    <asp:DataGrid ID="dg_Sch" runat="server" CssClass="font" Width="100%" AutoGenerateColumns="False"
                        AllowPaging="True">
                        <AlternatingItemStyle BackColor="White"></AlternatingItemStyle>
                        <ItemStyle BackColor="#EBF8FF"></ItemStyle>
                        <HeaderStyle ForeColor="White" BackColor="#2AAFC0"></HeaderStyle>
                        <Columns>
                            <asp:TemplateColumn>
                                <ItemStyle HorizontalAlign="Center" Width="7%"></ItemStyle>
                                <ItemTemplate>
                                    <input id="radio1" value='<%# DataBinder.Eval(Container.DataItem,"tsn")%>' type="radio"
                                        name="tsn">
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn DataField="tsn" Visible="False"></asp:BoundColumn>
                            <asp:BoundColumn DataField="years" HeaderText="年度別">
                                <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="PlanName" HeaderText="訓練計畫">
                                <HeaderStyle HorizontalAlign="Center" Width="30%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="TenderCName" HeaderText="標案名稱">
                                <HeaderStyle HorizontalAlign="Center" Width="44%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="left"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="TenderSDate" HeaderText="投標日期">
                                <HeaderStyle HorizontalAlign="Center" Width="12%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                    <div align="center">
                        <asp:Button ID="send" runat="server" Text="送出"></asp:Button></div>
                </asp:Panel>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
