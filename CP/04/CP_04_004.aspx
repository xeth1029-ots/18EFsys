<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_04_004.aspx.vb" Inherits="WDAIIP.CP_04_004" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>CP_04_004</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../style.css" type="text/css" rel="stylesheet">
    <style type="text/css">
        .class_link A { color: #000000; }
            .class_link A:link { color: #0000ff; }
        A:visited { color: #0000ff; }
        .class_link A:hover { color: #0000ff; }
        A:active { color: #0000ff; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" width="600">
            <tr>
                <td>
                    <font class="font" size="2">首頁&gt;&gt;訓練查核與績效管理&gt;&gt;訓練資料查詢&gt;&gt;<font class="font" color="#800000">師資</font></font><font class="font" color="#800000" size="2">資料</font>
                </td>
            </tr>
        </table>
        <table class="font" id="Table1" style="width: 600px; height: 106px" cellspacing="1" cellpadding="1" width="600" border="1">
            <tr>
                <td style="width: 10%; height: 28px" bgcolor="#cccc66">
                    <font class="font" face="新細明體" size="2">年度<font color="crimson">*</font></font>
                </td>
                <td style="height: 28px">
                    <asp:DropDownList ID="yearlist" runat="server">
                    </asp:DropDownList>
                    <asp:RequiredFieldValidator ID="MustYear" runat="server" ErrorMessage="請選擇年度" Display="Dynamic" ControlToValidate="yearlist" CssClass="font"></asp:RequiredFieldValidator></FONT></FONT>
                </td>
            </tr>
            <tr>
                <td style="width: 10%" bgcolor="#cccc66">
                    <font class="font" face="新細明體" size="2">轄區</font>
                </td>
                <td>
                    <table id="Table2" style="width: 100%; height: 52px" cellspacing="1" cellpadding="1" width="536" border="0">
                        <tr>
                        </tr>
                    </table>
                    <asp:CheckBoxList ID="DistrictList" runat="server" CssClass="font" AutoPostBack="True" Height="11px" Width="512px" RepeatDirection="Horizontal">
                    </asp:CheckBoxList>
                </td>
            </tr>
            <tr>
                <td style="width: 10%" bgcolor="#cccc66">
                    <font class="font" face="新細明體" size="2">內外聘別</font>
                </td>
                <td>
                    <table id="Table3" style="width: 100%; height: 52px" cellspacing="1" cellpadding="1" width="536" border="0">
                        <tr>
                        </tr>
                    </table>
                    <asp:RadioButton ID="AllKindEngage" runat="server" GroupName="KindEngage" Text="全部"></asp:RadioButton><asp:RadioButton ID="KindEngage1" runat="server" GroupName="KindEngage" Text="內聘"></asp:RadioButton><asp:RadioButton ID="KindEngage2" runat="server" GroupName="KindEngage" Text="外聘"></asp:RadioButton>
                </td>
            </tr>
        </table>
        <table class="font" id="Table4" cellspacing="0" cellpadding="0" width="600" border="0">
            <tr>
                <td>
                    <font face="新細明體">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:Button ID="bt_search" runat="server" Text="查詢"></asp:Button><asp:Button ID="bt_reset" runat="server" Text="重新設定"></asp:Button></font>
                </td>
            </tr>
        </table>
        <table class="font" id="Table5" cellspacing="0" cellpadding="0" width="600" border="0">
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Height="224px" Width="600px" ShowFooter="True" AutoGenerateColumns="False" AllowPaging="True">
                        <FooterStyle BackColor="GreenYellow"></FooterStyle>
                        <AlternatingItemStyle BackColor="White"></AlternatingItemStyle>
                        <ItemStyle BackColor="#FFFDC7"></ItemStyle>
                        <HeaderStyle HorizontalAlign="Center" BackColor="#999900"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn DataField="Name" HeaderText="轄區" FooterText="合計"></asp:BoundColumn>
                            <asp:ButtonColumn DataTextField="Teacher_Count" HeaderText="老師數目">
                                <ItemStyle CssClass="class_link"></ItemStyle>
                            </asp:ButtonColumn>
                            <asp:TemplateColumn Visible="False" HeaderText="功能">
                                <HeaderStyle Width="100px"></HeaderStyle>
                                <ItemTemplate>
                                    <asp:Button runat="server" Text="詳細" CommandName="Company_Seqno" CausesValidation="false" ID="Button3"></asp:Button>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                    <div align="center">
                        <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                    </div>
                </td>
            </tr>
        </table>
        </FONT>
    <p>
        <font face="新細明體">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font>
        <asp:Label ID="NoData" runat="server" CssClass="font"></asp:Label>
        </p>
    </form>
</body>
</html>
