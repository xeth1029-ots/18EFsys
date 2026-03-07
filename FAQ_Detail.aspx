<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="FAQ_Detail.aspx.vb" Inherits="WDAIIP.FAQ_Detail" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>FAQ_Detail</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="css/style.css" type="text/css" rel="stylesheet">
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" width="600">
            <tr>
                <td>
                    <font class="font" size="2">首頁&gt;&gt;</font><font color="#800000">問題集</font>
                </td>
            </tr>
        </table>
        <p>
        </p>
        <table id="Table1" cellspacing="0" cellpadding="0" width="740" border="0">
            <tr>
                <td class="font" width="90%">功能項目：
				<asp:Label ID="FunctionName" runat="server" CssClass="font"></asp:Label>
                </td>
                <td class="font">共
				<asp:Label ID="Num" runat="server" CssClass="font"></asp:Label>&nbsp;筆
                </td>
            </tr>
        </table>
        <table width="740" cellpadding="0" cellspacing="0" border="0">
            <tr class="bluecol">
                <td align="center" width="100" class="bluecol">序號
                </td>
                <td align="center" class="bluecol">問與答
                </td>
            </tr>
        </table>
        <table id="Table5" cellspacing="0" cellpadding="0" width="740" align="left">
            <tbody>
                <tr>
                    <td>
                        <asp:DataGrid ID="DataGrid1" runat="server" ShowHeader="False" AllowPaging="True" AutoGenerateColumns="False" Width="100%">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <Columns>
                                <asp:TemplateColumn HeaderText="問題描述">
                                    <ItemTemplate>
                                        <table border="0" class="font" width="740">
                                            <tr>
                                                <td width="100" class="bluecol">
                                                    <div align="center">
                                                        Q
													<asp:Label ID="Label4" runat="server" Text='<%# (Me.DataGrid1.PageSize * Me.DataGrid1.CurrentPageIndex) + Container.ItemIndex + 1 %>'>：
                                                    </asp:Label>
                                                    </div>
                                                </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="Label3" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.Question") %>'>
                                                    </asp:Label>
                                                    (提問單位:
												<asp:Label ID="Label1" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.PostUnit") %>'>
                                                </asp:Label>
                                                    <asp:Label ID="Label5" runat="server" Text='<%# Formatdatetime(DataBinder.Eval(Container, "DataItem.PostDate"),2) %>'>
                                                    </asp:Label>)
                                                </td>
                                            </tr>
                                            <tr height="30">
                                                <td class="whitecol">
                                                    <div align="center">
                                                    </div>
                                                </td>
                                                <td>
                                                    <asp:Label ID="Label2" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.Deal") %>'>
                                                    </asp:Label>
                                                </td>
                                            </tr>
                                        </table>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                            </Columns>
                            <PagerStyle Visible="False"></PagerStyle>
                        </asp:DataGrid>
                        <div align="center">
                            <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                        </div>
                    </td>
                    <td>
                        <font face="新細明體"></font>
                    </td>
                </tr>
                <tr>
                    <td>
                        <div align="center">
                            <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                        </div>
                    </td>
                    <td>
                        <font face="新細明體"></font>
                    </td>
                </tr>
                <tr>
                    <td align="center">
                        <input id="Button1" type="button" value="回上一頁" name="Button1" runat="server" class="button_b_M">
                    </td>
                    <td align="center"></td>
                </tr>
            </tbody>
        </table>
        <input id="AcceptSearch" type="hidden" name="AcceptSearch" runat="server">
    </form>
</body>
</html>
