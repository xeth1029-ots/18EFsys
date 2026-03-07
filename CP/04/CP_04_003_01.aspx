<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_04_003_01.aspx.vb" Inherits="WDAIIP.CP_04_003_01" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>學員資料</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <style type="text/css">
        body { overflow-y: auto; }
    </style>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table id="Table2" cellspacing="0" cellpadding="0" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:Label ID="Label1" runat="server" CssClass="font">年度：</asp:Label>
                                <asp:Label ID="YearLabel" runat="server" CssClass="font"></asp:Label>
                            </td>
                            <td>
                                <asp:Label ID="Label2" runat="server" CssClass="font">訓練計畫：</asp:Label>
                                <asp:Label ID="TrainPlanLabel" runat="server" CssClass="font"></asp:Label>
                            </td>
                            <td>
                                <asp:Label ID="Label4" runat="server" CssClass="font">班別名稱</asp:Label>
                                <asp:Label ID="ClassNameLabel" runat="server" CssClass="font"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td align="right" colspan="3">
                                <input id="Button1" type="button" value="檢視此班級的學員參訓歷史" name="Button1" runat="server" class="button_b_L"></td>
                        </tr>
                        <tr>
                            <td colspan="3">
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AutoGenerateColumns="False" AllowPaging="True" CssClass="font" DataKeyField="IDNO" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                                    <HeaderStyle HorizontalAlign="Center" CssClass="head_navy" Width="20%"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn DataField="StudentID" HeaderText="學號"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="Name" HeaderText="姓名"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="Sex" HeaderText="性別"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="Birthday" HeaderText="出生日期" ></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:Button ID="Button3" runat="server" Text="學員參訓歷史" CommandName="list" CssClass="asp_button_M"></asp:Button>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn Visible="False" DataField="StudentID" HeaderText="StudentID"></asp:BoundColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                       <%-- <tr>
                            <td colspan="3">
                                <div align="center">
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                </div>
                            </td>
                        </tr>--%>
                    </table>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                    </div>
                </td>
            </tr>
             <tr>
                <td align="center">
                    <input id="btnClose2" type="button" value="關閉視窗" name="Button1" runat="server" class="button_b_M"></td>
            </tr>

        </table>
    </form>
</body>
</html>
