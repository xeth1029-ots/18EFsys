<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_05_001_detail.aspx.vb" Inherits="WDAIIP.SYS_05_001_detail" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>上稿維護-FAQ_Detail</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        //重新調整頁面高度
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //parent.setMainFrameHeight();  //重新調整頁面高度
        if (window.top && window.top.setMainFrameHeight() != undefined) { window.top.setMainFrameHeight(); }
        //$(document).ready(function () { });
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <div style="width: 100%; padding: 0px; min-height: 860px;">
            <table id="Table1" cellspacing="0" cellpadding="0" width="100%" border="0">
                <tr>
                    <td class="font" width="90%">功能項目：<asp:Label ID="FunctionName" runat="server" CssClass="font"></asp:Label></td>
                    <td class="font">共<asp:Label ID="Num" runat="server" CssClass="font"></asp:Label>&nbsp;筆</td>
                </tr>
            </table>
            <table cellspacing="0" cellpadding="0" width="100%" border="0">
                <tr class="head_navy">
                    <td align="center" style="width: 10%">序號</td>
                    <td align="center" style="width: 70%">問與答</td>
                    <td align="center" style="width: 20%">功能</td>
                </tr>
            </table>
            <table id="Table5" cellspacing="0" cellpadding="0" width="100%" align="left">
                <tbody>
                    <tr>
                        <td>
                            <asp:DataGrid ID="DataGrid1" runat="server" OnEditCommand="ECmd" OnCancelCommand="CCmd" OnDeleteCommand="DCmd" OnUpdateCommand="UCmd" AutoGenerateColumns="False" ShowHeader="false" Width="100%" AllowPaging="True">
                                <AlternatingItemStyle BackColor="#F5F5F5" />
                                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                <Columns>
                                    <asp:BoundColumn Visible="False" DataField="SeqNO" HeaderText="序號" HeaderStyle-Width="10%"></asp:BoundColumn>
                                    <asp:TemplateColumn HeaderText="問題描述" FooterStyle-Width="80%">
                                        <ItemTemplate>
                                            <table class="font" width="100%" border="0">
                                                <tr>
                                                    <td width="10%" class="head_navy">
                                                        <div align="center">Q<asp:Label ID="SeqNOlabel" runat="server" Text="<%# (Me.DataGrid1.PageSize * Me.DataGrid1.CurrentPageIndex) + Container.ItemIndex + 1 %>">：</asp:Label></div>
                                                    </td>
                                                    <td>
                                                        <asp:Label ID="question" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.Question") %>'></asp:Label>(提問單位:
                                                        <asp:Label ID="postunit" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.PostUnit") %>'></asp:Label>
                                                        <asp:Label ID="postdate" runat="server" Text='<%# FormatDateTime(DataBinder.Eval(Container, "DataItem.PostDate"), 2) %>'></asp:Label>)
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td>
                                                        <div align="center"></div>
                                                    </td>
                                                    <td>
                                                        <asp:Label ID="answer" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.Deal") %>'></asp:Label></td>
                                                </tr>
                                            </table>
                                        </ItemTemplate>
                                        <EditItemTemplate>
                                            <table class="font" id="Table2" width="100%" border="0">
                                                <tr>
                                                    <td style="width: 10%">
                                                        <div align="center">Q<asp:Label ID="Label8" runat="server" Text="<%# (Me.DataGrid1.PageSize * Me.DataGrid1.CurrentPageIndex) + Container.ItemIndex + 1 %>">：</asp:Label></div>
                                                    </td>
                                                    <td>
                                                        <asp:TextBox ID="QuestionTextBox" runat="server" Width="90%" Text='<%# DataBinder.Eval(Container, "DataItem.Question") %>' TextMode="MultiLine"></asp:TextBox>
                                                        <br />
                                                        (提問<asp:TextBox ID="PostUnitTextBox" runat="server" Width="10%" Text='<%# DataBinder.Eval(Container, "DataItem.PostUnit") %>'></asp:TextBox>)
                                                        <asp:Label ID="Label1" runat="server">功能項目:</asp:Label>
                                                        <asp:DropDownList ID="FunctionList" runat="server"></asp:DropDownList>
                                                        <asp:RequiredFieldValidator ID="MustQuestion" runat="server" Display="Dynamic" ControlToValidate="QuestionTextBox" ErrorMessage="請輸入問題"></asp:RequiredFieldValidator>&nbsp;
                                                        <asp:RequiredFieldValidator ID="MustPostUnit" runat="server" ControlToValidate="PostUnitTextBox" ErrorMessage="請輸入單位"></asp:RequiredFieldValidator>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td>
                                                        <div align="center"></div>
                                                    </td>
                                                    <td>
                                                        <asp:TextBox ID="AnswerTextBox" runat="server" Width="100%" Text='<%# DataBinder.Eval(Container, "DataItem.Deal") %>' TextMode="MultiLine"></asp:TextBox>
                                                        <asp:RequiredFieldValidator ID="MustAnswer" runat="server" Display="Dynamic" ControlToValidate="AnswerTextBox" ErrorMessage="請輸入解答"></asp:RequiredFieldValidator>
                                                    </td>
                                                </tr>
                                            </table>
                                        </EditItemTemplate>
                                    </asp:TemplateColumn>
                                    <asp:EditCommandColumn ButtonType="PushButton" UpdateText="更新" CancelText="取消" EditText="編輯" ItemStyle-CssClass="whitecol2" ItemStyle-Font-Size="Small" ItemStyle-HorizontalAlign="Center"></asp:EditCommandColumn>
                                    <asp:ButtonColumn Text="刪除" ButtonType="PushButton" CommandName="Delete" ItemStyle-CssClass="whitecol2" ItemStyle-Font-Size="Small" ItemStyle-HorizontalAlign="Center"></asp:ButtonColumn>
                                </Columns>
                                <PagerStyle Visible="False"></PagerStyle>
                            </asp:DataGrid>
                            <div align="center">
                                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                            </div>
                        </td>
                    </tr>
                    <tr>
                        <td class="style1" colspan="2">
                            <div align="center">
                                <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                            </div>
                        </td>
                    </tr>
                    <tr>
                        <td align="center" colspan="2" class="whitecol">
                            <input id="Button1" type="button" value="回上一頁" name="Button1" runat="server" class="asp_button_M"></td>
                    </tr>
                </tbody>
            </table>
        </div>
    </form>
</body>
</html>
