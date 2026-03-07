<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_04_003_01_History.aspx.vb" Inherits="WDAIIP.CP_04_003_01_History" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>學員參訓歷史</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
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
                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" AutoGenerateColumns="False" Width="100%" PageSize="20">
                                    <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號">
                                            <HeaderStyle HorizontalAlign="Center" Width="30px"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Name" HeaderText="姓名">
                                            <HeaderStyle Width="60px"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="IDNO" HeaderText="身分證號碼">
                                            <HeaderStyle Width="65px"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Birthday" HeaderText="出生日期"  >
                                            <HeaderStyle Width="55px"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <%--<asp:BoundColumn DataField="DistName" HeaderText="轄區&lt;BR&gt;中心">--%>
                                        <asp:BoundColumn DataField="DistName" HeaderText="轄區&lt;BR&gt;分署">
                                            <HeaderStyle HorizontalAlign="Center" Width="50px"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構">
                                            <HeaderStyle Width="100px"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TMID" HeaderText="訓練職類">
                                            <HeaderStyle Width="100px"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassName" HeaderText="班別">
                                            <HeaderStyle Width="100px"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="THours" HeaderText="受訓&lt;BR&gt;時數">
                                            <HeaderStyle Width="30px"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TRound" HeaderText="受訓期間">
                                            <HeaderStyle HorizontalAlign="Center" Width="60px"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="SkillName" HeaderText="技能檢定">
                                            <HeaderStyle Width="100px"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TFlag" HeaderText="訓練&lt;BR&gt;狀態">
                                            <HeaderStyle Width="30px"></HeaderStyle>
                                        </asp:BoundColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></td>
            </tr>
            <tr>
                <td align="center">
                    <input id="btnClose2" type="button" value="關閉視窗" name="Button1" runat="server" class="button_b_M"></td>
            </tr>
        </table>
    </form>
</body>
</html>
