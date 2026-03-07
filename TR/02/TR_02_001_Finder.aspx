<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_02_001_Finder.aspx.vb" Inherits="WDAIIP.TR_02_001_Finder" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>引用事業單位</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function checkRadio(num) {
            var mytable = document.getElementById('DataGrid1');
            for (var i = 1; i < mytable.rows.length; i++) {
                var myradio = mytable.rows[i].cells[0].children[0];
                myradio.checked = (num == i) ? true : false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="500" border="0">
            <tr>
                <td align="center">
                    <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>首頁&gt;&gt;訓練與就業需求管理&gt;&gt;企業需求訪視表&gt;&gt;<font color="#990000">引用事業單位</font>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td width="80" bgcolor="#2aafc0">
                                <font face="新細明體">&nbsp;&nbsp;&nbsp; <font color="#ffffff">公司名稱</font></font>
                            </td>
                            <td bgcolor="#ecf7ff">
                                <font face="新細明體">
                                    <asp:TextBox ID="Uname" runat="server" Columns="15"></asp:TextBox></font>
                            </td>
                            <td width="80" bgcolor="#2aafc0">
                                <font face="新細明體">&nbsp;&nbsp;&nbsp; <font color="#ffffff">統一編號</font></font>
                            </td>
                            <td bgcolor="#ecf7ff">
                                <asp:TextBox ID="Intaxno" runat="server" Columns="15"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td bgcolor="#2aafc0">
                                <font face="新細明體">&nbsp;&nbsp;&nbsp; <font color="#ffffff">縣市</font></font>
                            </td>
                            <td bgcolor="#ecf7ff">
                                <font face="新細明體">
                                    <asp:TextBox ID="TBCity" runat="server" Columns="15"></asp:TextBox><input id="Button3" type="button" value="..." name="Button3" runat="server"><input id="city_code" style="width: 26px; height: 22px" type="hidden" name="city_code" runat="server"><input id="zip_code" style="width: 26px; height: 22px" type="hidden" name="zip_code" runat="server"></font>
                            </td>
                            <td bgcolor="#2aafc0">
                                <font face="新細明體"></font><font face="新細明體">&nbsp;&nbsp;&nbsp; <font color="#ffffff">保險證號</font></font>
                            </td>
                            <td bgcolor="#ecf7ff">
                                <asp:TextBox ID="Ubno" runat="server" Columns="15"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="4">
                                <font face="新細明體"></font><font face="新細明體">
                                    <asp:Button ID="Button1" runat="server" Text="查詢"></asp:Button></font>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" AutoGenerateColumns="False" Width="100%" CssClass="font" AllowPaging="True">
                                    <AlternatingItemStyle BackColor="White"></AlternatingItemStyle>
                                    <ItemStyle BackColor="#ECF7FF"></ItemStyle>
                                    <HeaderStyle ForeColor="White" BackColor="#2AAFC0"></HeaderStyle>
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="20px"></HeaderStyle>
                                            <ItemTemplate>
                                                <input id="Radio1" type="radio" value="Radio1" runat="server">
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn HeaderText="序號">
                                            <HeaderStyle Width="25px"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Uname" HeaderText="事業單位名稱"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="Intaxno" HeaderText="統一編號"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="Ename" HeaderText="負責人姓名"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="Addr" HeaderText="通訊地址"></asp:BoundColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" style="height: 18px">
                                <font face="新細明體">
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                </font>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <font face="新細明體">
                                    <asp:Button ID="Button2" runat="server" Text="送出"></asp:Button></font>
                            </td>
                        </tr>
                    </table>
                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
