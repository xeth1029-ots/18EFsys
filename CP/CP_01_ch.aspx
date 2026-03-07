<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_01_ch.aspx.vb" Inherits="WDAIIP.CP_01_ch" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>職類班級選擇</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function chkdata() {
            var msg = ''
            if (document.getElementById('CyclType').value != '' && !isUnsignedInt(document.getElementById('CyclType').value)) msg += '期別請輸入數字\n'
            if (msg != '') {
                window.alert(msg);
                return false;
            }
        }

        function ClearTMID() {
            //debugger;
            document.getElementById('TB_career_id').value = '';
            document.getElementById('trainValue').value = '';
            //document.getElementById('ClassID').value='';
            //document.getElementById('CyclType').value='';
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <div style="overflow-y: auto; height: 630px;">
            <%--<table id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0"><tr><td></td></tr></table>--%>
            <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                <tr>
                    <td width="20%" class="bluecol">班別代碼</td>
                    <td width="30%" class="whitecol">
                        <asp:TextBox ID="ClassID" runat="server" Columns="15" MaxLength="30"></asp:TextBox></td>
                    <td width="20%" class="bluecol">訓練職類</td>
                    <td width="30%" class="whitecol">
                        <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                        <input id="trainValue" type="hidden" name="trainValue" runat="server" />
                        <input onclick="wopen('../Common/TrainJob.aspx?field=TB_career_id&amp;TMID=' + document.getElementById('trainValue').value, 'TrainJob', 700, 550, 0);" type="button" value="..." />
                        <input id="Button1" onclick="ClearTMID();" type="button" value="清除" name="Button1" />
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">訓練時段</td>
                    <td class="whitecol" width="30%">
                        <asp:DropDownList ID="HourRan" runat="server"></asp:DropDownList></td>
                    <td class="bluecol">期別</td>
                    <td class="whitecol" width="30%">
                        <asp:TextBox ID="CyclType" runat="server" Columns="5" MaxLength="5" Width="30%"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol">班級名稱</td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="ClassCName" runat="server" Columns="50" Width="50%" MaxLength="100"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol">班級範圍</td>
                    <td class="whitecol" colspan="3">
                        <asp:RadioButtonList ID="ClassRound" runat="server" RepeatDirection="Horizontal" CssClass="font">
                            <asp:ListItem Value="開訓二週前" Selected="True">開訓二週前</asp:ListItem>
                            <asp:ListItem Value="已開訓">已開訓</asp:ListItem>
                            <asp:ListItem Value="已結訓">已結訓</asp:ListItem>
                            <asp:ListItem Value="未開訓">未開訓</asp:ListItem>
                            <asp:ListItem Value="全部">全部</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td colspan="4">
                        <div align="center" class="whitecol">
                            <asp:Button ID="search_but" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td colspan="4">
                        <div align="center">
                            <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                        </div>
                    </td>
                </tr>
            </table>
            <div style="max-height: 380px; overflow-y: auto;">
                <table id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                    <tr>
                        <td>
                            <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" AllowPaging="True" AutoGenerateColumns="False" Width="100%" CellPadding="8">
                                <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                <%--<ItemStyle BackColor="#FFECEC"></ItemStyle>--%>
                                <%--現在這版不要用紅色了所以註解掉上面那行--%>
                                <Columns>
                                    <asp:TemplateColumn>
                                        <HeaderStyle CssClass="head_navy" Width="10%"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        <ItemTemplate>
                                            <input id="radio1" value='<%# DataBinder.Eval(Container.DataItem, "OCID")%>' type="radio" name="class1">
                                        </ItemTemplate>
                                    </asp:TemplateColumn>
                                    <asp:BoundColumn DataField="TrainName2" HeaderText="訓練職類">
                                        <HeaderStyle HorizontalAlign="Center" CssClass="head_navy" Width="20%"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="ClassID" HeaderText="班別代碼">
                                        <HeaderStyle HorizontalAlign="Center" CssClass="head_navy" Width="12%"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="CLASSCNAME2" HeaderText="班級名稱">
                                        <HeaderStyle HorizontalAlign="Center" CssClass="head_navy" Width="22%"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="STDate" HeaderText="開訓日期">
                                        <HeaderStyle HorizontalAlign="Center" CssClass="head_navy" Width="12%"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    </asp:BoundColumn>
                                    <%--<asp:BoundColumn DataField="IsApplic2" HeaderText="志願班別" Visible="false">
                                        <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    </asp:BoundColumn>--%>
                                    <asp:BoundColumn DataField="HOURRANNAME" HeaderText="訓練時段">
                                        <HeaderStyle HorizontalAlign="Center" CssClass="head_navy" Width="12%"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    </asp:BoundColumn>
                                </Columns>
                                <PagerStyle Visible="False"></PagerStyle>
                            </asp:DataGrid>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <div align="center">
                                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                            </div>
                        </td>
                    </tr>
                </table>
            </div>
            <div align="center" class="whitecol">
                <asp:Button ID="BTN_send" runat="server" Text="送出" CssClass="asp_button_M"></asp:Button>
            </div>
        </div>
    </form>
</body>
</html>
