<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_02_ch1.aspx.vb" Inherits="WDAIIP.SD_02_ch1" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD html 4.0 Transitional//EN">
<html>
<head>
    <title>職類班級選擇</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        //if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        function chkdata() {
            var msg = ''

            if (document.form1.syear.selectedIndex == 0) {
                msg = msg + '請選擇年度!\n';
            }
            //if (document.form1.trainValue.value=='') msg=msg+'請選擇訓練職類!';
            if (msg != '') {
                window.alert(msg);
                return false;
            }
        }

        function returnNum() {
            window.opener.form1.TMID1.value = document.form1.class1.value;
        }
    </script>
    
</head>
<body>
    <form id="form1" method="post" runat="server">
        <div style="overflow-y: auto; height: 630px;">
            <table class="font" id="table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                <tr>
                    <td width="20%" class="bluecol_need">年度</td>
                    <td width="30%" class="whitecol">
                        <asp:DropDownList ID="syear" runat="server"></asp:DropDownList></td>
                    <td width="20%" class="bluecol">訓練職類</td>
                    <td width="30%" class="whitecol">
                        <asp:TextBox ID="tb_career_id" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                        <input id="trainvalue" type="hidden" name="trainvalue" runat="server" />
                        <input onclick="opentrain(document.getelementbyid('trainvalue').value);" type="button" value="..." class="asp_button_Mini" />
                    </td>
                </tr>
                <tr>
                    <td width="20%" class="bluecol">訓練時段</td>
                    <td width="30%" class="whitecol">
                        <asp:DropDownList ID="hourran" runat="server"></asp:DropDownList></td>
                    <td width="20%" class="bluecol">期別</td>
                    <td width="30%" class="whitecol">
                        <asp:TextBox ID="cycltype" runat="server" Columns="5" MaxLength="5" Width="30%"></asp:TextBox></td>
                </tr>
                <tr>
                    <td width="20%" class="bluecol">班級名稱</td>
                    <td width="80%" class="whitecol" colspan="3">
                        <asp:TextBox ID="classcname" runat="server" Columns="50" Width="60%" MaxLength="100"></asp:TextBox></td>
                </tr>
                <tr>
                    <td width="20%" class="bluecol">班別代碼</td>
                    <td width="80%" class="whitecol" colspan="3">
                        <asp:TextBox ID="classid" runat="server" Columns="15" Width="40%" MaxLength="30"></asp:TextBox></td>
                </tr>
                <tr>
                    <td width="20%" class="bluecol">班級範圍</td>
                    <td class="whitecol" colspan="3">
                        <asp:RadioButtonList ID="classround" runat="server" RepeatDirection="horizontal" CssClass="font" RepeatLayout="flow">
                            <asp:ListItem Value="開訓二週前" Selected="true">開訓二週前</asp:ListItem>
                            <asp:ListItem Value="已開訓">已開訓</asp:ListItem>
                            <asp:ListItem Value="已結訓">已結訓</asp:ListItem>
                            <asp:ListItem Value="未開訓">未開訓</asp:ListItem>
                            <asp:ListItem Value="全部">全部</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <%--<tr>
                <td width="100%" colspan="4" align="center" class="whitecol"></td>
            </tr>--%>
            </table>
            <div align="center" class="whitecol">
                <asp:Button ID="search_but" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
            </div>
            <div align="center">
                <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
            </div>
            <table id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                <tr>
                    <td>
                        <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Width="100%" AutoGenerateColumns="false" AllowPaging="true">
                            <AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
                            <HeaderStyle CssClass="head_navy" />
                            <Columns>
                                <asp:TemplateColumn>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    <ItemTemplate>
                                        <input id="radio1" value='<%#DataBinder.Eval(Container.DataItem, "ocid")%>' type="radio" name="class1">
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:BoundColumn DataField="TrainName2" HeaderText="訓練職類">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="left"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="classid" HeaderText="班別代碼">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="CLASSCNAME2" HeaderText="班級名稱">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="left"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="stdate" HeaderText="開訓日期">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="isapplic" HeaderText="志願班別" Visible="false">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                            </Columns>
                            <PagerStyle Visible="false"></PagerStyle>
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
                <tr>
                    <td align="center">
                        <asp:Button ID="send" runat="server" Text="送出" CssClass="asp_button_M"></asp:Button></td>
                </tr>
            </table>
        </div>
        <asp:HiddenField ID="Hid_SSSDTRID" runat="server" />
    </form>
</body>
</html>
