<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="EXAM_03_ch.aspx.vb" Inherits="WDAIIP.EXAM_03_ch" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>職類班級選擇</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script src="../../js/common.js"></script>
    <script language="javascript">
        function chkdata() {
            var msg = ''
            if (document.getElementById('CyclType').value != '' && !isUnsignedInt(document.getElementById('CyclType').value)) {
                msg += '期別請輸入數字\n';
            }
            if (msg != '') {
                window.alert(msg);
                return false;
            }
        }

        function returnNum() {
            window.opener.form1.TMID1.value = document.form1.class1.value;
        }

        function ClearTMID() {
            document.getElementById('TB_career_id').value = '';
            document.getElementById('trainValue').value = '';
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch" id="Table1" cellspacing="1" cellpadding="1">
                        <tr id="YearsTR" runat="server">
                            <td class="bluecol" style="width: 20%">年度
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:DropDownList ID="Years" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" style="width: 20%">班別代碼
                            </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="ClassID" runat="server" Columns="15" Width="40%"></asp:TextBox>
                            </td>
                            <td class="bluecol" style="width: 20%">訓練職類
                            </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="TB_career_id" runat="server" Columns="15" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="...">
                                <input id="Button1" onclick="ClearTMID();" type="button" value="清除" name="Button1">
                                <input id="trainValue" type="hidden" name="trainValue" runat="server">
                                <input id="jobValue" type="hidden" name="jobValue" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練時段
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="HourRan" runat="server">
                                </asp:DropDownList>
                            </td>
                            <td class="bluecol">期別
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="CyclType" runat="server" Columns="5" Width="40%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">班級範圍
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="ClassRound" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="開訓二週前" Selected="True">開訓二週前</asp:ListItem>
                                    <asp:ListItem Value="已開訓">已開訓</asp:ListItem>
                                    <asp:ListItem Value="已結訓">已結訓</asp:ListItem>
                                    <asp:ListItem Value="未開訓">未開訓</asp:ListItem>
                                    <asp:ListItem Value="全部">全部</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td>
                                <p align="center" class="whitecol">
                                    <asp:Button ID="search_but" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                </p>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <p align="center">
                                </p>
                                <p align="center">
                                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                                </p>
                            </td>
                        </tr>
                    </table>
                    <table id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" AllowPaging="True" AutoGenerateColumns="False" Width="100%" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="5%" />
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <input id="radio1" value='<%# DataBinder.Eval(Container.DataItem,"OCID")%>' type="radio" name="class1">
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="TrainName" HeaderText="訓練職類">
                                            <HeaderStyle HorizontalAlign="Center" Width="19%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassID" HeaderText="班級代碼">
                                            <HeaderStyle HorizontalAlign="Center" Width="19%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassCName" HeaderText="班級名稱">
                                            <HeaderStyle HorizontalAlign="Center" Width="19%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDate" HeaderText="開訓日期" DataFormatString="{0:d}">
                                            <HeaderStyle HorizontalAlign="Center" Width="19%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="IsApplic" HeaderText="志願班別">
                                            <HeaderStyle HorizontalAlign="Center" Width="19%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <p align="center">
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                </p>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <p align="center" class="whitecol">
                                    <asp:Button ID="send" runat="server" Text="送出" CssClass="asp_button_M"></asp:Button>
                                </p>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
