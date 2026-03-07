<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="GetClass.aspx.vb" Inherits="WDAIIP.GetClass" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>選擇班別</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function GetOCID(flag, OCIDValue) {
            if (flag) {
                if (document.form1.OCID.value != '') { document.form1.OCID.value += ',' }
                document.form1.OCID.value += OCIDValue;
            }
            else {
                if (document.form1.OCID.value.indexOf(',' + OCIDValue) != -1) {
                    document.form1.OCID.value = document.form1.OCID.value.replace(',' + OCIDValue, '')
                }
                else if (document.form1.OCID.value.indexOf(OCIDValue + ',') == 0) {
                    document.form1.OCID.value = document.form1.OCID.value.replace(OCIDValue + ',', '')
                }
                else if (document.form1.OCID.value.indexOf(OCIDValue) == 0) {
                    document.form1.OCID.value = document.form1.OCID.value.replace(OCIDValue, '')
                }
            }
        }
        function select_all(flag) {
            var mytable = document.getElementById('DataGrid1');

            for (var i = 1; i < mytable.rows.length; i++) {
                var mycheck = mytable.rows[i].cells[0].children[0];
                mycheck.checked = flag;
                GetOCID(flag, mycheck.value);
            }
        }

        function ClearTMID() {
            document.getElementById('TB_career_id').value = '';
            document.getElementById('trainValue').value = '';
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td align="center">
                    <table class="table_nw" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td width="80" class="bluecol">訓練職類
                            </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Width="310px"></asp:TextBox>
                                <input id="trainValue" type="hidden" name="trainValue" runat="server" />
                                <input onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." class="asp_button_Mini" />
                                <input id="Button3" onclick="ClearTMID();" type="button" value="清除" name="Button1" class="asp_button_S">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label>
                            </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()"></asp:TextBox>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="asp_button_Mini" />
                                <input id="cjobValue" type="hidden" name="cjobValue" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練時段
                            </td>
                            <td width="200" class="whitecol">
                                <asp:DropDownList ID="TPeriod" runat="server">
                                </asp:DropDownList>
                            </td>
                            <td width="80" class="bluecol">期別
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="CyclType" runat="server" Columns="5"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">班級代碼
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="ClassID" runat="server" Columns="15"></asp:TextBox>
                            </td>
                            <td class="bluecol">班級名稱
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="ClassName" runat="server" Columns="20" MaxLength="20"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">班級範圍
                            </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="ClassRound" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="開訓二週前" Selected="True">開訓二週前</asp:ListItem>
                                    <asp:ListItem Value="已開訓">已開訓</asp:ListItem>
                                    <asp:ListItem Value="已結訓">已結訓</asp:ListItem>
                                    <asp:ListItem Value="未開訓">未開訓</asp:ListItem>
                                    <asp:ListItem Value="全部">全部</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="4" class="whitecol">
                                <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" AllowPaging="True" AutoGenerateColumns="False" Width="100%">
                                    <AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <HeaderTemplate>
                                                <input id="Checkbox2" type="checkbox" runat="server">
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <input id="Checkbox1" type="checkbox" runat="server">
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="TrainName" HeaderText="訓練職類">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassID" HeaderText="班級代碼">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="CLASSCNAME2" HeaderText="班級名稱">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDate" HeaderText="開訓日期">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="IsApplic" HeaderText="志願班別">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <asp:Button ID="Button2" runat="server" Text="送出" CssClass="asp_button_S"></asp:Button>
                                <input id="OCID" type="hidden" runat="server" />
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
