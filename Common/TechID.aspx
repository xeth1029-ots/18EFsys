<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TechID.aspx.vb" Inherits="WDAIIP.TechID" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>請選擇老師</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <%--<LINK href="../style.css" type="text/css" rel="stylesheet">--%>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        function ReturnTechID(TechID, TechName) {
            opener.document.getElementById(getParamValue('ValueField')).value = TechID;
            opener.document.getElementById(getParamValue('TextField')).value = TechName;
            window.close();
        }

        function ReturnTechID2() {
            opener.document.getElementById(getParamValue('ValueField')).value = document.getElementById('TeachID').value;
            opener.document.getElementById(getParamValue('TextField')).value = document.getElementById('TeachName').value;
            window.close();
        }

        function OpenProMenu(num) {
            document.getElementById('State').value = num;
            if (num == 1) {
                document.getElementById('ProTR1').style.display = '';
                document.getElementById('ProTR2').style.display = 'none';
            }
            else {
                document.getElementById('ProTR1').style.display = 'none';
                document.getElementById('ProTR2').style.display = '';
            }
        }

        function SelectTechID(Flag, TechID, TechName) {
            if (Flag) {
                if (document.getElementById('TeachID').value == '') {
                    document.getElementById('TeachID').value = TechID;
                    document.getElementById('TeachName').value = TechName;
                }
                else {
                    document.getElementById('TeachID').value += ',' + TechID;
                    document.getElementById('TeachName').value += ',' + TechName;
                }
            }
            else {
                if (document.getElementById('TeachID').value.indexOf(',' + TechID + ',') != -1) {
                    document.getElementById('TeachID').value = document.getElementById('TeachID').value.replace(',' + TechID, '')
                    document.getElementById('TeachName').value = document.getElementById('TeachName').value.replace(',' + TechName, '')
                }
                else if (document.getElementById('TeachID').value.indexOf(',' + TechID) != -1) {
                    document.getElementById('TeachID').value = document.getElementById('TeachID').value.replace(',' + TechID, '')
                    document.getElementById('TeachName').value = document.getElementById('TeachName').value.replace(',' + TechName, '')
                }
                else if (document.getElementById('TeachID').value.indexOf(TechID + ',') != -1) {
                    document.getElementById('TeachID').value = document.getElementById('TeachID').value.replace(TechID + ',', '')
                    document.getElementById('TeachName').value = document.getElementById('TeachName').value.replace(TechName + ',', '')
                }
                else if (document.getElementById('TeachID').value.indexOf(TechID) != -1) {
                    document.getElementById('TeachID').value = document.getElementById('TeachID').value.replace(TechID, '')
                    document.getElementById('TeachName').value = document.getElementById('TeachName').value.replace(TechName, '')
                }
            }
        }
    </script>
    <%--<style type="text/css"> .auto-style1 { height: 22px; } </style>--%>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr id="ProTR1" runat="server">
                <td>
                    <table class="table_nw" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="bluecol" style="width: 20%">講師代碼： </td>
                            <td class="whitecol" style="width: 30%"><asp:TextBox ID="TeacherID" runat="server"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol" style="width: 20%">講師姓名： </td>
                            <td class="whitecol" style="width: 30%"><asp:TextBox ID="TeachCName" runat="server"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol" style="width: 20%">內外聘： </td>
                            <td class="whitecol" style="width: 80%">
                                <asp:DropDownList ID="KindEngage1" runat="server">
                                    <asp:ListItem Value="%">全部</asp:ListItem>
                                    <asp:ListItem Value="1">內</asp:ListItem>
                                    <asp:ListItem Value="2">外</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="2" class="whitecol">
                                <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                <input onclick="ReturnTechID('', '')" type="button" value="清除" class="asp_button_M">
                            </td>
                        </tr>
                    </table>
                    <asp:HyperLink ID="Close" runat="server" ForeColor="Blue">關閉進階搜尋</asp:HyperLink>
                </td>
            </tr>
            <tr id="ProTR2" runat="server">
                <td align="left">
                    <table class="table_nw" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="bluecol" style="width: 20%">內外聘： </td>
                            <td class="whitecol" style="width: 80%">
                                <asp:DropDownList ID="KindEngage" runat="server" AutoPostBack="True">
                                    <asp:ListItem Value="%">全部</asp:ListItem>
                                    <asp:ListItem Value="1">內</asp:ListItem>
                                    <asp:ListItem Value="2">外</asp:ListItem>
                                </asp:DropDownList>
                                <input onclick="ReturnTechID('', '')" type="button" value="清除" class="asp_button_M">
                            </td>
                        </tr>
                    </table>
                    <asp:HyperLink ID="Open" runat="server" ForeColor="Blue">進階搜尋</asp:HyperLink>
                </td>
            </tr>
            <tr>
                <td align="center" class="whitecol">
                    <div style="overflow-y: auto; height: 400px;">
                        <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                            <AlternatingItemStyle BackColor="WhiteSmoke" />
                            <HeaderStyle HorizontalAlign="Center" CssClass="head_navy" />
                            <Columns>
                                <asp:TemplateColumn HeaderStyle-Width="10%">
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    <ItemTemplate>
                                        <input id="Radio1" type="radio" value="Radio1" runat="server"><input id="Checkbox1" type="checkbox" runat="server">
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:BoundColumn DataField="KindEngage" HeaderText="內外聘" HeaderStyle-Width="30%"></asp:BoundColumn>
                                <asp:BoundColumn DataField="TeacherID" HeaderText="講師代碼" HeaderStyle-Width="30%"></asp:BoundColumn>
                                <asp:BoundColumn DataField="TeachCName" HeaderText="講師姓名" HeaderStyle-Width="30%"></asp:BoundColumn>
                            </Columns>
                        </asp:DataGrid>
                    </div>
                    <input id="Button2" type="button" value="送出" name="Button2" runat="server" onclick="ReturnTechID2();" class="asp_button_M">
                </td>
            </tr>
        </table>
        <input id="State" type="hidden" value="0" runat="server">
        <input id="TeachID" type="hidden" name="TeachID" runat="server">
        <input id="TeachName" type="hidden" name="TeachName" runat="server">
        <input id="modifytype" type="hidden" name="modifytype" runat="server" size="1">
    </form>
</body>
</html>