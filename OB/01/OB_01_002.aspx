 

<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="OB_01_002.aspx.vb" Inherits="WDAIIP.OB_01_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>OB_01_002</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/common.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script language="JavaScript">

        function ClearOrg() {
            //debugger;
            var center = document.getElementById('center');
            var Org = document.getElementById('Org');
            var RIDValue = document.getElementById('RIDValue');
            var orgid_value = document.getElementById('orgid_value');
            center.value = '';
            RIDValue.value = '';
            orgid_value.value = '';
            //return false;
        }
			
    </script>
    <style type="text/css">
        .style1
        {
            height: 28px;
        }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
        <tr>
            <td>
                <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            <asp:Label ID="TitleLab1" runat="server"></asp:Label><asp:Label ID="TitleLab2" runat="server">
										<FONT face="新細明體">首頁&gt;&gt;委外訓練管理&gt;&gt;<font color="#990000">工作小組成員資料查詢</font></FONT>
                            </asp:Label><font color="#000000">(<font face="新細明體"><font color="#ff0000">*</font>為必填欄位</font>)</font>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <table class="table_sch" id="TableLay2" cellspacing="1" cellpadding="1">
                    <tr>
                        <td class="bluecol">
                            服務單位名稱
                        </td>
                        <td colspan="3" class="whitecol">
                            <asp:TextBox ID="center" runat="server"></asp:TextBox>
                            <input id="Org" type="button" value="..." name="Org" runat="server" 
                                class="button_b_Mini">
                            <input id="RIDValue" style="width: 32px; height: 22px" type="hidden" name="RIDValue"
                                runat="server"><input id="orgid_value" type="hidden" name="orgid_value"
                                    runat="server">
                            <input id="OrgClear" onclick="ClearOrg();" type="button" value="清除" name="OrgClear"
                                runat="server" class="button_b_S">
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            服務部門
                        </td>
                        <td colspan="3" class="whitecol">
                            <asp:TextBox ID="DeptName" runat="server" MaxLength="20" Width="300px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            成員姓名
                        </td>
                        <td colspan="3" class="whitecol">
                            <asp:TextBox ID="memName" runat="server" MaxLength="20" Width="150px"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td width="100" class="bluecol">
                            具備採購法證照
                        </td>
                        <td colspan="3" class="whitecol">
                            <asp:RadioButtonList ID="rblQualified" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                <asp:ListItem Value="N">無</asp:ListItem>
                                <asp:ListItem Value="B">基礎</asp:ListItem>
                                <asp:ListItem Value="A">進階</asp:ListItem>
                                <asp:ListItem Value="I" Selected="True">不區分</asp:ListItem>
                            </asp:RadioButtonList>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <p align="center">
                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label><asp:TextBox
                        ID="TxtPageSize" runat="server" MaxLength="2" Width="23px">10</asp:TextBox>
                    <asp:Button ID="btnQuery" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button><font
                        face="新細明體">&nbsp;</font>
                    <asp:Button ID="btnAdd" runat="server" Text="新增" CssClass="asp_button_S"></asp:Button></p>
            </td>
        </tr>
        <tr>
            <td>
                <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td align="center">
                            <p>
                                <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" width="100%"
                                    border="0" runat="server">
                                    <tr>
                                        <td>
                                            <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AllowSorting="True" PagerStyle-HorizontalAlign="Left"
                                                PagerStyle-Mode="NumericPages" AllowPaging="True" AutoGenerateColumns="False"
                                                CssClass="font">
                                                <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                <Columns>
                                                    <asp:BoundColumn HeaderText="序號">
                                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="msn" HeaderText="工作小組序號">
                                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="OrgName" HeaderText="服務單位">
                                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="DeptName" HeaderText="服務部門">
                                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="memName" HeaderText="成員姓名">
                                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:TemplateColumn HeaderText="功能">
                                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        <ItemTemplate>
                                                            <asp:Button ID="btn_edit" runat="server" Text="修改" CommandName="edit"></asp:Button>
                                                            <asp:Button ID="btn_del" runat="server" Text="刪除" CommandName="del"></asp:Button>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
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
                                </table>
                            </p>
                            <p>
                                <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label></p>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
