<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_01_003.aspx.vb" Inherits="WDAIIP.SYS_01_003" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>帳號審核</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;帳號維護管理&gt;&gt;帳號審核</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="FrameTable1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td align="center">
                    <%--<table class="font" width="100%"><tr><td class="font">首頁&gt;&gt;系統管理&gt;&gt;帳號維護管理&gt;&gt;帳號審核</td></tr></table>--%>
                    <table id="TB_Condition" runat="server" class="table_sch" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" style="width: 20%">申請種類
                            </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:RadioButtonList ID="ApplyType" runat="server" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="Account" Selected="True">帳號申請審核</asp:ListItem>
                                    <asp:ListItem Value="Plan">計畫申請審核</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                            <td class="bluecol" style="width: 20%">審核結果
                            </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:DropDownList ID="Resultsrh" runat="server">
                                    <asp:ListItem Value="Y">通過</asp:ListItem>
                                    <asp:ListItem Value="N">不通過</asp:ListItem>
                                    <asp:ListItem Value="X" Selected="True">尚未審核</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">帳號
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="nameid" runat="server" MaxLength="15" Width="40%"></asp:TextBox>
                            </td>
                            <td class="bluecol">姓名
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="namefield" runat="server" MaxLength="15" Width="40%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練機構
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="TBplan" Width="40%" runat="server" MaxLength="25" onfocus="this.blur()"></asp:TextBox>
                                <input id="choice_button" type="button" value="選擇" name="choice_button" runat="server" class="asp_button_M">
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td class="whitecol">
                                <div align="center">
                                    <asp:Button ID="but_search" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                </div>
                            </td>
                        </tr>
                    </table>
                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr id="TR_Acc1" runat="server">
                            <td style="color: red; font-size: 12pt" align="left">* 說明欄中未審核的計畫數量，不一定等同於權限內可審核的計畫數量。
                            </td>
                        </tr>
                        <tr id="TR_Account" runat="server">
                            <td align="center">
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn DataField="Account" HeaderText="帳號" HeaderStyle-Width="14%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="Name" HeaderText="姓名" HeaderStyle-Width="14%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="IDNO" HeaderText="身分證號" HeaderStyle-Width="14%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="RoleName" HeaderText="角色" HeaderStyle-Width="14%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="OrgName" SortExpression="OrgID" HeaderText="所屬單位" HeaderStyle-Width="16%"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="說明" HeaderStyle-Width="14%">
                                            <ItemTemplate>
                                                <asp:LinkButton ID="AccountNote" runat="server" ForeColor="Blue" CommandName="AuditPlan"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="功能" HeaderStyle-Width="14%">
                                            <HeaderStyle CssClass="whitecol" />
                                            <ItemStyle CssClass="whitecol" />
                                            <HeaderTemplate>
                                                <asp:DropDownList ID="AuditAllAccount" runat="server" CssClass="whitecol">
                                                    <asp:ListItem></asp:ListItem>
                                                    <asp:ListItem Value="Y">通過</asp:ListItem>
                                                    <asp:ListItem Value="N">不通過</asp:ListItem>
                                                </asp:DropDownList>
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:DropDownList ID="AuditListAccount" runat="server" CssClass="whitecol">
                                                    <asp:ListItem></asp:ListItem>
                                                    <asp:ListItem Value="Y">通過</asp:ListItem>
                                                    <asp:ListItem Value="N">不通過</asp:ListItem>
                                                </asp:DropDownList>
                                                <asp:Label ID="AuditAccountStatus" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle NextPageText="  下一頁&amp;gt;&amp;gt;" PrevPageText="&amp;lt;&amp;lt;上一頁  " HorizontalAlign="center" Mode="NumericPages"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr id="TR_AuditAccount" runat="server">
                            <td align="center" class="whitecol">
                                <asp:Button ID="AuditAccont" runat="server" Text="確認" CssClass="asp_button_M"></asp:Button>&nbsp;
                            <asp:Button ID="Btn_Cancel" runat="server" Text="取消" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                        <tr id="TR_Plan" runat="server">
                            <td align="center">
                                <asp:DataGrid ID="Datagrid2" runat="server" Width="100%" AutoGenerateColumns="False" AllowPaging="True" CssClass="font" CellPadding="8">
                                    <AlternatingItemStyle HorizontalAlign="Center" BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn DataField="Account" HeaderText="申請人帳號" HeaderStyle-Width="18%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="Name" HeaderText="申請人姓名" HeaderStyle-Width="18%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="所屬單位" HeaderStyle-Width="18%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="PlanName" HeaderText="申請計畫代碼" HeaderStyle-Width="18%"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="說明" HeaderStyle-Width="18%">
                                            <ItemTemplate>
                                                <asp:LinkButton ID="PlanNote" runat="server" ForeColor="Blue" CommandName="AuditAccount"></asp:LinkButton>
                                                <asp:Label ID="OthPlanNote" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="功能" HeaderStyle-Width="10%">
                                            <HeaderTemplate>
                                                <asp:DropDownList ID="AuditAllPlan" runat="server">
                                                    <asp:ListItem></asp:ListItem>
                                                    <asp:ListItem Value="Y">通過</asp:ListItem>
                                                    <asp:ListItem Value="N">不通過</asp:ListItem>
                                                </asp:DropDownList>
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:DropDownList ID="AuditListPlan" runat="server">
                                                    <asp:ListItem></asp:ListItem>
                                                    <asp:ListItem Value="Y">通過</asp:ListItem>
                                                    <asp:ListItem Value="N">不通過</asp:ListItem>
                                                </asp:DropDownList>
                                                <asp:Label ID="AuditPlanStatus" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn Visible="False" DataField="PlanID"></asp:BoundColumn>
                                        <asp:BoundColumn Visible="False" DataField="DistID"></asp:BoundColumn>
                                        <asp:BoundColumn Visible="False" DataField="Shared"></asp:BoundColumn>
                                    </Columns>
                                    <PagerStyle HorizontalAlign="Right" Mode="NumericPages"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr id="TR_AuditPlan" runat="server">
                            <td align="center" class="whitecol">
                                <asp:Button ID="AuditPlan" runat="server" Text="確認" CssClass="asp_button_M"></asp:Button>&nbsp;
                            <asp:Button ID="Btn_Cancel2" runat="server" Text="取消" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                </td>
            </tr>
        </table>
        &nbsp;
    <input id="YearsValue" type="hidden" name="YearsValue" runat="server"><input id="DistValue" type="hidden" name="DistValue" runat="server">
        <input id="PlanIDValue" style="width: 56px; height: 22px" type="hidden" size="4" name="PlanIDValue" runat="server">
        <input id="RIDValue" style="width: 56px; height: 22px" type="hidden" size="4" name="RIDValue" runat="server">
        <input id="OrgIDValue" type="hidden" name="OrgIDValue" runat="server">
    </form>
</body>
</html>
