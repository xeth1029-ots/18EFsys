<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_03_025.aspx.vb" Inherits="WDAIIP.SYS_03_025" %>

<!DOCTYPE html PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>帳號群組與功能設定</title>
    <meta content="true" name="vs_showGrid" />
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function Show_SelectAll(tmpName1, tmpName2, tmpCnt) {
            if (document.getElementById(tmpName1)) {
                for (i = 0; i < tmpCnt; i++) {
                    if (document.getElementById(tmpName2.replace("ctl2", "ctl" + (i + 2)))) {
                        if (document.getElementById(tmpName2.replace("ctl2", "ctl" + (i + 2))).disabled == false) {
                            document.getElementById(tmpName2.replace("ctl2", "ctl" + (i + 2))).checked = document.getElementById(tmpName1).checked;
                        }
                    }
                }
            }
        }
    </script>
    <%--<style type="text/css">
        .auto-style1 { color: Black; text-align: right; padding: 4px 6px; background-color: #f1f9fc; border-right: 3px solid #49cbef; height: 29px; }
        .auto-style2 { color: #333333; padding: 4px; height: 29px; }
    </style>--%>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">
				            首頁&gt;&gt;系統管理&gt;&gt;功能權限管理&gt;&gt;帳號群組與功能設定
                    </asp:Label>
                </td>
            </tr>
        </table>

        <%--<table class="font" width="100%">
		<tr>
			<td class="font">
				首頁&gt;&gt;系統管理&gt;&gt;功能權限管理&gt;&gt;<font color="#990000"><font color="#990000">帳號群組與功能設定<font face="Times New Roman" size="2"></font></font></font>
			</td>
		</tr>
	</table>--%>
        <table class="table_nw" id="tb_Query" cellspacing="1" cellpadding="1" width="100%" runat="server">
            <tr>
                <td class="bluecol" width="20%">年度
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="list_Years" runat="server" AutoPostBack="true">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">轄區
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="list_DistID" runat="server" AutoPostBack="true">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">計畫代碼
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="list_PlanID" runat="server" AutoPostBack="true">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">訓練單位
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="list_OrgID" runat="server" AutoPostBack="true">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">帳號
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="txt_Account" runat="server" Width="20%"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">姓名
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="txt_Name" runat="server" Width="20%"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="whitecol" align="center" colspan="2">
                    <asp:Button ID="btn_Query" runat="server" Text="查詢" CssClass="asp_Export_M"></asp:Button>
                    &nbsp;
                    <asp:Button ID="btn_Print" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                </td>
            </tr>
            <tr>
                <td class="whitecol" align="center" colspan="2">
                    <div id="Div1" runat="server">
                        <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AllowPaging="true" AutoGenerateColumns="False" CssClass="font" CellPadding="8">
                            <AlternatingItemStyle BackColor="#F5F5F5" />
                            <HeaderStyle CssClass="head_navy" />
                            <Columns>
                                <asp:TemplateColumn HeaderText="序號">
                                    <HeaderStyle Width="6%" />
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:TemplateColumn>
                                <asp:TemplateColumn HeaderText="帳號">
                                    <HeaderStyle Width="14%" />
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    <ItemTemplate>
                                        <asp:Label ID="lab_Account" runat="server"></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:TemplateColumn HeaderText="姓名">
                                    <HeaderStyle Width="12%" />
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    <ItemTemplate>
                                        <asp:Label ID="lab_Name" runat="server"></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:TemplateColumn HeaderText="群組功能">
                                    <HeaderStyle Width="38%" />
                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    <ItemTemplate>
                                        <asp:Label ID="lab_Group" runat="server"></asp:Label>
                                        <input id="hide_GID" type="hidden" runat="server" name="hide_GID" />
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:TemplateColumn HeaderText="功能">
                                    <HeaderStyle Width="25%" />
                                    <ItemStyle HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                                    <ItemTemplate>
                                        <asp:LinkButton ID="btn_EditGroup" runat="server" Text="群組設定" CommandName="Group" CssClass="linkbutton"></asp:LinkButton>
                                        <asp:LinkButton ID="btn_EditOption" runat="server" Text="個人設定" CommandName="Option" CssClass="linkbutton"></asp:LinkButton>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                            </Columns>
                            <PagerStyle HorizontalAlign="Right" Mode="NumericPages"></PagerStyle>
                        </asp:DataGrid>
                    </div>
                    <asp:Label ID="lab_Msg" runat="server" ForeColor="Red" Visible="False">查無資料!</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="tb_Group" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td class="bluecol" style="width: 20%">帳號
                </td>
                <td class="whitecol" style="width: 30%">
                    <asp:Label ID="lab2UserID" runat="server"></asp:Label>
                </td>
                <td class="bluecol" style="width: 20%">姓名
                </td>
                <td class="whitecol" style="width: 30%">
                    <asp:Label ID="lab2Name" runat="server"></asp:Label>
                </td>
            </tr>
            <tr>
                <td class="bluecol">建檔單位</td>
                <td class="whitecol"><asp:DropDownList ID="ddlGDist2" runat="server" AutoPostBack="True"></asp:DropDownList></td>
                <td class="bluecol">群組階層</td>
                <td class="whitecol"><asp:DropDownList ID="ddlGroupType2" runat="server" AutoPostBack="True"></asp:DropDownList></td>
            </tr>
            <tr>
                <td align="center" colspan="4">
                    <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" CellPadding="8">
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <HeaderStyle CssClass="head_navy" />
                        <Columns>
                            <asp:TemplateColumn HeaderText="選用">
                                <HeaderStyle Width="5%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:CheckBox ID="chk_GroupValid" runat="server"></asp:CheckBox>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="建檔單位">
                                <HeaderStyle Width="20%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lab_DistName" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="群組階層">
                                <HeaderStyle Width="12%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lab_GroupType" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="群組名稱">
                                <HeaderStyle Width="28%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lab_GroupName" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="備註">
                                <HeaderStyle Width="15%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lab_GroupNote" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="使用計畫">
                                <HeaderStyle Width="15%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lplanName" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="功能">
                                <HeaderStyle Width="5%" />
                                <ItemStyle HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                                <ItemTemplate>
                                    <asp:LinkButton ID="btnPlanEdit" Text="使用計畫設定" runat="server" CssClass="linkbutton" CommandName="PlanEdit"></asp:LinkButton>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                        <PagerStyle HorizontalAlign="Right" Mode="NumericPages"></PagerStyle>
                    </asp:DataGrid><asp:Label ID="lab_Msg2" runat="server" ForeColor="Red" Visible="False">無群組可賦予!</asp:Label>
                </td>
            </tr>
            <tr>
                <td align="center" colspan="4" class="whitecol">
                    <asp:Button ID="btn_SaveGroup" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                    <asp:Button ID="btn_CancelGroup" runat="server" Text="取消" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
        </table>
        <table class="font" id="tb_GroupFun" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td class="bluecol" style="width: 20%">帳號
                </td>
                <td class="whitecol">
                    <asp:Label ID="lab3UserID" runat="server"></asp:Label>
                </td>
                <td class="bluecol">姓名
                </td>
                <td class="whitecol">
                    <asp:Label ID="lab3Name" runat="server"></asp:Label>
                </td>
            </tr>
            <tr>
                <td colspan="4">
                    <asp:DataGrid ID="DataGrid3" runat="server" AutoGenerateColumns="False" CssClass="font" Width="100%" CellPadding="8">
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <HeaderStyle CssClass="head_navy" />
                        <Columns>
                            <asp:TemplateColumn HeaderText="建檔單位" HeaderStyle-Width="20%" ItemStyle-HorizontalAlign="Center">
                                <ItemTemplate>
                                    <asp:Label ID="labUntName" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="群組階層" HeaderStyle-Width="12%" ItemStyle-HorizontalAlign="Center">
                                <ItemTemplate>
                                    <asp:Label ID="labGType" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="群組名稱" HeaderStyle-Width="20%" ItemStyle-HorizontalAlign="left">
                                <ItemTemplate>
                                    <asp:Label ID="labGName" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="設定備註說明" HeaderStyle-Width="38%">
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="labMemo" runat="server">已設定</asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="功能" HeaderStyle-Width="10%" ItemStyle-HorizontalAlign="Center" ItemStyle-Font-Size="Small">
                                <ItemTemplate>
                                    <asp:LinkButton ID="btnEdit" Text="設定" runat="server" CssClass="linkbutton" CommandName="Edit1"></asp:LinkButton>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                    </asp:DataGrid>
                </td>
            </tr>
            <tr>
                <td align="center" colspan="4" class="whitecol">
                    <asp:Button ID="btn_CancelOption1" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
        </table>
        <table class="table_nw" id="tb_Function" cellspacing="1" cellpadding="1" width="100%" runat="server">
            <tr>
                <td class="bluecol" style="width: 20%">帳號
                </td>
                <td class="whitecol" style="width: 30%">
                    <asp:Label ID="lab4UserID" runat="server"></asp:Label>
                </td>
                <td class="bluecol" style="width: 20%">姓名
                </td>
                <td class="whitecol" style="width: 30%">
                    <asp:Label ID="lab4Name" runat="server"></asp:Label>
                </td>
            </tr>
            <tr>
                <td class="bluecol" style="width: 20%">功能類別
                </td>
                <td class="whitecol" colspan="3">
                    <asp:DropDownList ID="ddlFun" runat="server" AutoPostBack="true">
                        <asp:ListItem Value="">全部</asp:ListItem>
                        <%--<asp:ListItem Value="TC">[TC]訓練機構管理</asp:ListItem>
					<asp:ListItem Value="SD">[SD]學員動態管理</asp:ListItem>
					<asp:ListItem Value="CP">[CP]查核績效管理</asp:ListItem>
					<asp:ListItem Value="TR">[TR]訓練需求管理</asp:ListItem>
					<asp:ListItem Value="CM">[CM]訓練經費控管</asp:ListItem>
					<asp:ListItem Value="FM">[FM]設備預算管理</asp:ListItem>
					<asp:ListItem Value="SE">[SE]技能檢定管理</asp:ListItem>
					<asp:ListItem Value="EXAM">[EXAM]招生甄試管理</asp:ListItem>
					<asp:ListItem Value="SV">[SV]問卷管理</asp:ListItem>
					<asp:ListItem Value="OB">[OB]委外訓練管理</asp:ListItem>
					<asp:ListItem Value="SYS">[SYS]系統管理</asp:ListItem>
					<asp:ListItem Value="FAQ">[FAQ]問答集</asp:ListItem>
					<asp:ListItem Value="OO">[OO]其他系統</asp:ListItem>--%>
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td align="center" colspan="4" class="whitecol">
                    <asp:DataGrid ID="Datagrid4" runat="server" AutoGenerateColumns="False" CssClass="font" Width="100%" CellPadding="8">
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <HeaderStyle CssClass="head_navy" />
                        <Columns>
                            <asp:TemplateColumn HeaderText="功能類別">
                                <HeaderStyle Width="15%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:TextBox ID="txtFunID" Visible="False" runat="server"></asp:TextBox>
                                    <asp:Label ID="lab_MainMenu" runat="server" ForeColor="#00007F"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="功能項目">
                                <HeaderStyle Width="55%"></HeaderStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lab_FunName" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn DataField="Memo" HeaderText="備註">
                                <HeaderStyle Width="20%"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="選用">
                                <HeaderStyle Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <HeaderTemplate>
                                    <asp:CheckBox ID="chk_EnableAll" runat="server" Text="選用"></asp:CheckBox>
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <asp:CheckBox ID="chk_Enable" runat="server" Text="選用"></asp:CheckBox>&nbsp;
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                    </asp:DataGrid><asp:Label ID="lab_Msg4" runat="server" ForeColor="Red" Visible="False">無功能可設定!</asp:Label>
                </td>
            </tr>
            <tr>
                <td align="center" colspan="4" class="whitecol">
                    <asp:Button ID="btn_SaveOption" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                    <asp:Button ID="btn_CancelOption2" runat="server" Text="取消" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
        </table>
        <table class="table_nw" id="tb_GroupTPlan" cellspacing="1" cellpadding="1" width="100%" runat="server">
            <tr>
                <td class="bluecol" style="width: 20%">帳號
                </td>
                <td class="whitecol" style="width: 30%">
                    <asp:Label ID="Lab5UserID" runat="server"></asp:Label>
                </td>
                <td class="bluecol" style="width: 20%">姓名
                </td>
                <td class="whitecol" style="width: 30%">
                    <asp:Label ID="Lab5Name" runat="server"></asp:Label>
                </td>
            </tr>
            <tr>
                <td class="bluecol" style="width: 20%">群組名稱
                </td>
                <td class="whitecol" colspan="3">
                    <asp:Label ID="LabGroupName" runat="server"></asp:Label>
                    <input id="HidGID" type="hidden" runat="server" name="HidGID" />
                </td>
            </tr>
            <tr>
                <td align="center" colspan="4" class="whitecol">
                    <asp:DataGrid ID="Datagrid5" runat="server" AutoGenerateColumns="False" CssClass="font" Width="100%" CellPadding="8">
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <HeaderStyle CssClass="head_navy" />
                        <Columns>
                            <asp:TemplateColumn HeaderText="選用">
                                <HeaderStyle Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:CheckBox ID="ChkBox1" runat="server" />
                                    <input id="hid_TPlanID" type="hidden" runat="server" name="hid_TPlanID" />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="計畫名稱">
                                <HeaderStyle Width="90%"></HeaderStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lab_PlanName" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                    </asp:DataGrid>
                    <asp:Label ID="lab_Msg5" runat="server" ForeColor="Red" Visible="False">無計畫可設定!</asp:Label>
                </td>
            </tr>
            <tr>
                <td align="center" colspan="4" class="whitecol">
                    <asp:Button ID="btnSave5" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                    <asp:Button ID="btnCancel5" runat="server" Text="取消" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
        </table>
        <input id="hid_GidTPlanIDs" type="hidden" runat="server" />
        <input id="Hid_PLANID" type="hidden" runat="server" />
    </form>
</body>
</html>
