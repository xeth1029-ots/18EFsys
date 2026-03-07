<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_03_022.aspx.vb" Inherits="WDAIIP.SYS_03_022" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>群組設定</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function Check_Data() {
            var msg = '';
            var txt_GroupName = document.getElementById("txt_GroupName");
            var ddlGtype = document.getElementById("ddlGtype");

            if (ddlGtype.value == '') { msg += '請選擇使用單位!\n'; }

            if (txt_GroupName.value == '') { msg += '請輸入群組名稱!\n'; }

            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function Show_Select(obj, idx) {
            var myTable = document.getElementById('DataGrid2');
            var intCntCell = myTable.rows[idx].cells.length;

            var chkEnable;
            var chkSch;
            var chkEdit;
            var chkPrt;

            if (intCntCell == 7) {
                if (myTable.rows[idx].cells[3].children[0]) chkEnable = myTable.rows[idx].cells[3].children[0];
                if (myTable.rows[idx].cells[4].children[0]) chkSch = myTable.rows[idx].cells[4].children[0];
                if (myTable.rows[idx].cells[5].children[0]) chkEdit = myTable.rows[idx].cells[5].children[0];
                if (myTable.rows[idx].cells[6].children[0]) chkPrt = myTable.rows[idx].cells[6].children[0];

            } else {
                if (myTable.rows[idx].cells[2].children[0]) chkEnable = myTable.rows[idx].cells[2].children[0];
                if (myTable.rows[idx].cells[3].children[0]) chkSch = myTable.rows[idx].cells[3].children[0];
                if (myTable.rows[idx].cells[4].children[0]) chkEdit = myTable.rows[idx].cells[4].children[0];
                if (myTable.rows[idx].cells[5].children[0]) chkPrt = myTable.rows[idx].cells[5].children[0];
            }

            switch (obj) {
                case "enable":
                    if (chkEnable.checked == false) {
                        if (chkSch) chkSch.checked = false;
                        if (chkEdit) chkEdit.checked = false;
                        if (chkPrt) chkPrt.checked = false;
                    }
                    break;
                default:
                    if (chkSch.checked || chkEdit.checked || chkPrt.checked) chkEnable.checked = true;
                    break;
            }
        }

        function Show_SelectAll(objTable, objChk, idx) {
            var myTable = document.getElementById(objTable);
            var chk = document.getElementById(objChk);
            var intCntCell = 0;

            for (i = 1; i <= myTable.rows.length - 1; i++) {
                intCntCell = myTable.rows[i].cells.length;

                if (intCntCell == 7) {
                    if (chk.checked) {
                        if (myTable.rows[i].cells[idx].children[0]) myTable.rows[i].cells[idx].children[0].checked = true;
                    } else {
                        if (myTable.rows[i].cells[idx].children[0]) myTable.rows[i].cells[idx].children[0].checked = false;

                        if (idx == 3) {
                            if (myTable.rows[i].cells[4].children[0]) myTable.rows[i].cells[4].children[0].checked = false;
                            if (myTable.rows[i].cells[5].children[0]) myTable.rows[i].cells[5].children[0].checked = false;
                            if (myTable.rows[i].cells[6].children[0]) myTable.rows[i].cells[6].children[0].checked = false;
                        }
                    }

                } else {
                    if (chk.checked) {
                        if (myTable.rows[i].cells(idx - 1).children[0]) myTable.rows[i].cells(idx - 1).children[0].checked = true;
                    } else {
                        if (myTable.rows[i].cells(idx - 1).children[0]) myTable.rows[i].cells(idx - 1).children[0].checked = false;

                        if (idx == 3) {
                            if (myTable.rows[i].cells[3].children[0]) myTable.rows[i].cells[3].children[0].checked = false;
                            if (myTable.rows[i].cells[4].children[0]) myTable.rows[i].cells[4].children[0].checked = false;
                            if (myTable.rows[i].cells[5].children[0]) myTable.rows[i].cells[5].children[0].checked = false;
                        }
                    }
                }
            }

            if (idx != 3) {
                var flagChk = false;

                for (i = 1; i <= myTable.rows.length - 1; i++) {
                    flagChk = false;
                    intCntCell = myTable.rows[i].cells.length;

                    if (intCntCell == 7) {
                        if (myTable.rows[i].cells[4].children[0]) {
                            if (myTable.rows[i].cells[4].children[0].checked) flagChk = true;
                        } else {
                            if (chk.checked == true) flagChk = true;
                        }

                        if (myTable.rows[i].cells[5].children[0]) { if (myTable.rows[i].cells[5].children[0].checked) flagChk = true; }
                        if (myTable.rows[i].cells[6].children[0]) { if (myTable.rows[i].cells[6].children[0].checked) flagChk = true; }
                        myTable.rows[i].cells[3].children[0].checked = flagChk;

                    } else {
                        if (myTable.rows[i].cells[3].children[0]) {
                            if (myTable.rows[i].cells[3].children[0].checked) flagChk = true;
                        } else {
                            if (chk.checked == true) flagChk = true;
                        }

                        if (myTable.rows[i].cells[4].children[0]) { if (myTable.rows[i].cells[4].children[0].checked) flagChk = true; }
                        if (myTable.rows[i].cells[5].children[0]) { if (myTable.rows[i].cells[5].children[0].checked) flagChk = true; }
                        myTable.rows[i].cells[2].children[0].checked = flagChk;
                    }
                }
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server"> 首頁&gt;&gt;系統管理&gt;&gt;功能權限管理&gt;&gt;群組設定 </asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="FrameTable2" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>

                    <asp:Panel ID="tb_Query" runat="server">
                        <table class="table_nw" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol" style="width: 20%">建檔單位
                                </td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="ddlQDistID" runat="server">
                                    </asp:DropDownList>
                                    <input type="hidden" id="hidIsSys" runat="server">
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">群組名稱
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_QGroupName" runat="server" Width="30%"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">群組階層
                                </td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="ddlQType" runat="server">
                                        <asp:ListItem Value="">不拘</asp:ListItem>
                                        <asp:ListItem Value="0">署</asp:ListItem>
                                        <asp:ListItem Value="1">分署</asp:ListItem>
                                        <asp:ListItem Value="2">委訓</asp:ListItem>
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">啟用狀態
                                </td>
                                <td class="whitecol">
                                    <asp:RadioButtonList ID="rdoQVailid" runat="server" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="" Selected="True">不拘</asp:ListItem>
                                        <asp:ListItem Value="1">啟用</asp:ListItem>
                                        <asp:ListItem Value="0">停用</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                        </table>
                        <table width="100%">
                            <tr>
                                <td align="center" class="whitecol">
                                    <asp:Button ID="btn_Query" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                    &nbsp;<asp:Button ID="btn_Add" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>
                                   
                                </td>
                            </tr>
                            <tr>
                                <td align="center" class="font">
                                    <asp:Label ID="lab_Msg" runat="server" Visible="False" ForeColor="Red">查無資料</asp:Label>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                    <table class="font" id="tb_List" cellspacing="0" cellpadding="0" width="100%" border="0" runat="server">
                        <tr>
                            <td align="center">
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:TemplateColumn HeaderText="序號">
                                            <HeaderStyle Width="4%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_SNo" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="建檔單位">
                                            <HeaderStyle Width="15%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_GroupDistID" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="群組階層">
                                            <HeaderStyle Width="8%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_GroupType" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="群組名稱">
                                            <HeaderStyle Width="18%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_GroupName" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="建檔者">
                                            <HeaderStyle Width="9%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_GroupCUsr" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="最後修改者">
                                            <HeaderStyle Width="9%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_GroupMUsr" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="備註">
                                            <HeaderStyle Width="16%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_GroupNote" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="啟用">
                                            <ItemStyle Width="5%" HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_Enable" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Button ID="btn_Edit" runat="server" Text="修改" CommandName="EDIT" CssClass="asp_button_M"></asp:Button>
                                                <asp:Button ID="btn_Copy" runat="server" Text="複製" CommandName="COPY" CssClass="asp_button_M"></asp:Button>
                                                <br>
                                                <asp:Button ID="btn_Del" runat="server" Text="刪除" CommandName="DEL" CssClass="asp_button_M"></asp:Button>
                                                <asp:Button ID="btnExe1" runat="server" Text="被賦予者" CommandName="EXE1" CssClass="asp_button_M"></asp:Button>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                    </table>
                    <asp:Panel ID="Panel_Exe1" runat="server">
                        <table class="table_nw" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol_need" style="width: 20%">建檔單位<font color="red">*</font>
                                </td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="ddlDistID3" runat="server">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">群組階層<font color="red">*</font>
                                </td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="ddlGtype3" runat="server">
                                        <asp:ListItem Value="">請選擇</asp:ListItem>
                                        <asp:ListItem Value="0">署</asp:ListItem>
                                        <asp:ListItem Value="1">分署</asp:ListItem>
                                        <asp:ListItem Value="2">委訓</asp:ListItem>
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">群組名稱<font color="red">*</font>
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_GroupName3" runat="server" MaxLength="25" Columns="60" Width="30%"></asp:TextBox>
                                    <%--<input id="Hidden1" style="width: 25px; height: 22px" type="hidden" name="hide_GID" runat="server">--%>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">備註
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_GroupNote3" runat="server" MaxLength="50" Columns="60" Width="30%"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td align="center" class="whitecol" colspan="2">
                                    <asp:Label ID="lab_Msg3" runat="server" Visible="False" ForeColor="Red">查無資料</asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol" colspan="2">
                                    <%--<table class="font" cellspacing="0" cellpadding="0" width="100%" border="0"></table>--%>
                                    <asp:DataGrid ID="DataGrid3" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <Columns>
                                            <asp:BoundColumn HeaderText="序號">
                                                <HeaderStyle Width="25%" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="orgname" HeaderText="單位名稱">
                                                <HeaderStyle Width="25%" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="name" HeaderText="使用者姓名">
                                                <HeaderStyle Width="25%" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="account" HeaderText="使用者帳號">
                                                <HeaderStyle Width="25%" />
                                            </asp:BoundColumn>
                                        </Columns>
                                    </asp:DataGrid>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol" align="center" colspan="2">
                                    <asp:Button ID="btn_LoadBack3" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                    <asp:Panel ID="tb_Edit" runat="server">
                        <table class="table_nw" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol_need" style="width: 20%">建檔單位<font color="red">*</font>
                                </td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="ddlDistID" runat="server">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">群組階層<font color="red">*</font>
                                </td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="ddlGtype" runat="server">
                                        <asp:ListItem Value="">請選擇</asp:ListItem>
                                        <asp:ListItem Value="0">署</asp:ListItem>
                                        <asp:ListItem Value="1">中心</asp:ListItem>
                                        <asp:ListItem Value="2">委訓</asp:ListItem>
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">群組名稱<font color="red">*</font>
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_GroupName" runat="server" MaxLength="25" Columns="60" Width="50%"></asp:TextBox>
                                    <input id="hide_GID" type="hidden" name="hide_GID" runat="server" />
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">群組備註</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_GroupNote" runat="server" MaxLength="50" Columns="60" Width="50%"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">啟用狀態
                                </td>
                                <td class="whitecol">
                                    <asp:CheckBox ID="chk_Valid" runat="server" Checked="True" Text="啟用"></asp:CheckBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">功能顯示
                                </td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="list_MainMenu" runat="server">
                                        <asp:ListItem Value="">全部</asp:ListItem>
                                        <asp:ListItem Value="TC">[TC]訓練機構管理</asp:ListItem>
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
                                        <asp:ListItem Value="OO">[OO]其他系統</asp:ListItem>
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" style="width: 20%">功能項目<br />
                                    (關鍵字)
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txtFunName" runat="server" MaxLength="50" Width="30%"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol" colspan="2" align="center">
                                    <asp:Button ID="btnSearch3" runat="server" Text="查詢" CssClass="asp_button_M" />
                                </td>
                            </tr>
                            <tr>
                               <td class="whitecol" colspan="2" align="center">
                                     <asp:Button ID="btn_Save_2" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                                     &nbsp;<asp:Button ID="btn_LoadBack_2" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                            <tr>
                                <td align="left" colspan="2" class="whitecol">
                                    <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <Columns>
                                            <asp:TemplateColumn HeaderText="功能類別" HeaderStyle-Width="12%" ItemStyle-HorizontalAlign="Center">
                                                <ItemTemplate>
                                                    <asp:TextBox ID="txtFunID" Visible="False" runat="server"></asp:TextBox>
                                                    <asp:Label ID="lab_MainMenu" runat="server" ForeColor="#00007F"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="功能項目" HeaderStyle-Width="40%">
                                                <ItemTemplate>
                                                    <asp:Label ID="lab_FunName" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:BoundColumn DataField="Memo" HeaderText="備註"></asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="選用" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="8%">
                                                <HeaderTemplate>
                                                    <asp:CheckBox ID="chk_EnableAll" runat="server" Text="選用"></asp:CheckBox>
                                                </HeaderTemplate>
                                                <ItemTemplate>
                                                    <asp:CheckBox ID="chk_Enable" runat="server" Text="選用"></asp:CheckBox>&nbsp;
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="查詢" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="8%">
                                                <HeaderTemplate>
                                                    <asp:CheckBox ID="chk_SchAll" runat="server" Text="查詢"></asp:CheckBox>
                                                </HeaderTemplate>
                                                <ItemTemplate>
                                                    <asp:CheckBox ID="chk_Sch" runat="server" Text="查詢"></asp:CheckBox>&nbsp;
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="維護" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="8%">
                                                <HeaderTemplate>
                                                    <asp:CheckBox ID="chk_EditAll" runat="server" Text="維護"></asp:CheckBox>
                                                </HeaderTemplate>
                                                <ItemTemplate>
                                                    <asp:CheckBox ID="chk_Edit" runat="server" Text="維護"></asp:CheckBox>&nbsp;
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="列印" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="8%">
                                                <HeaderTemplate>
                                                    <asp:CheckBox ID="chk_PrtAll" runat="server" Text="列印"></asp:CheckBox>
                                                </HeaderTemplate>
                                                <ItemTemplate>
                                                    <asp:CheckBox ID="chk_Prt" runat="server" Text="列印"></asp:CheckBox>&nbsp;
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                    </asp:DataGrid>
                                </td>
                            </tr>
                        </table>
                        <table width="100%">
                            <tr>
                                <td align="center" class="whitecol">
                                    <asp:Button ID="btn_Save" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                                    &nbsp;<asp:Button ID="btn_LoadBack" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
