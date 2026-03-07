<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_06_002.aspx.vb" Inherits="WDAIIP.SYS_06_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>LOG查詢</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function SelectAll(Flag) {
            var MyTable = document.getElementById('DataGrid1');
            for (i = 1; i < MyTable.rows.length; i++) {
                //debugger; //SelectItem(Flag, MyTable.rows[i].cells[0].children[0].value);
                if (MyTable.rows[i].cells[0].children[2]) { MyTable.rows[i].cells[0].children[2].checked = Flag; }
            }
        }
        function SelectOne1(Flag) {
            var DG1_checkbox3 = document.getElementById('DG1_checkbox3');
            if (DG1_checkbox3 && DG1_checkbox3.checked && Flag == false) { DG1_checkbox3.checked = false; }
        }
    </script>
</head>
<body>
    <form id="Form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;LOG查詢</asp:Label>
                </td>
            </tr>
        </table>

        <table id="FrameTable3" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>

                    <table id="Table3" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol" style="width: 20%">使用者帳號</td>
                            <td class="whitecol">
                                <asp:TextBox ID="tUserID" runat="server" MaxLength="15"></asp:TextBox>
                            </td>
                            <td class="bluecol" style="width: 20%">身分證號</td>
                            <td class="whitecol">
                                <asp:TextBox ID="tIDNO" runat="server" MaxLength="15"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" style="width: 20%">功能名稱</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="ddlKind" runat="server" AutoPostBack="True"></asp:DropDownList>
                                <asp:DropDownList ID="ddlFunP" runat="server" AutoPostBack="True"></asp:DropDownList>
                                <asp:DropDownList ID="ddlFunC" runat="server"></asp:DropDownList>
                                <%--<asp:DropDownList ID="Kind" runat="server" AutoPostBack="True">
                                        <asp:ListItem Value="">==請選擇==</asp:ListItem>
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
                                    <asp:DropDownList ID="FunParent" runat="server" AutoPostBack="True">
                                    </asp:DropDownList>
                                    <asp:DropDownList Style="z-index: 0" ID="ddlFunID" runat="server" AutoPostBack="True">
                                    </asp:DropDownList>--%>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" style="width: 20%">作業方式</td>
                            <td class="whitecol">
                                <asp:DropDownList ID="ddlWorkMethod" runat="server" AutoPostBack="True" AppendDataBoundItems="true">
                                </asp:DropDownList>
                            </td>
                            <td class="bluecol" style="width: 20%">作業日期</td>
                            <td class="whitecol">
                                <span runat="server">
                                    <asp:TextBox ID="WorkDate1" runat="server" MaxLength="10" Columns="6"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= WorkDate1.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif">~
								        <asp:TextBox ID="WorkDate2" runat="server" MaxLength="10" Columns="6"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= WorkDate2.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif">
                                </span>
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
                        </tr>
                        <tr id="tr_ddl_INQUIRY_S" runat="server">
                            <td class="bluecol">查詢原因</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="ddl_INQUIRY_Sch" runat="server"></asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" style="width: 20%">作業顯示模式</td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList Style="z-index: 0" ID="rblWorkMode" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="0">不區分</asp:ListItem>
                                    <asp:ListItem Value="1" Selected="True">模糊顯示</asp:ListItem>
                                    <asp:ListItem Value="2">正常顯示</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4" align="center">
                                <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                <asp:Button ID="btnQuery" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>&nbsp;
                                <asp:Button ID="btnExport" runat="server" Text="匯出查詢紀錄單" CssClass="asp_Export_M"></asp:Button>&nbsp;
                                <asp:Label ID="lab_TOP_MAX_ROWS" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4" align="center">
                                <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table id="DataGridTable1" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td align="center">
                                <asp:DataGrid Style="z-index: 0" ID="DataGrid1" runat="server" Width="100%" AllowPaging="True" AutoGenerateColumns="False" CssClass="font">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn HeaderText="編號">
                                            <HeaderStyle HorizontalAlign="Center" Width="30px"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <HeaderTemplate>
                                                選取<input id="DG1_checkbox3" onclick="SelectAll(this.checked);" type="checkbox">
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_SNO" runat="server"></asp:Label>
                                                <asp:HiddenField ID="Hid_LAID" runat="server" />
                                                <input id="CB_SNO" type="checkbox" runat="server" onclick="SelectOne1(this.checked);" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="ACCOUNT" HeaderText="使用者帳號"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="CName" HeaderText="姓名"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="FunName" HeaderText="功能名稱"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="WorkMethod" HeaderText="作業方式"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="WorkMode" HeaderText="作業顯示模式"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="WorkDate" HeaderText="作業日期"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="INQNO_N" HeaderText="查詢原因"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="RESCNT" HeaderText="查詢筆數"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="Note" HeaderText="作業範圍"></asp:BoundColumn>
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
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
