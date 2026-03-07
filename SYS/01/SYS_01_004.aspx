<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_01_004.aspx.vb" Inherits="WDAIIP.SYS_01_004" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>單位-計畫賦予設定</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function returnValue(RID, Name) {
            document.form1.OrgName.value = Name;
            document.form1.OrgRID.value = RID;
        }

        function setAccValue(acc) {
            document.getElementById('hid_acc').value = acc;
        }

        function openView() {

        }

        function switch_chk(tbname, hidname) {
            var chk = document.getElementById("" + hidname + "");
            if (chk.value == '0') chk.value = '1';
            else chk.value = '0';

            chkall(tbname, chk.value);
        }

        function chkall(tbname, chkval) {
            var MyTable = document.getElementById("" + tbname + "");

            for (var i = 1; i < MyTable.rows.length; i++) {
                if (chkval == '0') MyTable.rows[i].cells[0].children[0].checked = true;
                else MyTable.rows[i].cells[0].children[0].checked = false;
            }

            showList(tbname);
        }

        function showList(obj) {
            var myTable = document.getElementById(obj);
            var myChk;
            var strName = '';
            var objListBox;

            if (obj == 'Datagrid1') objListBox = document.getElementById('lsbAccount');
            else objListBox = document.getElementById('lsbPlan');

            for (var i = objListBox.options.length - 1; i >= 0; i--) {
                objListBox.remove(i);
            }

            for (var i = 1; i < myTable.rows.length; i++) {
                myChk = myTable.rows[i].cells[0].children[0];
                strName = myTable.rows[i].cells[1].innerText;
                if (myChk.checked && myChk.disabled == false) objListBox.options.add(new Option(strName, ''));
            }
        }
    </script>
    <style type="text/css">
        .style1 { width: 90px; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;帳號維護管理&gt;&gt;單位-計畫賦予</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
            <%-- <tr>
            <td>
                首頁&gt;&gt;系統管理&gt;&gt;帳號維護管理&gt;&gt;單位-計畫賦予設定
            </td>
        </tr>--%>
            <tr>
                <td>
                    <table class="table_sch" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol" style="width: 20%">訓練單位
                            </td>
                            <td class="whitecol" style="width: 80%">
                                <asp:TextBox ID="center" Columns="40" onfocus="this.blur()" runat="server" Width="60%"></asp:TextBox>
                                <input id="btu_org" type="button" value="選擇" name="btu_org" runat="server" class="asp_button_M">
                                <asp:Button ID="but_search" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                <asp:LinkButton ID="lnkBtn" runat="server"></asp:LinkButton>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                <input id="hid_acc" type="hidden" name="hid_acc" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">年度
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="ddl_years" runat="server" AutoPostBack="True">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">帳號角色
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="rdo_role" runat="server" AutoPostBack="True" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">狀態
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="rdoIsUsed" runat="server" AutoPostBack="True" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                                    <asp:ListItem Value="A">不區分</asp:ListItem>
                                    <asp:ListItem Value="Y" Selected="True">啟用中</asp:ListItem>
                                    <asp:ListItem Value="N">停用中</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <table class="font" id="tbSch" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td width="50%">
                    <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr align="center" class="head_navy">
                            <td style="width: 15%">
                                <asp:CheckBox ID="chkAll1" runat="server"></asp:CheckBox>
                            </td>
                            <td style="width: 25%">姓名帳號
                            </td>
                            <td style="width: 60%">備註
                            </td>
                        </tr>
                    </table>
                    <div style="padding-right: 0px; margin-top: 0px; padding-left: 0px; margin-bottom: 0px; padding-bottom: 0px; margin-left: 0px; overflow-y: scroll; width: 100%; padding-top: 0px; height: 200px">
                        <asp:DataGrid ID="Datagrid1" runat="server" CssClass="font" HeaderStyle-CssClass="TD_TD1" AutoGenerateColumns="False" Width="100%" Visible="true" CellPadding="8">
                            <AlternatingItemStyle BackColor="WhiteSmoke" />
                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                            <Columns>
                                <asp:TemplateColumn HeaderStyle-Width="15%">
                                    <ItemTemplate>
                                        <asp:CheckBox ID="chk1" runat="server"></asp:CheckBox>
                                        <input type="hidden" id="IsUsed" runat="server" name="IsUsed" />
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:BoundColumn DataField="Name" HeaderText="姓名帳號" HeaderStyle-Width="25%"></asp:BoundColumn>
                                <asp:TemplateColumn HeaderText="備註" HeaderStyle-Width="60%">
                                    <ItemTemplate>
                                        <asp:Label ID="note1" runat="server" ForeColor="#0000cc"></asp:Label>
                                        <input type="hidden" id="hid_account" runat="server" name="hid_account" />
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                            </Columns>
                        </asp:DataGrid>
                    </div>
                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                </td>
                <td align="center">
                    <b>已選取姓名帳號</b><br>
                    <br style="line-height: 5px">
                    <asp:ListBox ID="lsbAccount" runat="server" Width="90%" SelectionMode="Multiple" Rows="12"></asp:ListBox>
                </td>
            </tr>
            <tr>
                <td>&nbsp;
                </td>
            </tr>
            <tr>
                <td>
                    <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr class="head_navy">
                            <td style="width: 15%">
                                <asp:CheckBox ID="chkAll2" runat="server"></asp:CheckBox>
                            </td>
                            <td style="width: 40%">計畫名稱
                            </td>
                            <td style="width: 45%">備註
                            </td>
                        </tr>
                    </table>
                    <div style="padding-right: 0px; margin-top: 0px; padding-left: 0px; margin-bottom: 0px; padding-bottom: 0px; margin-left: 0px; overflow-y: scroll; width: 100%; padding-top: 0px; height: 200px">
                        <asp:DataGrid ID="Datagrid2" runat="server" CssClass="font" HeaderStyle-CssClass="TD_TD1" AutoGenerateColumns="False" Width="100%" Visible="true" DataKeyField="PlanID" CellPadding="8">
                            <AlternatingItemStyle BackColor="WhiteSmoke" />
                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                            <Columns>
                                <asp:TemplateColumn HeaderStyle-Width="15%">
                                    <ItemTemplate>
                                        <asp:CheckBox ID="chk2" runat="server"></asp:CheckBox>
                                        <input type="hidden" id="IsUsed2" runat="server" name="IsUsed2" />
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:TemplateColumn HeaderText="計畫名稱" HeaderStyle-Width="40%">
                                    <ItemTemplate>
                                        <asp:Label ID="lb_PlanName" runat="server"></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:BoundColumn Visible="false" DataField="planid" HeaderText="planid"></asp:BoundColumn>
                                <asp:TemplateColumn HeaderText="備註" HeaderStyle-Width="45%">
                                    <ItemTemplate>
                                        <asp:Label ID="note2" runat="server" ForeColor="#0000cc"></asp:Label>
                                        <input type="hidden" id="hid_RID" runat="server" name="hid_RID" />
                                        <input type="hidden" id="hid_planid" runat="server" name="hid_planid" />
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                            </Columns>
                        </asp:DataGrid>
                    </div>
                    <asp:Label ID="msg2" runat="server" ForeColor="Red"></asp:Label>
                </td>
                <td align="center">
                    <b>已選取計畫名稱</b><br>
                    <br style="line-height: 5px">
                    <asp:ListBox ID="lsbPlan" runat="server" Width="90%" SelectionMode="Multiple" Rows="12"></asp:ListBox>
                </td>
            </tr>
            <tr id="tr_save" runat="server">
                <td align="center" colspan="2" class="whitecol">
                    <br>
                    <asp:Button ID="bt_save" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
        </table>
        <input id="chkall_1" type="hidden" runat="server">
        <input id="chkall_2" type="hidden" runat="server">
    </form>
</body>
</html>
