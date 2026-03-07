<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_01_002.aspx.vb" Inherits="WDAIIP.SYS_01_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>帳號計畫賦予</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function openrefer() {
            //document.form1.
            var RIDValue = document.getElementById('RIDValue');
            var Lis_acc = document.getElementById('Lis_acc');
            var Years = document.getElementById('Years');
            var msg = '';
            if (!RIDValue) { msg += '請選擇機構(不可為空)!\n'; }
            if (!Lis_acc) { msg += '請選擇帳號(不可為空)!\n'; }
            if (!Years) { msg += '請選擇年度(不可為空)!\n'; }
            if (msg != '') {
                alert(msg);
                return false;
            }
            if (RIDValue.value == '') { msg += '請選擇機構!\n'; }
            if (Lis_acc.value == '') { msg += '請選擇帳號!\n'; }
            if (Years.value == '') { msg += '請選擇年度!\n'; }
            if (msg != '') {
                alert(msg);
                return false;
            }
            window.open('SYS_01_002_f.aspx?AN=' + Lis_acc.value + '&RID=' + RIDValue.value + '&Years=' + Years.value, '', 'width=610,height=520,location=0,status=0,menubar=0,scrollbars=1,resizable=0');
            return false;
        }

        function chkaccount(num) {
            var Lis_acc = document.getElementById('Lis_acc');
            var center = document.getElementById('center');
            var isBlack = document.getElementById('isBlack');
            var OrgName = document.getElementById('OrgName');
            var isBlack2 = document.getElementById('isBlack2');
            var OrgRID = document.getElementById('OrgRID');

            var msg = '';
            var orgname_val = '';
            var msg1 = '';
            var msg2 = '';

            if (Lis_acc.value == '') msg += '請選擇帳號!\n';
            //debugger;
            if (num == 1) {
                orgname_val = center.value;
                if (isBlack.value == 'Y') {
                    msg1 = orgname_val + "，已列入處分名單，是否確定繼續？"
                    msg2 = orgname_val + "，已列入處分名單!!"
                    if (!confirm(msg1)) {
                        msg += msg2;
                    }
                }
            }
            if (num == 2) {
                orgname_val = OrgName.value;
                if (isBlack2.value == 'Y') {
                    msg1 = orgname_val + "，已列入處分名單，是否確定繼續？"
                    msg2 = orgname_val + "，已列入處分名單!!"
                    if (!confirm(msg1)) {
                        msg += msg2;
                    }
                }
            }

            if (num == 2) { if (OrgRID.value == '') msg += '請選擇隸屬機構!\n'; }

            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        /*
        function returnValue(RID,Name){
        document.form1.OrgName.value=Name;
        document.form1.OrgRID.value=RID;
        }
        function ShowFrame(){
        document.getElementById('FrameObj').style.display=document.getElementById('HistoryList2').style.display;
        }
        */

        function chkName() {
            var HidDefInput = document.getElementById('HidDefInput');
            var txtSchAccount = document.getElementById('txtSchAccount'); // document.form1.txtSchAccount;
            var Cst_DefInput = HidDefInput.value; // '請輸入姓名關鍵字';

            if (txtSchAccount.value == '') {
                txtSchAccount.value = Cst_DefInput;
                txtSchAccount.style.color = '#858585';

            } else if (txtSchAccount.value == Cst_DefInput) {
                txtSchAccount.value = '';
                txtSchAccount.style.color = 'black';
            }
        }

        function showList() {
            var myTable = document.getElementById('Datagrid2');
            var lsbShow = document.getElementById('lsbShow');
            //var myChk = '';
            //var strName = '';
            //var str = '';

            for (var i = lsbShow.options.length - 1; i >= 0; i--) {
                lsbShow.remove(i);
            }

            for (var i = 1; i < myTable.rows.length; i++) {
                var myChk = myTable.rows[i].cells[0].children[0];
                var strName = myTable.rows[i].cells[1].innerText;

                if (myChk.id == '') {
                    var str = myTable.rows[i].cells[0].children[0].innerHTML;
                    str = str.substring(str.indexOf('id=') + 3);
                    str = str.substring(0, str.indexOf(' '));
                    if (str.indexOf('"') > -1) {
                        str = str.replace(/\"/g, '');
                    }
                    myChk = document.getElementById(str);
                }

                if (myChk.checked) lsbShow.options.add(new Option(strName, ''));
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;帳號維護管理&gt;&gt;帳號-計畫賦予</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" style="width: 20%">狀態
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="rdoIsUsed" runat="server" AutoPostBack="True" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="A">不區分</asp:ListItem>
                                    <asp:ListItem Value="Y" Selected="True">啟用中</asp:ListItem>
                                    <asp:ListItem Value="N">停用中</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練單位
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="btu_org" type="button" value="選擇" name="btu_org" runat="server" class="asp_button_M">
                                <asp:Button ID="but_search" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                <input id="orgid_Level" type="hidden" name="orgid_Level" runat="server">
                            </td>
                        </tr>
                    </table>
                    <table id="ShowTable" cellspacing="0" cellpadding="0" class="font" border="0" width="100%" runat="server">
                        <tr>
                            <td>
                                <%--"AccountTable"--%>
                                <table class="table_sch" cellspacing="1" cellpadding="1" width="100%">
                                    <tr>
                                        <td class="bluecol" style="width: 20%">帳號
                                        </td>
                                        <td colspan="3" class="whitecol">
                                            <asp:DropDownList ID="Lis_acc" runat="server" AutoPostBack="True">
                                            </asp:DropDownList>
                                            <asp:TextBox ID="txtSchAccount" runat="server" Columns="22"></asp:TextBox><asp:Button ID="btnSchAccount" runat="server" Text="搜尋" CssClass="asp_button_M"></asp:Button>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">年度
                                        </td>
                                        <td colspan="2" class="whitecol">
                                            <asp:DropDownList ID="Years" runat="server" AutoPostBack="True">
                                            </asp:DropDownList>
                                            <asp:DropDownList ID="Lis_Plan" runat="server" AutoPostBack="True" Visible="false">
                                            </asp:DropDownList>
                                            <asp:Button ID="btu_add" runat="server" Text="新增" Visible="False" CssClass="asp_button_M"></asp:Button>
                                            <asp:Button ID="butRefer" runat="server" Text="參考其他年度設定" CssClass="asp_button_M"></asp:Button>
                                            <asp:Button ID="btn_QUERY2" runat="server" Text="跨年度查詢" CssClass="asp_button_M"></asp:Button>
                                        </td>
                                        <%--<td class="whitecol">
										<asp:Button ID="btu_add" runat="server" Text="新增" Visible="False" CssClass="asp_button_M"></asp:Button>
										<asp:Button ID="butRefer" runat="server" Text="參考其他年度設定" CssClass="asp_button_M"></asp:Button>
									</td>--%>
                                    </tr>
                                </table>
                                <table id="OrgTable" class="table_sch" runat="server" cellspacing="1" cellpadding="1">
                                    <tr>
                                        <td class="bluecol" style="width: 20%">隸屬機構
                                        </td>
                                        <td colspan="3" class="whitecol">
                                            <asp:TextBox ID="OrgName" runat="server" onfocus="this.blur()" Columns="40" Width="30%"></asp:TextBox>
                                            <input id="OrgRID" type="hidden" name="OrgRID" runat="server">
                                            <input id="Button2" type="button" value="選擇" name="Button2" runat="server" class="asp_button_M">
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr id="trView" runat="server">
                            <td align="left" colspan="3">
                                <font style="font-size: 12px; color: #FF0000">*刪除計劃時請謹慎</font>
                                <table cellspacing="0" cellpadding="0" width="100%">
                                    <tr>
                                        <td width="49%">
                                            <div style="padding-bottom: 0px; margin-top: 0px; padding-left: 0px; width: 100%; padding-right: 0px; margin-bottom: 0px; height: 200px; margin-left: 0px; padding-top: 0px; overflow: scroll;">
                                                <asp:DataGrid ID="Datagrid2" CssClass="font" runat="server" Width="100%" HeaderStyle-CssClass="TD_TD1" AutoGenerateColumns="False" OnDeleteCommand="CL_Delete" DataKeyField="PlanID" CellPadding="8">
                                                    <AlternatingItemStyle BackColor="WhiteSmoke" />
                                                    <HeaderStyle CssClass="head_navy" HorizontalAlign="Center"></HeaderStyle>
                                                    <Columns>
                                                        <asp:TemplateColumn>
                                                            <ItemStyle HorizontalAlign="Center" />
                                                            <ItemTemplate>
                                                                <asp:CheckBox ID="chk1" runat="server"></asp:CheckBox>
                                                                <asp:HiddenField ID="Hid_RID" runat="server" />
                                                                <asp:HiddenField ID="Hid_planid" runat="server" />
                                                                <asp:HiddenField ID="Hid_tplanid" runat="server" />
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:BoundColumn DataField="PlanName" HeaderText="訓練計畫" ItemStyle-Width="90%"></asp:BoundColumn>
                                                        <%--<asp:BoundColumn Visible="false" DataField="RID" HeaderText="RID"></asp:BoundColumn>
                                                        <asp:BoundColumn Visible="false" DataField="planid" HeaderText="planid"></asp:BoundColumn>--%>
                                                    </Columns>
                                                </asp:DataGrid>
                                            </div>
                                        </td>
                                        <td width="2%">&nbsp;
                                        </td>
                                        <td width="49%">
                                            <asp:ListBox ID="lsbShow" runat="server" Width="100%" Rows="12" SelectionMode="Multiple"></asp:ListBox>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="3" class="whitecol">
                                <asp:Button ID="bt_save" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">本功能屬權限管理範圍，因權限開放跨區選擇訓練單位、帳號<br>
                                故請謹慎選擇新增計畫名稱、隸屬機構之參數，避免計畫名稱、隸屬機構錯誤之困擾。<br>
                                PS:依過去資料之經驗判斷，大致上每個使用者在同一年度之計畫只會屬同一隸屬機構
                            </td>
                        </tr>
                        <tr>
                            <td align="left">
                                <asp:DataGrid ID="DataGrid1" CssClass="font" runat="server" Width="100%" Visible="False" HeaderStyle-CssClass="TD_TD1" AutoGenerateColumns="False" OnDeleteCommand="CL_Delete" DataKeyField="PlanID" CellPadding="8">
                                    <AlternatingItemStyle BackColor="WhiteSmoke" />
                                    <HeaderStyle CssClass="head_navy" HorizontalAlign="Center"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn DataField="PlanName" HeaderText="計畫代碼　(補助地方政府)" HeaderStyle-Width="40%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構" HeaderStyle-Width="40%"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderStyle-Width="20%">
                                            <ItemStyle HorizontalAlign="Center" />
                                            <HeaderTemplate>
                                                功能
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:Button ID="But_Att" runat="server" Text="委訓單位歸屬" CausesValidation="False" CommandName="Att" CssClass="asp_button_M"></asp:Button>
                                                <asp:Button ID="But_Del" runat="server" Text="刪除" CausesValidation="False" CommandName="Delete" CssClass="asp_button_M"></asp:Button>
                                                <asp:Label ID="Label1" runat="server"></asp:Label>
                                                <asp:Button ID="But_AttB" runat="server" Text="委訓單位歸屬全選" CausesValidation="False" CommandName="AttB" CssClass="asp_button_M"></asp:Button>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn Visible="False" DataField="RID" HeaderText="RID"></asp:BoundColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr align="center">
                            <td>
                                <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <input id="isBlack" type="hidden" name="isBlack" runat="server" />
        <input id="isBlack2" type="hidden" name="isBlack2" runat="server" />
        <input id="hidRoleId" type="hidden" name="hidRoleId" runat="server" />
        <asp:HiddenField ID="HidDefInput" runat="server" />
        <asp:HiddenField ID="Hid_RoleID" runat="server" />
    </form>
</body>
</html>
