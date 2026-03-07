<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_02_001.aspx.vb" Inherits="WDAIIP.TR_02_001" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>企業需求訪視表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function SelectAll() {
            var MyValue = getCheckBoxListValue('SCTID');
            var MyAllCheck = document.getElementById('SCTID_' + 0);

            if (document.getElementById('HidObj').value != MyValue.charAt(0)) {
                document.getElementById('HidObj').value = MyValue.charAt(0);
                for (var i = 1; i < MyValue.length; i++) {
                    var MyCheck = document.getElementById('SCTID_' + i);
                    MyCheck.checked = MyAllCheck.checked;
                }
            }
        }
        function CheckData() {
            var msg = '';
            if (!isChecked(document.getElementsByName('VisitKind'))) msg += '請選擇訪視類型\n';
            //if(document.form1.Uname.value=='') msg+='請輸入受訪單位名稱\n';
            if (document.form1.BDID.value == '') msg += '請選擇受訪單位\n';
            //if(document.form1.Intaxno.value=='') msg+='請輸入統一編號\n';
            //if(document.form1.Zip.value=='') msg+='請輸入地址(郵遞區號)\n';
            //if(document.form1.Addr.value=='') msg+='請輸入地址\n';
            if (document.form1.TradeID.selectedIndex == 0) msg += '請選擇行業別\n';
            //if(!isChecked(document.getElementsByName('KEID'))) msg+='請選擇現有員工數\n';
            //if(!isChecked(document.getElementsByName('Labor'))) msg+='請選擇勞保\n';
            if (document.form1.VisitedName.value == '') msg += '請輸入受訪者姓名\n';
            if (document.form1.VisitedTel.value == '') msg += '請輸入受訪者電話\n';
            if (!isChecked(document.getElementsByName('BVCID'))) msg += '請選擇訪視情形\n';
            else {
                if (document.getElementsByName('BVCID')[document.getElementsByName('BVCID').length - 1].checked && document.form1.VisitOther.value == '') msg += '請輸入訪視情形[其他]欄位\n'
            }
            if (!isChecked(document.getElementsByName('BPKind'))) msg += '請選擇後續處理方式\n';
            else {
                if (document.getElementsByName('BPKind')[document.getElementsByName('BPKind').length - 1].checked) {
                    if (document.form1.BPKind_Year.value == '') msg += '請輸入預計訪視日期[年]\n'
                    else if (!isUnsignedInt(document.form1.BPKind_Year.value)) msg += '預計訪視日期[年]必須為數字\n'
                    if (document.form1.BPKind_Mon.value == '') msg += '請輸入預計訪視日期[月]\n'
                    else if (!isUnsignedInt(document.form1.BPKind_Mon.value)) msg += '預計訪視日期[月]必須為數字\n'
                    if (document.form1.BPKind_Day.value == '') msg += '請輸預計訪視日期[日]\n'
                    else if (!isUnsignedInt(document.form1.BPKind_Day.value)) msg += '預計訪視日期[日]必須為數字\n'
                }
            }

            if (msg != '') {
                alert(msg);
                return false;
            }
        }
        function CheckAdd() {
            var msg = '';
            if (document.form1.KPID.value == '') msg += '請輸入職缺\n';
            if (document.form1.DegreeID.selectedIndex == 0) msg += '請選擇學歷\n';
            if (document.form1.ARID.selectedIndex == 0) msg += '請選擇年齡\n';
            if (document.form1.WYID.selectedIndex == 0) msg += '請選擇工作年資\n';
            if (document.form1.MilitaryID.selectedIndex == 0) msg += '請選擇兵役狀況\n';
            if (document.form1.RPNum.value == '') msg += '請輸入需求人數\n';
            else if (!isUnsignedInt(document.form1.RPNum.value)) msg += '需求人數必須為數字\n';
            if (document.form1.ProYear.value == '' && document.form1.ProMonth.value == '') msg += '請輸入估計晉用時間\n';
            else {
                if (!CheckInt(document.form1.ProYear.value)) msg += '晉用時間的年份必須為數字\n';
                if (!CheckInt(document.form1.ProMonth.value)) msg += '晉用時間的月份必須為數字\n';
            }

            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function CheckInt(num) {
            if (num != '') {
                if (!isUnsignedInt(num)) { return false; }
            }
            return true;
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">

        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>首頁&gt;&gt;訓練與就業需求管理&gt;&gt;<font color="#990000">企業需求訪視表
                                    <asp:Label ID="ProType" runat="server"></asp:Label></font>
                            </td>
                        </tr>
                    </table>
                    <table id="SearchTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <table class="table_sch" id="Table2" cellspacing="1" cellpadding="1">
                                    <tr>
                                        <td class="bluecol" width="80">企業名稱
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="SUname" runat="server" Columns="40" MaxLength="50"></asp:TextBox>
                                        </td>
                                        <td class="bluecol">統一編號
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="SIntaxno" runat="server" MaxLength="12"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">訪視類型
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:RadioButtonList ID="SVisitKind" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal"
                                                CssClass="font">
                                                <asp:ListItem Value="1" Selected="True">一般計畫</asp:ListItem>
                                                <asp:ListItem Value="2">訓用合一</asp:ListItem>
                                            </asp:RadioButtonList>
                                            、與企業合作辦理職前訓練、推動營造業事業單位辦理職前培訓
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">縣市
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:CheckBoxList ID="SCTID" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal"
                                                CssClass="font" RepeatColumns="8">
                                            </asp:CheckBoxList>
                                            <input id="HidObj" type="hidden" runat="server" name="HidObj">
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">訪視時間
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:TextBox ID="SVisitDate1" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer"
                                                onclick="javascript:show_calendar('SVisitDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif"
                                                align="top" width="30" height="30">~
                                            <asp:TextBox ID="SVisitDate2" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer"
                                                onclick="javascript:show_calendar('SVisitDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif"
                                                align="top" width="30" height="30">
                                        </td>
                                    </tr>
                                    <tr id="DistTr" runat="server">
                                        <td class="bluecol_need">轄區
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:DropDownList ID="SDistID" runat="server">
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                </table>
                                <p align="center">
                                    <asp:Button ID="Button1" runat="server" Text="查詢"
                                        CssClass="asp_button_S"></asp:Button>
                                </p>
                                <table class="font" id="ResultTable" cellspacing="0" cellpadding="0" width="100%"
                                    border="0" runat="server">
                                    <tr>
                                        <td>
                                            <table class="font" id="DataGridTable1" cellspacing="1" cellpadding="1" width="100%"
                                                border="0" runat="server">
                                                <tr>
                                                    <td>轄區:
                                                        <asp:Label ID="SDistName" runat="server"></asp:Label>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td>
                                                        <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Width="100%" AutoGenerateColumns="False"
                                                            AllowPaging="True">
                                                            <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                                                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                            <Columns>
                                                                <asp:BoundColumn HeaderText="序號">
                                                                    <HeaderStyle HorizontalAlign="Center" Width="25px"></HeaderStyle>
                                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                                </asp:BoundColumn>
                                                                <asp:BoundColumn DataField="DistName" HeaderText="轄區"></asp:BoundColumn>
                                                                <asp:TemplateColumn HeaderText="企業名稱(訪視次數)">
                                                                    <ItemTemplate>
                                                                        <asp:LinkButton ID="LinkButton1" runat="server" ForeColor="Blue" CommandName="view">LinkButton</asp:LinkButton>
                                                                    </ItemTemplate>
                                                                </asp:TemplateColumn>
                                                                <asp:BoundColumn DataField="Intaxno" HeaderText="統一編號"></asp:BoundColumn>
                                                                <asp:BoundColumn DataField="TradeName" HeaderText="行業分類"></asp:BoundColumn>
                                                                <asp:BoundColumn DataField="KEName" HeaderText="企業規模"></asp:BoundColumn>
                                                                <asp:TemplateColumn HeaderText="功能">
                                                                    <ItemTemplate>
                                                                        <asp:Button ID="Button2" runat="server" Text="新增" CommandName="add"></asp:Button>
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
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="center">
                                            <asp:Label ID="msg1" runat="server" ForeColor="Red"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="center">
                                            <asp:Button ID="Button3" runat="server" Text="新增新企業訪視表" CssClass="asp_button_M"></asp:Button>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="DetailTable" cellspacing="1" cellpadding="1" width="100%"  border="0" runat="server">
                        <tr>
                            <td>
                                <table class="table_sch" id="Table3">
                                    <tr>
                                        <td class="bluecol">訪視單位
                                        </td>
                                        <td class="whitecol" width="200">
                                            <asp:Label ID="OrgName" runat="server"></asp:Label><input id="RID" style="width: 78px; height: 22px"
                                                type="hidden" size="7" name="RID" runat="server">
                                        </td>
                                        <td class="bluecol">訪視人員
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="Visiter" runat="server"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need">訪視日期
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="VisitDate" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer"
                                                onclick="javascript:show_calendar('VisitDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif"
                                                align="top" width="30" height="30">
                                        </td>
                                        <td class="bluecol">訪視次數
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="VisitNum" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need">訪視類型
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:RadioButtonList ID="VisitKind" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">一般計畫</asp:ListItem>
                                                <asp:ListItem Value="2">訓用合一</asp:ListItem>
                                            </asp:RadioButtonList>
                                            、與企業合作辦理職前訓練、推動營造業事業單位辦理職前培訓
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need" colspan="4">受訪單位基本資料
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">受訪單位名稱
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="Uname" runat="server" onfocus="this.blur()"></asp:TextBox><input id="BDID"
                                                type="hidden" runat="server">
                                        </td>
                                        <td class="bluecol">統一編號
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="Intaxno" runat="server" Columns="10" onfocus="this.blur()"></asp:TextBox><input
                                                id="Button9" type="button" value="引用事業單位" name="Button9" runat="server">
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">地址
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:TextBox ID="City" runat="server" Columns="13" onfocus="this.blur()"></asp:TextBox><input
                                                id="Zip" style="width: 24px; height: 22px" type="hidden" runat="server"><input
                                                    type="button" value="..." id="Button14" name="Button14" runat="server" disabled="disabled">
                                            <asp:TextBox ID="Addr" runat="server" Columns="35" onfocus="this.blur()"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">行業別
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:DropDownList ID="TradeID" runat="server">
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">現有員工數
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:RadioButtonList ID="KEID" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal"
                                                CssClass="font" Enabled="False">
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">勞保
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:RadioButtonList ID="Labor" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal"
                                                CssClass="font" Enabled="False">
                                                <asp:ListItem Value="1">有</asp:ListItem>
                                                <asp:ListItem Value="0">無</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" colspan="4">受訪者資料
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need">姓名
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="VisitedName" runat="server"></asp:TextBox>
                                        </td>
                                        <td class="bluecol">職務
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="VisitedTitle" runat="server"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need">電話
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="VisitedTel" runat="server"></asp:TextBox>
                                        </td>
                                        <td class="bluecol">傳真
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="VistiedFax" runat="server"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">手機
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="VisitedMob" runat="server"></asp:TextBox>
                                        </td>
                                        <td class="bluecol">E-Mail
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="VisitedMail" runat="server"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need" colspan="4">訪視情形
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol" colspan="4">
                                            <asp:RadioButtonList ID="BVCID" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal"
                                                CssClass="font" RepeatColumns="3">
                                            </asp:RadioButtonList>
                                            <asp:TextBox ID="VisitOther" runat="server" MaxLength="100"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need" colspan="4">後續處理方式
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="TR_TD4" colspan="4">
                                            <table class="font" id="Table6" cellspacing="1" cellpadding="1" width="100%" border="0">
                                                <tr>
                                                    <td class="whitecol">
                                                        <asp:RadioButton ID="BPKind1" runat="server" Text="存檔備查" GroupName="BPKind"></asp:RadioButton>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">
                                                        <asp:RadioButton ID="BPKind2" runat="server" Text="持續聯繫，並e-mail或傳真相關資料提供參考。" GroupName="BPKind"></asp:RadioButton>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="whitecol">
                                                        <asp:RadioButton ID="BPKind3" runat="server" Text="需派員再次前往訪視。" GroupName="BPKind"></asp:RadioButton>(預計訪視日期西元
                                                        <asp:TextBox ID="BPKind_Year" runat="server" Columns="3"></asp:TextBox>年
                                                        <asp:TextBox ID="BPKind_Mon" runat="server" Columns="3"></asp:TextBox>月
                                                        <asp:TextBox ID="BPKind_Day" runat="server" Columns="3"></asp:TextBox>日)
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" colspan="4">用人及職訓需求資料
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">職缺(專業技能)
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:TextBox ID="ProName" runat="server" Columns="30" onfocus="this.blur()"></asp:TextBox><input
                                                id="KPID" style="width: 42px; height: 22px" type="hidden" runat="server"><input
                                                    id="Button12" type="button" value="挑選技能分類" name="Button12" runat="server">
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" colspan="4">資格條件
                                        </td>
                                    </tr>
                                    <tr>
                                        <td colspan="4">
                                            <table class="font" id="Table4" cellspacing="1" cellpadding="1" width="100%" border="0">
                                                <tr>
                                                    <td class="TR_TD4" colspan="3">
                                                        <table class="font" id="Table5" cellspacing="1" cellpadding="1" width="100%" border="0">
                                                            <tr>
                                                                <td class="bluecol">學歷
                                                                </td>
                                                                <td class="bluecol">年齡
                                                                </td>
                                                                <td class="bluecol">工作年資
                                                                </td>
                                                                <td class="bluecol">兵役狀況
                                                                </td>
                                                                <td class="bluecol">證照
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td class="whitecol">
                                                                    <asp:DropDownList ID="DegreeID" runat="server">
                                                                    </asp:DropDownList>
                                                                </td>
                                                                <td class="whitecol">
                                                                    <asp:DropDownList ID="ARID" runat="server">
                                                                    </asp:DropDownList>
                                                                </td>
                                                                <td class="whitecol">
                                                                    <asp:DropDownList ID="WYID" runat="server">
                                                                    </asp:DropDownList>
                                                                </td>
                                                                <td class="whitecol">
                                                                    <asp:DropDownList ID="MilitaryID" runat="server">
                                                                    </asp:DropDownList>
                                                                </td>
                                                                <td class="whitecol">
                                                                    <asp:TextBox ID="License" runat="server"></asp:TextBox>
                                                                </td>
                                                            </tr>
                                                        </table>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="bluecol" width="100">人數
                                                    </td>
                                                    <td class="whitecol" width="500" colspan="2">
                                                        <asp:TextBox ID="RPNum" runat="server" Columns="5"></asp:TextBox>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="bluecol">估計晉用時間
                                                    </td>
                                                    <td class="whitecol">
                                                        <asp:TextBox ID="ProYear" runat="server" Columns="5"></asp:TextBox>年
                                                        <asp:TextBox ID="ProMonth" runat="server" Columns="5"></asp:TextBox>月
                                                    </td>
                                                    <td class="whitecol" align="right" width="150">
                                                        <asp:Button ID="Button10" runat="server" Text="新增職類需求"></asp:Button>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td colspan="3">
                                                        <asp:DataGrid ID="DataGrid2" runat="server" CssClass="font" Width="100%" AutoGenerateColumns="False">
                                                            <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                                                            <ItemStyle BackColor="White"></ItemStyle>
                                                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                            <Columns>
                                                                <asp:BoundColumn HeaderText="職缺(專業技能)"></asp:BoundColumn>
                                                                <asp:BoundColumn HeaderText="學歷"></asp:BoundColumn>
                                                                <asp:BoundColumn HeaderText="年齡"></asp:BoundColumn>
                                                                <asp:BoundColumn HeaderText="工作年資"></asp:BoundColumn>
                                                                <asp:BoundColumn HeaderText="兵役"></asp:BoundColumn>
                                                                <asp:BoundColumn DataField="License" HeaderText="證照"></asp:BoundColumn>
                                                                <asp:BoundColumn DataField="RPNum" HeaderText="人數"></asp:BoundColumn>
                                                                <asp:BoundColumn HeaderText="預估晉用時間"></asp:BoundColumn>
                                                                <asp:TemplateColumn HeaderText="功能">
                                                                    <ItemTemplate>
                                                                        <asp:Button ID="Button5" runat="server" Text="刪除" CommandName="del"></asp:Button>
                                                                    </ItemTemplate>
                                                                </asp:TemplateColumn>
                                                            </Columns>
                                                        </asp:DataGrid>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <asp:Button ID="Button4" runat="server" Text="儲存訪視表" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="Button11"
                                    runat="server" Text="不儲存回上一頁" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="ListTable" cellspacing="1" cellpadding="1" width="100%" border="0"
                        runat="server">
                        <tr>
                            <td colspan="2">轄區：
                                <asp:Label ID="LDistName" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">機構名稱：
                                <asp:Label ID="LUname" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td width="250">行業分類：
                                <asp:Label ID="LTradeID" runat="server"></asp:Label>
                            </td>
                            <td width="450">機構規模：
                                <asp:Label ID="LKEID" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <asp:DataGrid ID="DataGrid3" runat="server" CssClass="font" Width="100%" AutoGenerateColumns="False">
                                    <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號">
                                            <HeaderStyle Width="25px"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="VisitDate" HeaderText="訪視日期" DataFormatString="{0:d}">
                                            <HeaderStyle Width="60px"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Name" HeaderText="訪視人員">
                                            <HeaderStyle Width="60px"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="訪視類型">
                                            <HeaderStyle Width="60px"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="BVCName" HeaderText="訪視結果">
                                            <HeaderStyle Width="150px"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="後續處理方式"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <HeaderStyle Width="90px"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Button ID="Button7" runat="server" Text="修改" CommandName="edit"></asp:Button>
                                                <asp:Button ID="Button8" runat="server" Text="刪除" CommandName="del"></asp:Button>
                                                <asp:Button ID="Button13" runat="server" Text="檢視" CommandName="view"></asp:Button>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="2">
                                <asp:Button ID="Button6" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>

    </form>
</body>
</html>
