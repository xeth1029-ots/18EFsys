<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_002_add.aspx.vb" Inherits="WDAIIP.TC_01_002_add" MaintainScrollPositionOnPostback="true" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訓練機構設定</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta http-equiv="Content-Type" content="text/html; charset=big5">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <%--  Page ValidateRequest="false" --%>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function ChangeMode(num) {
            var Table1 = document.getElementById('Table1');
            var HistoryTable = document.getElementById('HistoryTable');
            if (Table1 && HistoryTable) {
                if (num == 1) {
                    Table1.style.display = 'inline';
                    HistoryTable.style.display = 'none';
                }
                if (num == 2) {
                    Table1.style.display = 'none';
                    HistoryTable.style.display = 'inline';
                }
            }
        }

        function SetLastYearExeRate() {
            obj2 = document.form1.txtLastYearExeRate;
            obj3 = document.form1.ExeRate;
            if (obj3) {
                if (document.getElementById("LastYearExeRate").children[0].checked) {
                    obj2.disabled = true;
                    obj3.disabled = true;
                    //obj1.style.display='none';
                }
                else {
                    if (obj2.value == '') {
                        obj2.value = 0;
                    }
                    obj2.disabled = false;
                    obj3.disabled = false;
                    //obj1.style.display='inline';
                }
            }
            else {
                if (document.getElementById("LastYearExeRate").children[0].checked) {
                    obj2.disabled = true;
                    //obj1.style.display='none';
                }
                else {
                    if (obj2.value == '') {
                        obj2.value = 0;
                    }
                    obj2.disabled = false;
                }
            }
        }

        //統一編號檢查實作(統編 規則) 'isValidTWBID isTWBID TWBID COMIDNO true:OK false:NG
        function CheckTWBID(source, args) {
            args.IsValid = true;
            //請輸入正確的統一編號
            //欄位有效時再行驗證,無效時也要進行驗證
            var TBID = document.getElementById("TBID");
            var RIDValue = document.getElementById("RIDValue");
            var TBIDval = TBID.value;
            if (TBID.value.length > 8) { TBIDval = TBID.value.substr(0, 8); }
            //非委訓單位不檢核統編
            if (RIDValue.value != "" && RIDValue.value.length == 1) { return; }
            if (TBID.value == "00000000") { args.IsValid = false; }
            if (TBID.value.length < 8 || TBID.value.length > 10) { args.IsValid = false; }
            if (!isValidTWBID(TBIDval)) { args.IsValid = false; }
        }

        function ChcekRID(source, args) {
            //debugger;
            var TBplan = document.getElementById("TBplan");
            var RIDValue = document.getElementById("RIDValue");
            if (RIDValue.value == '' || TBplan.value == '') {
                args.IsValid = false;
            } else {
                args.IsValid = true;
            }
        }

        function wopen(url, name, width, height, k) {
            LeftPosition = (screen.width) ? (screen.width - width) / 2 : 0;
            TopPosition = (screen.availHeight) ? (screen.availHeight - height - 28) / 2 : 0;
            window.open(url, name, 'top=' + TopPosition + ',left=' + LeftPosition + ',width=' + width + ',height=' + height + ',resizable=0,scrollbars=' + k + ',status=0');
        }

        /*function chkdata(source, args){if (checkMaxLen(form1.ComSumm.value,300*2)){//若超過則傳回trueargs.IsValid = false;} else {args.IsValid = true;}}*/

        function chk_ActNo(source, args) {
            var TB_ActNo = document.getElementById("TB_ActNo");
            if (checkMaxLen(TB_ActNo.value, 20 * 2)) {
                //若超過則傳回true
                args.IsValid = false;
            } else {
                args.IsValid = true;
            }
        }

        function ClearBtn() {
            var RIDValue = document.getElementById("RIDValue");
            var PlanIDValue = document.getElementById("PlanIDValue");
            var TBplan = document.getElementById("TBplan");
            RIDValue.value = '';
            PlanIDValue.value = '';
            TBplan.value = '';
        }

        //2009-05-19 add 依需求只允許輸入整數(排除 00)
        function CheckZIPB3_1(source, args) {
            var ZipCODEB3 = document.getElementById("ZipCODEB3");
            if (!ZipCODEB3) { return; }
            if (isBlank(ZipCODEB3)) { args.IsValid = true; return; }
            if (isNaN(parseInt(trim(ZipCODEB3.value), 10))) { args.IsValid = false; return; }
            if (!isUnsignedInt(trim(ZipCODEB3.value))) { args.IsValid = false; return; }
            if (parseInt(trim(ZipCODEB3.value), 10) < 1) { args.IsValid = false; return; }
            args.IsValid = true; return;
        }

        //2009-05-20 add 依需求只允許輸入兩碼
        function CheckZIPB3_2(source, args) {
            var ZipCODEB3 = document.getElementById("ZipCODEB3");
            if (!ZipCODEB3) { return; }
            if (trim(ZipCODEB3.value) == "") { args.IsValid = true; return; }
            if (trim(ZipCODEB3.value).length == 2) { args.IsValid = true; return; }
            if (trim(ZipCODEB3.value).length == 3) { args.IsValid = true; return; }
            args.IsValid = false; return;
        }

        //2009-05-19 add 依需求只允許輸入整數(排除 00)
        function CheckZIPB3_1b(source, args) {
            var ZipCODEB3 = document.getElementById("ZipCODEB3_Org");
            if (!ZipCODEB3) { return; }
            if (isBlank(ZipCODEB3)) { args.IsValid = true; return; }
            if (isNaN(parseInt(trim(ZipCODEB3.value), 10))) { args.IsValid = false; return; }
            if (!isUnsignedInt(trim(ZipCODEB3.value))) { args.IsValid = false; return; }
            if (parseInt(trim(ZipCODEB3.value), 10) < 1) { args.IsValid = false; return; }
            args.IsValid = true; return;
        }

        //2009-05-20 add 依需求只允許輸入兩碼
        function CheckZIPB3_2b(source, args) {
            var ZipCODEB3 = document.getElementById("ZipCODEB3_Org");
            if (!ZipCODEB3) { return; }
            if (trim(ZipCODEB3.value) == "") { args.IsValid = true; return; }
            if (trim(ZipCODEB3.value).length == 2) { args.IsValid = true; return; }
            if (trim(ZipCODEB3.value).length == 3) { args.IsValid = true; return; }
            args.IsValid = false; return;
        }

        //檢查zipcode(City欄位名,Zip欄位名,Zip輸入內容)
        function getZipName(CityID, ZipID, ZipValue) {
            if (!isBlank(ZipID)) {
                if (isUnsignedInt(ZipValue) && ZipValue.length == 3) {
                    ifmCheckZip.document.form1.hidCityID.value = CityID;
                    ifmCheckZip.document.form1.hidZipID.value = ZipID.id;
                    ifmCheckZip.document.form1.hidValue.value = ZipValue;
                    ifmCheckZip.document.form1.submit();
                } else {
                    ZipID.value = '';
                    document.getElementById(CityID).value = '';
                    ZipID.focus();
                    alert('查無' + ZipValue + '郵遞區號!');
                }
            } else {
                document.getElementById(CityID).value = '';
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" runat="server" class="font" cellspacing="0" cellpadding="0" border="0" width="100%">
            <tr>
                <td>
                    <asp:CustomValidator ID="CustomValidator4" runat="server" CssClass="font" ClientValidationFunction="chk_ActNo" ErrorMessage="保險證號請輸入20字以內" Display="None" ControlToValidate="TB_ActNo"></asp:CustomValidator>
                   <%--<table class="font" id="Table2" width="100%"><tr><td><asp:Label ID="TitleLab1" runat="server"></asp:Label><asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;基本資料設定&gt;&gt;訓練機構設定</asp:Label>- <font color="#990000"><asp:Label ID="lblProecessType" runat="server"></asp:Label></font> <font color="#000000">(<font face="新細明體"><font color="#ff0000">*</font>為必填欄位</font>)</font></td></tr></table>--%>
                </td>
            </tr>
            <tr>
                <td>
                    <table id="MenuTable" runat="server" class="font" style="cursor: pointer" height="20" cellspacing="0" cellpadding="0" border="0">
                        <tr>
                            <td onclick="ChangeMode(1);" width="1" background="../../images/BookMark_01.gif"><font size="2"></font></td>
                            <td onclick="ChangeMode(1);" align="center" width="100" background="../../images/BookMark_02.gif"><font size="2">基本資料</font></td>
                            <td onclick="ChangeMode(1);" width="11" background="../../images/BookMark_03.gif"><font size="2"></font></td>
                            <td onclick="ChangeMode(2);" width="1" background="../../images/BookMark_01.gif"><font size="2"></font></td>
                            <td onclick="ChangeMode(2);" align="center" width="100" background="../../images/BookMark_02.gif"><font size="2">辦訓記錄</font></td>
                            <td onclick="ChangeMode(2);" width="11" background="../../images/BookMark_03.gif"><font size="2"></font></td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table class="table_nw" id="Table1" width="100%" runat="server" cellpadding="1" cellspacing="1">
                        <tr>
                            <td class="table_title" colspan="4">訓練機構共同資料<asp:Label ID="GWOrgKind" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td width="10%" class="bluecol_need">轄區分署 </td>
                            <td width="50%" colspan="3" class="whitecol">
                                <asp:DropDownList ID="DistrictList" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">機構名稱 全銜</td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="TBtitle" runat="server" Width="80%" MaxLength="100"></asp:TextBox>
                                <asp:Button ID="BT_CHGORG" runat="server" Text="變更"></asp:Button>
                                <asp:RequiredFieldValidator ID="RFValidator1" runat="server" ErrorMessage="請輸入機構名稱全銜" Display="None" ControlToValidate="TBtitle"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">機構別 </td>
                            <td colspan="3" class="whitecol">
                                <asp:DropDownList ID="OrgKindList" runat="server" CssClass="font"></asp:DropDownList>
                                <asp:RequiredFieldValidator ID="R_OrgKindList" runat="server" ErrorMessage="請選擇機構別" Display="None" ControlToValidate="OrgKindList"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">統一編號 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TBID" runat="server" MaxLength="10" Width="60%"></asp:TextBox>
                                <asp:CustomValidator ID="CustomValidator1" runat="server" ClientValidationFunction="CheckTWBID" ErrorMessage="請輸入正確的統一編號" Display="None"></asp:CustomValidator>
                                <asp:RegularExpressionValidator ID="IDValue" runat="server" ErrorMessage="統一編號請填寫八位任意數字" Display="None" ControlToValidate="TBID" ValidationExpression="[0-9]{8,10}"></asp:RegularExpressionValidator>
                                <asp:RequiredFieldValidator ID="IDa" runat="server" ErrorMessage="請輸入統一編號" Display="None" ControlToValidate="TBID"></asp:RequiredFieldValidator>
                            </td>
                            <td class="bluecol_need">立案證號 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TBseqno" runat="server" Width="60%" MaxLength="100"></asp:TextBox>
                                <asp:RequiredFieldValidator ID="seqno" runat="server" ErrorMessage="請輸入立案登記編號" Display="None" ControlToValidate="TBseqno"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr id="TPlanID28A" runat="server">
                            <td class="bluecol">&nbsp;<asp:Label ID="LabLastYear" runat="server"></asp:Label></td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="LastYearExeRate" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="-1" Selected="True">否</asp:ListItem>
                                    <asp:ListItem Value="1">是</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                            <td width="100" class="bluecol">&nbsp;<asp:Label ID="LabLastYear2" runat="server"></asp:Label></td>
                            <td class="whitecol">
                                <asp:TextBox ID="txtLastYearExeRate" runat="server" Width="25%"></asp:TextBox>%
                                <input id="LastYear" type="hidden" runat="server" size="1">
                                <asp:Button ID="ExeRate" runat="server" Text="計算執行率" CausesValidation="False"></asp:Button>
                            </td>
                        </tr>
                        <tr id="TPlanID28B" runat="server">
                            <td class="bluecol">通過訓練品質評核版本 </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="IsConTTQS" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="0" Selected="True">訓練機構版</asp:ListItem>
                                    <asp:ListItem Value="1">外訓版</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                            <td colspan="2" class="whitecol"><b>若機構別非勞工團體，則不必輸入銀行、分行、金融機構帳號、銀行戶名</b> </td>
                        </tr>
                        <tr id="TPlanID28C" runat="server">
                            <td class="bluecol">&nbsp;<asp:Label ID="Label1" runat="server">銀行(庫局)</asp:Label></td>
                            <td class="whitecol">
                                <asp:TextBox ID="BankName" runat="server" Width="40%" MaxLength="100"></asp:TextBox></td>
                            <td class="bluecol">&nbsp;<asp:Label ID="Label3" runat="server">分行(支庫局)</asp:Label></td>
                            <td class="whitecol">
                                <asp:TextBox ID="ExBankName" runat="server" Width="40%" MaxLength="100"></asp:TextBox></td>
                        </tr>
                        <tr id="TPlanID28D" runat="server">
                            <td class="bluecol">&nbsp;<asp:Label ID="Label2" runat="server">金融機構帳號</asp:Label></td>
                            <td class="whitecol">
                                <asp:TextBox ID="AccNo" runat="server" Width="40%" MaxLength="100"></asp:TextBox></td>
                            <td class="bluecol">&nbsp;<asp:Label ID="Label4" runat="server">銀行戶名</asp:Label></td>
                            <td class="whitecol">
                                <asp:TextBox ID="AccName" runat="server" Width="40%" MaxLength="100"></asp:TextBox></td>
                        </tr>
                        <tr style="display: none">
                            <td class="font" align="center" colspan="4">
                                <table class="table_nw" id="Table3" width="100%" runat="server">
                                    <tr>
                                        <td class="table_title" colspan="5">機構年度評鑑資料[<asp:Label ID="LYears" runat="server"></asp:Label>] </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">單位能力指標星等 </td>
                                        <td colspan="2" class="whitecol">
                                            <asp:Label ID="Point01A" runat="server"></asp:Label>&nbsp; </td>
                                        <td class="bluecol">單位能力指標分數 </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Point01B" runat="server"></asp:Label>&nbsp; </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">就業表現指標星等 </td>
                                        <td colspan="2" class="whitecol">
                                            <asp:Label ID="Point02A" runat="server"></asp:Label>&nbsp; </td>
                                        <td class="bluecol">就業表現指標分數 </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Point02B" runat="server"></asp:Label>&nbsp; </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">學員問卷滿意度<br>
                                            指標星等</td>
                                        <td colspan="2" class="whitecol">
                                            <asp:Label ID="Point03A" runat="server"></asp:Label>&nbsp;</td>
                                        <td class="bluecol">學員問卷滿意度<br>
                                            指標分數</td>
                                        <td class="whitecol">
                                            <asp:Label ID="Point03B" runat="server"></asp:Label>&nbsp;</td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td class="table_title" colspan="4">訓練機構承辦人資料 </td>
                        </tr>
                        <tr id="tr_level_list" runat="server">
                            <td class="bluecol_need">階層</td>
                            <td colspan="3" class="whitecol">
                                <asp:DropDownList ID="level_list" runat="server">
                                    <asp:ListItem Value="2">委訓</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">隸屬機構 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="TBplan" runat="server" Width="76%" onfocus="this.blur()"></asp:TextBox>
                                <input id="choice_button" onclick="javascript: wopen('../../Common/LevPlan.aspx', '計畫階段', 850, 570, 1)" type="button" value="選擇" name="choice_button" runat="server">
                                <input id="btn_clear" onclick="ClearBtn();" type="button" value="清除" name="Button3" runat="server">
                                <asp:CustomValidator ID="CustomValidator2" runat="server" ClientValidationFunction="ChcekRID" ErrorMessage="請選擇管控單位" Display="None"></asp:CustomValidator>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">是否為管控單位 </td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="IsConUnit" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="1">是</asp:ListItem>
                                    <asp:ListItem Value="0" Selected="True">否</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">分支單位名稱 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="TB_OrgPName" runat="server" Width="76%" MaxLength="100"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">地址 </td>
                            <td colspan="3" class="whitecol">
                                <input id="city_code" name="city_code" runat="server" maxlength="3" />－
                                <input id="ZipCODEB3" maxlength="3" name="ZipCODEB3" runat="server" />
                                <input id="hidZipCODE6W" type="hidden" runat="server" /><input id="hidZipCODEB3_N" type="hidden" runat="server" />
                                <asp:Literal ID="LitZipCODE" runat="server"></asp:Literal><br /><%--郵遞區號查詢--%>
                                <asp:TextBox ID="TBCity" runat="server" Width="20%" onfocus="this.blur()"></asp:TextBox>
                                <input id="Bt1_city_zip" type="button" value="..." name="city_zip" runat="server" />
                                <asp:TextBox ID="TBaddress" runat="server" Width="60%" MaxLength="200"></asp:TextBox>
                                <asp:RequiredFieldValidator ID="city" runat="server" ErrorMessage="請選擇地址縣市" Display="None" ControlToValidate="TBCity"></asp:RequiredFieldValidator>
                                <asp:RequiredFieldValidator ID="address" runat="server" ErrorMessage="請輸入地址" Display="None" ControlToValidate="TBaddress"></asp:RequiredFieldValidator>
                                <asp:RequiredFieldValidator ID="RequiredFieldValidator10" runat="server" ErrorMessage="「地址郵遞區號 後2碼/後3碼」為必填欄位" Display="None" ControlToValidate="ZipCODEB3"></asp:RequiredFieldValidator>
                                <asp:CustomValidator ID="cvCheckZIPB3_1" runat="server" ClientValidationFunction="CheckZIPB3_1" ErrorMessage="「地址郵遞區號 後2碼/後3碼」必須為數字，且不得輸入 0" Display="None"></asp:CustomValidator>
                                <asp:CustomValidator ID="cvCheckZIPB3_2" runat="server" ClientValidationFunction="CheckZIPB3_2" ErrorMessage="「地址郵遞區號 後2碼/後3碼」長度必須為 2碼或3碼(例 01 或 001)" Display="None"></asp:CustomValidator>
                            </td>
                        </tr>
                        <tr id="TPlanID28" runat="server">
                            <td class="bluecol" style="width: 20%">計畫主持人 </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="PlanMaster" runat="server" Width="60%"></asp:TextBox></td>
                            <td class="bluecol" style="width: 20%">主持人電話 </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="PlanMasterPhone" runat="server" Width="60%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">負責人姓名 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TBm_name" runat="server" Width="60%"></asp:TextBox></td>
                            <td class="bluecol">負責人電話 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TBm_Phone" runat="server" Width="60%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">聯絡人姓名 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TBContactName" runat="server" Width="60%"></asp:TextBox>
                                <asp:RequiredFieldValidator ID="ContactName" runat="server" ErrorMessage="請輸入聯絡人姓名" Display="None" ControlToValidate="TBContactName"></asp:RequiredFieldValidator>
                            </td>
                            <td class="bluecol_need">聯絡人電話 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TBtel" runat="server" Width="60%"></asp:TextBox>
                                <asp:RequiredFieldValidator ID="Re_tel" runat="server" ErrorMessage="請輸入聯絡人電話" Display="None" ControlToValidate="TBtel"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">聯絡人行動電話 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TBcontact_cellphone" runat="server" Width="60%"></asp:TextBox></td>
                            <td class="bluecol"></td>
                            <td class="whitecol"></td>
                        </tr>
                        <tr>
                            <td class="bluecol">聯絡人E-MAIL </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TBmail" runat="server" Width="85%"></asp:TextBox>
                                <asp:RegularExpressionValidator ID="mail" runat="server" ErrorMessage="請重新輸入 聯絡人E-MAIL" Display="None" ControlToValidate="TBmail" ValidationExpression="\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"></asp:RegularExpressionValidator>
                            </td>
                            <td class="bluecol">聯絡人傳真 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="ContactFax" runat="server" Width="60%"></asp:TextBox></td>
                        </tr>
                        <tr id="TrTPlanID28F1" runat="server">
                            <td class="bluecol_need">個人資料檔案保管人員 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="staffName" runat="server" Width="60%"></asp:TextBox>
                                <asp:RequiredFieldValidator ID="Req_staffName" runat="server" ErrorMessage="個人資料檔案保管人員 為必填欄位" Display="None" ControlToValidate="staffName"></asp:RequiredFieldValidator>
                            </td>
                            <td class="bluecol_need">個人資料檔案保管人員電話 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="staffPhone" runat="server" Width="60%"></asp:TextBox>
                                <asp:RequiredFieldValidator ID="Req_staffPhone" runat="server" ErrorMessage="個人資料檔案保管人員電話 為必填欄位" Display="None" ControlToValidate="staffPhone"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr id="TrTPlanID28F2" runat="server">
                            <td class="bluecol_need">個人資料檔案保管人員電子郵件 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="staffEmail" runat="server" Width="66%" MaxLength="200"></asp:TextBox>
                                <asp:RequiredFieldValidator ID="Req_staffEmail2" runat="server" ErrorMessage="個人資料檔案保管人員電子郵件 為必填欄位" Display="None" ControlToValidate="staffEmail"></asp:RequiredFieldValidator>
                                <asp:RegularExpressionValidator ID="Reg_staffEmail1" runat="server" ErrorMessage="請重新輸入 個人資料檔案保管人員電子郵件" Display="None" ControlToValidate="staffEmail" ValidationExpression="\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"></asp:RegularExpressionValidator>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">保險證號 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="TB_ActNo" runat="server" Width="44%"></asp:TextBox>
                                <asp:RequiredFieldValidator ID="Re_Field_ActNo" runat="server" ErrorMessage="請輸入保險證號" Display="None" ControlToValidate="TB_ActNo"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練容量 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TB_TrainCap" runat="server" Width="60%"></asp:TextBox></td>
                            <td class="bluecol">消防安檢狀況 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TB_FireControlState" runat="server" Width="60%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">專長訓練職類 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="TB_ProTrainKind" runat="server" Width="66%"></asp:TextBox>
                                <%--<asp:RequiredFieldValidator ID="Re_Field_ProTrainKind" runat="server" ErrorMessage="請輸入專長訓練職類" Display="None" ControlToValidate="TB_ProTrainKind"></asp:RequiredFieldValidator>--%>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">機構簡介 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="ComSumm" runat="server" Width="66%" TextMode="MultiLine" Rows="5" Columns="20"></asp:TextBox>
                                <%--<asp:RequiredFieldValidator ID="Re_Field_ComSumm" runat="server" ErrorMessage="請輸入機構簡介" Display="None" ControlToValidate="ComSumm"></asp:RequiredFieldValidator>--%>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">會員人數 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="MemberNum" runat="server" MaxLength="10" Columns="6" Width="60%"></asp:TextBox>
                                <asp:RegularExpressionValidator ID="chkMemberNum" runat="server" ControlToValidate="MemberNum" Display="None" ErrorMessage="「會員人數」請輸入數字" ValidationExpression="[0-9]{1,10}"></asp:RegularExpressionValidator>
                            </td>
                            <td class="bluecol">勞工保險加保人數<br>
                                - 會員人數 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="ActMemberNum" runat="server" MaxLength="10" Columns="6" Width="60%"></asp:TextBox>
                                <asp:RegularExpressionValidator ID="chkActMemberNum" runat="server" ControlToValidate="ActMemberNum" Display="None" ErrorMessage="「勞工保險加保人數-會員人數」請輸入數字" ValidationExpression="[0-9]{1,10}"></asp:RegularExpressionValidator>
                            </td>

                        </tr>
                        <tr>
                            <td class="bluecol">勞工保險加保人數<br>
                                -員工人數 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="ActStaffNum" runat="server" MaxLength="10" Columns="6" Width="60%"></asp:TextBox>
                                <asp:RegularExpressionValidator ID="chkActStaffNum" runat="server" ControlToValidate="ActStaffNum" Display="None" ErrorMessage="「勞工保險加保人數-員工人數」請輸入數字" ValidationExpression="[0-9]{1,10}"></asp:RegularExpressionValidator>
                            </td>
                            <td class="bluecol"></td>
                            <td class="whitecol"></td>
                        </tr>
                        <tr>
                            <td class="bluecol">&nbsp;無障礙訓練環境說明 </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="LabAccessible" Style="z-index: 0" runat="server" ForeColor="Red">註：無障礙訓練環境說明，勾選是者，請附相關佐證資料。</asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">&nbsp;是否提供無障礙空間 </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="Accessible" Style="z-index: 0" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="Y">是</asp:ListItem>
                                    <asp:ListItem Value="N">否</asp:ListItem>
                                    <asp:ListItem Value="NG" Selected="True">未選擇</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">&nbsp;是否提供適當教材 </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="Textbook" Style="z-index: 0" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="Y">是</asp:ListItem>
                                    <asp:ListItem Value="N">否</asp:ListItem>
                                    <asp:ListItem Value="NG" Selected="True">未選擇</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">&nbsp;是否提供教學輔具 </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="TeachAids" Style="z-index: 0" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="Y">是</asp:ListItem>
                                    <asp:ListItem Value="N">否</asp:ListItem>
                                    <asp:ListItem Value="NG" Selected="True">未選擇</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">&nbsp;是否提供人力協助 </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="HumanHelp" Style="z-index: 0" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="Y">是</asp:ListItem>
                                    <asp:ListItem Value="N">否</asp:ListItem>
                                    <asp:ListItem Value="NG" Selected="True">未選擇</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table class="font" id="HistoryTable" width="100%" runat="server">
                        <tr>
                            <td align="center" colspan="4">
                                <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="DistName" HeaderText="轄區&lt;BR&gt;分署">
                                            <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="PlanYear" HeaderText="年度">
                                            <HeaderStyle Width="15%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="PlanName" HeaderText="訓練計畫">
                                            <HeaderStyle Width="15%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TrainName" HeaderText="訓練職類">
                                            <HeaderStyle Width="15%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassName" HeaderText="班別名稱">
                                            <HeaderStyle Width="15%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TRound" HeaderText="受訓期間">
                                            <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="4">
                                <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr id="TrTPlanID28OrgType1" runat="server">
                <td>
                    <%-- 2018 add  --%>
                    <table class="table_nw" id="tbOrgType1" width="100%" runat="server" cellpadding="1" cellspacing="1">
                        <tr>
                            <td class="table_title" colspan="4">訓練機構屬性設定</td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="20%">計畫別 </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="dl_typeid1" runat="server" AutoPostBack="True">
                                    <asp:ListItem Value="">==請選擇==</asp:ListItem>
                                    <asp:ListItem Value="1">產業人才投資計畫</asp:ListItem>
                                    <asp:ListItem Value="2">提升勞工自主學習計畫</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">機構別 </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="dl_typeid2" runat="server">
                                    <asp:ListItem Value="">==請選擇==</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">立案地址/會址 </td>
                            <td class="whitecol">
                                <input id="city_code_org" maxlength="3" name="city_code_org" runat="server" />－
                                <input id="ZipCODEB3_Org" maxlength="3" name="ZipCODEB3_Org" runat="server" />
                                <input id="hidZipCODE6W_Org" type="hidden" runat="server" /><input id="hidZipCODEB3_Org_N" type="hidden" runat="server" />
                                <asp:Literal ID="LitZipCODEOrg" runat="server"></asp:Literal><br /><%--郵遞區號查詢--%>
                                <asp:TextBox ID="TBCity_Org" runat="server" Width="20%" onfocus="this.blur()"></asp:TextBox>
                                <input id="Bt2_city_zip_org" type="button" value="..." runat="server" />
                                <asp:TextBox ID="TBaddress_Org" runat="server" Width="60%"></asp:TextBox>
                                <asp:RequiredFieldValidator ID="city_org" runat="server" ErrorMessage="請選擇立案縣市" Display="None" ControlToValidate="TBCity_Org"></asp:RequiredFieldValidator>
                                <asp:RequiredFieldValidator ID="address_org" runat="server" ErrorMessage="請輸入立案地址" Display="None" ControlToValidate="TBaddress_Org"></asp:RequiredFieldValidator>
                                <asp:RequiredFieldValidator ID="city_zip_valid" runat="server" ErrorMessage="「立案郵遞區號 後2碼/後3碼」為必填欄位" Display="None" ControlToValidate="ZipCODEB3_Org"></asp:RequiredFieldValidator>
                                <asp:CustomValidator ID="cvCheckZIPB3_1b" runat="server" ClientValidationFunction="CheckZIPB3_1b" ErrorMessage="「立案郵遞區號 後2碼/後3碼」必須為數字，且不得輸入 0/00/000" Display="None"></asp:CustomValidator>
                                <asp:CustomValidator ID="cvCheckZIPB3_2b" runat="server" ClientValidationFunction="CheckZIPB3_2b" ErrorMessage="「立案郵遞區號 後2碼/後3碼」長度必須為 2碼或3碼(例 01 或 001)" Display="None"></asp:CustomValidator>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td align="center" colspan="5" class="whitecol">
                    <div align="center">
                        <br />
                        <asp:ValidationSummary ID="Summary" runat="server" DisplayMode="List" ShowSummary="False" ShowMessageBox="True"></asp:ValidationSummary>
                        <asp:Button ID="bt_save" runat="server" Text="儲存" CssClass="asp_Export_M"></asp:Button>
                        <asp:Button ID="Button1" runat="server" Text="回上一頁" CssClass="asp_button_M" CausesValidation="False"></asp:Button>
                    </div>
                </td>
            </tr>
        </table>
        <br />
        <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
        <input id="PlanIDValue" type="hidden" name="PlanIDValue" runat="server" />
        <input id="Re_ID" type="hidden" name="Re_ID" runat="server" />
        <input id="OrgIDValue" type="hidden" size="5" runat="server" />
        <input id="HidComidno" type="hidden" size="5" runat="server" />
        <iframe id="ifmCheckZip" name="ifmCheckZip" src="../../common/CheckZip.aspx" width="0%" height="0%" title="X"></iframe>
    </form>
</body>
</html>

