<%@ Page Language="vb" AutoEventWireup="false" EnableEventValidation="true" CodeBehind="TC_01_027_add.aspx.vb" Inherits="WDAIIP.TC_01_027_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>師資資料維護</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker2.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        //中翻英
        function sentName() {
            var msg = '';
            var TeachCName = document.getElementById('TeachCName');
            var strName = TeachCName.value;
            if (isBlank(TeachCName)) {
                msg = '請先輸入中文姓名!';
            } else if (strName.indexOf(" ") >= 0 || strName.indexOf("　") >= 0) {
                msg = '中文姓名不可輸入空格!';
            } else if (strName.length > 8) {
                msg = '中文姓名不可超過8個字!';
            }
            if (msg != '') {
                alert(msg);
            } else {
                window.open('../../common/Translation.aspx?&sn=teach&name=' + escape(strName) + '&field=TeachEName', "sch", 'width=750,height=600,top=200,left=450,location=0,status=0,menubar=0,scrollbars=1,resizable=0,scrollbars=0');
            }
            return false;
        }

        //查詢身分證button
        function wopen(url, name, width, height, k) {
            height = screen.availHeight ? (screen.availHeight * 8 / 10) : height;
            width = screen.width ? (screen.width * 5 / 10) : width;
            LeftPosition = (screen.width) ? (screen.width - width) / 2 : 0;
            TopPosition = (screen.availHeight) ? (screen.availHeight - height - 28) / 2 : 0;
            window.open(url, name, 'top=' + TopPosition + ',left=' + LeftPosition + ',width=' + width + ',height=' + height + ',resizable=0,scrollbars=' + k + ',status=0');
        }

        //判斷性別
        function SexChoice() {
            var xSex = document.form1.IDNO.value;
            switch (xSex.substr(1, 1)) {
                case '1':
                    document.form1.Sex[0].checked = true;
                    break;
                case '2':
                    document.form1.Sex[1].checked = true;
                    break;
                default:
                    break;
            }
        }

        //職稱選擇用txtbox輸入
        function change() {
            var IVID = document.getElementById('IVID');
            var R2 = document.getElementById('R2');
            var R1 = document.getElementById('R1');
            var Invest1 = document.getElementById('Invest1');
            if (R2.checked = true) {
                R1.checked = false;
                if (IVID) {
                    IVID.selectedIndex = 0;
                    IVID.disabled = true;
                }
                Invest1.disabled = false;
            }
        }

        //職稱選擇用選單輸入
        function change2() {
            var IVID = document.getElementById('IVID');
            var R1 = document.getElementById('R1');
            var R2 = document.getElementById('R2');
            var Invest1 = document.getElementById('Invest1');
            if (R1.checked = true) {
                R2.checked = false;
                Invest1.value = '';
                Invest1.disabled = true;
                if (IVID) {
                    IVID.disabled = false;
                }
            }
        }

        //判斷師資別選項
        function KindIDChoice() {
            document.getElementById('KindID_TD2').title = document.form1.KindID(document.form1.KindID.selectedIndex).text;
            document.form1.KindID.title = document.form1.KindID(document.form1.KindID.selectedIndex).text;
        }

        //限定textbox的欄位長度
        function checkTextLength(obj, length) {
            if (obj.value.length > length) {
                obj.value = obj.value.substring(0, length);
                alert("限欄位長度不能大於" + length + "個字元(含空白字元)，超出字元將自動截斷");
            }
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
                    //alert('查無' + ZipValue + '郵遞區號!');
                }
            } else {
                document.getElementById(CityID).value = '';
            }
        }

        function PL_focusState1() {
            //ForeColor
            //var msgX1 = "請優先填寫與課程相關之專業技術類證照，若無證照資訊請填寫「無」。";
            var hid_PLMsgX1 = document.getElementById("hid_PLMsgX1");
            var msgX1 = hid_PLMsgX1.value;
            var ProLicense1 = document.getElementById("ProLicense1");
            if (ProLicense1.value == msgX1) {
                ProLicense1.value = "";
                ProLicense1.style.color = "#000000";
            }
            else if (ProLicense1.value == "") {
                ProLicense1.value = msgX1;
                ProLicense1.style.color = "#666666";
            }
        }

        function PL_focusState2() {
            //ForeColor
            //var msgX1 = "請優先填寫與課程相關之專業技術類證照，若無證照資訊請填寫「無」。";
            var hid_PLMsgX1 = document.getElementById("hid_PLMsgX1");
            var msgX1 = hid_PLMsgX1.value;
            var ProLicense2 = document.getElementById("ProLicense2");
            if (ProLicense2.value == msgX1) {
                ProLicense2.value = "";
                ProLicense2.style.color = "#000000";
            }
            else if (ProLicense2.value == "") {
                ProLicense2.value = msgX1;
                ProLicense2.style.color = "#666666";
            }
        }

        //檢查儲存
        function chkdata() {
            var msg = '';
            //var msgX1 = "請優先填寫與課程相關之專業技術類證照，若無證照資訊請填寫「無」。";
            var hid_PLMsgX1 = document.getElementById("hid_PLMsgX1");
            var msgX1 = hid_PLMsgX1.value;
            var ProLicense1 = document.getElementById('ProLicense1');
            var ProLicense2 = document.getElementById('ProLicense2');
            var TPlanID = document.getElementById("TPlanID");
            var ExpUnit1 = document.getElementById("ExpUnit1");
            var ExpYears1 = document.getElementById("ExpYears1");
            var ExpMonths1 = document.getElementById("ExpMonths1");
            var Specialty1 = document.getElementById("Specialty1");
            var tINV1 = document.getElementById("tINV1");
            var tINV1_PLval = tINV1.attributes["placeholder"].value;
            var TeacherID = document.getElementById("TeacherID");
            var TeachCName = document.getElementById("TeachCName");
            var IDNO = document.getElementById("IDNO");
            var birthday = document.getElementById("birthday");
            var TB_career_id = document.getElementById("TB_career_id");
            var SchoolName = document.getElementById("SchoolName");
            var Department = document.getElementById("Department");

            if (ExpUnit1.value == '') msg += '請輸入 服務單位1\n';
            if (ExpYears1.value == '') msg += '請輸入 年資的年1\n';
            if (ExpMonths1.value == '') msg += '請輸入 年資的月1\n';
            if (Specialty1.value == '') msg += '請輸入 專長1\n';
            if (tINV1.value == '') msg += '請輸入 職稱1\n';
            if (msg == '' && tINV1_PLval != '' && tINV1.value == tINV1_PLval) msg += '請輸入 職稱1\n';
            if (TeacherID.value == '') msg = msg + '請輸入講師代碼!\n';
            if (TeachCName.value == '') msg = msg + '請輸入講師姓名!\n';
            if (IDNO.value == '') {
                msg += '請輸入身分證號碼!\n';
            } else {
                if (document.form1.PassPortNO.selectedIndex == 0) {
                    if (!checkId(IDNO.value)) {
                        if (!confirm('身分證號碼錯誤，是否確定要儲存?')) {
                            msg += '身分證號碼錯誤\n';
                        }
                    }
                }
            }
            if (birthday.value == '') {
                msg += '請輸入出生日期!\n';  //edit，by:20181113
            }
            else {
                var flag_NG_birth = (bl_rocYear == "Y") ? !(!checkRocDate(birthday.value)) : (!checkDate(birthday.value));
                if (flag_NG_birth) { msg += '出生日期格式不正確!\n'; }
            }

            if (TB_career_id.value == '') {
                msg = msg + '請選擇主要職類!\n';
            }
            //假如是委訓
            if (document.form1.LID.value == '2') {
                if (TPlanID.value != '28') {
                    //企訓專用
                    if (document.form1.Invest1) {
                        if (document.form1.Invest1.value == '') msg += '請輸入職稱\n';
                    }
                }
            } else {
                var IVID = document.getElementById('IVID');
                if (IVID) {
                    if (document.form1.KindEngage.selectedIndex != 2 || TPlanID.value == '28') {
                        if (IVID.selectedIndex == 0) msg += '請選擇職稱\n';
                    } else {
                        if (IVID.selectedIndex == 0 && document.form1.Invest1.value == '') msg += '請選擇職稱\n';
                    }
                }
            }
            if (document.form1.KindEngage.selectedIndex == 0) msg = msg + '請選擇內外聘!\n';
            if (document.form1.LID.value != '2') {
                if (document.form1.KindID.selectedIndex == '0') msg = msg + '請選擇師資別!\n';
            }
            if (document.form1.DegreeID.selectedIndex == '0') msg = msg + '請選擇最高學歷!\n';
            if (document.form1.GraduateStatus.selectedIndex == '0') msg = msg + '請選擇畢業狀況!\n';
            if (SchoolName.value == "") { msg += '請輸入 學校名稱!\n'; }
            if (Department.value == "") { msg += '請輸入 科系名稱!\n'; }
            if (document.form1.Phone.value == '') msg += '請輸入聯絡電話\n';

            $('#city_code').val($.trim($('#city_code').val()));
            $('#AddressZIPB3').val($.trim($('#AddressZIPB3').val()));
            $('#Address').val($.trim($('#Address').val()));
            //'city_code'AddressZIPB3'hidAddressZIP6W'Litcity_code'TBCity'bt_openZip1'Address
            if ($('#city_code').val() == '') msg += '請輸入通訊地址的郵遞區號前3碼\n';
            //checkzip23 郵遞區號
            msg += checkzip23(true, '通訊地址', 'AddressZIPB3');
            if ($('#Address').val() == '') msg += '請輸入通訊地址\n';

            if (document.form1.WorkOrg.value == '') msg += '請輸入服務單位\n';
            if (document.form1.ExpYears.value != '') {
                if (!isUnsignedInt(document.form1.ExpYears.value))
                    msg += '年資必須為數字\n';
            }
            if (document.form1.WorkPhone.value == '') msg += '請輸入服務單位電話\n';

            //'city_code1'WorkZIPB3'hidWorkZIP6W'Litcity_code1'TBCity1'bt_openZip2'WorkAddr
            $('#city_code1').val($.trim($('#city_code1').val()));
            $('#WorkZIPB3').val($.trim($('#WorkZIPB3').val()));
            $('#WorkAddr').val($.trim($('#WorkAddr').val()));
            if ($('#city_code1').val() == '') msg += '請輸入服務單位地址 郵遞區號前3碼\n';
            //checkzip23 郵遞區號
            msg += checkzip23(true, '服務單位地址', 'WorkZIPB3');
            if ($('#WorkAddr').val() == '') msg += '請輸入服務單位地址\n';

            if (document.form1.ExpYears1.value != '') {
                if (!isUnsignedInt(document.form1.ExpYears1.value))
                    msg += '經歷年資1必須為數字\n';
            }
            if (document.form1.ExpYears2.value != '') {
                if (!isUnsignedInt(document.form1.ExpYears2.value))
                    msg += '經歷年資2必須為數字\n';
            }
            if (document.form1.ExpYears3.value != '') {
                if (!isUnsignedInt(document.form1.ExpYears3.value))
                    msg += '經歷年資3必須為數字\n';
            }
            if (document.form1.ExpSDate1.value != '') {
                if (bl_rocYear == "Y") {
                    if (!checkRocDate(document.form1.ExpSDate1.value)) msg += '服務期間1起始日期不正確\n';  //edit，by:20181022
                }
                else {
                    if (!checkDate(document.form1.ExpSDate1.value)) msg += '服務期間1起始日期不正確\n';
                }
            }
            if (document.form1.ExpEDate1.value != '') {
                if (bl_rocYear == "Y") {
                    if (!checkRocDate(document.form1.ExpEDate1.value)) msg += '服務期間1終至日期不正確\n';  //edit，by:20181022
                }
                else {
                    if (!checkDate(document.form1.ExpEDate1.value)) msg += '服務期間1終至日期不正確\n';
                }
            }
            if (document.form1.ExpSDate2.value != '') {
                if (bl_rocYear == "Y") {
                    if (!checkRocDate(document.form1.ExpSDate2.value)) msg += '服務期間2起始日期不正確\n';  //edit，by:20181022
                }
                else {
                    if (!checkDate(document.form1.ExpSDate2.value)) msg += '服務期間2起始日期不正確\n';
                }
            }
            if (document.form1.ExpEDate2.value != '') {
                if (bl_rocYear == "Y") {
                    if (!checkRocDate(document.form1.ExpEDate2.value)) msg += '服務期間2終至日期不正確\n';  //edit，by:20181022
                }
                else {
                    if (!checkDate(document.form1.ExpEDate2.value)) msg += '服務期間2終至日期不正確\n';
                }
            }
            if (document.form1.ExpSDate3.value != '') {
                if (bl_rocYear == "Y") {
                    if (!checkRocDate(document.form1.ExpSDate3.value)) msg += '服務期間3起始日期不正確\n';  //edit，by:20181022
                }
                else {
                    if (!checkDate(document.form1.ExpSDate3.value)) msg += '服務期間3起始日期不正確\n';
                }
            }
            if (document.form1.ExpEDate3.value != '') {
                if (bl_rocYear == "Y") {
                    if (!checkRocDate(document.form1.ExpEDate3.value)) msg += '服務期間3終至日期不正確\n';  //edit，by:20181022
                }
                else {
                    if (!checkDate(document.form1.ExpEDate3.value)) msg += '服務期間3終至日期不正確\n';
                }
            }
            if (document.form1.WorkStatus.selectedIndex == '0') {
                msg = msg + '請選擇任職狀況!\n';
            }
            if (checkMaxLen2(document.getElementById('Specialty1').value, 250)) {
                msg += '【專業1】長度不可超過250字元\n';
            }
            if (checkMaxLen2(document.getElementById('Specialty2').value, 250)) {
                msg += '【專業2】長度不可超過250字元\n';
            }
            if (checkMaxLen2(document.getElementById('Specialty3').value, 250)) {
                msg += '【專業3】長度不可超過250字元\n';
            }
            if (checkMaxLen2(document.getElementById('Specialty4').value, 250)) {
                msg += '【專業4】長度不可超過250字元\n';
            }
            if (checkMaxLen2(document.getElementById('Specialty5').value, 250)) {
                msg += '【專業5】長度不可超過250字元\n';
            }
            if (checkMaxLen2(document.getElementById('TransBook').value, 100)) {
                msg += '【譯著】長度不可超過100字元\n';
            }
            if (checkMaxLen2(ProLicense1.value, 200)) {
                msg += '【專業證照-政府機關辦理相關證照或檢定】長度不可超過200字元\n';
            }
            if (ProLicense1.value == msgX1 || ProLicense1.value == "") {
                msg += '請輸入 專業證照-政府機關辦理相關證照或檢定!\n';
            }
            if (checkMaxLen2(ProLicense2.value, 200)) {
                msg += '【專業證照-其他證照或檢定】長度不可超過200字元\n';
            }
            if (ProLicense2.value == msgX1 || ProLicense2.value == "") {
                msg += '請輸入 專業證照-其他證照或檢定!\n';
            }
            if (msg != '') {
                window.alert(msg);
                return false;
            }
        }

        /*
		function KindOfTeacher(selectedID) {
		var KindEngage = document.getElementById('KindEngage');
		var parms = "[['KindEngage','" + KindEngage.value + "']]";  //透過selectControl傳遞給SQLMap的年度查詢條件,格式請參考selectControl定義說明
		selectControl('QueryIDKindOfTeacher', 'KindID', 'KindName', 'KindID', '請選擇', selectedID, parms);
		}
		*/
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
            <tbody>
                <tr>
                    <td>

                        <br>
                        <table class="table_sch" id="Table3" cellpadding="1" cellspacing="1" width="100%">
                            <tbody>
                                <tr>
                                    <%--<td class="bluecol_need" width="20%">計劃階層<asp:Label ID="Proecess" runat="server" ForeColor="#990000" Visible="false"></asp:Label></td>--%>
                                    <td class="bluecol_need" width="20%">訓練機構<asp:Label ID="Proecess" runat="server" ForeColor="#990000" Visible="false"></asp:Label></td>
                                    <td class="whitecol" width="30%">
                                        <asp:TextBox ID="TBplanOrgName" runat="server" onfocus="this.blur()" Width="80%"></asp:TextBox>
                                        <input style="display: none" onclick="javascript: wopen('../../Common/LevPlan.aspx', '計畫階段', 850, 400, 1)" type="button" value="選擇">
                                    </td>
                                    <td class="bluecol_need" width="20%">講師代碼<br />
                                        (最多十碼) </td>
                                    <td class="whitecol" width="30%">
                                        <asp:TextBox ID="TeacherID" runat="server" MaxLength="10" Width="60%"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">講師姓名 </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="TeachCName" runat="server" MaxLength="20" Width="60%"></asp:TextBox>
                                        <input id="btnSchEng" onclick="sentName();" type="button" value="英譯">
                                    </td>
                                    <td class="bluecol">講師英文姓名 </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="TeachEName" runat="server" MaxLength="30" Width="60%"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">身分別 </td>
                                    <td colspan="3" class="whitecol">
                                        <asp:RadioButtonList ID="PassPortNO" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                            <asp:ListItem Value="1" Selected="True">本國</asp:ListItem>
                                            <asp:ListItem Value="2">外籍(含大陸人士)</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">身分證號碼 </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="IDNO" runat="server" MaxLength="12" Width="50%"></asp:TextBox>
                                        <input id="Button3" type="button" value="載入" name="Button3" runat="server" class="asp_button_M">
                                    </td>
                                    <td class="bluecol_need">出生日期 </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="birthday" runat="server" onfocus="this.blur()" Width="40%" MaxLength="10"></asp:TextBox>
                                        <span id="span1" runat="server">
                                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= birthday.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">性別 </td>
                                    <td colspan="3" class="whitecol">
                                        <asp:RadioButtonList ID="Sex" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                            <asp:ListItem Value="M">男</asp:ListItem>
                                            <asp:ListItem Value="F">女</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">主要職類 </td>
                                    <td id="TD1" runat="server" class="whitecol">
                                        <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Width="80%"></asp:TextBox>
                                        <input onclick="openTrain2(document.getElementById('trainValue').value);" type="button" value="..." class="button_b_Mini">
                                        <input id="trainValue" type="hidden" name="trainValue" runat="server" />
                                        <input id="jobValue" type="hidden" name="jobValue" runat="server" />
                                    </td>
                                    <td id="Invest1_TD1" runat="server" class="bluecol_need">職稱 </td>
                                    <td id="Invest1_TD2" align="left" runat="server" class="whitecol">
                                        <asp:RadioButton ID="R1" runat="server"></asp:RadioButton>
                                        <asp:DropDownList ID="IVID" runat="server"></asp:DropDownList><br>
                                        <asp:RadioButton ID="R2" runat="server"></asp:RadioButton>
                                        <asp:TextBox ID="Invest1" runat="server" MaxLength="50" Width="60%"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr id="TR1" runat="server">
                                    <td id="TD4" runat="server" class="bluecol_need">內外聘 </td>
                                    <td id="TD5" runat="server" class="whitecol">
                                        <asp:DropDownList ID="KindEngage" runat="server" AutoPostBack="True">
                                            <asp:ListItem Value="0">請選擇</asp:ListItem>
                                            <asp:ListItem Value="1">內聘(專任)</asp:ListItem>
                                            <asp:ListItem Value="2">外聘(兼任)</asp:ListItem>
                                        </asp:DropDownList>
                                    </td>
                                    <td id="KindID_TD1" runat="server" class="bluecol_need">師資別 </td>
                                    <td id="KindID_TD2" runat="server" class="whitecol">
                                        <asp:DropDownList ID="KindID" runat="server">
                                            <asp:ListItem Value="0">請選擇內外聘</asp:ListItem>
                                        </asp:DropDownList>
                                        <%--
                                        <asp:Label ID="lab_KindEngageNG" runat="server" Text="請選擇內外聘"></asp:Label>
                                        <asp:DropDownList ID="KindID_E1" runat="server">
                                            <asp:ListItem Value="0">請選擇內外聘</asp:ListItem>
                                        </asp:DropDownList>
                                        <asp:DropDownList ID="KindID_E2" runat="server">
                                            <asp:ListItem Value="0">請選擇內外聘</asp:ListItem>
                                        </asp:DropDownList><FONT color="#cc66ff">說明</FONT>--%>
                                    </td>
                                </tr>
                                <tr id="tr_techtype12" runat="server">
                                    <td class="bluecol_need">類別 </td>
                                    <td colspan="5" class="whitecol">
                                        <asp:CheckBox Style="z-index: 0" ID="cb_techtype1" runat="server" CssClass="font" Text="講師"></asp:CheckBox>
                                        <asp:CheckBox Style="z-index: 0" ID="cb_techtype2" runat="server" CssClass="font" Text="助教"></asp:CheckBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">最高學歷 </td>
                                    <td class="whitecol">
                                        <asp:DropDownList ID="DegreeID" runat="server"></asp:DropDownList></td>
                                    <td class="bluecol_need">畢業狀況 </td>
                                    <td class="whitecol">
                                        <asp:DropDownList ID="GraduateStatus" runat="server"></asp:DropDownList></td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">學校名稱 </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="SchoolName" runat="server" MaxLength="50" Width="60%"></asp:TextBox></td>
                                    <td class="bluecol_need">科系名稱 </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="Department" runat="server" MaxLength="50" Width="60%"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">聯絡電話 </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="Phone" runat="server" MaxLength="20" Width="50%"></asp:TextBox></td>
                                    <td class="bluecol">行動電話 </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="Mobile" runat="server" MaxLength="20" Width="50%"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td class="bluecol">電子郵件 </td>
                                    <td colspan="3" class="whitecol">
                                        <asp:TextBox ID="Email" runat="server" MaxLength="64" Width="60%"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">通訊地址 </td>
                                    <td colspan="3" class="whitecol">
                                        <input id="city_code" maxlength="3" runat="server" />－
                                        <input id="AddressZIPB3" maxlength="3" runat="server" />
                                        <input id="hidAddressZIP6W" type="hidden" runat="server" />
                                        <asp:Literal ID="Litcity_code" runat="server"></asp:Literal><%--郵遞區號查詢--%>
                                        <br />
                                        <asp:TextBox ID="TBCity" runat="server" onfocus="this.blur()" Width="20%"></asp:TextBox>
                                        <input id="bt_openZip1" type="button" value="..." name="bt_openZip1" runat="server" class="button_b_Mini" />
                                        <asp:TextBox ID="Address" runat="server" Width="30%"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="4" class="table_title">服務單位資料 </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">服務單位名稱 </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="WorkOrg" runat="server" Width="60%"></asp:TextBox></td>
                                    <td class="bluecol">年資 </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="ExpYears" runat="server" Width="15%"></asp:TextBox>年
                                        <asp:DropDownList ID="ExpMonths" runat="server"></asp:DropDownList>個月 <font color="#ff0000">(請輸入數字)</font>
                                    </td>
                                </tr>
                                <tr class="whitecol">
                                    <td class="bluecol">服務部門 </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="ServDept" runat="server" Width="60%"></asp:TextBox></td>
                                    <td id="Invest2_TD1" runat="server" class="bluecol">職稱 </td>
                                    <td id="Invest2_TD2" runat="server" class="whitecol">
                                        <asp:TextBox ID="Invest2" runat="server" MaxLength="50" Width="60%"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">服務單位電話 </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="WorkPhone" runat="server" Width="60%"></asp:TextBox></td>
                                    <td class="bluecol">服務單位傳真 </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="Fax" runat="server" Width="60%"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td class="bluecol">服務單位地址 </td>
                                    <td colspan="3" class="whitecol">
                                        <input id="city_code1" maxlength="3" runat="server" />－
                                        <input id="WorkZIPB3" maxlength="3" runat="server" />
                                        <input id="hidWorkZIP6W" type="hidden" runat="server" />
                                        <asp:Literal ID="Litcity_code1" runat="server"></asp:Literal><%--郵遞區號查詢--%>
                                        <br />
                                        <asp:TextBox ID="TBCity1" runat="server" onfocus="this.blur()" Width="20%"></asp:TextBox>
                                        <input id="bt_openZip2" type="button" value="..." name="bt_openZip2" runat="server" class="button_b_Mini">
                                        <asp:TextBox ID="WorkAddr" runat="server" Width="30%" MaxLength="50"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="center" colspan="4" class="table_title">經歷 </td>
                                </tr>
                                <tr>
                                    <td class="bluecol" width="20%">服務單位 </td>
                                    <td colspan="3" class="whitecol">
                                        <table class="font" border="0" width="100%">
                                            <tr>
                                                <td class="whitecol" width="3%">
                                                    <asp:Label ID="star1" runat="server"><FONT color="#ff0000">*</FONT></asp:Label>1. </td>
                                                <td class="whitecol" width="25%">
                                                    <asp:TextBox ID="ExpUnit1" runat="server" MaxLength="50" Width="80%"></asp:TextBox></td>
                                                <td class="whitecol" width="7%">，<asp:Label ID="star2" runat="server"><FONT color="#ff0000">*</FONT></asp:Label>年資 </td>
                                                <td class="whitecol" width="20%">
                                                    <asp:TextBox ID="ExpYears1" runat="server" Width="55%" MaxLength="4"></asp:TextBox>年
                                                    <asp:DropDownList ID="ExpMonths1" runat="server"></asp:DropDownList>個月<br />
                                                    <font color="#ff0000">(年資請輸入數字)</font>
                                                </td>
                                                <td class="whitecol" width="5%">
                                                    <asp:Label ID="star4" runat="server"><FONT color="#ff0000">*</FONT></asp:Label>職稱： </td>
                                                <td class="whitecol" width="20%">
                                                    <asp:TextBox ID="tINV1" runat="server" MaxLength="50" placeholder="請輸入職稱1" Width="60%"></asp:TextBox></td>
                                            </tr>
                                            <tr>
                                                <td class="whitecol" width="5%">2. </td>
                                                <td class="whitecol" width="25%">
                                                    <asp:TextBox ID="ExpUnit2" runat="server" MaxLength="50" Width="80%"></asp:TextBox></td>
                                                <td class="whitecol" width="5%">，年資 </td>
                                                <td class="whitecol">
                                                    <asp:TextBox ID="ExpYears2" runat="server" Width="55%" MaxLength="4"></asp:TextBox>年
                                                    <asp:DropDownList ID="ExpMonths2" runat="server"></asp:DropDownList>個月
                                                </td>
                                                <td class="whitecol" width="5%">職稱： </td>
                                                <td>
                                                    <asp:TextBox ID="tINV2" runat="server" MaxLength="50" placeholder="請輸入職稱2" Width="60%"></asp:TextBox></td>
                                            </tr>
                                            <tr>
                                                <td class="whitecol">3. </td>
                                                <td class="whitecol">
                                                    <asp:TextBox ID="ExpUnit3" runat="server" MaxLength="50" Width="80%"></asp:TextBox></td>
                                                <td class="whitecol" width="5%">，年資 </td>
                                                <td class="whitecol">
                                                    <asp:TextBox ID="ExpYears3" runat="server" Width="55%" MaxLength="4"></asp:TextBox>年
                                                    <asp:DropDownList ID="ExpMonths3" runat="server"></asp:DropDownList>個月
                                                </td>
                                                <td class="whitecol" width="5%">職稱： </td>
                                                <td>
                                                    <asp:TextBox ID="tINV3" runat="server" MaxLength="50" placeholder="請輸入職稱3" Width="60%"></asp:TextBox></td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol">服務期間 </td>
                                    <td colspan="3" class="whitecol">
                                        <table class="whitecol" border="0">
                                            <tr>
                                                <td width="10%">&nbsp;&nbsp;&nbsp; 1. </td>
                                                <td width="40%">
                                                    <asp:TextBox ID="ExpSDate1" runat="server" Width="60%" MaxLength="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('ExpSDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></td>
                                                <td width="5%">～ </td>
                                                <td width="45%">
                                                    <asp:TextBox ID="ExpEDate1" runat="server" Width="60%" MaxLength="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('ExpEDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></td>
                                            </tr>
                                            <tr>
                                                <td width="5%">&nbsp;&nbsp;&nbsp; 2. </td>
                                                <td width="45%">
                                                    <asp:TextBox ID="ExpSDate2" runat="server" Width="60%" MaxLength="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('ExpSDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></td>
                                                <td width="5%">～ </td>
                                                <td width="45%">
                                                    <asp:TextBox ID="ExpEDate2" runat="server" Width="60%" MaxLength="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('ExpEDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></td>
                                            </tr>
                                            <tr>
                                                <td width="5%">&nbsp;&nbsp;&nbsp; 3. </td>
                                                <td width="45%">
                                                    <asp:TextBox ID="ExpSDate3" runat="server" Width="60%" MaxLength="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('ExpSDate3','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></td>
                                                <td width="5%">～ </td>
                                                <td width="45%">
                                                    <asp:TextBox ID="ExpEDate3" runat="server" Width="60%" MaxLength="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('ExpEDate3','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol">專長1<asp:Label ID="star3" runat="server"><FONT color="#ff0000">*</FONT></asp:Label></td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="Specialty1" onblur="checkTextLength(this,500)" onkeyup="checkTextLength(this,500);" runat="server" Width="50%" TextMode="MultiLine" Height="80px" onchange="checkTextLength(this,250)" MaxLength="500"></asp:TextBox></td>
                                    <td class="bluecol">專長2 </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="Specialty2" onblur="checkTextLength(this,500)" onkeyup="checkTextLength(this,500);" runat="server" Width="50%" TextMode="MultiLine" Height="80px" onchange="checkTextLength(this,500)" MaxLength="500"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td class="bluecol">專長3 </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="Specialty3" onblur="checkTextLength(this,500)" onkeyup="checkTextLength(this,500);" runat="server" Width="50%" TextMode="MultiLine" Height="80px" onchange="checkTextLength(this,500)" MaxLength="500"></asp:TextBox></td>
                                    <td class="bluecol">專長4 </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="Specialty4" onblur="checkTextLength(this,500)" onkeyup="checkTextLength(this,500);" runat="server" Width="50%" TextMode="MultiLine" Height="80px" onchange="checkTextLength(this,500)" MaxLength="500"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td class="bluecol">專長5 </td>
                                    <td colspan="3" class="whitecol">
                                        <asp:TextBox ID="Specialty5" onblur="checkTextLength(this,500)" onkeyup="checkTextLength(this,500);" runat="server" Width="50%" TextMode="MultiLine" Height="80px" onchange="checkTextLength(this,500)" MaxLength="500"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td class="bluecol">譯著 </td>
                                    <td colspan="3" class="whitecol">
                                        <asp:TextBox ID="TransBook" onblur="checkTextLength(this,100)" onkeyup="checkTextLength(this,100);" runat="server" Width="50%" TextMode="MultiLine" Height="80px" onchange="checkTextLength(this,100)" Columns="50" MaxLength="100"></asp:TextBox></td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need" width="20%">專業證照 </td>
                                    <td colspan="3" class="whitecol">
                                        <table class="whitecol" border="0" width="100%">
                                            <tr>
                                                <td width="30%">
                                                    <asp:Label ID="labPL1" runat="server" Text="政府機關辦理相關證照或檢定" ForeColor="Red"></asp:Label></td>
                                                <td>
                                                    <asp:TextBox ID="ProLicense1" runat="server" Columns="50" MaxLength="200" Width="60%"></asp:TextBox></td>
                                            </tr>
                                            <tr>
                                                <td width="30%">
                                                    <asp:Label ID="labPL2" runat="server" Text="其他證照或檢定" ForeColor="Red"></asp:Label></td>
                                                <td>
                                                    <asp:TextBox ID="ProLicense2" runat="server" Columns="50" MaxLength="200" Width="60%"></asp:TextBox></td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">排課使用 </td>
                                    <td colspan="3" class="whitecol">
                                        <asp:DropDownList ID="WorkStatus" runat="server">
                                            <asp:ListItem Value="0">請選擇</asp:ListItem>
                                            <asp:ListItem Value="1">是</asp:ListItem>
                                            <asp:ListItem Value="2">否</asp:ListItem>
                                        </asp:DropDownList>
                                        <font color="#ff0000">※排課使用:若選[是]則【排課功能】的師資選單會顯示,若選[否]則不顯示!!</font>
                                    </td>
                                </tr>
                            </tbody>
                        </table>
                        <div align="center" class="whitecol">
                            <asp:Button ID="Button1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                            <input id="Button2" type="button" value="回上一頁" name="Button2" runat="server" class="asp_button_M">
                        </div>
                        <div align="center">
                            <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                        </div>
                    </td>
                </tr>
            </tbody>
        </table>
        <asp:Button ID="Button4" runat="server" Text="複製產生(隱藏)"></asp:Button>
        <input id="TechID" type="hidden" runat="server">
        <input id="HidTechID" type="hidden" runat="server">
        <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
        <input id="PlanIDValue" type="hidden" name="PlanIDValue" runat="server">
        <input id="TPlanID" type="hidden" runat="server">
        <input id="LID" type="hidden" runat="server">
        <input id="hid_PLMsgX1" type="hidden" runat="server" value="請優先填寫與課程相關之專業技術類證照，若無證照資訊請填寫「無」。">
        <input id="HidsearchBox" type="hidden" runat="server">
    </form>
    <iframe id="ifmCheckZip" name="ifmCheckZip" src="../../common/CheckZip.aspx" width="0%" height="0%"></iframe>
</body>
</html>
