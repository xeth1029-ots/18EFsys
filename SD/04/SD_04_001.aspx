<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_04_001.aspx.vb" Inherits="WDAIIP.SD_04_001" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>全期排課作業</title>
    <meta content="False" name="vs_snapToGrid">
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/OpenWin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        //980224 fix只管一班則自動帶入
        function GETvalue() {
            document.getElementById('Button6').click();
        }

        function CheckImportData() {
            var msg = '';
            var OCIDValue1 = document.getElementById('OCIDValue1');
            var OCIDValue2 = document.getElementById('OCIDValue2');
            if (OCIDValue1.value == '') msg += '請先選擇班級\n';
            if (OCIDValue2.value == '') msg += '請先選擇要載入的班級\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function chkintosort() {
            var intoSort = document.getElementById('intoSort');
            var BTNintoSort = document.getElementById('BTNintoSort');
            if (confirm('排入正式課程，將無法再做修正!!!')) {
                BTNintoSort.click();
                if (intoSort) { intoSort.disabled = true; }
                return true;
            } else {
                return false;
            }
        }

        function chkcancelsort() {
            if (confirm('確定要取消正式課程，將無法回復資料!!!')) {
                return true;
            } else {
                return false;
            }
        }

        function Course(opentype, fieldname) {
            var RIDValue = document.getElementById('RIDValue');
            wopen('Course.aspx?RID=' + RIDValue.value + '&type=' + opentype + '&fieldname=' + fieldname, 'CheckID', 550, 400, 1);
        }

        function LessonTeah1(opentype, fieldname, hiddenname) {
            var RIDValue = document.getElementById('RIDValue');
            wopen('LessonTeah1.aspx?RID=' + RIDValue.value + '&type=' + opentype + '&fieldname=' + fieldname + '&hiddenname=' + hiddenname, 'LessonTeah1', 750, 600, 1);
        }

        function LessonTeah2(opentype, fieldname, hiddenname) {
            wopen('LessonTeah2.aspx?RID=' + document.form1.RIDValue.value + '&type=' + opentype + '&fieldname=' + fieldname + '&hiddenname=' + hiddenname, 'LessonTeah2', 750, 600, 1);
        }

        function CheckAddcourse() {
            if (isEmpty('OCIDValue1')) {
                alert('請先選擇班級');
                return false;
            }
            else {
                return CheckCourse('OItemID', 'OCourseID', 'OCalHours', 'OStartDate', 'OEndDate', 'ORoomID', 'OLessonTeah1', '', 'OS1', 'OE1', 'OS2', 'OE2', 'OS3', 'OE3', 'OS4', 'OE4', 'OS5', 'OE5', 'OS6', 'OE6', 'OS7', 'OE7', '')
            }
        }

        function CheckCourse(ItemID, CourseID, CalHours, StartDate, EndDate, RoomID, LessonTeah1, LessonTeah2, S1, E1, S2, E2, S3, E3, S4, E4, S5, E5, S6, E6, S7, E7, Recycle) {
            var msg = ''
            var ClassCount = 0;
            if (getValue(ItemID) == '') {
                msg += '請輸入項次!\n';
            } else {
                if (!isUnsignedInt(getValue(ItemID))) {
                    if (!isPositiveFloat(getValue(ItemID))) {
                        msg += '項次請輸入正整數或正小數點!\n';
                    }
                }
            }
            if (isEmpty(CourseID)) { msg += "請選擇課程代碼!\n"; }
            if (isEmpty(CalHours)) msg += '請輸入時數\n';
            else if (!isUnsignedInt(document.getElementById(CalHours).value)) msg += '時數必須為數字\n';
            var STDate = document.getElementById('ClassStart').innerHTML;
            var FTDate = document.getElementById('ClassEnd').innerHTML;
            if (!isEmpty(StartDate)) {
                if (!checkDate(document.getElementById(StartDate).value)) msg += '排課起日不是正確的日期格式!\n';
                else if (compareDate(STDate, getValue(StartDate)) == 1) msg += '排課起日不能在開訓日之前\n';
            }
            if (!isEmpty(EndDate)) {
                if (!checkDate(document.getElementById(EndDate).value)) msg += '排課迄日不是正確的日期格式!\n';
                else if (compareDate(FTDate, getValue(EndDate)) == -1) msg += '排課迄日不能在結訓日之後\n';
            }
            if (!isEmpty(StartDate) && !isEmpty(EndDate)) {
                var MyStartDate = getValue(StartDate);
                var MyEndDate = getValue(EndDate);
                if (compareDate(MyStartDate, MyEndDate) == 1) msg += '排課起日不能超過排課迄日\n';
            }
            if (isEmpty(RoomID)) { msg += '請輸入教室!\n'; }
            if (isEmpty(LessonTeah1)) { msg += '請選擇教師!\n'; }
            if ((getValue(S1) == '' && getValue(E1) != '') || (getValue(S1) != '' && getValue(E1) == '')) { msg += '週一節次起迄均要輸入!\n'; }
            if ((getValue(S1) != '' && !isUnsignedInt(getValue(S1))) || (getValue(E1) != '' && !isUnsignedInt(getValue(E1)))) msg += '週一節次必須輸入數字\n';
            if ((getValue(S2) == '' && getValue(E2) != '') || (getValue(S2) != '' && getValue(E2) == '')) { msg += '週二節次起迄均要輸入!\n'; }
            if ((getValue(S2) != '' && !isUnsignedInt(getValue(S2))) || (getValue(E2) != '' && !isUnsignedInt(getValue(E2)))) msg += '週二節次必須輸入數字\n';
            if ((getValue(S3) == '' && getValue(E3) != '') || (getValue(S3) != '' && getValue(E3) == '')) { msg += '週三節次起迄均要輸入!\n'; }
            if ((getValue(S3) != '' && !isUnsignedInt(getValue(S3))) || (getValue(E3) != '' && !isUnsignedInt(getValue(E3)))) msg += '週三節次必須輸入數字\n';
            if ((getValue(S4) == '' && getValue(E4) != '') || (getValue(S4) != '' && getValue(E4) == '')) { msg += '週四節次起迄均要輸入!\n'; }
            if ((getValue(S4) != '' && !isUnsignedInt(getValue(S4))) || (getValue(E4) != '' && !isUnsignedInt(getValue(E4)))) msg += '週四節次必須輸入數字\n';
            if ((getValue(S5) == '' && getValue(E5) != '') || (getValue(S5) != '' && getValue(E5) == '')) { msg += '週五節次起迄均要輸入!\n'; }
            if ((getValue(S5) != '' && !isUnsignedInt(getValue(S5))) || (getValue(E5) != '' && !isUnsignedInt(getValue(E5)))) msg += '週五節次必須輸入數字\n';
            if ((getValue(S6) == '' && getValue(E6) != '') || (getValue(S6) != '' && getValue(E6) == '')) { msg += '週六節次起迄均要輸入!\n'; }
            if ((getValue(S6) != '' && !isUnsignedInt(getValue(S6))) || (getValue(E6) != '' && !isUnsignedInt(getValue(E6)))) msg += '週六節次必須輸入數字\n';
            if ((getValue(S7) == '' && getValue(E7) != '') || (getValue(S7) != '' && getValue(E7) == '')) { msg += '週日節次起迄均要輸入!\n'; }
            if ((getValue(S7) != '' && !isUnsignedInt(getValue(S7))) || (getValue(E7) != '' && !isUnsignedInt(getValue(E7)))) msg += '週七節次必須輸入數字\n';
            var MyClassType = document.form1.ClassType.value;
            if (!CheckMyValue(S1, MyClassType) || !CheckMyValue(S2, MyClassType) || !CheckMyValue(S3, MyClassType) || !CheckMyValue(S4, MyClassType) || !CheckMyValue(S5, MyClassType) || !CheckMyValue(S6, MyClassType) || !CheckMyValue(S7, MyClassType) || !CheckMyValue(E1, MyClassType) || !CheckMyValue(E2, MyClassType) || !CheckMyValue(E3, MyClassType) || !CheckMyValue(E4, MyClassType) || !CheckMyValue(E5, MyClassType) || !CheckMyValue(E6, MyClassType) || !CheckMyValue(E7, MyClassType)) {
                switch (document.form1.ClassType.value) {
                    case '01':
                    case '05':
                        msg += "日間班不得排晚間時間。\n"
                        break;
                    case '02':
                        msg += "夜間班不得排日間時間。\n"
                        break;
                }
            }
            if (MyClassType == '04')
                if (getValue(S1) != '' || getValue(S2) != '' || getValue(S3) != '' || getValue(S4) != '' || getValue(S5) != '' || getValue(E1) != '' || getValue(E2) != '' || getValue(E3) != '' || getValue(E4) != '' || getValue(E5) != '') { msg += "假日班不得排日間時間。\n"; }
            if (msg == '') {
                return true;
            } else {
                alert(msg);
                return false;
            }
        }

        //檢查訓練種類
        function CheckMyValue(obj, num) {
            if (getValue(obj) != '') {
                switch (num) {
                    case '01':
                    case '05':
                        if (getValue(obj) > 8) return false;
                        break;
                    case '02':
                        if (getValue(obj) < 9) return false;
                        break;
                    case '04':
                        break;
                }
            }
            return true;
        }

        function GetCourseID(CourseID, TextField, ValueField, Tech1Field, TechName1Field, Tech2Field, TechName2Field, RoomField) {
            var RIDValue = document.getElementById('RIDValue');
            var Doc = frames['iframe1'];
            Doc.document.getElementById('CourseID').value = CourseID;
            Doc.document.getElementById('TextField').value = TextField;
            Doc.document.getElementById('ValueField').value = ValueField;
            Doc.document.getElementById('Tech1Field').value = Tech1Field;
            Doc.document.getElementById('TechName1Field').value = TechName1Field;
            Doc.document.getElementById('Tech2Field').value = Tech2Field;
            Doc.document.getElementById('TechName2Field').value = TechName2Field;
            Doc.document.getElementById('RoomField').value = RoomField;
            Doc.document.getElementById('RID').value = RIDValue.value;
            Doc.document.getElementById('Button1').click();
        }

        function choose_class(num) {
            HidInfo();
            openClass('../02/SD_02_ch.aspx?RWClass=1&RID=' + document.form1.RIDValue.value + '&special=2');
            /*if(num==1)
			openClass('../02/SD_02_ch.aspx?RID='+document.form1.RIDValue.value+'&special=2');
			else
			openClass('../02/SD_02_ch1.aspx?RID='+document.form1.RIDValue.value+'&special=2');*/
        }

        function choose_class2(num) {
            openClass('../02/SD_02_ch.aspx?RID=' + document.form1.RIDValue.value + '&AcctPlan=1&special=2&sort=2');
            /*if(num==1)
			openClass('../02/SD_02_ch.aspx?RID='+document.form1.RIDValue.value+'&AcctPlan=1&special=2&sort=2');
			else
			openClass('../02/SD_02_ch1.aspx?RID='+document.form1.RIDValue.value+'&AcctPlan=1&special=2&sort=2');*/
        }

        /*
		function openClaendar(obj){
		var STDate=document.getElementById('ClassStart').innerHTML;
		var FTDate=document.getElementById('ClassEnd').innerHTML;
		var NowDate=document.getElementById(obj).value;
		if (NowDate=='')
		NowDate=STDate;
		if(STDate!='')
		wopen('../../common/Calendar.aspx?STDate='+STDate+'&FTDate='+FTDate+'&NowDate='+NowDate+'&ValueField='+obj,'',350,230);
		}
        */

        function HidInfo() {
            document.getElementById('OCID1').value = '';
            document.getElementById('TMIDValue1').value = '';
            document.getElementById('OCIDValue1').value = '';
            document.getElementById('TMID1').value = '';
            document.getElementById('OCID2').value = '';
            document.getElementById('TMID2').value = '';
            document.getElementById('TMIDValue2').value = '';
            document.getElementById('OCIDValue2').value = '';
            document.getElementById('OCID').value = '';
            document.getElementById('ClassYear').innerHTML = '';
            document.getElementById('CyclType').innerHTML = '';
            document.getElementById('HourRan').innerHTML = '';
            document.getElementById('TranJob').innerHTML = '';
            document.getElementById('ClassStart').innerHTML = '';
            document.getElementById('ClassEnd').innerHTML = '';
            document.getElementById('ClassHours').innerHTML = '';
            document.getElementById('Totals').innerHTML = '';
            document.getElementById('ISinto').innerHTML = '';
            if (document.getElementById('DataGrid1')) document.getElementById('DataGrid1').style.display = 'none';
            document.getElementById('NoFormal').disabled = true;
            document.getElementById('intoSort').disabled = true;
            document.getElementById('Cancelinto').disabled = true;
            document.getElementById('print').disabled = true;
            document.getElementById('Button5').disabled = true;
        }

        //選擇機構
        /*
        function GetOrg(num){
		HidInfo();
		document.getElementById('Button6').click();
		if(num==1)
		openOrg('../../Common/LevOrg.aspx');
		else
		openOrg('../../Common/LevOrg1.aspx');
		}
        */

        function showHistory(obj) {
            if (document.getElementById(obj)) {
                if (document.getElementById(obj).style.display == 'none') {
                    document.getElementById(obj).style.display = '';
                }
                else {
                    document.getElementById(obj).style.display = 'none';
                }
            }
        }

        function returnValue1(ItemID, CourseID, CalHours, StartDate, EndDate, RoomID, LessonTeah1, LessonTeah2, S1, E1, S2, E2, S3, E3, S4, E4, S5, E5, S6, E6, S7, E7, Recycle, CourseIDValue, LessonTeah1Value, LessonTeah2Value) {
            document.getElementById('OItemID').value = ItemID;
            document.getElementById('OCourseID').value = CourseID;
            document.getElementById('OCalHours').value = CalHours;
            document.getElementById('OStartDate').value = StartDate;
            document.getElementById('OEndDate').value = EndDate;
            document.getElementById('ORoomID').value = RoomID;
            document.getElementById('OLessonTeah1').value = LessonTeah1;
            document.getElementById('OLessonTeah2').value = LessonTeah2;
            document.getElementById('OS1').value = S1;
            document.getElementById('OE1').value = E1;
            document.getElementById('OS2').value = S2;
            document.getElementById('OE2').value = E2;
            document.getElementById('OS3').value = S3;
            document.getElementById('OE3').value = E3;
            document.getElementById('OS4').value = S4;
            document.getElementById('OE4').value = E4;
            document.getElementById('OS5').value = S5;
            document.getElementById('OE5').value = E5;
            document.getElementById('OS6').value = S6;
            document.getElementById('OE6').value = E6;
            document.getElementById('OS7').value = S7;
            document.getElementById('OE7').value = E7;
            document.getElementById('ORecycle').value = Recycle;
            document.getElementById('OCourseIDValue').value = CourseIDValue;
            document.getElementById('OLessonTeah1Value').value = LessonTeah1Value;
            document.getElementById('OLessonTeah2Value').value = LessonTeah2Value;
            return false;
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;課程管理&gt;&gt;全期排課</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="99%" border="0">
            <tr>
                <td align="center">
                    <table id="xxx" cellpadding="1" cellspacing="1" width="100%" border="0">
                        <tr>
                            <td align="left">(此作業需先將課程代碼和師資資料先設定完成，再執行此作業)</td>
                        </tr>
                    </table>
                    <table id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Button ID="NoFormal" runat="server" Text="預覽正式課程" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="intoSort" runat="server" Text="排入正式課程" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="Cancelinto" runat="server" Text="取消正式課程" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="print" runat="server" Text="列印課程總表" CausesValidation="False" Enabled="False" CssClass="asp_Export_M"></asp:Button>
                                <asp:Button ID="Button5" runat="server" Text="刪除預排資料" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="Button7" runat="server" Text="重建時間配當" Enabled="False" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <table class="table_sch" id="MenuTable" cellspacing="1" cellpadding="1">
                                    <tr>
                                        <td class="bluecol" style="width: 20%">訓練機構</td>
                                        <td colspan="9" class="whitecol">
                                            <asp:TextBox ID="center" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox><input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                            <input id="Button3" onclick="GETvalue()" type="button" value="..." name="Button3" runat="server" class="button_b_Mini">
                                            <asp:Button ID="Button6" Style="display: none" runat="server" Text="Button6"></asp:Button>
                                            <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()"><asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table></span>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol" colspan="2"></td>
                                        <td class="bluecol">年度</td>
                                        <td class="whitecol"><asp:Label ID="ClassYear" runat="server"></asp:Label></td>
                                        <td class="bluecol">期別</td>
                                        <td class="whitecol"><asp:Label ID="CyclType" runat="server"></asp:Label></td>
                                        <td class="bluecol">訓練時段</td>
                                        <td class="whitecol"><asp:Label ID="HourRan" runat="server"></asp:Label></td>
                                        <td class="bluecol">訓練職類</td>
                                        <td class="whitecol"><asp:Label ID="TranJob" runat="server"></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">排課班級</td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" AutoPostBack="True" Columns="33" Width="250px"></asp:TextBox>
                                            <input id="Button1" onclick="choose_class()" type="button" value="..." name="Button1" runat="server" class="button_b_Mini">
                                            <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                            <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                            <input id="TMID1" type="hidden" runat="server"><br>
                                            <span id="HistoryList" style="position: absolute; display: none"><asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table></span>
                                        </td>
                                        <td class="bluecol">開訓日期
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="ClassStart" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol">結訓日期
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="ClassEnd" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol">訓練時數
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="ClassHours" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol">已排總時數
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Totals" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">載入班級
                                        </td>
                                        <td colspan="5" class="whitecol">
                                            <asp:TextBox ID="OCID2" runat="server" onfocus="this.blur()" AutoPostBack="True" Columns="33" Width="250px"></asp:TextBox>
                                            <input id="Button2" onclick="choose_class2()" type="button" value="..." name="Button2" runat="server" class="button_b_Mini"><input id="Button4" type="button" value="清除" name="Button4" runat="server" class="asp_button_M">
                                            <asp:Button ID="LoadIntoClass" runat="server" Text="載入排課班級" Width="150px" CssClass="asp_button_M"></asp:Button><input id="TMID2" type="hidden" name="TMID2" runat="server">
                                            <input id="TMIDValue2" type="hidden" name="TMIDValue2" runat="server">
                                            <input id="OCIDValue2" type="hidden" name="OCIDValue2" runat="server">
                                        </td>
                                        <td class="bluecol">系統提示</td>
                                        <td class="whitecol" colspan="3">
                                            <asp:Label ID="ISinto" runat="server" ForeColor="Red"></asp:Label>
                                            <%--<asp:Button ID="btnDelX2" runat="server" CssClass="asp_button_M" Text="清除全期排課框架" />--%>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <table class="font" style="border-collapse: collapse" cellspacing="0" cellpadding="0" width="100%" border="1">
                                    <tr class="head_navy">
                                        <td align="center">項次</td>
                                        <td align="center">課程名稱/代碼</td>
                                        <td align="center">時數</td>
                                        <td align="center">排課起日</td>
                                        <td align="center">排課迄日</td>
                                        <td align="center">教室</td>
                                        <td align="center">教師</td>
                                        <td align="center">助教</td>
                                        <td align="center" colspan="2">週一</td>
                                        <td align="center" colspan="2">週二</td>
                                        <td align="center" colspan="2">週三</td>
                                        <td align="center" colspan="2">週四</td>
                                        <td align="center" colspan="2">週五</td>
                                        <td align="center" colspan="2">週六</td>
                                        <td align="center" colspan="2">週日</td>
                                        <td align="center">循環</td>
                                        <td align="center">功能</td>
                                    </tr>
                                    <tr>
                                        <td align="center"><asp:TextBox ID="OItemID" runat="server" Width="25px" Font-Size="XX-Small"></asp:TextBox></td>
                                        <td align="center"><asp:TextBox ID="OCourseID" Style="cursor: pointer" runat="server" Width="90" ToolTip="點選兩下可以跳出視窗選擇課程名稱" MaxLength="40" Font-Size="XX-Small"></asp:TextBox></td>
                                        <td align="center"><asp:TextBox ID="OCalHours" runat="server" Width="25" MaxLength="3" Font-Size="XX-Small"></asp:TextBox></td>
                                        <td align="center"><asp:TextBox ID="OStartDate" Style="cursor: pointer" runat="server" Width="60" ToolTip="點選兩下可以跳出視窗選擇起迄日期" MaxLength="12" Font-Size="XX-Small"></asp:TextBox></td>
                                        <td align="center"><asp:TextBox ID="OEndDate" Style="cursor: pointer" runat="server" Width="60" ToolTip="點選兩下可以跳出視窗選擇起迄日期" MaxLength="12" Font-Size="XX-Small"></asp:TextBox></td>
                                        <td align="center"><asp:TextBox ID="ORoomID" runat="server" Width="70" MaxLength="8" Font-Size="XX-Small"></asp:TextBox></td>
                                        <td align="center"><asp:TextBox ID="OLessonTeah1" Style="cursor: pointer" runat="server" Width="60" ToolTip="點選兩下可以跳出視窗選擇教師" MaxLength="8" Font-Size="XX-Small"></asp:TextBox></td>
                                        <td align="center"><asp:TextBox ID="OLessonTeah2" Style="cursor: pointer" runat="server" Width="60" ToolTip="點選兩下可以跳出視窗選擇助教" MaxLength="8" Font-Size="XX-Small"></asp:TextBox></td>
                                        <td align="center"><asp:TextBox ID="OS1" runat="server" Width="20" MaxLength="2" Font-Size="XX-Small"></asp:TextBox></td>
                                        <td align="center"><asp:TextBox ID="OE1" runat="server" Width="20" MaxLength="2" Font-Size="XX-Small"></asp:TextBox></td>
                                        <td align="center"><asp:TextBox ID="OS2" runat="server" Width="20" MaxLength="2" Font-Size="XX-Small"></asp:TextBox></td>
                                        <td align="center"><asp:TextBox ID="OE2" runat="server" Width="20" MaxLength="2" Font-Size="XX-Small"></asp:TextBox></td>
                                        <td align="center"><asp:TextBox ID="OS3" runat="server" Width="20" MaxLength="2" Font-Size="XX-Small"></asp:TextBox></td>
                                        <td align="center"><asp:TextBox ID="OE3" runat="server" Width="20" MaxLength="2" Font-Size="XX-Small"></asp:TextBox></td>
                                        <td align="center"><asp:TextBox ID="OS4" runat="server" Width="20" MaxLength="2" Font-Size="XX-Small"></asp:TextBox></td>
                                        <td align="center"><asp:TextBox ID="OE4" runat="server" Width="20" MaxLength="2" Font-Size="XX-Small"></asp:TextBox></td>
                                        <td align="center"><asp:TextBox ID="OS5" runat="server" Width="20" MaxLength="2" Font-Size="XX-Small"></asp:TextBox></td>
                                        <td align="center"><asp:TextBox ID="OE5" runat="server" Width="20" MaxLength="2" Font-Size="XX-Small"></asp:TextBox></td>
                                        <td align="center"><asp:TextBox ID="OS6" runat="server" Width="20" MaxLength="2" Font-Size="XX-Small"></asp:TextBox></td>
                                        <td align="center"><asp:TextBox ID="OE6" runat="server" Width="20" MaxLength="2" Font-Size="XX-Small"></asp:TextBox></td>
                                        <td align="center"><asp:TextBox ID="OS7" runat="server" Width="20" MaxLength="2" Font-Size="XX-Small"></asp:TextBox></td>
                                        <td align="center"><asp:TextBox ID="OE7" runat="server" Width="20" MaxLength="2" Font-Size="XX-Small"></asp:TextBox></td>
                                        <td align="center"><asp:TextBox ID="ORecycle" runat="server" Width="20px" MaxLength="2" Font-Size="XX-Small"></asp:TextBox></td>
                                        <td align="center"><asp:ImageButton ID="but_add" runat="server" ImageUrl="../../images/Add.gif"></asp:ImageButton></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:TemplateColumn HeaderText="項次" FooterText="合計">
                                            <EditItemTemplate>
                                                <asp:TextBox ID="TItemID" runat="server" Width="20"></asp:TextBox>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="課程代碼">
                                            <HeaderTemplate>課程代碼</HeaderTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox ID="TCourseID" Style="cursor: pointer" runat="server" Width="90" MaxLength="40"></asp:TextBox>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="時數" ItemStyle-HorizontalAlign="Center">
                                            <HeaderTemplate>時數</HeaderTemplate>
                                            <ItemTemplate></ItemTemplate>
                                            <FooterStyle HorizontalAlign="Center"></FooterStyle>
                                            <EditItemTemplate>
                                                <asp:TextBox ID="TCalHours" runat="server" Width="25" MaxLength="3"></asp:TextBox>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="排課起日">
                                            <HeaderTemplate>排課起日</HeaderTemplate>
                                            <ItemTemplate></ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox ID="TStartDate" runat="server" Width="60" MaxLength="12" Style="cursor: pointer"></asp:TextBox>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="排課迄日">
                                            <HeaderTemplate>排課迄日</HeaderTemplate>
                                            <ItemTemplate></ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox ID="TEndDate" runat="server" Width="60" MaxLength="12" Style="cursor: pointer"></asp:TextBox>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="教室">
                                            <HeaderTemplate>教室</HeaderTemplate>
                                            <ItemTemplate></ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox ID="TRoomID" runat="server" Width="70" MaxLength="8"></asp:TextBox>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="教師">
                                            <HeaderTemplate>教師</HeaderTemplate>
                                            <ItemTemplate></ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:HiddenField ID="HidTLessonTeah1" runat="server" />
                                                <asp:TextBox ID="TLessonTeah1" Style="cursor: pointer" runat="server" MaxLength="8" Width="60"></asp:TextBox>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="助教">
                                            <HeaderTemplate>助教</HeaderTemplate>
                                            <ItemTemplate></ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:HiddenField ID="HidTLessonTeah2" runat="server" />
                                                <asp:TextBox ID="TLessonTeah2" Style="cursor: pointer" runat="server" MaxLength="8" Width="60"></asp:TextBox>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="週一">
                                            <HeaderTemplate>週一</HeaderTemplate>
                                            <ItemTemplate></ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox ID="TS1" runat="server" MaxLength="2" Width="20"></asp:TextBox>
                                                <asp:TextBox ID="TE1" runat="server" MaxLength="2" Width="20"></asp:TextBox>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="週二">
                                            <HeaderTemplate>週二</HeaderTemplate>
                                            <ItemTemplate></ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox ID="TS2" runat="server" MaxLength="2" Width="20"></asp:TextBox>
                                                <asp:TextBox ID="TE2" runat="server" MaxLength="2" Width="20"></asp:TextBox>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="週三">
                                            <HeaderTemplate>週三</HeaderTemplate>
                                            <ItemTemplate></ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox ID="TS3" runat="server" MaxLength="2" Width="20"></asp:TextBox>
                                                <asp:TextBox ID="TE3" runat="server" MaxLength="2" Width="20"></asp:TextBox>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="週四">
                                            <HeaderTemplate>週四</HeaderTemplate>
                                            <ItemTemplate></ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox ID="TS4" runat="server" MaxLength="2" Width="20"></asp:TextBox>
                                                <asp:TextBox ID="TE4" runat="server" MaxLength="2" Width="20"></asp:TextBox>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="週五">
                                            <HeaderTemplate>週五</HeaderTemplate>
                                            <ItemTemplate></ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox ID="TS5" runat="server" MaxLength="2" Width="20"></asp:TextBox>
                                                <asp:TextBox ID="TE5" runat="server" MaxLength="2" Width="20"></asp:TextBox>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="週六">
                                            <HeaderTemplate>週六</HeaderTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox ID="TS6" runat="server" Width="20" MaxLength="2"></asp:TextBox>
                                                <asp:TextBox ID="TE6" runat="server" Width="20" MaxLength="2"></asp:TextBox>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="週日">
                                            <HeaderTemplate>週日</HeaderTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox ID="TS7" runat="server" Width="20" MaxLength="2"></asp:TextBox>
                                                <asp:TextBox ID="TE7" runat="server" Width="20" MaxLength="2"></asp:TextBox>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="循環">
                                            <EditItemTemplate>
                                                <asp:TextBox ID="TRecycle" runat="server" Width="20px"></asp:TextBox>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="RealHours" ItemStyle-HorizontalAlign="Center" HeaderText="實際時數">
                                            <FooterStyle HorizontalAlign="Center"></FooterStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="../../images/Edit.gif" CommandName="edit"></asp:ImageButton>
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:ImageButton ID="ImageButton2" runat="server" ImageUrl="../../images/Update.gif" CommandName="update"></asp:ImageButton>
                                                <asp:ImageButton ID="Imagebutton3" runat="server" ImageUrl="../../images/Cancel.gif" CommandName="cancel"></asp:ImageButton>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="刪除">
                                            <ItemStyle HorizontalAlign ="Center" />
                                            <ItemTemplate>
                                                <asp:ImageButton ID="Imagebutton4" runat="server" ImageUrl="../../images/Del.gif" CommandName="del"></asp:ImageButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="複製">
                                            <ItemStyle HorizontalAlign="Center" Font-Size="Small" />
                                            <ItemTemplate>
                                                <asp:LinkButton ID="lbtnCopy1" runat="server" Text="複製" CssClass="linkbutton"></asp:LinkButton>
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
        <input id="OCID" type="hidden" runat="server" />
        <input id="OCourseIDValue" type="hidden" runat="server" />
        <input id="OLessonTeah1Value" type="hidden" runat="server" />
        <input id="OLessonTeah2Value" type="hidden" runat="server" />

        <input id="ClassType" type="hidden" runat="server" />
        <input id="hidTmpTItemID" type="hidden" runat="server" name="hidTmpTItemID" />
        <input id="HidClassStartDate" type="hidden" runat="server" />
        <input id="HidClassEndDate" type="hidden" runat="server" />
        <input id="HidClassHours" type="hidden" runat="server" />
        <asp:Button ID="BTNintoSort" Style="display: none" runat="server" Text="BTNintoSort" />
    </form>
    <iframe id="iframe1" style="display: none" src="FindCourse.aspx"></iframe>
</body>
</html>