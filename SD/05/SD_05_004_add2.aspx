<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_004_add2.aspx.vb" Inherits="WDAIIP.SD_05_004_add2" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>離退訓作業</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var cst_inline1 = "";
        function ShowOrg(vStatus) {
            //'RTReasonID'
            /*
			'04:患病或遇意外傷害
			'03:遇家庭等災變事故
			'07:奉召服兵役
			'02:提前就業(訓期滿1/2以上)
			'98:其他(職前訓練須經分署/縣市政府專案認定)
			
			'01:缺課時數超過規定
			'13:參訓期間行為不檢情節重大
			'14:訓期未滿1/2找到工作
			'99:其他
			01:缺課時數超過規定			02:提前就業			06:訓練成績不合格
			*/
            var rdoObjStatus = document.getElementById('StudStatus');
            //var StudStatus = document.getElementsByName('StudStatus');
            var HidRTReasonID = document.getElementById('HidRTReasonID');
            var HidvStatus = document.getElementById('HidvStatus');
            var RTReasonID22 = document.getElementById('RTReasonID22');
            var spRTReasonID22 = document.getElementById('spRTReasonID22');

            var rdoObj = null;
            var rdoObjN = null; //其他不選擇
            if (vStatus == '2') rdoObj = document.getElementById('RTReasonID2');
            if (vStatus == '2') rdoObjN = document.getElementById('RTReasonID3');

            if (vStatus == '3') rdoObj = document.getElementById('RTReasonID3');
            if (vStatus == '3') rdoObjN = document.getElementById('RTReasonID2');
            HidvStatus.value = vStatus;

            rdoObjStatus.disabled = false;
            if (vStatus == '2' || vStatus == '3') {
                var rdoList = rdoObjStatus.getElementsByTagName('input');
                for (var i = 0; i < rdoList.length; i++) {
                    if (rdoList[i].value == vStatus) {
                        rdoList[i].checked = true;
                        break;
                    }
                }
                rdoObjStatus.disabled = true;
            }

            if (rdoObj) {
                var rdoList = rdoObj.getElementsByTagName('input');
                for (var i = 0; i < rdoList.length; i++) {
                    //02:提前就業
                    if (rdoList[i].checked == true) {
                        HidRTReasonID.value = rdoList[i].value;
                        break;
                    }
                }
            }
            if (rdoObjN) {
                var rdoList = rdoObjN.getElementsByTagName('input');
                for (var i = 0; i < rdoList.length; i++) {
                    // 02:提前就業
                    if (rdoList[i].checked == true) {
                        rdoList[i].checked = false;
                        break;
                    }
                }
            }

            //提前就業
            var trOrgData1 = document.getElementById('trOrgData1');
            if (trOrgData1) { trOrgData1.style.display = 'none'; }
            if (HidRTReasonID && HidRTReasonID.value == '02') {
                if (trOrgData1) { trOrgData1.style.display = cst_inline1; } //'inline'; 
            }
            if (spRTReasonID22) { spRTReasonID22.style.display = 'none'; }
            if (HidRTReasonID && HidRTReasonID.value == '98') {
                if (spRTReasonID22) { spRTReasonID22.style.display = cst_inline1; } //'inline'; 
            }
        }

        function search() {
            var OCIDValue1 = document.getElementById('OCIDValue1');
            if (OCIDValue1.value == '') {
                alert('請選擇職類班別!');
                return false;
            }
        }

        function choose_class() {
            openClass('../02/SD_02_ch.aspx?special=2');
            //window.open('../02/SD_02_ch.aspx?special=2&BtnName=Button5','','width=550,height=520,location=0,status=0,menubar=0,scrollbars=1,resizable=0');
        }

        function chkdata() {
            //document.form1.
            var HidUseCanOff = document.getElementById('HidUseCanOff');
            var HidCanOffStudExists = document.getElementById('HidCanOffStudExists');
            var SumOfPay = document.getElementById('SumOfPay');
            var HadPay = document.getElementById('HadPay');
            var OCIDValue1 = document.getElementById('OCIDValue1');
            var SOCID = document.getElementById('SOCID');
            //var StudStatus = document.getElementById('StudStatus');
            var StudStatus = document.getElementsByName('StudStatus');

            var RejectTDate = document.getElementById('RejectTDate');
            var RTReasonID2 = document.getElementsByName('RTReasonID2');
            var RTReasonID22 = document.getElementsByName('RTReasonID22');
            var RTReasoOther2 = document.getElementById('RTReasoOther2');

            var RTReasonID3 = document.getElementsByName('RTReasonID3');
            var RTReasoOther3 = document.getElementById('RTReasoOther3');
            var hidTHoours = document.getElementById('hidTHoours');
            var TrainHours = document.getElementById('TrainHours');
            var NeedPay = document.getElementById('NeedPay');
            var SumOfPay = document.getElementById('SumOfPay');
            //var HadPay = document.getElementById('HadPay');
            var PayStatus = document.getElementById('PayStatus');
            var NoClose = document.getElementById('NoClose');
            var Other = document.getElementById('Other');
            var msg = '';

            var SOCIDvalue = ""; //目前使用者所選擇的學號。
            var HidUseCanOff = HidUseCanOff.value; //可使用離退判斷
            var HidCanOffStudExists = HidCanOffStudExists.value; //可以離退的學號

            var sumofpay = parseInt(SumOfPay.value, 10);
            var hadpay = parseInt(HadPay.value, 10);
            if (OCIDValue1.value == '') msg += '請選擇原始職類班別\n';
            if (SOCID.selectedIndex == 0 && !SOCID.disabled) {
                msg += '請選擇學員\n';
            }
            else {
                //目前使用者所選擇的學號。
                SOCIDvalue = SOCID.value.split("&")[0];
                //判斷是否已經離退訓 //目前使用者所選擇的學號的在訓狀況。
                var State = SOCID.value.split("&")[1];
                if (State == '2' || State == '3') msg += '不能選擇離退訓的學員\n'
            }
            if (HidUseCanOff == "1") {
                if (HidCanOffStudExists.indexOf(SOCIDvalue) == -1) {
                    //查無可離退學號。
                    msg += '請先至「學員資料維護」將欲辦理離退訓作業學員資料[確實依最新資料更新維護]後，再修改預算別為不補助及補助比例為0%。\n';
                }
            }

            //if (!isChecked(StudStatus)) msg += '請選擇離訓或退訓\n';
            if (RejectTDate.value == '' && !RejectTDate.disabled) msg += '請輸入離退訓日期\n';
            if (RejectTDate.value != '' && !checkDate(RejectTDate.value)) msg += '退訓日期格式不正確\n';

            var rtother = ''; //離退訓原因選擇其他說明必填。
            if (getRadioValue(RTReasonID2) == '98' && getRadioValue(RTReasonID22) == '') {
                //20,21,22,23
                msg += '選了「其他(職前訓練須經分署/縣市政府專案認定)」，請再選擇原因\n';
            }
            //if (getRadioValue(RTReasonID2) == '98') rtother = 'Y2'; //離退訓原因選擇其他說明必填。
            if (getRadioValue(RTReasonID22) == '23') rtother = 'Y2'; //離退訓原因選擇其他說明必填。20,21,22,23
            if (getRadioValue(RTReasonID3) == '99') rtother = 'Y3'; //離退訓原因選擇其他說明必填。
            if (rtother == 'Y2' && RTReasoOther2.value == '') {
                msg += '離訓原因選擇其他說明必填，請輸入\n';
            }
            if (rtother == 'Y3' && RTReasoOther3.value == '') {
                msg += '退訓原因選擇其他說明必填，請輸入\n';
            }
            //debugger;//getValue("RTReasonID");getRadioValue(document.form1.RTReasonID)
            //if (!isChecked(RTReasonID)) msg += '請選擇離退訓原因\n';
            //else if (rtother == 'Y' && RTReasoOther.value == '') msg += '離退訓原因選擇其他說明必填，請輸入\n';

            var HidRTReasonID = document.getElementById('HidRTReasonID');
            //if (document.form1.RTReasonID[1].checked == true)
            //if (document.getElementById('trOrgData1').style.display == 'inline')
            //if (HidRTReasonID.value == '01')
            /*
			01:缺課時數超過規定
			02:提前就業
			06:訓練成績不合格
			*/
            var JobDate = document.getElementById('JobDate');
            var GetJob1 = document.getElementById('GetJob1');
            var JobOrgName = document.getElementById('JobOrgName');
            var JobTel = document.getElementById('JobTel');
            var JobZipCode = document.getElementById('JobZipCode');
            var Jobaddress = document.getElementById('Jobaddress');
            var GetJobCode1 = document.getElementById('GetJobCode1');
            var trJOBFIELD = document.getElementById('trJOBFIELD');
            //var JobCode5 = document.getElementById('JobCode5');
            //var SpecTrace = document.getElementById('SpecTrace');
            //If (vGetJobCode1=='05') {}.SelectedValue = "05" Then

            if (HidRTReasonID.value == '02') {
                //就業單位到職日 JobDate
                //切結對象 GetJob1
                //就業單位名稱 JobOrgName
                //事業單位地址 JobCity    JobZipCode Jobaddress
                //事業單位電話 JobTel
                //薪資級距 JobSalID
                var vGetJob1 = getValue("GetJob1");
                var vJobSalID = getValue("JobSalID");
                var vGetJobCode1 = getValue("GetJobCode1");
                //if (vGetJob1 != '') alert('vGetJob1::' + vGetJob1);
                //if (vJobSalID != '') alert('vJobSalID::' + vJobSalID);
                if (JobDate.value == '') { msg += '請輸入就業單位到職日\n' }
                if (JobDate.value != '' && !checkDate(JobDate.value)) { msg += '【就業單位到職日】不是正確的日期格式\n'; }
                //if (GetJob1.value == '') msg += '請選擇切結對象\n'
                if (vGetJob1 == '') msg += '請選擇切結對象\n'
                if (vGetJob1 != '') {
                    if (vGetJob1 != '1' && vGetJob1 != '2') msg += '切結對象只能選擇 雇主切結或學員切結!\n'
                }
                if (JobOrgName.value == '') msg += '請輸入就業單位名稱\n'
                if (JobZipCode.value == '') msg += '請選擇就業單位郵遞區號\n'
                if (Jobaddress.value == '') msg += '請輸入就業單位地址\n'
                if (JobTel.value == '') msg += '請輸入就業單位電話\n'
                //debugger;
                if (vJobSalID == '') msg += '請選擇就業薪資級距\n'
                if (vGetJobCode1 == '') msg += '請選擇就業原因\n'
                msg += ChkJobRelateYN();
                if (trJOBFIELD) {
                    if (getValue("ddlJOBFIELD") == "") { msg += '訓後就業場域 為必填(請選擇)\n'; }
                }
            }

            //if (TrainHours.value != '' && !isUnsignedInt(TrainHours.value)) msg += '實際受訓時數必須為數字\n';
            let hidthours1 = convertToNumber(hidTHoours.value);
            let trainhours1 = convertToNumber(TrainHours.value);
            if (TrainHours.value != '') {
                if (!isValidNumberUsingRegex(TrainHours.value)) { msg += '實際受訓時數必須為數字\n'; }
                else if (!isDivisibleByHalf(trainhours1)) { msg += '實際受訓時數必須為可整除0.5的數字\n'; }
                else if (trainhours1 < 0) { msg += '實際受訓時數必須為大於等於0的數字\n'; }
                else if (hidthours1 > 0 && trainhours1 > hidthours1) { msg += '實際受訓時數必須為小於等於訓練時數 ' + hidthours1 + '\n'; }
                else { TrainHours.value = trainhours1; }
            }

            if (NeedPay.selectedIndex == 0) msg += '請選擇是否賠償\n';
            if (NeedPay.selectedIndex == 1 && SumOfPay.value == '') msg += '請輸入賠償金額\n';

            if (!isUnsignedInt(SumOfPay.value) && SumOfPay.value != '') {
                msg += '賠償金額必須為數字\n';
                SumOfPay.focus();
                SumOfPay.select();
            }
            if (!isUnsignedInt(HadPay.value) && HadPay.value != '') {
                msg += '已賠金額必須為數字\n';
                HadPay.focus();
                HadPay.select();
            }

            if (isUnsignedInt(SumOfPay.value) && isUnsignedInt(HadPay.value)) {
                if (hadpay > sumofpay) {
                    msg += '已賠償金額必須小於賠償金額\n';
                    HadPay.focus();
                    HadPay.select();
                }
                if (hadpay == 0 && sumofpay != 0) {
                    msg += '賠償金額目前為「0」；已賠償金額不為「0」\n';
                    HadPay.focus();
                    HadPay.select();
                }
            }
            if (getValue(PayStatus) == 2) {
                if (NoClose.value == '') {
                    msg += '請選擇追償狀況_未結案原因\n';
                }
            }

            if (getValue(PayStatus) == 3) {
                if (Other.value == '') {
                    msg += '請選擇追償狀況_其他原因\n';
                }
            }

            if (msg != '') {
                alert(msg);
                return false;
            }

        }

        function NeedPays() {
            var NeedPay = document.getElementById('NeedPay');
            var SumOfPay = document.getElementById('SumOfPay');
            var HadPay = document.getElementById('HadPay');

            if (NeedPay.selectedIndex == 1) {
                SumOfPay.disabled = false;
                HadPay.disabled = false;
                //SumOfPay.value=SumOfPay1.value;	
            }
            else if (NeedPay.selectedIndex == 2) {
                SumOfPay.disabled = true;
                HadPay.disabled = true;
                SumOfPay.value = '0';
                HadPay.value = '0';
            }
        }

        /*
		function NeedPays(){				
		if (document.form1.NeedPay.selectedIndex==1){
		document.form1.SumOfPay.disabled=false;	
		document.form1.HadPay.disabled=false;
				 
		}
		else if (document.form1.NeedPay.selectedIndex==2){
		document.form1.SumOfPay.disabled=true;
		document.form1.HadPay.disabled=true;
		document.form1.SumOfPay.value='';
		}
		}			
		*/
        /*
		01:缺課時數超過規定
		02:提前就業
		06:訓練成績不合格
		*/
        function setRdoDisabled(id) {
            var rdoObj = document.getElementById(id);
            var rdoList = rdoObj.getElementsByTagName('input');
            for (var i = 0; i < rdoList.length; i++) {
                if (rdoList[i].value == '01' || rdoList[i].value == '02' || rdoList[i].value == '06') {
                    rdoList[i].disabled = "disabled";
                }
                else {
                    rdoList[i].disabled = "";
                }
            }
        }

        //function Button2_onclick() {}

        function GetJobRelateYN() {
            var rbJobRelateY = document.getElementById('rbJobRelateY');
            var rbJobRelateN = document.getElementById('rbJobRelateN');
            var obj1 = document.getElementById('cbl_SD12008LB1');
            if (obj1) {
                obj1.disabled = true;
                if (rbJobRelateY.checked) { obj1.disabled = false; }
                if (rbJobRelateN.checked) { setValue(obj1, "00"); }
            }
        }

        function ChkJobRelateYN() {
            var msg1 = "";
            var obj1 = document.getElementById('cbl_SD12008LB1');
            if (!obj1) { return msg1 }; //無此物件，返回true
            var rbJobRelateY = document.getElementById('rbJobRelateY');
            var rbJobRelateN = document.getElementById('rbJobRelateN');
            if (!rbJobRelateY.checked && !rbJobRelateN.checked) {
                msg1 += "請選擇 就業關聯性(訓後工作內容與參訓職類關聯性)\n";
            }
            if (msg1 == "" && rbJobRelateY.checked && isEmpty('cbl_SD12008LB1')) {
                msg1 += "就業關聯性 選擇有關聯，請勾選子項\n";
            }
            return msg1;
        }

    </script>
    <%--<style type="text/css">.style1 {height: 30px;}</style>--%>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="bluecol_need" width="20%">職類/班別 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input id="btn_OCID" onclick="choose_class()" type="button" value="..." runat="server" class="button_b_Mini">
                                <input id="TMIDValue1" type="hidden" runat="server">
                                <input id="OCIDValue1" type="hidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="20%">學員姓名 </td>
                            <td class="whitecol" width="30%">
                                <asp:DropDownList ID="SOCID" runat="server">
                                </asp:DropDownList>
                                <input id="SLTID" type="hidden" runat="server">
                            </td>
                            <td width="20%" class="bluecol_need">離退訓日期 </td>
                            <td class="whitecol" width="30%">
                                <asp:TextBox ID="RejectTDate" runat="server" onfocus="this.blur()" Columns="20"></asp:TextBox><img id="IMG1" style="cursor: pointer" onclick="javascript:show_calendar('RejectTDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30">
                            </td>
                        </tr>
                        <tr id="trRejectDayIn14" runat="server">
                            <td class="bluecol_need">遞補期限內離退訓 </td>
                            <td colspan="3" class="whitecol">
                                <asp:CheckBox ID="cbRejectDayIn14" runat="server" Text="(兩週內)離退訓"></asp:CheckBox>
                                <asp:Label ID="labmsg1" runat="server"></asp:Label><asp:Label ID="labMakeSOCID" runat="server"></asp:Label>
                                <asp:CheckBox ID="cbRejectDayIn14_N" runat="server" Text="否"></asp:CheckBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">離訓原因 </td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="RTReasonID2" runat="server" Width="100%" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="5">
                                </asp:RadioButtonList>
                                <%--for選了其他(職前訓練須經分署/縣市政府專案認定)--%>
                                <span id="spRTReasonID22" runat="server" style="display: none">
                                    <asp:RadioButtonList ID="RTReasonID22" runat="server" Width="100%" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="4">
                                    </asp:RadioButtonList>
                                    <asp:TextBox ID="RTReasoOther2" runat="server" MaxLength="100" Width="55%"></asp:TextBox>(若選其他，其他說明為必填。)<br>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">退訓原因 </td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="RTReasonID3" runat="server" Width="100%" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="4">
                                </asp:RadioButtonList>
                                <asp:TextBox ID="RTReasoOther3" runat="server" MaxLength="100" Width="55%"></asp:TextBox>(若選其他，其他說明為必填。)<br>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">離退訓種類 </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="StudStatus" runat="server" CssClass="font" RepeatDirection="Horizontal" Enabled="False">
                                    <asp:ListItem Value="2">離訓<font color=red>(由學員提出申請並經核定者)</font></asp:ListItem>
                                    <asp:ListItem Value="3">退訓<font color=red>(經訓練單位勒令退訓者)</font></asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">離退訓原因說明 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="RTReasonThat" runat="server" Width="80%" TextMode="MultiLine" MaxLength="256" Rows="8"></asp:TextBox>
                            </td>
                        </tr>
                        <tr id="trOrgData1" runat="server">
                            <td colspan="4">
                                <%--<table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
								<tr>
									<td width="120" class="bluecol_need">
										就業單位名稱
									</td>
									<td colspan="3" class="whitecol">
										<asp:TextBox ID="OrgName" runat="server" MaxLength="50"></asp:TextBox>
									</td>
								</tr>
								<tr>
									<td class="bluecol_need">
										就業單位電話
									</td>
									<td colspan="3" class="whitecol">
										<asp:TextBox ID="JobTel" runat="server" MaxLength="25"></asp:TextBox>
									</td>
								</tr>
								<tr>
									<td class="bluecol_need">
										就業單位地址
									</td>
									<td colspan="3" class="whitecol">
										<asp:TextBox ID="JobCity" runat="server" onfocus="this.blur()" Columns="22"></asp:TextBox><input id="JobZipCode" type="hidden" name="JobZipCode" runat="server">
										<input id="btnGetZip" onclick="getZip('../../js/Openwin/zipcode.aspx', 'JobCity', 'JobZipCode')" type="button" value="..." name="btnGetZip" runat="server" class="button_b_Mini">
										<asp:TextBox ID="Jobaddress" runat="server" MaxLength="100" Columns="55"></asp:TextBox>
									</td>
								</tr>
								<tr>
									<td class="bluecol_need">
										就業單位到職日
									</td>
									<td colspan="3" class="whitecol">
										<asp:TextBox ID="JobDate" runat="server" Columns="10"></asp:TextBox><img id="JobDate1" style="cursor: pointer" onclick="javascript:show_calendar('JobDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30">
									</td>
								</tr>
								<tr>
									<td class="bluecol_need">
										就業單位薪資級距
									</td>
									<td colspan="3" class="whitecol">
										<asp:RadioButtonList ID="JobSalID" runat="server" RepeatColumns="3" CellPadding="0" CellSpacing="0" RepeatDirection="Horizontal" CssClass="font">
										</asp:RadioButtonList>
										<asp:Button ID="btnClearJobSalID" runat="server" Text="清除薪資級距" CssClass="asp_button_M"></asp:Button>
									</td>
								</tr>
							</table>--%>
                                <table class="table_nw" cellspacing="1" cellpadding="1" width="99%" border="0">
                                    <tr>
                                        <td class="bluecol_need" width="20%">就業單位到職日 </td>
                                        <td colspan="3" class="whitecol">
                                            <asp:TextBox ID="JobDate" runat="server" Columns="20" MaxLength="12"></asp:TextBox><img id="JobDate1" style="cursor: pointer" onclick="javascript:show_calendar('JobDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30">
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need">切結對象 </td>
                                        <td colspan="3" class="whitecol">
                                            <asp:RadioButtonList ID="GetJob1" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="4" CellPadding="0" CellSpacing="0">
                                                <asp:ListItem Value="1">雇主切結</asp:ListItem>
                                                <asp:ListItem Value="2">學員切結</asp:ListItem>
                                                <asp:ListItem Value="3">勞保勾稽</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need">就業單位名稱 </td>
                                        <td colspan="3" class="whitecol">
                                            <asp:TextBox ID="JobOrgName" runat="server" Columns="40" MaxLength="100"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">勞保証字號 </td>
                                        <td colspan="3" class="whitecol">
                                            <asp:TextBox ID="BusGNO" runat="server" MaxLength="30" Columns="20"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need">事業單位地址 </td>
                                        <td colspan="3" class="whitecol">
                                            <asp:TextBox ID="JobCity" runat="server" onfocus="this.blur()" Columns="30"></asp:TextBox>
                                            <input id="JobZipCode" type="hidden" name="JobZipCode" runat="server">
                                            <input id="btnGetZip" runat="server" type="button" value="..." name="btnGetZip" class="button_b_Mini">
                                            <asp:TextBox ID="Jobaddress" runat="server" MaxLength="100" Columns="66"></asp:TextBox>
                                            <%--<asp:TextBox ID="City" runat="server" Columns="13"></asp:TextBox>
										<input id="Button15" type="button" value="..." name="Button15" runat="server" class="button_b_Mini">
										<input id="BusZip" type="hidden" runat="server">
										<asp:TextBox ID="BusAddr" runat="server" Columns="35"></asp:TextBox>--%>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need">事業單位電話 </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:TextBox ID="JobTel" runat="server" Columns="20"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">事業單位傳真 </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:TextBox ID="BusFax" runat="server" Columns="20"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">職稱 </td>
                                        <td colspan="3" class="whitecol">
                                            <asp:TextBox ID="BusTitle" runat="server" Columns="30"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td id="tdJobSalID" runat="server" class="bluecol_need">薪資級距 </td>
                                        <td colspan="3" class="whitecol">
                                            <asp:RadioButtonList ID="JobSalID" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="3" CellPadding="0" CellSpacing="0">
                                            </asp:RadioButtonList>
                                            <asp:Button ID="btnClearJobSalID" runat="server" Text="清除薪資級距" CssClass="asp_button_M"></asp:Button>
                                            <br />
                                            <asp:Label ID="Label1" runat="server" Text="※勞保勾稽之薪資級距係指月投保薪資，雇主切結及學員切結之薪資級距係指月薪資總額。" ForeColor="Red"></asp:Label>
                                            <table class="font" id="MemoTable" onmouseover="document.getElementById('MemoTable').style.display='';" style="display: none; border-collapse: collapse" onmouseout="document.getElementById('MemoTable').style.display='none';" cellspacing="0" cellpadding="1" width="400" bgcolor="lemonchiffon" border="1">
                                                <tr>
                                                    <td>各訓練計畫之訓練目標、訓練對象及訓練地區均有不同，當採人工方式進行結訓學員就業調查時，可能因學員特殊性之就業樣態，而無法確切填報學員就業單位名稱、工作職稱等基本資料，<b>承辦單位可考量各該訓練計畫之屬性，將「同意承訓單位依學員就業效果之事實填報詳細說明資料，予以切結，進行就業認定」，納為規劃委訓之條件。</b>至學員切結之就業調查記錄相關文件，承訓單位於結案時，應同時影送主辦機關留存，並規定正本由承訓單位至少保存3年以上，俾供查驗。如各訓練計畫項下對於訓後就業認定及就業率核算標準另有其他(或補充)規範時，仍依該規定辦理(例:原住民訓練、保母及照顧服務人員訓練、訓用合一職前訓練)。 </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need">就業原因 </td>
                                        <td colspan="3" class="whitecol">
                                            <asp:RadioButtonList ID="GetJobCode1" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="4" CellPadding="0" CellSpacing="0">
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr id="JobCode5" runat="server">
                                        <td class="bluecol">特殊屬性訓練班次<br />
                                            結訓學員就業<br />
                                            追蹤情形說明 </td>
                                        <td colspan="3" class="whitecol"><font onmouseover="document.getElementById('MemoTable').style.display='';" style="cursor: pointer" onmouseout="document.getElementById('MemoTable').style.display='none';" color="blue">有關「特殊屬性」說明</font>
                                            <br>
                                            <asp:TextBox ID="SpecTrace" runat="server" Columns="30" TextMode="MultiLine" Rows="5" Width="66%"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">行業類別 </td>
                                        <td colspan="3" class="whitecol">
                                            <asp:DropDownList ID="ddlSCJOB" runat="server">
                                                <asp:ListItem Value="===請選擇===">===請選擇===</asp:ListItem>
                                                <asp:ListItem Value="(尚未取得該資料)">(尚未取得該資料)</asp:ListItem>
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">是否為公法救助關係 </td>
                                        <td colspan="3" class="whitecol">
                                            <asp:RadioButtonList ID="PublicRescue" runat="server" RepeatDirection="Horizontal" CssClass="font" CellPadding="0" CellSpacing="0">
                                                <asp:ListItem Value="Y">是</asp:ListItem>
                                                <asp:ListItem Value="N">否</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr id="trJobRelate1" runat="server">
                                        <td class="bluecol">就業關聯性 </td>
                                        <td colspan="3" class="whitecol">
                                            <asp:RadioButtonList ID="JobRelate" runat="server" RepeatDirection="Horizontal" CssClass="font" CellPadding="0" CellSpacing="0" RepeatLayout="Flow">
                                                <asp:ListItem Value="Y">是</asp:ListItem>
                                                <asp:ListItem Value="N">否</asp:ListItem>
                                            </asp:RadioButtonList>
                                            <asp:Button ID="btnSaveJobRelate" runat="server" CssClass="asp_button_M" Text="儲存就業關聯性" Visible="False" />
                                        </td>
                                    </tr>
                                    <tr id="trJobRelate2" runat="server">
                                        <td class="bluecol_need">就業關聯性<br />
                                            (訓後工作內容<br />
                                            與參訓職類關聯性) </td>
                                        <td colspan="3" class="whitecol">
                                            <table class="font">
                                                <tr>
                                                    <td>
                                                        <asp:RadioButton ID="rbJobRelateY" runat="server" GroupName="gnJobRelate1" Text="有關聯" />
                                                    </td>
                                                    <td>
                                                        <asp:CheckBoxList ID="cbl_SD12008LB1" runat="server" CssClass="font" RepeatLayout="Flow">
                                                        </asp:CheckBoxList>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td colspan="2">
                                                        <asp:RadioButton ID="rbJobRelateN" runat="server" GroupName="gnJobRelate1" Text="沒關聯" />
                                                    </td>
                                                    <%--<td></td><asp:Button ID="btnSaveJobRelate2" runat="server" CssClass="asp_button_M" Text="儲存就業關聯性" Visible="False" />--%>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr id="trJOBFIELD" runat="server">
                                        <td class="bluecol_need">訓後就業場域 </td>
                                        <td colspan="3" class="whitecol">
                                            <asp:DropDownList ID="ddlJOBFIELD" runat="server">
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">實際參訓時數<br />
                                (此時數將會顯示於<br />
                                受訓證明書) </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="TrainHours" runat="server" Columns="5"></asp:TextBox><asp:Label ID="LabTHours" runat="server"></asp:Label>

                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">是否賠償 </td>
                            <td colspan="3" class="whitecol">
                                <asp:DropDownList ID="NeedPay" runat="server">
                                    <asp:ListItem Value="0">請選擇</asp:ListItem>
                                    <asp:ListItem Value="Y">是</asp:ListItem>
                                    <asp:ListItem Value="N">否</asp:ListItem>
                                </asp:DropDownList>
                                <asp:LinkButton ID="LinkButton1" runat="server" ForeColor="Blue">依公式試算</asp:LinkButton>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">應賠金額 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="SumOfPay" runat="server" MaxLength="8"></asp:TextBox>
                                <input id="SumOfPay1" type="hidden" runat="server">
                            </td>
                            <td class="bluecol">已賠金額 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="HadPay" runat="server" MaxLength="8"></asp:TextBox>
                            </td>
                        </tr>
                        <tr id="Kind" runat="server">
                            <td class="bluecol">追償狀況 </td>
                            <td colspan="3" class="whitecol">
                                <table class="font" width="50%">
                                    <tr>
                                        <td rowspan="3" width="30%">
                                            <asp:RadioButtonList ID="PayStatus" runat="server" CssClass="font">
                                                <asp:ListItem Value="1">結案</asp:ListItem>
                                                <asp:ListItem Value="2">未結案</asp:ListItem>
                                                <asp:ListItem Value="3">其他</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td width="60%"></td>
                                    </tr>
                                    <tr>
                                        <td width="60%">
                                            <asp:DropDownList ID="NoClose" runat="server">
                                                <asp:ListItem Value="">請選擇</asp:ListItem>
                                                <asp:ListItem Value="1">處分送達</asp:ListItem>
                                                <asp:ListItem Value="2">行政訴訟</asp:ListItem>
                                                <asp:ListItem Value="3">強制執行</asp:ListItem>
                                                <asp:ListItem Value="4">列管追蹤</asp:ListItem>
                                                <asp:ListItem Value="5">其他</asp:ListItem>
                                            </asp:DropDownList>
                                            <asp:TextBox ID="NoClose_Desc" runat="server" Columns="22"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td width="60%">
                                            <asp:DropDownList ID="Other" runat="server">
                                                <asp:ListItem Value="">請選擇</asp:ListItem>
                                                <asp:ListItem Value="1">死亡</asp:ListItem>
                                                <%--<asp:ListItem Value="2">中心敗訴</asp:ListItem>--%>
                                                <asp:ListItem Value="2">分署敗訴</asp:ListItem>
                                                <asp:ListItem Value="3">其他</asp:ListItem>
                                            </asp:DropDownList>
                                            <asp:TextBox ID="OtherDesc" runat="server" Columns="22"></asp:TextBox>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">備註(處理進度) </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="tb_note" runat="server" Width="80%" TextMode="MultiLine" MaxLength="256" Rows="8"></asp:TextBox>
                            </td>
                        </tr>
                    </table>
                    <p align="center">
                        <asp:Button ID="Button1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>&nbsp;
					<asp:Button ID="Btn2back" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>&nbsp;
					<%--<input id="Button2" type="button" value="回上一頁" name="Button2" runat="server" class="button_b_S" onclick="return Button2_onclick()">&nbsp;--%>
                    </p>
                </td>
            </tr>
        </table>
        <input id="hidTHoours" type="hidden" name="hidTHoours" runat="server" />
        <input id="HidRejectDay" type="hidden" name="HidRejectDay" runat="server">
        <input id="HidCanOffStudExists" type="hidden" name="HidCanOffStudExists" runat="server" />
        <input id="HidUseCanOff" type="hidden" name="HidUseCanOff" runat="server" />
        <input id="HidRTReasonID" type="hidden" name="HidRTReasonID" runat="server" />
        <input id="HidvStatus" type="hidden" name="HidvStatus" runat="server" />
        <asp:HiddenField ID="hidSBID" runat="server" />
        <asp:HiddenField ID="HidSOCIDValue" runat="server" />
        <asp:HiddenField ID="Hid_STDate" runat="server" />
        <asp:HiddenField ID="Hid_FTDate" runat="server" />
    </form>
</body>
</html>
