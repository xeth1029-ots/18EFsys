<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_004_add.aspx.vb" Inherits="WDAIIP.SD_05_004_add" %>

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
        function ShowOrg() {
            //'RTReasonID'
            /* 01:缺課時數超過規定			02:提前就業			06:訓練成績不合格 */
            var HidRTReasonID = document.getElementById('HidRTReasonID');
            var rdoObj = document.getElementById('RTReasonID');
            var rdoList = rdoObj.getElementsByTagName('input');
            document.getElementById('org_TR').style.display = 'none';
            for (var i = 0; i < rdoList.length; i++) {
                // 02:提前就業
                if (rdoList[i].checked == true) {
                    HidRTReasonID.value = rdoList[i].value;
                    break;
                }
            }
            //提前就業
            if (HidRTReasonID.value == '02') {
                document.getElementById('org_TR').style.display = 'inline';
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
            //var HidCanOffStudExists = document.getElementById('HidCanOffStudExists');
            //var SumOfPay = document.getElementById('SumOfPay');
            //var HadPay = document.getElementById('HadPay');
            var OCIDValue1 = document.getElementById('OCIDValue1');
            var SOCID = document.getElementById('SOCID');
            //var StudStatus = document.getElementById('StudStatus');
            var StudStatus = document.getElementsByName('StudStatus');
            var RejectTDate = document.getElementById('RejectTDate');
            var RTReasonID = document.getElementsByName('RTReasonID');
            var RTReasoOther = document.getElementById('RTReasoOther');

            var hidTHoours = document.getElementById('hidTHoours');
            var TrainHours = document.getElementById('TrainHours');
            //var NeedPay = document.getElementById('NeedPay');
            //var SumOfPay = document.getElementById('SumOfPay');
            //var HadPay = document.getElementById('HadPay');
            var PayStatus = document.getElementById('PayStatus');
            var NoClose = document.getElementById('NoClose');
            var Other = document.getElementById('Other');
            var msg = '';

            var SOCIDvalue = ""; //目前使用者所選擇的學號。
            var HidUseCanOff = HidUseCanOff.value; //可使用離退判斷
            //var HidCanOffStudExists = HidCanOffStudExists.value; //可以離退的學號

            //var sumofpay = parseInt(SumOfPay.value, 10);
            //var hadpay = parseInt(HadPay.value, 10);
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
                //if (HidCanOffStudExists.indexOf(SOCIDvalue) == -1) {
                //    查無可離退學號。
                //    msg += '請先至「學員資料維護」將欲辦理離退訓作業學員資料[確實依最新資料更新維護]後，再修改預算別為不補助及補助比例為0%。\n';
                //    msg += "辦理離退訓作業，系統會自動將此學員之【預算別】改為不補助，【補助比例】改為0%，請確認!"
                //}
            }

            if (!isChecked(StudStatus)) msg += '請選擇離訓或退訓\n';
            if (RejectTDate.value == '' && !RejectTDate.disabled) msg += '請輸入離退訓日期\n';
            if (RejectTDate.value != '' && !checkDate(RejectTDate.value)) msg += '退訓日期格式不正確\n';

            var rtother = ''; //離退訓原因選擇其他說明必填。
            if (getRadioValue(RTReasonID) == '98' || getRadioValue(RTReasonID) == '99') {
                rtother = 'Y'; //離退訓原因選擇其他說明必填。
            }
            //debugger;//getValue("RTReasonID");getRadioValue(document.form1.RTReasonID)
            if (!isChecked(RTReasonID)) msg += '請選擇離退訓原因\n';
            else if (rtother == 'Y' && RTReasoOther.value == '') msg += '離退訓原因選擇其他說明必填，請輸入\n';

            var HidRTReasonID = document.getElementById('HidRTReasonID');
            //if (document.form1.RTReasonID[1].checked == true)
            //if (document.getElementById('org_TR').style.display == 'inline')
            //if (HidRTReasonID.value == '01')
            /*,01:缺課時數超過規定,02:提前就業,06:訓練成績不合格,*/
            var OrgName = document.getElementById('OrgName');
            var JobTel = document.getElementById('JobTel');
            var JobZipCode = document.getElementById('JobZipCode');
            var Jobaddress = document.getElementById('Jobaddress');
            var JobDate = document.getElementById('JobDate');

            if (HidRTReasonID.value == '02') {
                if (OrgName.value == '') msg += '請輸入就業單位名稱\n'
                if (JobTel.value == '') msg += '請輸入就業單位電話\n'
                if (JobZipCode.value == '') msg += '請選擇就業單位郵遞區號\n'
                if (Jobaddress.value == '') msg += '請輸入就業單位地址\n'
                if (JobDate.value == '') {
                    msg += '請輸入就業單位到職日\n'
                }
                else {
                    if (!checkDate(JobDate.value)) {
                        msg += '【就業單位到職日】不是正確的日期格式\n';
                    }
                }
                if (getValue("JobSalID") == '') msg += '請選擇就業薪資級距\n'
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
            //if (NeedPay.selectedIndex == 0) msg += '請選擇是否賠償\n';
            //if (NeedPay.selectedIndex == 1 && SumOfPay.value == '') msg += '請輸入賠償金額\n';

            if (getValue(PayStatus) == 2) {
                if (NoClose.value == '') { msg += '請選擇追償狀況_未結案原因\n'; }
            }
            if (getValue(PayStatus) == 3) {
                if (Other.value == '') { msg += '請選擇追償狀況_其他原因\n'; }
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
            msg += "辦理離退訓作業，系統會自動將此學員之【預算別】改為不補助，【補助比例】改為0%，請確認!"
            return confirm(msg);
        }

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

    </script>
    <%--<style type="text/css">.style1 {height: 30px;}.auto-style1 {color: Black;text-align: right;padding: 4px 6px;background-color: #f1f9fc;border-right: 3px solid #49cbef;height: 45px;}.auto-style2 {color: #333333;padding: 4px;height: 45px;}</style>--%>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch" id="Table3" cellpadding="1" cellspacing="1" width="100%">
                        <tr>
                            <td class="bluecol" style="width: 20%">職類/班別 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input id="btn_OCID" onclick="choose_class()" type="button" value="..." runat="server" class="button_b_Mini" />
                                <input id="TMIDValue1" type="hidden" runat="server" />
                                <input id="OCIDValue1" type="hidden" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" style="width: 20%">學員姓名 </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:DropDownList ID="SOCID" runat="server">
                                </asp:DropDownList>
                                <input id="SLTID" type="hidden" runat="server" />
                            </td>
                            <td class="bluecol" style="width: 20%">離退訓日期 </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="RejectTDate" runat="server" onfocus="this.blur()" Columns="20"></asp:TextBox>
                                <img id="IMG1" style="cursor: pointer" onclick="javascript:show_calendar('RejectTDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">離退訓種類 </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="StudStatus" runat="server" CssClass="font" RepeatDirection="Horizontal" AutoPostBack="True">
                                    <asp:ListItem Value="2">離訓<font color=red>(由學員提出申請並經核定者)</font></asp:ListItem>
                                    <asp:ListItem Value="3">退訓<font color=red>(經訓練單位勒令退訓者)</font></asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="trRejectDayIn14" runat="server">
                            <td class="bluecol">遞補期限內離退訓 </td>
                            <td colspan="3" class="whitecol">
                                <asp:CheckBox ID="cbRejectDayIn14" runat="server" Text="(兩週內)離退訓"></asp:CheckBox>
                                <asp:Label ID="labmsg1" runat="server"></asp:Label><asp:Label ID="labMakeSOCID" runat="server"></asp:Label>
                                <asp:CheckBox ID="cbRejectDayIn14_N" runat="server" Text="否"></asp:CheckBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">離退訓原因 </td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="RTReasonID" runat="server" Width="100%" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="4">
                                </asp:RadioButtonList>
                                <asp:TextBox ID="RTReasoOther" runat="server" MaxLength="100" Width="55%"></asp:TextBox>(若選其他，其他說明為必填。)<br>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">離退訓原因說明 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="RTReasonThat" runat="server" Width="80%" TextMode="MultiLine" MaxLength="256" Rows="8"></asp:TextBox>
                            </td>
                        </tr>
                        <tr id="org_TR" runat="server">
                            <td colspan="4">
                                <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td class="bluecol" style="width: 20%">就業單位名稱 </td>
                                        <td colspan="3" class="whitecol">
                                            <asp:TextBox ID="OrgName" runat="server" MaxLength="50" Columns="40"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">就業單位電話 </td>
                                        <td colspan="3" class="whitecol">
                                            <asp:TextBox ID="JobTel" runat="server" MaxLength="25" Columns="30"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">就業單位地址 </td>
                                        <td colspan="3" class="whitecol">
                                            <asp:TextBox ID="JobCity" runat="server" onfocus="this.blur()" Columns="30"></asp:TextBox>
                                            <input id="JobZipCode" type="hidden" runat="server" />
                                            <input id="btnGetZip" type="button" value="..." name="btnGetZip" runat="server" class="button_b_Mini">
                                            <asp:TextBox ID="Jobaddress" runat="server" MaxLength="100" Columns="66"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">就業單位到職日 </td>
                                        <td colspan="3" class="whitecol">
                                            <asp:TextBox ID="JobDate" runat="server" Columns="20"></asp:TextBox>
                                            <img id="JobDate1" style="cursor: pointer" onclick="javascript:show_calendar('JobDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30">
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">就業單位薪資級距 </td>
                                        <td colspan="3" class="whitecol">
                                            <asp:RadioButtonList ID="JobSalID" runat="server" RepeatColumns="3" CellPadding="0" CellSpacing="0" RepeatDirection="Horizontal" CssClass="font">
                                            </asp:RadioButtonList>
                                            <asp:Button ID="btnClearJobSalID" runat="server" Text="清除薪資級距" CssClass="asp_button_M"></asp:Button>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">實際參訓時數<br />
                                (此時數將會顯示於受訓證明書) </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="TrainHours" runat="server" Columns="5" Width="10%"></asp:TextBox><asp:Label ID="LabTHours" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <%--<tr>
						<td class="bluecol">是否賠償 </td>
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
						</td>
						<td class="bluecol">已賠金額 </td>
						<td class="whitecol">
							<asp:TextBox ID="HadPay" runat="server" MaxLength="8"></asp:TextBox>
						</td>
					</tr>--%>
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
                                        <td>
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
                                        <td>
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
                    <p align="center" class="whitecol">
                        <asp:Button ID="Button1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                        <input id="Button2" type="button" value="回上一頁" name="Button2" runat="server" class="asp_button_M">
                    </p>
                </td>
            </tr>
        </table>
        <%--<input id="SumOfPay1" type="hidden" runat="server" />--%>
        <input id="hidTHoours" type="hidden" runat="server" />
        <input id="HidRejectDay" type="hidden" runat="server" />
        <%--<input id="HidCanOffStudExists" type="hidden" runat="server" />--%>
        <input id="HidUseCanOff" type="hidden" runat="server" />
        <input id="HidRTReasonID" type="hidden" runat="server" />
    </form>
</body>
</html>
