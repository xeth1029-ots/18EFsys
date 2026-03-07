<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_004_add.aspx.vb" Inherits="WDAIIP.TC_01_004_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>開班資料設定</title>
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
    <script type="text/javascript" language="javascript">
        /*
        function CheckData(){
		    alert('111');
		    //檢查報名日期及開訓日期
		    var msg = '';
		    var SEnterDate2 = document.getElementById('TB_SEnterDate').value ;
		    var FEnterDate2 = document.getElementById('TB_FEnterDate').value ;
		    var STDate2 = document.getElementById('TB_STDate').value ;
		    if ((Date.parse(SEnterDate2 )).valueOf() >= (Date.parse(FEnterDate2)).valueOf())
		       {msg += '[報名結束日期]必須大於[報名開始日期]\n';}
		    if ((Date.parse(STDate2)).valueOf() <= (Date.parse(FEnterDate2)).valueOf())
		       {msg += '[開訓日期]必須大於[報名結束日期]\n';}
		    if(msg!=''){
		       alert(msg);
		       return false;
		    }
		}

        且「甄試日期」最快得安排於報名截止當日起2日後。
        另因已有「報名結束日期」及「甄試日期」，可計算報名登錄可作業的最晚時間，
        計算後帶出的項目「報名登錄最晚可作業日期」，
        該日期邏輯為：報名截止日後3日內 或 甄試日前2日取其離報名截止日較近者。(均為日曆日)
        「報名開始日期」、「報名結束日期」、「甄試日期」均由班級資料中帶入，不可修改，僅顯示。
        (ex：
        1.報名截止日：104/09/09，甄試日：104/09/11，則最晚可報名登錄作業：104/09/09
        2.報名截止日：104/09/09，甄試日：104/09/15，則最晚可報名登錄作業：104/09/12)。

        function auto_FEnterDate2() {
            var msg = "";
            var ExamDate = document.getElementById('ExamDate');
            var TB_FEnterDate = document.getElementById('TB_FEnterDate');
            var FEnterDate2 = document.getElementById('FEnterDate2');
            var vFEnterDate2_old = FEnterDate2.value;
            //alert('111');
            if (isEmpty('ExamDate')) {
                msg += '甄試日期 不可為空!\n';
            }
            if (isEmpty('TB_FEnterDate')) {
                msg += '報名結束日期 不可為空!\n';
            }
            if (!isEmpty('ExamDate') && !checkDate(ExamDate.value)) {
                msg += '甄試日期 日期格式有誤!\n';
            }
            if (!isEmpty('TB_FEnterDate') && !checkDate(TB_FEnterDate.value)) {
                msg += '報名結束日期 日期格式有誤!\n';
            }
            if (msg != "") {
                FEnterDate2.value = vFEnterDate2_old;
                //alert(msg);
                return false;
            }
            var flag = getDiffDay(YMDD2MDDY(TB_FEnterDate.value), YMDD2MDDY(ExamDate.value));
            if (flag == 0) { msg += '甄試日期不能和報名結束日期同一天\n'; }
            else if (flag <= 1) { msg += '「甄試日期」最快得安排於報名截止當日起2日後\n'; }
            if (msg != "") {
                FEnterDate2.value = vFEnterDate2_old;
                return false;
            }
            var vTB_FEnterDate3 = addDateByDay(TB_FEnterDate.value, 3);
            var vExamDate2 = addDateByDay(ExamDate.value, -2);
            var flag2 = getDiffDay(YMDD2MDDY(vTB_FEnterDate3), YMDD2MDDY(vExamDate2));
            if (flag2 >= 0) {
                FEnterDate2.value = vTB_FEnterDate3;
            }
            else {
                FEnterDate2.value = vExamDate2;
            }
        }       
        */

        function wopen(url, name, width, height, k) {
            LeftPosition = (screen.width) ? (screen.width - width) / 2 : 0;
            TopPosition = (screen.availHeight) ? (screen.availHeight - height - 28) / 2 : 0;
            window.open(url, name, 'top=' + TopPosition + ',left=' + LeftPosition + ',width=' + width + ',height=' + height + ',resizable=0,scrollbars=' + k + ',status=0');
        }

        function window_onload() {
            if (form1.TB_SEnterDate.disabled) {
                var imgdateobj = document.getElementById('date1');
                if (imgdateobj) {
                    imgdateobj.style.cursor = "";
                    imgdateobj.onclick = null;
                    imgdateobj.style.display = 'none';
                }
            }
            if (form1.TB_FEnterDate.disabled) {
                var imgdateobj = document.getElementById('date2');
                if (imgdateobj) {
                    imgdateobj.style.cursor = "";
                    imgdateobj.onclick = null;
                    imgdateobj.style.display = 'none';
                }
            }
            if (form1.TB_STDate.disabled) {
                var imgdateobj = document.getElementById('date3');
                if (imgdateobj) {
                    imgdateobj.style.cursor = "";
                    imgdateobj.onclick = null;
                    imgdateobj.style.display = 'none';
                }
            }
            if (form1.TB_FTDate.disabled) {
                var imgdateobj = document.getElementById('date4');
                if (imgdateobj) {
                    imgdateobj.style.cursor = "";
                    imgdateobj.onclick = null;
                    imgdateobj.style.display = 'none';
                }
            }
            if (form1.TB_CheckInDate.disabled) {
                var imgdateobj = document.getElementById('date5');
                if (imgdateobj) {
                    imgdateobj.style.cursor = "";
                    imgdateobj.onclick = null;
                    imgdateobj.style.display = 'none';
                }
            }

            //if (form1.TB_QaySDate.disabled) {var imgdateobj = document.getElementById('date6');if (imgdateobj) {    imgdateobj.style.cursor = "";    
            //imgdateobj.onclick = null;imgdateobj.style.display = 'none';}//}//if (form1.TB_QayFDate.disabled) {var imgdateobj = document.getElementById('date7');if (imgdateobj) {    imgdateobj.style.cursor = "";    
            //imgdateobj.onclick = null;    imgdateobj.style.display = 'none';}//}
            if (form1.ExamDate.disabled) {
                var imgdateobj = document.getElementById('ImgExamDate');
                if (imgdateobj) {
                    imgdateobj.style.cursor = "";
                    imgdateobj.onclick = null;
                    imgdateobj.style.display = 'none';
                }
            }
        }

        //Level//function check_add() {var message = "";if (isEmpty(document.form1.LevelName)) {message += "請選擇階段!!\n";}if (document.form1.LevelSDate.value == '') {message += "請輸入階段起始日!!\n";}
        //if (document.form1.LevelEDate.value == '') {message += "請輸入階段結束日!!\n";}if (document.form1.LevelHour.value == '') {message += "請輸入階段時數!!\n";} else {if (!isUnsignedInt(document.form1.LevelHour.value))
        //message += '階段時數必須為數字\n';}if (message != "") {alert(message);return false;}return true;//}
        function check_value() {
            if (document.form1.TPeriodList.value == '03') {
                alert("目前暫時不能選擇此項目");
                document.form1.TPeriodList.value = '';
            }
        }

        function words() {
            var msg = "";
            msg = "";
            msg += "全日制職業訓練，應符合下列條件：\n";
            msg += "一、訓練期間一個月以上。\n";
            msg += "二、每星期上課四次以上。\n";
            msg += "三、每次上課日間四小時以上。\n";
            msg += "四、每月總訓練時數一百小時以上。\n";
            alert(msg);
        }

        function check_unit(source, args) {
            if (document.form1.tb_TPlan_str.value == '15') {//計畫為學習券時才判斷
                if (document.form1.tb_class_unit.value == '') {//沒有選擇產生班級名稱
                    args.IsValid = false;
                } else {
                    args.IsValid = true;
                }
            }
        }

        //檢查假如勾選不開班 必須輸入原因
        function CheckNotOpenReason(source, args) {
            if (document.getElementById('CB_NotOpen').checked) {
                var MyReason = getCheckBoxListValue('NORID');
                if (parseInt(MyReason) == 0)
                    args.IsValid = false;
                else
                    args.IsValid = true;
            }
        }

        function CheckOther(source, args) {
            var MyReason = getCheckBoxListValue('NORID');
            if (MyReason.charAt(MyReason.length - 1) == '1') {
                if (document.getElementById('OtherReason').value == '') {
                    args.IsValid = false;
                }
            }
        }

        function setExamDate() {
            document.getElementById('ExamDate').value = '';
        }

        function ock_CheckBox1() {
            var CheckBox1 = document.getElementById('CheckBox1');
            if (!CheckBox1) { return false; }

            var city_code = document.getElementById('city_code');
            var TaddressZIPB3 = document.getElementById('TaddressZIPB3');
            var hidTaddressZIP6W = document.getElementById('hidTaddressZIP6W');
            var TBCity = document.getElementById('TBCity');
            var TBaddress = document.getElementById('TBaddress');

            var EZip_Code = document.getElementById('EZip_Code');
            var EADDRESSZIPB3 = document.getElementById('EADDRESSZIPB3');
            var hidEADDRESSZIP6W = document.getElementById('hidEADDRESSZIP6W');
            var ECity = document.getElementById('ECity');
            var EADDRESS = document.getElementById('EADDRESS');

            var hidEZip_Code = document.getElementById('hidEZip_Code');
            var hidEADDRESSZIPB3 = document.getElementById('hidEADDRESSZIPB3');
            var hidhidEADDRESSZIP6W = document.getElementById('hidhidEADDRESSZIP6W');
            var hidECity = document.getElementById('hidECity');
            var hidEADDRESS = document.getElementById('hidEADDRESS');
            if (CheckBox1.checked) {
                EZip_Code.value = city_code.value;
                EADDRESSZIPB3.value = TaddressZIPB3.value;
                hidEADDRESSZIP6W.value = hidTaddressZIP6W.value;
                ECity.value = TBCity.value;
                EADDRESS.value = TBaddress.value;
            } else {
                EZip_Code.value = hidEZip_Code.value;
                EADDRESSZIPB3.value = hidEADDRESSZIPB3.value;
                hidEADDRESSZIP6W.value = hidhidEADDRESSZIP6W.value;
                ECity.value = hidECity.value;
                EADDRESS.value = hidEADDRESS.value;
            }
            return true;
        }
    </script>
    <%--<style type="text/css">
        .auto-style1 { display: inline-block; padding: 6px 12px; border: 1px solid #2396b8; border-radius: 2px; background-color: #31b0d5; color: #FFF; margin: 2px; width: 62px; height: 24px; }
        .auto-style2 { color: #FF0000; text-align: right; padding: 4px 6px; background-color: #f1f9fc; border-right: 3px solid #49cbef; height: 49px; }
        .auto-style3 { color: #333333; padding: 4px; width: 251px; height: 49px; }
        .auto-style4 { color: #333333; padding: 4px; height: 49px; }
    </style>--%>
</head>
<body onload="window_onload();">
    <form id="form1" method="post" runat="server">
        <table class="font" width="100%" border="0" cellspacing="1" cellpadding="1">
            <tr>
                <td colspan="4">
                    <font color="#990000">
                        <asp:Label ID="lblProecessType" runat="server" Visible="false"></asp:Label></font>
                    <%--<font color="#000000">(<font color="#ff0000">*</font>為必填欄位)</font>--%>
                    <asp:CustomValidator ID="CustomValidator1" runat="server" Display="None" ErrorMessage="請選擇產生班級名稱" ClientValidationFunction="check_unit"></asp:CustomValidator>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need" width="15%">計畫階層</td>
                <td class="whitecol" width="35%">
                    <asp:TextBox ID="TBplan" runat="server" onfocus="this.blur()" Columns="33" Width="70%"></asp:TextBox>
                    <input id="choice_button" disabled="disabled" onclick="javascript: wopen('../../Common/LevPlan.aspx', '計畫階段', 850, 400, 1)" type="button" value="選擇" name="choice_button" runat="server" class="button_b_S">
                    <asp:RequiredFieldValidator ID="plan" runat="server" Display="None" ErrorMessage="請選擇計畫階段" ControlToValidate="TBplan"></asp:RequiredFieldValidator>
                </td>
                <td class="bluecol_need" width="15%">班別代碼</td>
                <td class="whitecol" width="35%">
                    <asp:TextBox ID="TBclass_id" runat="server" Width="128px" onfocus="this.blur()"></asp:TextBox>&nbsp;
                    <input id="Choice_Button2" onclick="javascript: wopen('TC_01_004_class.aspx', '班別代碼', 1200, 660, 1)" type="button" value="選擇" name="choice_button" runat="server" class="button_b_S">
                    <asp:RequiredFieldValidator ID="class_id" runat="server" Display="None" ErrorMessage="請選擇班別代碼" ControlToValidate="TBclass_id"></asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr>
                <td id="td2" runat="server" class="bluecol_need">班級中文名稱</td>
                <td class="whitecol">
                    <asp:TextBox ID="TB_ClassName" runat="server" Columns="33" Width="50%"></asp:TextBox>
                    <input id="class_unit_button" onclick="javascript: wopen('TC_01_004_unit.aspx?textField=TB_ClassName&amp;valueField=tb_class_unit&amp;classunit=' + document.form1.tb_class_unit.value + '', '班級名稱', 250, 250, 1)" type="button" value="產生班級名稱" name="class_unit_button" runat="server" visible="false" class="button_b_M">
                    <asp:RequiredFieldValidator ID="class_name" runat="server" Display="None" ErrorMessage="請輸入班級中文名稱" ControlToValidate="TB_ClassName"></asp:RequiredFieldValidator>
                    <input id="tb_class_unit" type="hidden" name="tb_class_unit" runat="server" />
                    <input id="tb_TPlan_str" type="hidden" name="tb_TPlan_str" runat="server" />
                </td>
                <td class="bluecol">期別</td>
                <td class="whitecol">
                    <asp:TextBox ID="TB_CyclType" runat="server" Width="40%" MaxLength="5"></asp:TextBox>
                    <%--<asp:RequiredFieldValidator ID="Cycl_Type" runat="server" Display="None" ErrorMessage="請輸入期別" ControlToValidate="TB_CyclType"></asp:RequiredFieldValidator>--%>                    <%--<asp:RegularExpressionValidator ID="R_Cycl_Type" runat="server" Display="None" ErrorMessage="期別請輸入兩碼數字" ControlToValidate="TB_CyclType" ValidationExpression="[0-9]{2}"></asp:RegularExpressionValidator>--%>
                </td>
            </tr>
            <tr>
                <td id="td_ClassEngName" runat="server" class="bluecol_need">班級英文名稱</td>
                <td colspan="3" class="whitecol">
                    <asp:TextBox ID="ClassEngName" runat="server" Columns="50" Width="70%"></asp:TextBox>
                    <%--<asp:RequiredFieldValidator ID="Class_EName" runat="server" Display="None" ErrorMessage="請輸入班級英文名稱" ControlToValidate="ClassEngName"></asp:RequiredFieldValidator>--%>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">訓練職類</td>
                <td class="whitecol">
                    <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                    <input id="career" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." name="career" runat="server" class="button_b_Mini">
                    <input id="trainValue" type="hidden" name="trainValue" runat="server" />&nbsp;
                    <input id="jobValue" type="hidden" name="jobValue" runat="server" />
                    <asp:RequiredFieldValidator ID="Tcareerid" runat="server" Display="None" ErrorMessage="請選擇訓練職類" ControlToValidate="TB_career_id"></asp:RequiredFieldValidator>
                </td>
                <td class="bluecol_need">訓練性質</td>
                <td class="whitecol">
                    <asp:Label ID="lab_TPropertyID1" runat="server" Text="在職"></asp:Label>
                    <asp:HiddenField ID="Hid_RB_TPropertyID1" runat="server" Value="1" />
                </td>

                <%--<asp:ListItem Value="0">職前</asp:ListItem>--%>                <%-- <td class="bluecol_need">訓練性質</td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="RB_TPropertyID" runat="server" RepeatColumns="2" CssClass="font">
                        <asp:ListItem Value="1">在職</asp:ListItem>
                    </asp:RadioButtonList>
                    <asp:RequiredFieldValidator ID="TPropertyID" runat="server" Display="None" ErrorMessage="請選擇訓練性質" ControlToValidate="RB_TPropertyID"></asp:RequiredFieldValidator>
                </td>--%>
            </tr>
            <tr>
                <td class="bluecol_need">訓練課程類型
                </td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="rblADVANCE" runat="server" CssClass="font" RepeatDirection="Horizontal">
                        <asp:ListItem Value="01">基礎</asp:ListItem>
                        <asp:ListItem Value="02">進階</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
                <td class="whitecol"></td>
                <td class="whitecol"></td>
            </tr>
            <tr>
                <td class="bluecol_need">
                    <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label></td>
                <td colspan="3" class="whitecol">
                    <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="33" Width="60%"></asp:TextBox>
                    <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="button_b_Mini">
                    <asp:RequiredFieldValidator ID="fill1b" runat="server" Display="None" ErrorMessage="請選擇通俗職類" ControlToValidate="txtCJOB_NAME"></asp:RequiredFieldValidator>
                    <input id="cjobValue" type="hidden" name="cjobValue" runat="server">
                </td>
            </tr>
            <tr id="CompanyTR" runat="server">
                <td class="bluecol_need">企業名稱</td>
                <td colspan="3" class="whitecol">
                    <asp:TextBox ID="Companyname" runat="server" Columns="60" Width="77%"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="bluecol_need">報名開始日期</td>
                <td class="whitecol">
                    <asp:TextBox ID="TB_SEnterDate" runat="server" onfocus="this.blur()" Width="40%" Columns="20"></asp:TextBox>
                    <img id="date1" style="cursor: pointer" onclick="javascript:show_calendar('TB_SEnterDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30">
                    <asp:RequiredFieldValidator ID="SEnterDate" runat="server" Display="None" ErrorMessage="請選擇報名開始日期" ControlToValidate="TB_SEnterDate"></asp:RequiredFieldValidator><br>
                    <asp:DropDownList ID="HR1" runat="server"></asp:DropDownList>時：
                    <asp:DropDownList ID="MM1" runat="server"></asp:DropDownList>分
                </td>
                <td class="bluecol_need">報名結束日期</td>
                <td class="whitecol">
                    <asp:TextBox ID="TB_FEnterDate" runat="server" onfocus="this.blur()" Width="40%" Columns="20"></asp:TextBox>
                    <img id="date2" style="cursor: pointer" onclick="javascript:show_calendar('TB_FEnterDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30">
                    <asp:RequiredFieldValidator ID="FEnterDate" runat="server" Display="None" ErrorMessage="請選擇報名結束日期" ControlToValidate="TB_FEnterDate"></asp:RequiredFieldValidator><br>
                    <asp:DropDownList ID="HR2" runat="server"></asp:DropDownList>時：
                    <asp:DropDownList ID="MM2" runat="server"></asp:DropDownList>分
                </td>
            </tr>
            <tr>
                <td id="td_ExamDate" runat="server" class="bluecol_need">甄試日期</td>
                <td class="whitecol">
                    <asp:TextBox ID="ExamDate" runat="server" Width="40%" MaxLength="11" Columns="20"></asp:TextBox>
                    <img id="ImgExamDate" style="cursor: pointer" onclick="javascript:setExamDate();show_calendar('ExamDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" runat="server">
                    <asp:DropDownList ID="ExamPeriod" runat="server"></asp:DropDownList>
                    <asp:Label ID="lab_msg_ExamDate" runat="server"></asp:Label>
                    <span id="spExamDateTime" runat="server">
                        <br>
                        <asp:DropDownList ID="HR6" runat="server"></asp:DropDownList>時：                   
                        <asp:DropDownList ID="MM6" runat="server"></asp:DropDownList>分
                    </span>
                </td>
                <td class="bluecol_need">報名登錄最晚<br />
                    可作業時間</td>
                <td class="whitecol">
                    <asp:TextBox ID="FEnterDate2" runat="server" onfocus="this.blur()" Width="40%" MaxLength="11" Columns="20"></asp:TextBox>
                    <img id="Img_FEnterDate2" style="cursor: pointer" onclick="javascript:show_calendar('FEnterDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30">
                    <%-- Width="40%" onfocus="this.blur()"
                        &nbsp;<input id="btnFEnterDate2" style="width: 32px; height: 24px; /*display: none;*/" onclick="javascript: auto_FEnterDate2();" type="button" value="計算" runat="server" class="button_b_S">--%><br>
                    <asp:DropDownList ID="HR5" runat="server"></asp:DropDownList>時：
                    <asp:DropDownList ID="MM5" runat="server"></asp:DropDownList>分
                </td>
            </tr>
            <%--<tr>
                <td class="bluecol_need">問卷調查開始日期</td>
                <td class="whitecol">
                    <asp:TextBox ID="TB_QaySDate" runat="server" onfocus="this.blur()" Width="40%"></asp:TextBox>
                    <img id="date6" style="cursor: pointer" onclick="javascript:show_calendar('TB_QaySDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                    <asp:RequiredFieldValidator ID="QaySDate" runat="server" Display="None" ErrorMessage="請選擇問卷調查開始日期" ControlToValidate="TB_QaySDate"></asp:RequiredFieldValidator><br>
                    <asp:DropDownList ID="HR3" runat="server"></asp:DropDownList>時：
                    <asp:DropDownList ID="MM3" runat="server"></asp:DropDownList>分
                </td>
                <td class="bluecol_need">問卷調查結束日期</td>
                <td class="whitecol">
                    <asp:TextBox ID="TB_QayFDate" runat="server" onfocus="this.blur()" Width="40%"></asp:TextBox>
                    <img id="date7" style="cursor: pointer" onclick="javascript:show_calendar('TB_QayFDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                    <asp:RequiredFieldValidator ID="QayFDate" runat="server" Display="None" ErrorMessage="請選擇問卷調查結束日期" ControlToValidate="TB_QayFDate"></asp:RequiredFieldValidator><br>
                    <asp:DropDownList ID="HR4" runat="server"></asp:DropDownList>時：
                    <asp:DropDownList ID="MM4" runat="server"></asp:DropDownList>分
                </td>
            </tr>--%>
            <tr>
                <td class="bluecol_need">課程內容</td>
                <td class="whitecol">
                    <asp:TextBox ID="TB_Content" runat="server" Columns="25" TextMode="MultiLine" Rows="5" Width="88%"></asp:TextBox>
                    <asp:RequiredFieldValidator ID="Content" runat="server" Display="None" ErrorMessage="請輸入課程內容" ControlToValidate="TB_Content"></asp:RequiredFieldValidator>
                </td>
                <td class="bluecol_need">訓練人數</td>
                <td class="whitecol">
                    <asp:TextBox ID="TB_TNum" runat="server" Width="30%" MaxLength="5"></asp:TextBox>
                    <%--<asp:RequiredFieldValidator ID="TNum" runat="server" Display="None" ErrorMessage="請輸入訓練人數" ControlToValidate="TBaddress"></asp:RequiredFieldValidator><asp:RegularExpressionValidator ID="Re_num" runat="server" Display="None" ErrorMessage="請輸入數字" ControlToValidate="TB_TNum" ValidationExpression="[0-9]+"></asp:RegularExpressionValidator>--%>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">課程目標</td>
                <td class="whitecol">
                    <asp:TextBox ID="TB_Purpose" runat="server" Columns="25" TextMode="MultiLine" Rows="5" Width="88%"></asp:TextBox>
                    <asp:RequiredFieldValidator ID="Purpose" runat="server" Display="None" ErrorMessage="請輸入課程目標" ControlToValidate="TB_Purpose"></asp:RequiredFieldValidator>
                </td>
                <td class="bluecol_need">訓練期限</td>
                <td class="whitecol">
                    <asp:DropDownList ID="TDeadline_List" runat="server"></asp:DropDownList>
                    <asp:RequiredFieldValidator ID="Re_TDeadline_List" runat="server" Display="None" ErrorMessage="請選擇訓練期限" ControlToValidate="TDeadline_List"></asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">上課地點</td>
                <td colspan="3" class="whitecol">
                    <input id="city_code" onfocus="this.blur()" runat="server" maxlength="3" />－
                    <input id="TaddressZIPB3" maxlength="3" runat="server" />
                    <input id="hidTaddressZIP6W" type="hidden" runat="server" />
                    <asp:Literal ID="Litcity_code" runat="server"></asp:Literal><%--郵遞區號查詢--%>
                    <br />
                    <asp:TextBox ID="TBCity" runat="server" Width="20%" onfocus="this.blur()"></asp:TextBox>
                    <input id="Bt1_city_zip" type="button" value="..." name="city_zip" runat="server" class="button_b_Mini" disabled="disabled" />
                    <asp:TextBox ID="TBaddress" runat="server" Width="60%"></asp:TextBox>
                    <%--<asp:RequiredFieldValidator ID="reqFVcity" runat="server" Display="None" ErrorMessage="請選擇上課地點縣市" ControlToValidate="TBCity"></asp:RequiredFieldValidator>
                    <asp:RequiredFieldValidator ID="reqFVaddress" runat="server" Display="None" ErrorMessage="請輸入上課地點地址" ControlToValidate="TBaddress"></asp:RequiredFieldValidator>--%>
                </td>
            </tr>
            <tr id="trEADDRESS" runat="server">
                <td id="td_EADDRESS" runat="server" class="bluecol_need">甄試地點</td>
                <td colspan="3" class="whitecol">
                    <asp:CheckBox ID="CheckBox1" runat="server" Text="同上課地點"></asp:CheckBox><br />
                    <input id="EZip_Code" onfocus="this.blur()" runat="server" maxlength="3" />－
                    <input id="EADDRESSZIPB3" maxlength="3" runat="server" />
                    <input id="hidEADDRESSZIP6W" type="hidden" runat="server" />
                    <asp:Literal ID="LitEZipCode" runat="server"></asp:Literal><%--郵遞區號查詢--%>
                    <asp:Label ID="lab_msg_EADDRESS" runat="server"></asp:Label>
                    <br />
                    <asp:TextBox ID="ECity" runat="server" Width="20%" onfocus="this.blur()"></asp:TextBox>
                    <input id="Ezipbtn" type="button" value="..." runat="server" class="button_b_Mini" />
                    <asp:TextBox ID="EADDRESS" runat="server" Width="60%"></asp:TextBox>
                    <asp:RequiredFieldValidator ID="reqFVcity2" runat="server" Display="None" ErrorMessage="請選擇甄試地點縣市" ControlToValidate="ECity"></asp:RequiredFieldValidator>
                    <asp:RequiredFieldValidator ID="reqFVaddress2" runat="server" Display="None" ErrorMessage="請輸入甄試地點地址" ControlToValidate="EADDRESS"></asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr>
                <td class="bluecol">訓練單位地址</td>
                <td colspan="3" class="whitecol">
                    <asp:Label ID="LabAdd" runat="server" Width="440px"></asp:Label></td>
            </tr>
            <tr>
                <td class="bluecol_need">訓練時數</td>
                <td class="whitecol">
                    <asp:TextBox ID="TB_THours" runat="server" Width="50%" MaxLength="5"></asp:TextBox>
                    <asp:RequiredFieldValidator ID="THours" runat="server" Display="None" ErrorMessage="請輸入訓練時數" ControlToValidate="TB_THours"></asp:RequiredFieldValidator><asp:RegularExpressionValidator ID="Re_hours" runat="server" Display="None" ErrorMessage="請輸入數字" ControlToValidate="TB_THours" ValidationExpression="[0-9]+"></asp:RegularExpressionValidator>
                </td>
                <td class="whitecol">&nbsp;</td>
                <td class="whitecol">&nbsp;&nbsp;&nbsp;&nbsp;</td>
            </tr>
            <tr>
                <td class="bluecol_need">訓練時段</td>
                <td class="whitecol" colspan="3">
                    <asp:DropDownList ID="TPeriodList" runat="server"></asp:DropDownList>
                    <asp:RequiredFieldValidator ID="Re_TPeriod_List" runat="server" Display="None" ErrorMessage="請選擇訓練時段" ControlToValidate="TPeriodList"></asp:RequiredFieldValidator>
                    <span id="trTB_NOTE3" runat="server">
                        <br />
                        填寫方式:(每/隔) 週(一~日)00:00~23:59<br />
                        <asp:TextBox ID="TB_NOTE3" runat="server" Columns="25" TextMode="MultiLine" Rows="5" Width="77%"></asp:TextBox>
                    </span>
                </td>
            </tr>
            <%--<tr>
                <td class="bluecol_need">是否為法定全日制</td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="Radio_isfulldate" runat="server" Width="143px" CssClass="font" RepeatDirection="Horizontal">
                        <asp:ListItem Value="Y">是</asp:ListItem>
                        <asp:ListItem Value="N">否</asp:ListItem>
                    </asp:RadioButtonList>
                    <asp:Label ID="Label1" runat="server" ForeColor="Blue" Font-Underline="True">(定義說明)</asp:Label>
                    <asp:RequiredFieldValidator ID="Re_isfulldate" runat="server" Display="None" ErrorMessage="請選擇是否為法定全日制" ControlToValidate="Radio_isfulldate"></asp:RequiredFieldValidator>
                </td>
            </tr>--%>
            <tr>
                <td class="bluecol_need">開訓日期</td>
                <td class="whitecol">
                    <span id="span1" runat="server">
                        <asp:TextBox ID="TB_STDate" runat="server" onfocus="this.blur()" Width="40%"></asp:TextBox>
                        <img id="date3" style="cursor: pointer" onclick="javascript:show_calendar('<%= TB_STDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                        <asp:RequiredFieldValidator ID="STDate" runat="server" Display="None" ErrorMessage="請選擇開訓日期" ControlToValidate="TB_STDate"></asp:RequiredFieldValidator>
                    </span>

                </td>
                <td class="bluecol_need">結訓日期</td>
                <td class="whitecol">
                    <span id="span2" runat="server">
                        <asp:TextBox ID="TB_FTDate" runat="server" onfocus="this.blur()" Width="40%"></asp:TextBox>
                        <img id="date4" style="cursor: pointer" onclick="javascript:show_calendar('<%=TB_FTDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                        <asp:RequiredFieldValidator ID="FTDate" runat="server" Display="None" ErrorMessage="請選擇結訓日期" ControlToValidate="TB_FTDate"></asp:RequiredFieldValidator>
                    </span>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">報到日期</td>
                <td class="whitecol">
                    <span id="span3" runat="server">
                        <asp:TextBox ID="TB_CheckInDate" runat="server" onfocus="this.blur()" Width="40%"></asp:TextBox>
                        <img id="date5" style="cursor: pointer" onclick="javascript:show_calendar('<%= TB_CheckInDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                        <asp:RequiredFieldValidator ID="CheckInDate" runat="server" Display="None" ErrorMessage="請選擇報到日期" ControlToValidate="TB_CheckInDate"></asp:RequiredFieldValidator>
                    </span>
                </td>
                <td class="bluecol">導師名稱</td>
                <td class="whitecol">
                    <asp:TextBox ID="CTName" runat="server" Width="50%" MaxLength="40"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="bluecol">納入志願</td>
                <td class="whitecol">
                    <asp:CheckBox ID="CB_IsApplic" runat="server" CssClass="font"></asp:CheckBox></td>
                <td class="bluecol">不開班</td>
                <td class="whitecol">
                    <asp:CheckBox ID="CB_NotOpen" runat="server" CssClass="font" Enabled="False"></asp:CheckBox><asp:CustomValidator ID="CheckReason" runat="server" Display="None" ErrorMessage="請選擇不開班原因" ClientValidationFunction="CheckNotOpenReason"></asp:CustomValidator></td>
            </tr>
            <tr>
                <td class="bluecol">不開班原因</td>
                <td colspan="3" class="whitecol">
                    <asp:CheckBoxList ID="NORID" runat="server" RepeatDirection="Horizontal" CellSpacing="0" CellPadding="0" RepeatLayout="Flow"></asp:CheckBoxList>
                    <asp:TextBox ID="OtherReason" runat="server" Width="50%"></asp:TextBox>
                    <input id="NORIDValue" type="hidden" runat="server">
                    <asp:CustomValidator ID="CustomValidator2" runat="server" Display="None" ErrorMessage="請輸入不開班原因[其他]" ClientValidationFunction="CheckOther"></asp:CustomValidator>
                </td>
            </tr>
            <%--<tr>
                <td colspan="4">
                    <table width="100%" id="tb_CLASSLEVEL" runat="server">
                        <tr>
                            <td colspan="4" align="center" class="table_title">課程階段</td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">課程階段</td>
                            <td colspan="3" class="whitecol">
                                <asp:DropDownList ID="TB_LevelType" runat="server" AutoPostBack="True">
                                    <asp:ListItem Value="0">無</asp:ListItem>
                                    <asp:ListItem Value="1">一</asp:ListItem>
                                    <asp:ListItem Value="2">二</asp:ListItem>
                                    <asp:ListItem Value="3">三</asp:ListItem>
                                    <asp:ListItem Value="4">四</asp:ListItem>
                                </asp:DropDownList>
                                <asp:RequiredFieldValidator ID="LevelType" runat="server" Display="None" ErrorMessage="請選擇課程階段" ControlToValidate="TB_LevelType"></asp:RequiredFieldValidator>(請依此處選擇的課程階段輸入下列階段起訖日及時數)
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">階段</td>
                            <td class="whitecol">
                                <asp:DropDownList ID="LevelName" runat="server"></asp:DropDownList></td>
                            <td class="bluecol_need">階段起始日</td>
                            <td class="whitecol">
                                <span id="span4" runat="server">
                                    <asp:TextBox ID="LevelSDate" runat="server" onfocus="this.blur()" Width="40%"></asp:TextBox>
                                    <img id="Img2" style="cursor: pointer" onclick="javascript:show_calendar('<%= LevelSDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">階段結束日</td>
                            <td class="whitecol" align="left">
                                <span id="span5" runat="server">
                                    <asp:TextBox ID="LevelEDate" runat="server" onfocus="this.blur()" Width="40%"></asp:TextBox>
                                    <img id="Img3" style="cursor: pointer" onclick="javascript:show_calendar('<%= LevelEDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                </span>
                            </td>
                            <td class="bluecol">階段時數</td>
                            <td class="whitecol">
                                <asp:TextBox ID="LevelHour" runat="server" Width="50%"></asp:TextBox>
                                <asp:Button ID="add_but" runat="server" Text="新增" CausesValidation="False" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4">
                                <asp:DataGrid ID="DG_ClassLevel" runat="server" CssClass="font" AutoGenerateColumns="False" Width="100%" CellPadding="8">
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <Columns>
                                        <asp:TemplateColumn HeaderText="階段名稱">
                                            <ItemTemplate>
                                                <asp:HiddenField ID="HidCCLID" runat="server" />
                                                <asp:HiddenField ID="HidLevelName" runat="server" />
                                                <asp:Label ID="LevelName" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="階段起始日">
                                            <ItemTemplate>
                                                <asp:Label ID="LevelSDate" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="階段結束日">
                                            <ItemTemplate>
                                                <asp:Label ID="LevelEDate" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="階段時數">
                                            <ItemTemplate>
                                                <asp:Label ID="LevelHour" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <ItemStyle HorizontalAlign="Center" Font-Size="Small" />
                                            <ItemTemplate>
                                                <asp:LinkButton ID="btnDel" runat="server" CssClass="asp_Export_M" CommandName="Del">刪除</asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>&nbsp;
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>--%>
        </table>

        <table width="100%">
            <tr>
                <td align="center" class="whitecol" width="100%">
                    <asp:Button ID="bt_save" runat="server" Text="儲存" CssClass="asp_Export_M"></asp:Button>
                    <asp:Button ID="next_class" runat="server" CssClass="asp_Export_M" Visible="False" Enabled="False" Text="維護下一班" CausesValidation="False"></asp:Button>
                    <input onclick="javascript: window.open('../03/TC_03_oper.aspx', '', 'width=1200,height=660,location=0,status=0,menubar=0,scrollbars=0,resizable=0');" type="button" value="時數迄日換算" class="asp_Export_M">
                    <asp:Button ID="Button1" runat="server" Text="回上一頁" CausesValidation="False" CssClass="asp_Export_M"></asp:Button>
                    <asp:ValidationSummary ID="Summary" runat="server" ShowMessageBox="True" ShowSummary="False" DisplayMode="List"></asp:ValidationSummary>
                </td>
            </tr>
        </table>
        <input id="PlanIDValue" type="hidden" name="PlanIDValue" runat="server" />
        <input id="P_ComIDNO" type="hidden" name="P_ComIDNO" runat="server" />
        <input id="P_SeqNO" type="hidden" name="P_SeqNO" runat="server" />
        <input id="P_Years" type="hidden" name="P_Years" runat="server" />
        <input id="P_Relship" type="hidden" name="P_Relship" runat="server" />
        <input id="Re_ID" type="hidden" name="Re_ID" runat="server" />
        <input id="clsid" type="hidden" name="clsid" runat="server" />
        <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
        <input id="OldClassID" type="hidden" runat="server" />
        <input id="ClassEng" type="hidden" name="ClassEng" runat="server" />
        <input id="TBclass" type="hidden" name="TBclass" runat="server" />
        <input id="hid_classnum" type="hidden" name="hid_classnum" runat="server" />
        <asp:HiddenField ID="Hid_RID1" runat="server" />
        <asp:HiddenField ID="hidEZip_Code" runat="server" />
        <asp:HiddenField ID="hidEADDRESSZIPB3" runat="server" />
        <asp:HiddenField ID="hidhidEADDRESSZIP6W" runat="server" />
        <asp:HiddenField ID="hidECity" runat="server" />
        <asp:HiddenField ID="hidEADDRESS" runat="server" />
        <asp:HiddenField ID="Hid_PERC100" runat="server" />
    </form>
</body>
</html>
