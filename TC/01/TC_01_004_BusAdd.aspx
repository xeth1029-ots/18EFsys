<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_004_BusAdd.aspx.vb" Inherits="WDAIIP.TC_01_004_BusAdd" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>開班資料(企訓專用)</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/OpenWin/openwin.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        //變更不開班理由
        function ChangeReason() {
            var OtherReason = document.getElementById('OtherReason');
            var NORID = document.getElementById('NORID');
            if (NORID.value == '99') {
                //OtherReason.style.display = 'inline';
                OtherReason.style.display = '';
            }
            else {
                OtherReason.value = '';
                OtherReason.style.display = 'none';
            }
        }

        //檢查增加課程
        function CheckAddClass() {
            var msg = '';
            var Weeks = document.getElementById('Weeks');
            var Times = document.getElementById('Times');
            if (Weeks.selectedIndex == 0) msg += '請選擇星期\n';
            if (Times.value == '') msg += '請選擇上課時段\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        //檢查資料正確性
        function CheckData() {
            var msg = '';

            if (document.getElementById('TBclass_id').value == '') msg += '請選擇班別代碼\n';
            if (document.getElementById('ClassCName').value == '') msg += '請輸入課程中文名稱\n';
            //if (document.getElementById('CyclType').value == '') msg += '請輸入期別\n';
            //if(document.getElementById('ClassEngName').value=='') msg+='請輸入課程英文名稱\n';
            if (document.getElementById('TB_career_id').value == '') msg += '請選擇訓練職類\n';
            if (document.all("TPropertyID") != null) {
                if (!isChecked(document.getElementsByName('TPropertyID'))) msg += '請選擇訓練性質\n';
            }
            if (document.getElementById('TNum').value == '') msg += '請輸入訓練人數\n';
            else if (!isUnsignedInt(document.getElementById('TNum').value)) msg += '訓練人數必須為數字\n'
            if (document.getElementById('SEnterDate').value == '') msg += '請輸入開始報名日期\n';
            else if (!checkDate(document.getElementById('SEnterDate').value)) msg += '開始報名日期不是正確的日期格式\n';
            if (document.getElementById('FEnterDate').value == '') msg += '請輸入結束報名日期\n';
            else if (!checkDate(document.getElementById('FEnterDate').value)) msg += '結束報名日期不是正確的日期格式\n';
            //if(document.getElementById('Content').value=='') msg+='請輸入課程內容\n';
            //if(document.getElementById('Purpose').value=='') msg+='請輸入課程目標\n';
            //if (document.getElementById('ExamDate').value != '' && !checkDate(document.getElementById('ExamDate').value)) msg += '甄試日期不是正確的日期格式\n';
            //if (document.getElementById('ExamDate').value != '' && checkDate(document.getElementById('ExamDate').value)) {
            //    if (document.getElementById('ExamPeriod').selectedIndex == 0) msg += '請選擇甄試時段\n';
            //}
            if (document.getElementById('TDeadline').selectedIndex == 0) msg += '請選擇訓練期限\n';
            //if(document.getElementById('TBCity').value=='') msg+='請輸入訓練地點[地區]\n';
            //if(document.getElementById('TAddress').value=='') msg+='請輸入訓練地點[地址]\n';
            if (document.getElementById('THours').value == '') msg += '請輸入訓練時數\n';
            else if (!isUnsignedInt(document.getElementById('THours').value)) msg += '訓練時數必須為數字\n';
            //if(document.getElementById('TPeriod').selectedIndex==0) msg+='請選擇上課時段\n';
            /*
            if (document.all("IsFullDate")!= null){
            if(!isChecked(document.getElementsByName('IsFullDate'))) msg+='請選擇是否為全日制\n';
            }*/
            if (document.getElementById('STDate').value == '') msg += '請輸入開訓日期\n';
            else if (!checkDate(document.getElementById('STDate').value)) msg += '結訓日期不是正確的日期格式\n';
            if (document.getElementById('FTDate').value == '') msg += '請輸入結訓日期\n';
            else if (!checkDate(document.getElementById('FTDate').value)) msg += '結訓日期不是正確的日期格式\n';
            //if(document.getElementById('CheckInDate').value=='') msg+='請輸入報到日期\n';
            //else if(!checkDate(document.getElementById('CheckInDate').value)) msg+='報到日期不是正確的日期格式\n';
            if (document.getElementById('CheckInDate').value != '' && !checkDate(document.getElementById('CheckInDate').value)) msg += '報到日期不是正確的日期格式\n';
            //if(document.getElementById('TechName').value=='') msg+='請選擇師資\n';
            /*if(isChecked(document.getElementById('NotOpen'))){
            var MyNORID=getCheckBoxListValue('NORID');
            if(parseInt(MyNORID)==0) msg+='請選擇不開班原因\n';
            else if(MyNORID.charAt(MyNORID.length-1)=='1'){
            if(document.getElementById('OtherReason').value=='') msg+='請輸入不開班原因[其他]\n';
            }
            }*/
            if (document.getElementById('CredPoint').value == '') msg += '請輸入學分數\n';
            /*
            if(document.getElementById('RoomName').value=='') msg+='請輸入上課教室名稱\n';
            if(!isChecked(document.getElementsByName('FactMode'))) msg+='請選擇場地類型\n';
            if(getRadioValue(document.getElementsByName('FactMode'))=='99' && document.getElementById('FactModeOther').value=='') msg+='請輸入場地類型[其他]\n';
            if(document.getElementById('ConNum').value=='') msg+='請輸入容納人數\n';
            */
            if (document.getElementById('ClassCate').selectedIndex == 0) msg += '請選擇訓練職能\n'; //msg+='請選擇課程類別\n';
            /*
            if (document.all("IsBusiness")!= null){
            if(isChecked(document.getElementsByName('IsBusiness')) && document.getElementById('EnterpriseName').value=='') msg+='請填寫企業包班名稱\n';
            }
            */
            //檢查報名日期及開訓日期
            var SEnterDate2 = document.getElementById('SEnterDate').value;  //報名開始日期
            var FEnterDate2 = document.getElementById('FEnterDate').value;  //報名結束日期
            var STDate2 = document.getElementById('STDate').value;          //開訓日期 
            var FTDate2 = document.getElementById('FTDate').value;          //結訓日期
            if ((Date.parse(SEnterDate2)).valueOf() >= (Date.parse(FEnterDate2)).valueOf()) { msg += '[報名結束日期]必須大於[報名開始日期]\n'; }
            if ((Date.parse(FTDate2)).valueOf() <= (Date.parse(STDate2)).valueOf()) { msg += '[結訓日期]必須大於[開訓日期]\n'; }
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function getNewDate(dd, dadd) {
            var a = new Date(dd)
            a = a.valueOf()
            a = a + dadd * 24 * 60 * 60 * 1000
            a = new Date(a)
            return (a.getFullYear() + "/" + (a.getMonth() + 1) + "/" +
                a.getDate())
        }

        //限制TextBox在MultiLine時的字數
        function checkTextLength(obj, long) {
            var maxlength = new Number(long); // Change number to your max length.

            if (obj.value.length > maxlength) {
                obj.value = obj.value.substring(0, maxlength);
                alert("限欄位長度不能大於" + maxlength + "個字元(含空白字元)，超出字元將自動截斷");
            }
        }

        //function setExamDate() {
        //    document.getElementById('ExamDate').value = '';
        //}
    </script>
    <%-- <style type="text/css">
        .auto-style4 { color: Black; text-align: right; padding: 4px 6px; background-color: #f1f9fc; border-right: 3px solid #49cbef; height: 76px; }
        .auto-style5 { color: #333333; padding: 4px; height: 76px; }
    </style>--%>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <%--<table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">
                                    首頁&gt;&gt;訓練機構管理&gt;&gt;開班資料設定
                                </asp:Label>
                                <font color="#990000">-
								<asp:Label ID="ProecessType" runat="server"></asp:Label></font> </td>
                        </tr>
                    </table>--%>
                    <asp:Label ID="ProecessType" runat="server" Visible="false"></asp:Label>
                    <asp:Label ID="Label1" runat="server" Visible="false"></asp:Label>
                    <table class="table_nw" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="bluecol" width="15%">訓練機構<font color="red">*</font> </td>
                            <td class="whitecol" width="35%">
                                <asp:TextBox ID="OrgName" runat="server" onfocus="this.blur()" Width="80%"></asp:TextBox><input id="RIDValue" type="hidden" name="RIDValue" runat="server"></td>
                            <td class="bluecol" width="15%">班別代碼<font color="#ff0000">*</font> </td>
                            <td class="whitecol" width="35%">
                                <asp:TextBox ID="TBclass_id" runat="server" onfocus="this.blur()" Columns="10" Width="60%"></asp:TextBox>
                                <input id="Choice_Button2" onclick="javascript: wopen('TC_01_004_Class.aspx', '班別代碼', 400, 400, 1)" type="button" value="選擇" name="choice_button" runat="server" class="asp_button_M">
                                <input id="clsid" type="hidden" name="clsid" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" align="center">班級中文名稱<font color="red">*</font> </td>
                            <td class="whitecol">
                                <asp:TextBox ID="ClassCName" runat="server" onfocus="this.blur()" Width="80%"></asp:TextBox></td>
                            <td class="bluecol" align="center">期別</td>
                            <td class="whitecol">
                                <asp:TextBox ID="CyclType" runat="server" onfocus="this.blur()" Columns="3" MaxLength="2" Width="30%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol" align="center">班級英文名稱</td>
                            <td class="whitecol">
                                <asp:TextBox ID="ClassEngName" runat="server" Width="80%"></asp:TextBox></td>
                            <td class="bluecol" align="center">
                                <asp:Label ID="LabTMID" runat="server">訓練職類</asp:Label><font color="red">*</font> </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Width="70%"></asp:TextBox><input id="career" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." name="career" runat="server"><input id="trainValue" style="width: 18px; height: 22px" type="hidden" name="trainValue" runat="server"><input id="jobValue" style="width: 8px; height: 22px" type="hidden" name="jobValue" runat="server"></td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類<FONT color="red">*</FONT></asp:Label></td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Width="30%" Columns="30"></asp:TextBox><input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server"><asp:RequiredFieldValidator ID="fill1b" runat="server" ErrorMessage="請選擇通俗職類" Display="None" ControlToValidate="txtCJOB_NAME"></asp:RequiredFieldValidator><input id="cjobValue" type="hidden" name="cjobValue" runat="server"></td>
                        </tr>
                        <tr>
                            <td class="bluecol" align="center">訓練人數<font color="red">*</font> </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TNum" runat="server" Columns="5" Width="20%"></asp:TextBox>人 </td>
                            <td class="bluecol">申請階段<font color="red">*</font></td>
                            <td class="whitecol">
                                <asp:DropDownList ID="AppStage" runat="server"></asp:DropDownList></td>
                            <%--<td class="bluecol" id="TPropertyID_TD" align="center" runat="server">訓練性質<font color="red">*</font> </td>
                            <td class="whitecol"><asp:RadioButtonList ID="TPropertyID" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow"></asp:RadioButtonList></td>--%>
                        </tr>
                        <tr>
                            <td class="bluecol" align="center">報名開始日期<font color="red">*</font> </td>
                            <td class="whitecol">
                                <asp:TextBox ID="SEnterDate" runat="server" Columns="10" onfocus="this.blur()" Width="40%"></asp:TextBox>
                                <img id="Img1" style="cursor: pointer" onclick="javascript:show_calendar('SEnterDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30"><br>
                                <asp:DropDownList ID="HR1" runat="server"></asp:DropDownList>時：<asp:DropDownList ID="MM1" runat="server"></asp:DropDownList>分
                            </td>
                            <td class="bluecol" align="center">報名結束日期</td>
                            <td class="whitecol">
                                <asp:TextBox ID="FEnterDate" runat="server" Columns="10" onfocus="this.blur()" Width="40%"></asp:TextBox><img id="Img2" style="cursor: pointer" onclick="javascript:show_calendar('FEnterDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30"><br />
                                <asp:DropDownList ID="HR2" runat="server"></asp:DropDownList>時：<asp:DropDownList ID="MM2" runat="server"></asp:DropDownList>分
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" align="center">上架日期 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="OnShellDate" runat="server" Columns="10" onfocus="this.blur()" Width="40%"></asp:TextBox><img id="Img_OnShellDate" style="cursor: pointer" onclick="javascript:show_calendar('OnShellDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30"><br />
                                <asp:DropDownList ID="OnShellDate_HR" runat="server"></asp:DropDownList>時：
                                <asp:DropDownList ID="OnShellDate_MI" runat="server"></asp:DropDownList>分
                            </td>
                            <td class="bluecol" align="center"></td>
                            <td class="whitecol"></td>
                        </tr>
                        <%--
						<TR>
							<TD class="bluecol">&nbsp;&nbsp;&nbsp; 課程內容<FONT color="red">*</FONT></TD>
							<TD colSpan="3"><asp:textbox id="Content" runat="server" Columns="50" Width="424px" TextMode="MultiLine" Rows="5"></asp:textbox></TD>
						</TR>
                        --%>
                        <tr>
                            <td colspan="4" width="100%">
                                <table class="font" id="Datagrid3Table" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                    <tr>
                                        <td class="table_title" align="center">課程大綱</td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:DataGrid ID="Datagrid3" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" CellPadding="8">
                                                <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                <ItemStyle CssClass="whitecol" />
                                                <Columns>
                                                    <asp:TemplateColumn HeaderText="日期">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <asp:Label ID="STrainDateLabel" runat="server"></asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="授課時間">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <asp:Label ID="PNameLabel" runat="server"></asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="時數">
                                                        <HeaderStyle Width="5%"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <asp:Label ID="PHourLabel" runat="server"></asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="課程進度／內容">
                                                        <HeaderStyle Width="25%"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <asp:TextBox ID="PContText" runat="server" onfocus="this.blur()" Width="100%" Columns="50" TextMode="MultiLine" Rows="5" Enabled="False" Height="58px"></asp:TextBox>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="學／術科">
                                                        <HeaderStyle Width="15%"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <asp:DropDownList ID="drpClassification1" runat="server" Enabled="False" AutoPostBack="True">
                                                                <asp:ListItem Value="1">學科</asp:ListItem>
                                                                <asp:ListItem Value="2">術科</asp:ListItem>
                                                            </asp:DropDownList>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="上課地點">
                                                        <HeaderStyle Width="15%"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <asp:DropDownList ID="drpPTID" runat="server" Enabled="False">
                                                            </asp:DropDownList>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="任課教師">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <input id="Tech1Value" type="hidden" name="Tech1Value" runat="server">
                                                            <asp:TextBox ID="Tech1Text" runat="server" onfocus="this.blur()" Columns="5" Enabled="False" Width="100%"></asp:TextBox>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="助教">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <input id="Tech2Value" type="hidden" name="Tech2Value" runat="server">
                                                            <asp:TextBox ID="Tech2Text" runat="server" onfocus="this.blur()" Columns="5" Enabled="False" Width="100%"></asp:TextBox>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                </Columns>
                                            </asp:DataGrid>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">課程目標<font color="red">*</font> </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="Purpose" runat="server" Columns="50" TextMode="MultiLine" Rows="5" Width="30%"></asp:TextBox></td>
                        </tr>
                        <%-- <tr>
                            <td class="bluecol">甄試日期 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="ExamDate" runat="server" Columns="10" onfocus="this.blur()" Width="40%"></asp:TextBox><img id="Img3" style="cursor: pointer" onclick="setExamDate();javascript:show_calendar('ExamDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30">
                                <asp:DropDownList ID="ExamPeriod" runat="server"></asp:DropDownList>
                            </td>
                            <td align="left"></td>
                            <td></td>
                        </tr>--%>
                        <tr>
                            <td class="bluecol">訓練時數<font color="red">*</font> </td>
                            <td class="whitecol">
                                <asp:TextBox ID="THours" runat="server" Columns="5" Width="40%"></asp:TextBox>小時 </td>
                            <td class="bluecol">訓練期限<font color="red">*</font> </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="TDeadline" runat="server"></asp:DropDownList>
                            </td>

                            <%--
							<TD class="bluecol">&nbsp;&nbsp;&nbsp; 上課時段<FONT color="red"></FONT><FONT color="red">*</FONT></TD>
							<TD><asp:dropdownlist id="TPeriod" runat="server"></asp:dropdownlist></TD>
                            --%>
                        </tr>
                        <tr id="IsFullDate_TR" runat="server">
                            <td class="bluecol">是否為<br>
                                法定全日制<font color="red">*</font> </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="IsFullDate" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="Y">是</asp:ListItem>
                                    <asp:ListItem Value="N">否</asp:ListItem>
                                </asp:RadioButtonList>
                                <span style="cursor: pointer" onclick="alert('全日制職業訓練，應符合下列條件：\n一、訓練期間一個月以上。\n二、每星期上課四次以上。\n三、每次上課日間四小時以上。\n四、每月總訓練時數一百小時以上。\n');"><font color="blue">(定義說明)</font></span> </td>
                            <td align="left"></td>
                            <td></td>
                        </tr>
                        <tr>
                            <td class="bluecol">開訓日期<font color="red">*</font> </td>
                            <td class="whitecol">
                                <asp:TextBox ID="STDate" runat="server" onfocus="this.blur()" Columns="10" Width="40%"></asp:TextBox><img id="Img4" style="cursor: pointer" onclick="javascript:show_calendar('STDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30"></td>
                            <td class="bluecol">結訓日期<font color="red">*</font> </td>
                            <td class="whitecol">
                                <asp:TextBox ID="FTDate" runat="server" onfocus="this.blur()" Columns="10" Width="40%"></asp:TextBox><img id="Img5" style="cursor: pointer" onclick="javascript:show_calendar('FTDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30"></td>
                        </tr>
                        <tr>
                            <td class="bluecol">報到日期 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="CheckInDate" runat="server" Columns="10" onfocus="this.blur()" Width="40%"></asp:TextBox><img id="Img6" style="cursor: pointer" onclick="javascript:show_calendar('CheckInDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30"></td>
                            <td class="bluecol">師資 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TechName" onblur="checkTextLength(this,40)" onkeyup="checkTextLength(this,40)" runat="server" onfocus="this.blur()" Columns="15" onChange="checkTextLength(this,40)" Width="60%"></asp:TextBox><input id="CTName" type="hidden" runat="server"><input id="Button4" type="button" value="選擇" name="Button4" runat="server" class="asp_button_M"></td>
                        </tr>
                        <tr id="NotOpenTR" runat="server">
                            <td class="bluecol">不開班 </td>
                            <td class="whitecol">
                                <asp:CheckBox ID="NotOpen" runat="server"></asp:CheckBox></td>
                            <td class="bluecol" id="IsApplic_TD" runat="server">納入志願 </td>
                            <td class="whitecol">
                                <asp:CheckBox ID="IsApplic" runat="server"></asp:CheckBox></td>
                        </tr>
                        <tr id="NORIDTR" runat="server">
                            <td class="bluecol">不開班原因 </td>
                            <td class="whitecol" colspan="3">
                                <asp:CheckBoxList ID="NORID" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow"></asp:CheckBoxList>
                                <asp:TextBox ID="OtherReason" runat="server" MaxLength="50" Width="30%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">學分數<font color="red">*</font> </td>
                            <td class="whitecol">
                                <asp:TextBox ID="CredPoint" runat="server" Columns="5" onfocus="this.blur()" Width="25%"></asp:TextBox></td>
                            <td class="bluecol" id="RoomNameTD" align="center" runat="server">上課教室<br>
                                名稱<font color="red">*</font> </td>
                            <td class="whitecol">
                                <asp:TextBox ID="RoomName" runat="server" Width="40%"></asp:TextBox></td>
                        </tr>
                        <tr id="oldPlace" runat="server">
                            <td class="bluecol">場地類型<font color="red">*</font> </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="FactMode" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="1">教室</asp:ListItem>
                                    <asp:ListItem Value="2">演講廳</asp:ListItem>
                                    <asp:ListItem Value="3">會議室</asp:ListItem>
                                    <asp:ListItem Value="99">其他(請說明)</asp:ListItem>
                                </asp:RadioButtonList>
                                <asp:TextBox ID="FactModeOther" runat="server" Width="40%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">學科場地1 </td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="SciPlaceID" runat="server" Enabled="False"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol">術科場地1 </td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="TechPlaceID" runat="server" Enabled="False"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol">學科場地2 </td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="SciPlaceID2" runat="server" Enabled="False"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol">術科場地2 </td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="TechPlaceID2" runat="server" Enabled="False"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol">學科上課<br>
                                地址 </td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="AddressSciPTID" runat="server" Enabled="False"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol">術科上課<br>
                                地址 </td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="AddressTechPTID" runat="server" Enabled="False"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol">容納人數<font color="red">*</font> </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="ConNum" runat="server" Columns="10" Width="15%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">聯絡人 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="ContactName" runat="server" MaxLength="50" Width="60%"></asp:TextBox></td>
                            <td class="whitecol"></td>
                            <td class="whitecol"></td>
                            <%--<td class="bluecol">電話 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="ContactPhone" runat="server" MaxLength="50" Width="60%"></asp:TextBox></td>--%>
                        </tr>
                        <tr id="trContactPhone_2024_N1" runat="server">
                            <td class="bluecol_need">辦公室電話</td>
                            <td class="whitecol">
                                <asp:TextBox ID="ContactPhone_1" runat="server" MaxLength="10" Width="18%" ToolTip="區碼(2~4碼)" placeholder="區碼(0開頭)"></asp:TextBox>-
                                <asp:TextBox ID="ContactPhone_2" runat="server" MaxLength="10" Width="30%" ToolTip="電話(8碼內)" placeholder="電話(8碼)"></asp:TextBox>#
                                <asp:TextBox ID="ContactPhone_3" runat="server" MaxLength="10" Width="18%" ToolTip="分機(8碼內)" placeholder="分機"></asp:TextBox>
                            </td>
                            <td class="bluecol_need">行動電話</td>
                            <td class="whitecol">
                                <asp:TextBox ID="ContactMobile_1" runat="server" MaxLength="10" Width="18%" ToolTip="手機號碼前4碼" placeholder="手機前4碼(0開頭)"></asp:TextBox>-
                                <asp:TextBox ID="ContactMobile_2" runat="server" MaxLength="10" Width="30%" ToolTip="手機號碼後6碼" placeholder="手機後6碼"></asp:TextBox>
                            </td>
                        </tr>
                        <tr id="trContactPhone_2024_N2" runat="server">
                            <td class="whitecol"></td>
                            <td class="whitecol">
                                <asp:Label ID="lab_ContactPhone_m1" runat="server" Text="(【辦公室電話】、【行動電話】至少須擇一填寫)" ForeColor="Red"></asp:Label>
                            </td>
                            <td class="whitecol"></td>
                            <td class="whitecol">
                                <asp:Label ID="lab_ContactMobile_m2" runat="server" Text="(【辦公室電話】、【行動電話】至少須擇一填寫)" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">電子郵件 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="ContactEmail" runat="server" MaxLength="64" Width="80%"></asp:TextBox></td>
                            <td class="bluecol">傳真 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="ContactFax" runat="server" MaxLength="64" Width="60%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練職能<font color="red">*</font> </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="ClassCate" runat="server"></asp:DropDownList></td>
                            <td class="bluecol">是否為企業包班<br>
                                <font color="#660033">企業包班名稱</font></td>
                            <td class="whitecol">
                                <asp:CheckBox ID="IsBusiness" runat="server" Text="企業包班"></asp:CheckBox><br>
                                &nbsp;<asp:TextBox ID="EnterpriseName" runat="server" Width="60%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="head_navy" align="center" colspan="4">上課時間 </td>
                        </tr>
                        <tr>
                            <td colspan="4">
                                <table class="font" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td class="bluecol" width="20%">星期 </td>
                                        <td class="bluecol" width="70%">時間 </td>
                                        <td class="bluecol" width="10%"></td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol" align="center">
                                            <asp:DropDownList ID="Weeks" runat="server"></asp:DropDownList></td>
                                        <td class="whitecol" align="center">
                                            <asp:TextBox ID="Times" runat="server" Columns="50" onfocus="this.blur()" Width="70%"></asp:TextBox></td>
                                        <td class="whitecol" align="center">
                                            <asp:Button ID="Button5" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_M"></asp:Button></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4">
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:TemplateColumn HeaderText="星期">
                                            <HeaderStyle Width="100px"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="Weeks1" runat="server"></asp:Label>
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:DropDownList ID="Weeks2" runat="server">
                                                </asp:DropDownList>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="上課時段">
                                            <ItemTemplate>
                                                <asp:Label ID="Times1" runat="server"></asp:Label>
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox ID="Times2" runat="server"></asp:TextBox>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <HeaderStyle Width="20%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:Button ID="Button6" runat="server" Text="修改" CausesValidation="False" CommandName="edit" CssClass="asp_button_M"></asp:Button>
                                                <asp:Button ID="Button7" runat="server" Text="刪除" CausesValidation="False" CommandName="del" CssClass="asp_button_M"></asp:Button>
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <itemstyle horizontalalign="Center" />
                                                <asp:Button ID="Button8" runat="server" Text="儲存" CausesValidation="False" CommandName="save" CssClass="asp_button_M"></asp:Button>
                                                <asp:Button ID="Button9" runat="server" Text="取消" CausesValidation="False" CommandName="cancel" CssClass="asp_button_M"></asp:Button>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="4" class="whitecol">
                                <asp:Button ID="Button1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="Button3" runat="server" Text="回查詢頁面" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <input id="PlanID" type="hidden" runat="server">
        <input id="ComIDNO" type="hidden" runat="server">
        <input id="SeqNO" type="hidden" runat="server">
        <input id="Relship" type="hidden" runat="server">
        <input id="Years" type="hidden" runat="server">
        <input id="OldClassID" type="hidden" runat="server">
        <input id="ClassEng" type="hidden" runat="server">
        <input id="TBclass" type="hidden" runat="server">
        <input id="ClassCount" type="hidden" runat="server">
        <asp:HiddenField ID="HidSYSDATE" runat="server" />
        <asp:HiddenField ID="hid_SEnterDate_old" runat="server" />
        <asp:HiddenField ID="hid_FEnterDate_old" runat="server" />
    </form>
</body>
</html>
