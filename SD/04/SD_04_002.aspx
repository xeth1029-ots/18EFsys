<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_04_002.aspx.vb" Inherits="WDAIIP.SD_04_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>單月排課作業</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        var cst_inline = '';  //'inline';
        //btnShowCourse
        function OpShowCourse() {
            var RIDValue = document.getElementById('RIDValue');
            var OCIDValue1 = document.getElementById('OCIDValue1');
            wopen('SD_04_002_c1.aspx?OCID=' + OCIDValue1.value + '&RID=' + RIDValue.value, 'OpShowCourse1', 900, 850, 1);
        }

        function aloader2on() {
            var construction2 = document.getElementById("construction2");
            var form1 = document.getElementById("form1");
            form1.style.display = "none";                 //不顯示
            //construction2.style.display = "block";      //顯示     //(20180907 由於此遮罩屬於TIMS功能，因此先將此遮罩拿掉)
        }

        function aloader2off() {
            var construction2 = document.getElementById("construction2");
            construction2.style.display = "none";  //不顯示
        }

        function LessonTeah3(opentype, st, fieldname, hiddenname) {
            var RIDValue = document.getElementById('RIDValue');
            var sUrl1 = ""; //var sUrl1 = "../../SD/04/";
            if (st == '1') {
                wopen(sUrl1 + 'LessonTeah1.aspx?RID=' + RIDValue.value + '&type=' + opentype + '&fieldname=' + fieldname + '&hiddenname=' + hiddenname, 'LessonTeah1', 900, 850, 1);
            }
            if (st == '2') {
                wopen(sUrl1 + 'LessonTeah2.aspx?RID=' + RIDValue.value + '&type=' + opentype + '&fieldname=' + fieldname + '&hiddenname=' + hiddenname, 'LessonTeah2', 900, 850, 1);
            }
            //hiddenname
            if (st == '3') {
                wopen(sUrl1 + 'LessonTeah2.aspx?RID=' + RIDValue.value + '&type=' + opentype + '&fieldname=' + fieldname + '&hiddenname=' + hiddenname, 'LessonTeah3', 900, 850, 1);
            }
            //OLessonTeah4
            if (st == '4') {
                wopen(sUrl1 + 'LessonTeah2.aspx?RID=' + RIDValue.value + '&type=' + opentype + '&fieldname=' + fieldname + '&hiddenname=' + hiddenname, 'LessonTeah4', 900, 850, 1);
            }
        }

        /*
		function LessonTeah1(opentype, fieldname, hiddenname) {
		wopen('LessonTeah1.aspx?RID=' + document.form1.RIDValue.value + '&type=' + opentype + '&fieldname=' + fieldname + '&hiddenname=' + hiddenname, 'LessonTeah1', 400, 300, 1);
		}
		function LessonTeah2(opentype, st, fieldname, hiddenname) {
		wopen('LessonTeah2.aspx?RID=' + document.form1.RIDValue.value + '&type=' + opentype + '&fieldname=' + fieldname + '&hiddenname=' + hiddenname, 'LessonTeah2', 400, 300, 1);
		}
		*/

        function Course_search(Type, CourseValue, CourseName) {
            //Cst_notepad : notepad
            var RIDValue = document.getElementById('RIDValue');
            wopen('SD_04_002_Course.aspx?Type=' + Type + '&RID=' + RIDValue.value + '&CourseValue=' + CourseValue + '&CourseName=' + CourseName, 'CheckID', 900, 850, 1);
        }

        function choose_class2(num) {
            var RIDValue = document.getElementById('RIDValue');
            var DayCount = document.getElementById('DayCount'); //DayCount
            //openClass('../02/SD_02_ch.aspx?RID='+document.form1.RIDValue.value+'&AcctPlan=1&special=2&sort=2&DayCount='+document.form1.DayCount.value);
            //openClass('../02/SD_02_ch.aspx?Test=true&RID='+document.form1.RIDValue.value+'&AcctPlan=1&special=2&sort=2&DayCount='+document.form1.DayCount.value);
            if (document.form1.DayCount.value != '') {
                openClass('../02/SD_02_ch.aspx?RID=' + RIDValue.value + '&AcctPlan=1&special=2&sort=2&DayCount=' + DayCount.value);
            }
            else {
                openClass('../02/SD_02_ch.aspx?RID=' + RIDValue.value + '&AcctPlan=1&special=2&sort=2');
            }
            /*
            if(num==1)
			openClass('../02/SD_02_ch.aspx?RID='+document.form1.RIDValue.value+'&AcctPlan=1&special=2&sort=2');
			else
			openClass('../02/SD_02_ch1.aspx?RID='+document.form1.RIDValue.value+'&AcctPlan=1&special=2&sort=2');
            */
        }

        function CheckImportData() {
            var msg = '';
            if (document.getElementById('OCIDValue1').value == '') msg += '請先選擇班級\n';
            if (document.getElementById('OCIDValue2').value == '') msg += '請先選擇要載入的班級\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function chkdata() {
            var OCID = document.getElementById('OCID');
            var CSDate = document.getElementById('CSDate');
            var CFDate = document.getElementById('CFDate');
            var msg = '';
            if (OCID.selectedIndex == 0) msg += '請先選擇班級職類\n';
            if (CSDate.value == '') msg += '請輸入排課區間起始日期\n';
            else if (!checkDate(CSDate.value)) msg += '排課區間起始日期不是正確的日期格式\n';
            //else if(compareDate(document.form1.STDate.value,document.form1.CSDate.value)==1) msg+='排課起始日期不能早於開訓日期\n';
            if (CFDate.value == '') msg += '請輸入排課區間結束日期\n';
            else if (!checkDate(CFDate.value)) msg += '排課區間結束日期不是正確的日期格式\n';
            //else if(compareDate(document.form1.CFDate.value,document.form1.FTDate.value)==1) msg+='排課結束日期不能晚於結訓日期\n';
            //alert(document.form1.CFDate.value);
            //alert(document.form1.FTDate.value);
            if (msg != '') {
                alert(msg);
                return false;
            }
            aloader2on();
            return true;
        }

        /*
		function GetClass(num) {
		document.form1.OCID1.value = '';
		document.form1.OCIDValue1.value = '';
		document.form1.CSDate.value = '';
		document.form1.CFDate.value = '';
		document.getElementById('CheckBox1').checked = false;
		document.getElementById('CourseTable').style.display = 'none';
		document.getElementById('Years').innerHTML = '';
		document.getElementById('CyclType').innerHTML = '';
		document.getElementById('labTDate').innerHTML = '';
		document.getElementById('THours').innerHTML = '';
		document.getElementById('TPeriod').innerHTML = '';
		wopen('../02/SD_02_ch.aspx?RWClass=1&RID=' + document.form1.RIDValue.value, 'Class', 540, 520, 1);
		//xxxx
		if (num==1)
		wopen('../02/SD_02_ch.aspx?RID='+document.form1.RIDValue.value,'Class',540,520,1);
		else
		wopen('../02/SD_02_ch1.aspx?RID='+document.form1.RIDValue.value,'Class',540,520,1);
		}
		*/

        function GetClassTime(num) {
            switch (num) {
                case 1:
                    for (var j = 0; j < 4; j++) {
                        document.getElementById('ClassSort1_' + j).checked = document.getElementById('ClassSort2').checked;
                    }
                    if (!document.getElementById('ClassSort2').checked) document.getElementById('ClassSort4').checked = false;
                    break;
                case 2:
                    for (var j = 4; j < 8; j++) {
                        document.getElementById('ClassSort1_' + j).checked = document.getElementById('ClassSort3').checked;
                    }
                    if (!document.getElementById('ClassSort3').checked) document.getElementById('ClassSort4').checked = false;
                    break;
                case 3:
                    for (var j = 0; j < 8; j++) {
                        document.getElementById('ClassSort1_' + j).checked = document.getElementById('ClassSort4').checked;
                    }
                    document.getElementById('ClassSort2').checked = document.getElementById('ClassSort4').checked;
                    document.getElementById('ClassSort3').checked = document.getElementById('ClassSort4').checked;
                    break;
                case 4:
                    for (var j = 8; j < 12; j++) {
                        document.getElementById('ClassSort1_' + j).checked = document.getElementById('ClassSort5').checked;
                    }
                    break;
            }
        }

        function DelALLClass() {
            var OCID = document.getElementById('OCID');
            if (OCID.selectedIndex == '') {
                alert('請選擇班級!');
                return false;
            }
            else {
                return confirm('確定要刪除這個班級全部的課程資料嗎?\n\n(同意後將直接刪除)\n\n');
            }
        }

        function CheckClassTime() {
            var myvalue = getCheckBoxListValue('ClassSort1');
            var Result = true;
            Result = true;
            for (var i = 0; i < 4; i++) {
                if (myvalue.charAt(i) == '0') { Result = false; }
            }
            document.getElementById('ClassSort2').checked = Result;
            Result = true;
            for (var i = 4; i < 8; i++) {
                if (myvalue.charAt(i) == '0') { Result = false; }
            }
            document.getElementById('ClassSort3').checked = Result;
            Result = true;
            for (var i = 0; i < 8; i++) {
                if (myvalue.charAt(i) == '0') { Result = false; }
            }
            document.getElementById('ClassSort4').checked = Result;
            Result = true;
            for (var i = 8; i < 12; i++) {
                if (myvalue.charAt(i) == '0') { Result = false; }
            }
            document.getElementById('ClassSort5').checked = Result;
        }

        //新增排課
        function CheckNewCourse() {
            var TypeRadio_0 = document.getElementById('TypeRadio_0');
            var CourseIDValue = document.getElementById('CourseIDValue'); //CourseIDValue
            var Room = document.getElementById('Room'); //Room
            var OLessonTeah1Value = document.getElementById('OLessonTeah1Value'); //OLessonTeah1Value
            var TPeriodValue = document.getElementById('TPeriodValue');
            var msg = '';
            var myvalue = getCheckBoxListValue('ClassSort1');
            if (!TypeRadio_0) {
                alert('排課選項:異常(無此物件)!!');
                return false;
            }
            if (TypeRadio_0.checked) {
                if (CourseIDValue.value == '') msg += '請輸入課程代碼\n';
                if (Room.value == '') msg += '請輸入教室\n';
                if (OLessonTeah1Value.value == '') msg += '請輸入教師\n';
            }
            if (parseInt(myvalue) == 0) msg += '請選擇節次\n';
            switch (TPeriodValue.value) {
                case '01':
                    for (var i = 8; i < 12; i++) {
                        if (myvalue.charAt(i) == '1') {
                            msg += '日間班不能選擇9-12節\n';
                            break;
                        }
                    }
                    break;
                case '02':
                    for (var i = 0; i < 8; i++) {
                        if (myvalue.charAt(i) == '1') {
                            msg += '夜間班不能選擇1-8節\n';
                            break;
                        }
                    }
                    break;
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function DelClass() {
            var TPeriodValue = document.getElementById('TPeriodValue');
            var myvalue = getCheckBoxListValue('ClassSort1');
            var msg = '';
            if (parseInt(myvalue) == 0) {
                msg += '請選擇要刪除的節次\n';
            }
            else {
                if (confirm('確定要刪除這些課程嗎?\n\n(離開時請記得按儲存鈕)\n\n')) {
                    switch (TPeriodValue.value) {
                        case '01':
                            for (var i = 8; i < 12; i++) {
                                if (myvalue.charAt(i) == '1') {
                                    msg += '日間班不能選擇9-12節\n';
                                    break;
                                }
                            }
                            break;
                        case '02':
                            for (var i = 0; i < 8; i++) {
                                if (myvalue.charAt(i) == '1') {
                                    msg += '夜間班不能選擇1-8節\n';
                                    break;
                                }
                            }
                            break;
                    }
                }
                else {
                    return false;
                }
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        /*
		function Course(opentype, fieldname, hidden, tech1, tech2, room) {
		wopen('Course.aspx?type=' + opentype + '&fieldname=' + fieldname + '&hiddenname=' + hidden, 'CheckID', 550, 400, 1);
		}
		function Get_Teah(fieldname, hidden) {
		wopen('../../common/TechID.aspx?RID=' + document.form1.RIDValue.value + '&ValueField=' + hidden + '&TextField=' + fieldname, 'LessonTeah1', 350, 500, 1);
		}
		*/

        function ChangeDate(MyDate, STDate, FTDate) {
            var DecideDate = document.getElementById('DecideDate');
            if (!DecideDate) { return false; }
            DecideDate.value = window.showModalDialog('../../common/Calendar.aspx?NowDate=' + MyDate + '&STDate=' + STDate + '&FTDate=' + FTDate, window, 'dialogWidth:330px;dialogHeight:270px;');
            if (DecideDate.value != 'undefined')
                document.getElementById('Button13').click();
        }

        function GetCourseID(CourseID, TextField, ValueField, Tech1Field, TechName1Field, Tech2Field, TechName2Field, Tech3Field, TechName3Field, RoomField) {
            var RIDValue = document.getElementById('RIDValue');
            var Doc = frames['Iframe2']; //FindCourse.aspx
            Doc.document.getElementById('CourseID').value = CourseID;
            Doc.document.getElementById('TextField').value = TextField;
            Doc.document.getElementById('ValueField').value = ValueField;
            Doc.document.getElementById('Tech1Field').value = Tech1Field;
            Doc.document.getElementById('TechName1Field').value = TechName1Field;
            Doc.document.getElementById('Tech2Field').value = Tech2Field;
            Doc.document.getElementById('TechName2Field').value = TechName2Field;
            Doc.document.getElementById('Tech3Field').value = Tech3Field;
            Doc.document.getElementById('TechName3Field').value = TechName3Field;
            Doc.document.getElementById('RoomField').value = RoomField;
            Doc.document.getElementById('RID').value = RIDValue.value;
            Doc.document.getElementById('Button1').click();
        }

        //顯示單一課程時數
        var TimerID1;
        var TimerID2;
        function ShowCourseList(obj) {
            //alert('');
            //debugger;
            //var o = document.getElementById('DataGrid1');
            var DataGrid1 = document.getElementById('DataGrid1');
            var DataGrid3 = document.getElementById('DataGrid3');
            var CourseTable = document.getElementById('CourseTable');
            var CourseList = document.getElementById('CourseList');
            var LinkButton2 = document.getElementById('LinkButton2');
            if (DataGrid1 && DataGrid3 && CourseTable.style.display == cst_inline) {
                if (CourseList.style.display == 'none') {
                    CourseList.style.display = cst_inline; //'inline';
                    LinkButton2.innerHTML = '關閉目前各課程的排課時數';
                    if (DataGrid1.style.opacity == undefined) {
                        DataGrid1.style.filter = 'alpha(opacity=100)';
                    }
                    else {
                        DataGrid1.style.opacity = 100;
                    }
                    //document.getElementById('DataGrid1').style.filter = 'alpha(opacity=100)';
                    TimerID1 = setInterval("highlightit(30)", 50)
                }
                else {
                    CourseList.style.display = 'none';
                    LinkButton2.innerHTML = '檢視目前各課程的排課時數';
                    if (DataGrid1.style.opacity == undefined) {
                        DataGrid1.style.filter = 'alpha(opacity=30)';
                    }
                    else {
                        DataGrid1.style.opacity = 30;
                    }
                    //document.getElementById('DataGrid1').style.filter = 'alpha(opacity=30)';
                    TimerID2 = setInterval("highlightit(100)", 50)
                }
            }
            else {
                if (DataGrid1 && !DataGrid3 && CourseTable.style.display == cst_inline)
                    alert('目前尚未有任何的排課紀錄!');
                else {
                    alert('請先查詢 單月排課作業!');
                }
            }
        }

        function ShowCourseList4(obj) {
            //alert('');
            //debugger;
            //LinkButton4
            if (document.getElementById('DataGrid4')) {
                if (document.getElementById('CourseList2').style.display == 'none') {
                    document.getElementById('CourseList2').style.display = cst_inline; //'inline';
                    document.getElementById('LinkButton4').innerHTML = '關閉每月已排課時數';
                    //document.getElementById('DataGrid1').style.filter='alpha(opacity=100)';
                    //TimerID1=setInterval("highlightit(30)",50)
                }
                else {
                    document.getElementById('CourseList2').style.display = 'none';
                    document.getElementById('LinkButton4').innerHTML = '檢視每月已排課時數';
                    //document.getElementById('DataGrid1').style.filter='alpha(opacity=30)';
                    //TimerID2=setInterval("highlightit(100)",50)
                }
            }
            else {
                alert('查無該資料表!');
            }
        }

        function highlightit(num) {
            var o = document.getElementById('DataGrid1');
            if (num != 100) {			//表示要透明化
                if (o.style.opacity == undefined) {
                    if (o.filters.item('Alpha').Opacity > num)
                        o.filters.item('Alpha').Opacity -= 15; //Math.floor(15 * 100);
                    else
                        clearInterval(TimerID1);
                }
                else {
                    if (o.style.opacity > num)
                        o.style.opacity -= 15;
                    else
                        clearInterval(TimerID1);
                }
            }
            else {
                if (o.style.opacity == undefined) {
                    if (o.filters.item('Alpha').Opacity < num)
                        o.filters.item('Alpha').Opacity += 15; //Math.floor(15 * 100);
                    else
                        clearInterval(TimerID2);
                }
                else {
                    if (o.style.opacity < num)
                        o.style.opacity += 15;
                    else
                        clearInterval(TimerID2);
                }
            }
        }

        function ShowFrame() {
            var FrameObj = document.getElementById('FrameObj');
            var HistoryRID = document.getElementById('HistoryRID');
            var HistoryList2 = document.getElementById('HistoryList2');
            FrameObj.height = HistoryRID.rows.length * 22;
            if (FrameObj.height > 23) { FrameObj.height = 23; }
            FrameObj.style.display = HistoryList2.style.display;
        }

        /*
        function ChangeClassMode(){
		    if (document.form1.TypeRadio_0.checked)
		    {
		      document.getElementById('appoint').disabled = false;
		      document.getElementById('appButton').disabled = false;
		      document.getElementById('Button9').disabled = false;
		      document.getElementById('Button10').disabled = false;
		      document.getElementById('ClassSort2').disabled = false;
		      document.getElementById('ClassSort3').disabled = false;
		      document.getElementById('ClassSort4').disabled = false;
		      document.getElementById('ClassSort5').disabled = false;		        
		    }
		    else{
		      document.getElementById('appoint').disabled = true;
		        document.getElementById('appButton').disabled = true;
		        document.getElementById('Button9').disabled = true;
		        document.getElementById('Button10').disabled = true;
		        document.getElementById('ClassSort2').disabled = true;
		        document.getElementById('ClassSort3').disabled = true;
		        document.getElementById('ClassSort4').disabled = true;
		        document.getElementById('ClassSort5').disabled = true;			  
		    }
		} 
			
		function Locker_bt3() {
		    document.getElementById('Button3').disabled = true;
		    return true;
		}
			
		function Locker_bt3_open() {
		    document.getElementById('Button3').disabled = false;
		    return true;
		}
		*/
    </script>
</head>
<body>
    <div id="construction2" onclick="aloader2off();">
        <table width="100%" height="100%">
            <tr>
                <td align="center" valign="middle">
                    <img id="construction2-img" src="../../images/icon_construction-a.gif" alt="系統正在處理您的需求 請稍候.."></td>
            </tr>
        </table>
    </div>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;課程管理&gt;&gt;單月排課作業</asp:Label>
                </td>
            </tr>
        </table>
        <table width="100%" border="0" cellspacing="1" cellpadding="1">
            <tr>
                <td align="center">(此作業需先將課程代碼和師資資料先設定完成，再執行此作業)</td>
            </tr>
        </table>
        <table id="Table4" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td align="center">
                    <div align="left">
                        <asp:Panel ID="SearchTable" runat="server">
                            <table border="0" cellspacing="1" cellpadding="1" width="100%">
                                <tr>
                                    <td align="center">
                                        <div align="left">
                                            <table id="Table3" class="table_sch" cellpadding="1" cellspacing="1">
                                                <tr>
                                                    <td class="bluecol" style="width: 20%">訓練機構</td>
                                                    <td colspan="5" class="whitecol">
                                                        <asp:TextBox ID="center" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox>
                                                        <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                                        <input id="DayCount" type="hidden" name="DayCount" runat="server">
                                                        <input id="Button11" value="..." type="button" name="Button11" runat="server" class="button_b_Mini">
                                                        <%--'btnName=Button1 查詢班級--%>
                                                        <span style="z-index: 1; position: absolute; display: none" id="HistoryList2">
                                                            <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                                        </span>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="bluecol" style="width: 20%">班別關鍵字</td>
                                                    <td colspan="5" class="whitecol">
                                                        <asp:TextBox ID="CourKeyWord" runat="server" Width="50%" MaxLength="50"></asp:TextBox>
                                                        (可縮小下拉選擇)
                                                        <input id="Button1" type="button" value="重新查詢班級" runat="server" class="asp_button_M" />
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="bluecol_need" style="width: 20%">職類/班別 </td>
                                                    <td class="whitecol" style="table-layout: fixed; width: 20%;">
                                                        <asp:DropDownList ID="OCID" runat="server" AutoPostBack="True"></asp:DropDownList>
                                                        <iframe style="position: absolute; display: none; left: 25%" id="FrameObj" height="0" frameborder="0" width="80%" scrolling="no"></iframe>
                                                        <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                                    </td>
                                                    <td class="bluecol" style="width: 10%">年度 </td>
                                                    <td class="whitecol" style="width: 20%">
                                                        <asp:Label ID="Years" runat="server" CssClass="font"></asp:Label></td>
                                                    <td class="bluecol" style="width: 10%">期別 </td>
                                                    <td class="whitecol" style="width: 20%">
                                                        <asp:Label ID="CyclType" runat="server" CssClass="font"></asp:Label>
                                                        <input id="CyclTypeValue" type="hidden" name="CyclTypeValue" runat="server">
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="bluecol" width="10%">訓練時段 </td>
                                                    <td class="whitecol" width="15%">
                                                        <asp:Label ID="TPeriod" runat="server" CssClass="font"></asp:Label>
                                                        <input id="TPeriodValue" type="hidden" runat="server">
                                                    </td>
                                                    <td class="bluecol" width="10%">訓練期間 </td>
                                                    <td class="whitecol" width="15%">
                                                        <asp:Label ID="labTDate" runat="server" CssClass="font"></asp:Label></td>
                                                    <td class="bluecol" width="10%">訓練時數 </td>
                                                    <td class="whitecol" width="15%">
                                                        <asp:Label ID="THours" runat="server" CssClass="font"></asp:Label></td>
                                                </tr>
                                                <tr>
                                                    <td class="bluecol_need">排課區間 </td>
                                                    <td colspan="5" class="whitecol">
                                                        <asp:TextBox ID="CSDate" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                                        <span id="span1" runat="server">
                                                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= CSDate.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30"></span> ~
                                                        <asp:TextBox ID="CFDate" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                                        <span id="span2" runat="server">
                                                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= CFDate.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30"></span>&nbsp;&nbsp;
                                                        <asp:CheckBox ID="CheckBox1" runat="server" Text="訓練期間"></asp:CheckBox><asp:CheckBox ID="DesDate" runat="server" Text="指定日期排課(儲存後跳出日曆指定排課日期)" Checked="True"></asp:CheckBox>
                                                        <input id="STDate" type="hidden" name="STDate" runat="server">
                                                        <input id="FTDate" type="hidden" name="FTDate" runat="server">
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="bluecol">顯示節次 </td>
                                                    <td class="whitecol">
                                                        <asp:RadioButtonList ID="ShowClassNum" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                                            <asp:ListItem Value="全部">全部</asp:ListItem>
                                                            <asp:ListItem Value="第1-8節">第1-8節</asp:ListItem>
                                                            <asp:ListItem Value="第9-12節">第9-12節</asp:ListItem>
                                                        </asp:RadioButtonList>
                                                    </td>
                                                    <td colspan="4" class="whitecol" width="75%">
                                                        <asp:LinkButton ID="LinkButton2" runat="server" ForeColor="Blue">檢視目前各課程的排課時數</asp:LinkButton><br>
                                                        <div style="position: absolute; background-color: white; width: 65%; display: none; height: 70px" id="CourseList">
                                                            <asp:DataGrid ID="DataGrid3" runat="server" CssClass="font" AutoGenerateColumns="False" BorderColor="Black" CellPadding="2">
                                                                <ItemStyle BackColor="#ECF7FF"></ItemStyle>
                                                                <HeaderStyle ForeColor="White" BackColor="#2AAFC0"></HeaderStyle>
                                                                <Columns>
                                                                    <asp:BoundColumn DataField="CourseName" HeaderText="課程名稱"></asp:BoundColumn>
                                                                    <asp:BoundColumn DataField="MCourseName" HeaderText="主課程"></asp:BoundColumn>
                                                                    <asp:BoundColumn DataField="TotalHours" HeaderText="使用時數"></asp:BoundColumn>
                                                                </Columns>
                                                            </asp:DataGrid>
                                                        </div>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="bluecol" width="10%">系統提示 </td>
                                                    <td class="whitecol" colspan="5">
                                                        <asp:Label ID="SysInfo" runat="server" ForeColor="Red"></asp:Label></td>
                                                </tr>
                                                <tr>
                                                    <td class="bluecol">匯入排課作業 </td>
                                                    <td colspan="5" class="whitecol">
                                                        <input id="File1" disabled="disabled" size="50" type="file" name="File1" runat="server" accept=".xls,.ods" />
                                                        <asp:Button ID="Button14" runat="server" Text="匯入單月排課作業" CssClass="asp_button_M"></asp:Button>(必須為ods或xls格式)
                                                        <asp:HyperLink ID="HyperLink1" runat="server" CssClass="font" ForeColor="#8080FF" NavigateUrl="../../Doc/Class_Schedule_frm.zip">下載整批上載格式檔</asp:HyperLink>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="bluecol">載入班級 </td>
                                                    <td colspan="5" class="whitecol">
                                                        <asp:TextBox ID="OCID2" runat="server" onfocus="this.blur()" AutoPostBack="True" Width="40%"></asp:TextBox>
                                                        <input id="ChooseClassButton" onclick="choose_class2()" value="..." type="button" name="ChooseClassButton" runat="server" class="button_b_Mini">
                                                        <input id="ClearButton" value="清除" type="button" name="ClearButton" runat="server" class="asp_button_M">
                                                        <asp:Button ID="LoadIntoClass" Text="載入排課班級" runat="server" CssClass="asp_button_M"></asp:Button>
                                                        <input id="TMID2" type="hidden" name="TMID2" runat="server">
                                                        <input id="TMIDValue2" type="hidden" name="TMIDValue2" runat="server">
                                                        <input style="width: 6%" id="OCIDValue2" type="hidden" name="OCIDValue2" runat="server">
                                                        <asp:Label ID="Message" runat="server" ForeColor="Red"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                            <table width="100%">
                                                <tr>
                                                    <td colspan="6" align="center" class="whitecol">
                                                        <asp:CheckBox ID="CheckBox2" runat="server" Text="依排課區間刪除"></asp:CheckBox></td>
                                                </tr>
                                                <tr>
                                                    <td colspan="6" align="center" class="whitecol">
                                                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                                        <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                                        &nbsp;&nbsp;<asp:Button ID="Button2" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                                        &nbsp;<asp:Button ID="Button12" runat="server" Text="刪除排課資料" CssClass="asp_button_M" ToolTip="刪除排課-依排課區間刪除"></asp:Button>
                                                        &nbsp;<asp:Button ID="Button15" runat="server" Text="新增排課資料" CssClass="asp_button_M"></asp:Button>
                                                        &nbsp;<asp:Button ID="BtnExport" runat="server" CssClass="asp_Export_M" Text="匯出" />
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td colspan="6" align="center" class="whitecol">
                                                        <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label></td>
                                                </tr>
                                            </table>
                                        </div>
                                        <table id="CourseTable" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                            <tr>
                                                <td>
                                                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" PageSize="16" AllowPaging="True" CellPadding="8">
                                                        <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                        <Columns>
                                                            <asp:TemplateColumn HeaderText="日期(星期)">
                                                                <HeaderStyle Width="16%"></HeaderStyle>
                                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                                <ItemTemplate>
                                                                    <asp:LinkButton ID="LinkButton1" runat="server">LinkButton</asp:LinkButton>
                                                                </ItemTemplate>
                                                            </asp:TemplateColumn>
                                                            <asp:BoundColumn DataField="Class1" HeaderText="節次1">
                                                                <HeaderStyle HorizontalAlign="Center" Width="7%" />
                                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            </asp:BoundColumn>
                                                            <asp:BoundColumn DataField="Class2" HeaderText="節次2">
                                                                <HeaderStyle HorizontalAlign="Center" Width="7%" />
                                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            </asp:BoundColumn>
                                                            <asp:BoundColumn DataField="Class3" HeaderText="節次3">
                                                                <HeaderStyle HorizontalAlign="Center" Width="7%" />
                                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            </asp:BoundColumn>
                                                            <asp:BoundColumn DataField="Class4" HeaderText="節次4">
                                                                <HeaderStyle HorizontalAlign="Center" Width="7%" />
                                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            </asp:BoundColumn>
                                                            <asp:BoundColumn DataField="Class5" HeaderText="節次5">
                                                                <HeaderStyle HorizontalAlign="Center" Width="7%" />
                                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            </asp:BoundColumn>
                                                            <asp:BoundColumn DataField="Class6" HeaderText="節次6">
                                                                <HeaderStyle HorizontalAlign="Center" Width="7%" />
                                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            </asp:BoundColumn>
                                                            <asp:BoundColumn DataField="Class7" HeaderText="節次7">
                                                                <HeaderStyle HorizontalAlign="Center" Width="7%" />
                                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            </asp:BoundColumn>
                                                            <asp:BoundColumn DataField="Class8" HeaderText="節次8">
                                                                <HeaderStyle HorizontalAlign="Center" Width="7%" />
                                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            </asp:BoundColumn>
                                                            <asp:BoundColumn DataField="Class9" HeaderText="節次9">
                                                                <HeaderStyle HorizontalAlign="Center" Width="7%" />
                                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            </asp:BoundColumn>
                                                            <asp:BoundColumn DataField="Class10" HeaderText="節次10">
                                                                <HeaderStyle HorizontalAlign="Center" Width="7%" />
                                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            </asp:BoundColumn>
                                                            <asp:BoundColumn DataField="Class11" HeaderText="節次11">
                                                                <HeaderStyle HorizontalAlign="Center" Width="7%" />
                                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            </asp:BoundColumn>
                                                            <asp:BoundColumn DataField="Class12" HeaderText="節次12">
                                                                <HeaderStyle HorizontalAlign="Center" Width="7%" />
                                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            </asp:BoundColumn>
                                                            <asp:BoundColumn Visible="False" DataField="Vacation" HeaderText="Vacation"></asp:BoundColumn>
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
                            </table>
                        </asp:Panel>
                    </div>
                    <div align="left">
                        <table id="DetailTable" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                            <tr>
                                <td>
                                    <table id="Table1" class="table_nw" width="100%" cellpadding="1" cellspacing="1">
                                        <tr>
                                            <td class="bluecol" style="width: 20%">日期 </td>
                                            <td class="whitecol" style="width: 30%">
                                                <asp:Label ID="MyDate" runat="server" CssClass="font"></asp:Label>
                                                <asp:Label ID="MyWeek" runat="server" CssClass="font"></asp:Label>
                                                <input id="CSID" type="hidden" runat="server">
                                                <asp:Button ID="Button13" runat="server" Text="指定日期按鈕" CssClass="asp_button_M"></asp:Button>
                                            </td>
                                            <td class="bluecol" style="width: 20%">剩餘時數 </td>
                                            <td class="whitecol" style="width: 30%">
                                                <asp:Label ID="LeftHour" runat="server" CssClass="font"></asp:Label>
                                                <input id="TodayUseHour" type="hidden" runat="server">
                                                <asp:LinkButton ID="btnShowCourse" runat="server" ForeColor="Blue">檢視每月已排課時數</asp:LinkButton><br>
                                                <%-- <asp:datagrid id="DataGrid4" runat="server" AutoGenerateColumns="False" CellPadding="2"></asp:datagrid> --%>
                                                <div style="position: absolute; background-color: white; width: 65%; display: none; height: 70px" id="CourseList2">
                                                    <asp:DataGrid ID="DataGrid4" runat="server" CssClass="font" AutoGenerateColumns="False" BorderColor="Black" CellPadding="2">
                                                        <ItemStyle BackColor="#ECF7FF"></ItemStyle>
                                                        <HeaderStyle ForeColor="White" BackColor="#2AAFC0"></HeaderStyle>
                                                        <Columns>
                                                            <asp:BoundColumn DataField="ym1" HeaderText="年度月份"></asp:BoundColumn>
                                                            <asp:BoundColumn DataField="TotalHours" HeaderText="使用時數"></asp:BoundColumn>
                                                        </Columns>
                                                    </asp:DataGrid>
                                                </div>
                                            </td>
                                        </tr>
                                        <tr>
                                            <%-- <TD vAlign="top" rowSpan="7"><iframe id="IFRAME1" name="ClassFrame" src="SD_04_002_Course.aspx?TextField=CourseID&amp;HiddenField=CourseIDValue" frameBorder="0" width="100%" height="400" runat="server"></iframe></TD> --%>
                                            <td class="bluecol">課程名稱 </td>
                                            <td class="whitecol">
                                                <asp:TextBox Style="cursor: pointer" ID="CourseID" runat="server" Columns="45" ToolTip="輸入課程代碼可以自動轉換成課程名稱,點選兩下可以跳出視窗選擇課程名稱" EnableViewState="False"></asp:TextBox>
                                                <input id="CourseIDValue" type="hidden" runat="server">
                                            </td>
                                            <td class="bluecol">教室 </td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="Room" runat="server" Columns="30" MaxLength="30"></asp:TextBox></td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">
                                                <asp:Label ID="labTechN1" runat="server" Text="教師"></asp:Label></td>
                                            <td class="whitecol" colspan="3">
                                                <asp:TextBox ID="OLessonTeah1" runat="server" onfocus="this.blur()" Columns="22" ToolTip="點選兩下可以跳出視窗選擇教師"></asp:TextBox>
                                                <input id="OLessonTeah1Value" type="hidden" runat="server">
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">
                                                <asp:Label ID="labTechN2" runat="server" Text="助教1"></asp:Label></td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="OLessonTeah2" runat="server" onfocus="this.blur()" Columns="22" ToolTip="點選兩下可以跳出視窗選擇助教"></asp:TextBox>
                                                <input id="OLessonTeah2Value" type="hidden" runat="server">
                                            </td>
                                            <td class="bluecol">
                                                <asp:Label ID="labTechN3" runat="server" Text="助教2"></asp:Label></td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="OLessonTeah3" runat="server" onfocus="this.blur()" Columns="22" ToolTip="點選兩下可以跳出視窗選擇助教"></asp:TextBox>
                                                <input id="OLessonTeah3Value" type="hidden" runat="server">
                                            </td>
                                        </tr>
                                        <tr id="trlabTechN4" runat="server">
                                            <td class="bluecol">
                                                <asp:Label ID="labTechN4" runat="server" Text="助教1"></asp:Label></td>
                                            <td class="whitecol" colspan="3">
                                                <asp:TextBox ID="OLessonTeah4" runat="server" onfocus="this.blur()" Columns="22" ToolTip="點選兩下可以跳出視窗選擇教師"></asp:TextBox>
                                                <input id="OLessonTeah4Value" type="hidden" runat="server">
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">排課選項 </td>
                                            <td colspan="3" class="whitecol">
                                                <asp:RadioButtonList ID="TypeRadio" runat="server" RepeatDirection="Horizontal">
                                                    <asp:ListItem Value="0">一般</asp:ListItem>
                                                    <asp:ListItem Value="1">假日(若選假日排課,排課後課程名稱會顯示假日)</asp:ListItem>
                                                </asp:RadioButtonList>
                                            </td>
                                        </tr>
                                        <tr id="appButton" runat="server">
                                            <td rowspan="2" class="bluecol">節次 </td>
                                            <td colspan="3" class="whitecol">
                                                <asp:CheckBoxList ID="ClassSort1" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal" RepeatColumns="10">
                                                    <asp:ListItem Value="1">第1節</asp:ListItem>
                                                    <asp:ListItem Value="2">第2節</asp:ListItem>
                                                    <asp:ListItem Value="3">第3節</asp:ListItem>
                                                    <asp:ListItem Value="4">第4節</asp:ListItem>
                                                    <asp:ListItem Value="5">第5節</asp:ListItem>
                                                    <asp:ListItem Value="6">第6節</asp:ListItem>
                                                    <asp:ListItem Value="7">第7節</asp:ListItem>
                                                    <asp:ListItem Value="8">第8節</asp:ListItem>
                                                    <asp:ListItem Value="9">第9節</asp:ListItem>
                                                    <asp:ListItem Value="10">第10節</asp:ListItem>
                                                    <asp:ListItem Value="11">第11節</asp:ListItem>
                                                    <asp:ListItem Value="12">第12節</asp:ListItem>
                                                </asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="2" class="whitecol">
                                                <asp:CheckBox ID="ClassSort2" runat="server" Text="第1-4節" CssClass="font"></asp:CheckBox>
                                                <asp:CheckBox ID="ClassSort3" runat="server" Text="第5-8節" CssClass="font"></asp:CheckBox>
                                                <asp:CheckBox ID="ClassSort4" runat="server" Text="第1-8節" CssClass="font"></asp:CheckBox>
                                                <asp:CheckBox ID="ClassSort5" runat="server" Text="第9-12節" CssClass="font"></asp:CheckBox>
                                            </td>
                                            <td colspan="2" class="whitecol">
                                                <asp:Button ID="Button10" runat="server" Text="刪除排課-依節數" CssClass="asp_button_M" ToolTip="刪除排課-依節數"></asp:Button></td>
                                        </tr>
                                        <tr>
                                            <td colspan="4" align="center" class="whitecol" width="100%">
                                                <asp:Button ID="Button9" runat="server" Text="新增排課" CssClass="asp_button_M"></asp:Button></td>
                                        </tr>
                                        <tr id="appoint" runat="server">
                                            <td colspan="4" width="100%">
                                                <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <Columns>
                                                        <asp:BoundColumn DataField="ClassNum" HeaderText="節次">
                                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:TemplateColumn HeaderText="課程名稱">
                                                            <HeaderStyle Width="20%" />
                                                            <EditItemTemplate>
                                                                <asp:TextBox ID="CourseName" runat="server" Columns="20"></asp:TextBox>
                                                                <input id="CourseValue" type="hidden" runat="server">
                                                            </EditItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="教室">
                                                            <HeaderStyle Width="15%" />
                                                            <EditItemTemplate>
                                                                <asp:TextBox ID="ClassRoom" runat="server" Columns="20" MaxLength="30"></asp:TextBox>
                                                            </EditItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="教師">
                                                            <HeaderStyle Width="12%"></HeaderStyle>
                                                            <EditItemTemplate>
                                                                <asp:TextBox ID="Teacher1" runat="server" Columns="10"></asp:TextBox>
                                                                <input id="Teacher1Value" type="hidden" runat="server">
                                                            </EditItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="助教1">
                                                            <HeaderStyle Width="12%"></HeaderStyle>
                                                            <EditItemTemplate>
                                                                <asp:TextBox ID="Teacher2" runat="server" Columns="10"></asp:TextBox>
                                                                <input id="Teacher2Value" type="hidden" runat="server">
                                                            </EditItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="助教2">
                                                            <HeaderStyle Width="12%"></HeaderStyle>
                                                            <EditItemTemplate>
                                                                <asp:TextBox ID="Teacher3" runat="server" Columns="10"></asp:TextBox>
                                                                <input id="Teacher3Value" type="hidden" runat="server">
                                                            </EditItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="助教1">
                                                            <HeaderStyle Width="12%"></HeaderStyle>
                                                            <EditItemTemplate>
                                                                <asp:TextBox ID="Teacher4" runat="server" Columns="10"></asp:TextBox>
                                                                <input id="Teacher4Value" type="hidden" runat="server">
                                                            </EditItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="功能">
                                                            <HeaderStyle HorizontalAlign="Center" Width="9%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" CssClass="whitecol" />
                                                            <ItemTemplate>
                                                                <asp:Button ID="Button5" runat="server" Text="修改" CommandName="edit"></asp:Button>
                                                                <asp:Button ID="Button6" runat="server" Text="刪除" CommandName="del"></asp:Button>
                                                            </ItemTemplate>
                                                            <EditItemTemplate>
                                                                <asp:Button ID="Button7" runat="server" Text="更新" CommandName="update"></asp:Button>
                                                                <asp:Button ID="Button8" runat="server" Text="取消" CommandName="cancel"></asp:Button>
                                                            </EditItemTemplate>
                                                        </asp:TemplateColumn>
                                                    </Columns>
                                                </asp:DataGrid>
                                            </td>
                                        </tr>
                                    </table>
                                    <table width="100%">
                                        <tr>
                                            <td align="center" class="whitecol">
                                                <asp:Button ID="Button3" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                                                <%-- <asp:Label id="Labsave" runat="server" CssClass="font" ForeColor="Red" ToolTip="必免使用者重複按下儲存鈕消除">儲存鈕恢復中請稍後...</asp:Label> --%>
                                                <asp:Button ID="Button4" runat="server" Text="回排課列表" CssClass="asp_button_M"></asp:Button>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
        </table>
        <input id="DecideDate" type="hidden" runat="server" />
        <asp:HiddenField ID="hid_TNum" runat="server" />
    </form>
    <iframe style="display: none" id="Iframe2" src="FindCourse.aspx"></iframe>
</body>
</html>
