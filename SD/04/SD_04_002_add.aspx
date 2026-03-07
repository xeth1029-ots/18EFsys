<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_04_002_add.aspx.vb" Inherits="WDAIIP.SD_04_002_add" %>

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
        //if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript">
        function DelChoicClass() {
            return confirm('確定要刪除這個選擇的課程資料嗎?\n(同意後將直接刪除)\n');
        }

        function returnValue(CourseName, CourseValue, Tech1, TechName1, Tech2, TechName2, Tech3, TechName3, Room) {
            var oCourseID = document.getElementById('CourseID');
            var oCourseIDValue = document.getElementById('CourseIDValue');
            var oOLessonTeah1 = document.getElementById('OLessonTeah1');
            var oOLessonTeah1Value = document.getElementById('OLessonTeah1Value');
            var oOLessonTeah2 = document.getElementById('OLessonTeah2');
            var oOLessonTeah2Value = document.getElementById('OLessonTeah2Value');
            var oOLessonTeah3 = document.getElementById('OLessonTeah3');
            var oOLessonTeah3Value = document.getElementById('OLessonTeah3Value');
            var oRoom = document.getElementById('Room');
            if (oCourseID != null) oCourseID.value = CourseName;
            if (oCourseIDValue != null) oCourseIDValue.value = CourseValue;
            if (oOLessonTeah1 != null) oOLessonTeah1.value = TechName1;
            if (oOLessonTeah1Value != null) oOLessonTeah1Value.value = Tech1;
            if (oOLessonTeah2 != null) oOLessonTeah2.value = TechName2;
            if (oOLessonTeah2Value != null) oOLessonTeah2Value.value = Tech2;
            if (oOLessonTeah3 != null) oOLessonTeah3.value = TechName3;
            if (oOLessonTeah3Value != null) oOLessonTeah3Value.value = Tech3;
            if (oRoom != null) oRoom.value = Room;
            //if (window.document.form1.CourseID != null) window.document.form1.CourseID.value = CourseName;
            //if (window.document.form1.CourseIDValue != null) window.document.form1.CourseIDValue.value = CourseValue;
            //if (window.document.form1.OLessonTeah1 != null) window.document.form1.OLessonTeah1.value = TechName1;
            //if (window.document.form1.OLessonTeah1Value != null) window.document.form1.OLessonTeah1Value.value = Tech1;
            //if (window.document.form1.OLessonTeah2 != null) window.document.form1.OLessonTeah2.value = TechName2;
            //if (window.document.form1.OLessonTeah2Value != null) window.document.form1.OLessonTeah2Value.value = Tech2;
            //if (window.document.form1.Room != null) window.document.form1.Room.value = Room;
            //window.close();
        }

        function returnValue47(CourseName, CourseValue, Tech1, TechName1, Tech2, TechName2, Tech3, TechName3, Tech4, TechName4, Room) {
            var oCourseID = document.getElementById('CourseID');
            var oCourseIDValue = document.getElementById('CourseIDValue');
            var oOLessonTeah1 = document.getElementById('OLessonTeah1');
            var oOLessonTeah1Value = document.getElementById('OLessonTeah1Value');
            var oOLessonTeah2 = document.getElementById('OLessonTeah2');
            var oOLessonTeah2Value = document.getElementById('OLessonTeah2Value');
            var oOLessonTeah3 = document.getElementById('OLessonTeah3');
            var oOLessonTeah3Value = document.getElementById('OLessonTeah3Value');
            var oOLessonTeah4 = document.getElementById('OLessonTeah4');
            var oOLessonTeah4Value = document.getElementById('OLessonTeah4Value');
            var oRoom = document.getElementById('Room');
            if (oCourseID != null) oCourseID.value = CourseName;
            if (oCourseIDValue != null) oCourseIDValue.value = CourseValue;
            if (oOLessonTeah1 != null) oOLessonTeah1.value = TechName1;
            if (oOLessonTeah1Value != null) oOLessonTeah1Value.value = Tech1;
            if (oOLessonTeah2 != null) oOLessonTeah2.value = TechName2;
            if (oOLessonTeah2Value != null) oOLessonTeah2Value.value = Tech2;
            if (oOLessonTeah3 != null) oOLessonTeah3.value = TechName3;
            if (oOLessonTeah3Value != null) oOLessonTeah3Value.value = Tech3;
            if (oOLessonTeah4 != null) oOLessonTeah4.value = TechName3;
            if (oOLessonTeah4Value != null) oOLessonTeah4Value.value = Tech4;
            if (oRoom != null) oRoom.value = Room;
            //window.close();
        }

        function returnValue68(CourseName, CourseValue, Tech1, TechName1, Tech2, TechName2, Room, Classification1) {
            var cst_inline = ''; //'inline';
            var cst_none = 'none';
            var oCourseID = document.getElementById('CourseID');
            var oCourseIDValue = document.getElementById('CourseIDValue');
            var oOLessonTeah1 = document.getElementById('OLessonTeah1');
            var oOLessonTeah1Value = document.getElementById('OLessonTeah1Value');
            var oOLessonTeah2 = document.getElementById('OLessonTeah2');
            var oOLessonTeah2Value = document.getElementById('OLessonTeah2Value');
            //var oOLessonTeah3 = document.getElementById('OLessonTeah3');
            //var oOLessonTeah3Value = document.getElementById('OLessonTeah3Value');
            var oRoom = document.getElementById('Room');
            var olabTechN2 = document.getElementById('labTechN2');
            if (oCourseID != null) oCourseID.value = CourseName;
            if (oCourseIDValue != null) oCourseIDValue.value = CourseValue;
            if (oOLessonTeah1 != null) oOLessonTeah1.value = TechName1;
            if (oOLessonTeah1Value != null) oOLessonTeah1Value.value = Tech1;
            oOLessonTeah2.style.display = cst_inline;
            olabTechN2.style.display = cst_inline;
            if (Classification1 == '1') {
                //1個
                if (oOLessonTeah2 != null) oOLessonTeah2.value = '';
                if (oOLessonTeah2Value != null) oOLessonTeah2Value.value = '';
                oOLessonTeah2.style.display = cst_none;
                olabTechN2.style.display = cst_none;
            }
            if (Classification1 == '2') {
                //2個
                if (oOLessonTeah2 != null) oOLessonTeah2.value = TechName2;
                if (oOLessonTeah2Value != null) oOLessonTeah2Value.value = Tech2;
            }
            if (oRoom != null) oRoom.value = Room;
            //window.close();
        }

        function Course_search(Type) {
            //Cst_notepad : notepad
            var RIDValue = document.getElementById('RIDValue');
            var CourseValue = document.getElementById('CourseIDValue'); //CourseValue
            var CourseName = document.getElementById('CourseID'); //CourseName
            //'Edit', '" & CourseValue.ClientID & "', '" & CourseName.ClientID & "'
            wopen('SD_04_002_Course.aspx?Type=' + Type + '&RID=' + RIDValue.value + '&CourseValue=' + CourseValue.id + '&CourseName=' + CourseName.id, 'CheckID', 900, 850, 1);
        }

        //CourseID.Attributes("onchange")
        function GetCourseID(CourseID, TextField, ValueField, Tech1Field, TechName1Field, Tech2Field, TechName2Field, Tech3Field, TechName3Field, Tech4Field, TechName4Field, RoomField) {
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
            Doc.document.getElementById('Tech4Field').value = Tech4Field;
            Doc.document.getElementById('TechName4Field').value = TechName4Field;
            Doc.document.getElementById('RoomField').value = RoomField;
            Doc.document.getElementById('RID').value = RIDValue.value;
            Doc.document.getElementById('Button1').click();
        }

        function CheckNewCourse() {
            //TypeRadio
            var TypeRadio_0 = document.getElementById('TypeRadio_0');
            var TypeRadio_1 = document.getElementById('TypeRadio_1');
            var CourseIDValue = document.getElementById('CourseIDValue');
            var Room = document.getElementById('Room');
            var OLessonTeah1Value = document.getElementById('OLessonTeah1Value');
            var TPeriodValue = document.getElementById('TPeriodValue');
            var msg = '';
            var myvalue = getCheckBoxListValue('ClassSort1');
            //debugger;
            if (!TypeRadio_0) {
                alert('排課選項:異常(無此物件_0)!!');
                return false;
            }
            if (!TypeRadio_1) {
                alert('排課選項:異常(無此物件_1)!!');
                return false;
            }
            if (!TypeRadio_0.checked && !TypeRadio_1.checked) {
                msg += '請選擇排課選項\n';
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

        function LessonTeah3(opentype, st, fieldname, hiddenname) {
            var RIDValue = document.getElementById('RIDValue');
            var sUrl1 = ""; //var sUrl1 = "../../SD/04/";
            //alert("test!!");
            if (st == '1') {
                wopen(sUrl1 + 'LessonTeah1.aspx?RID=' + RIDValue.value + '&type=' + opentype + '&fieldname=' + fieldname + '&hiddenname=' + hiddenname, 'LessonTeah1', 900, 850, 1);
            }
            if (st == '2') {
                wopen(sUrl1 + 'LessonTeah2.aspx?RID=' + RIDValue.value + '&type=' + opentype + '&fieldname=' + fieldname + '&hiddenname=' + hiddenname, 'LessonTeah2', 900, 850, 1);
            }
            if (st == '3') {
                //hiddenname
                wopen(sUrl1 + 'LessonTeah2.aspx?RID=' + RIDValue.value + '&type=' + opentype + '&fieldname=' + fieldname + '&hiddenname=' + hiddenname, 'LessonTeah3', 900, 850, 1);
            }
            if (st == '4') {
                //hiddenname
                wopen(sUrl1 + 'LessonTeah2.aspx?RID=' + RIDValue.value + '&type=' + opentype + '&fieldname=' + fieldname + '&hiddenname=' + hiddenname, 'LessonTeah4', 900, 850, 1);
            }
        }

        /*
		function LessonTeah1(opentype, fieldname, hiddenname) {
		wopen('LessonTeah1.aspx?RID=' + document.form1.RIDValue.value + '&type=' + opentype + '&fieldname=' + fieldname + '&hiddenname=' + hiddenname, 'LessonTeah1', 400, 300, 1);
		}
		function LessonTeah2(opentype, fieldname, hiddenname) {
		wopen('LessonTeah2.aspx?RID=' + document.form1.RIDValue.value + '&type=' + opentype + '&fieldname=' + fieldname + '&hiddenname=' + hiddenname, 'LessonTeah2', 400, 300, 1);
		}
		*/

        function CheckClassTime() {
            var myvalue = getCheckBoxListValue('ClassSort1');
            var Result = true;
            for (var i = 0; i < 4; i++) {
                if (myvalue.charAt(i) == '0')
                    Result = false;
            }
            document.getElementById('ClassSort2').checked = Result;
            Result = true;
            for (var i = 4; i < 8; i++) {
                if (myvalue.charAt(i) == '0')
                    Result = false;
            }
            document.getElementById('ClassSort3').checked = Result;
            Result = true;
            for (var i = 0; i < 8; i++) {
                if (myvalue.charAt(i) == '0')
                    Result = false;
            }
            document.getElementById('ClassSort4').checked = Result;
            Result = true;
            for (var i = 8; i < 12; i++) {
                if (myvalue.charAt(i) == '0')
                    Result = false;
            }
            document.getElementById('ClassSort5').checked = Result;
        }

        function GetClassTime(num) {
            switch (num) {
                case 1:
                    for (var j = 0; j < 4; j++)
                        document.getElementById('ClassSort1_' + j).checked = document.getElementById('ClassSort2').checked;
                    if (!document.getElementById('ClassSort2').checked) document.getElementById('ClassSort4').checked = false;
                    break;
                case 2:
                    for (var j = 4; j < 8; j++)
                        document.getElementById('ClassSort1_' + j).checked = document.getElementById('ClassSort3').checked;
                    if (!document.getElementById('ClassSort3').checked) document.getElementById('ClassSort4').checked = false;
                    break;
                case 3:
                    for (var j = 0; j < 8; j++)
                        document.getElementById('ClassSort1_' + j).checked = document.getElementById('ClassSort4').checked;
                    document.getElementById('ClassSort2').checked = document.getElementById('ClassSort4').checked;
                    document.getElementById('ClassSort3').checked = document.getElementById('ClassSort4').checked;
                    break;
                case 4:
                    for (var j = 8; j < 12; j++)
                        document.getElementById('ClassSort1_' + j).checked = document.getElementById('ClassSort5').checked;
                    break;
            }
        }

        //顯示單一課程時數
        var TimerID1;
        var TimerID2;
        function ShowCourseList(obj) {
            //var cst_inline = 'inline';
            var cst_inline = '';
            var cst_none = 'none';
            var DataGrid1 = document.getElementById('DataGrid1');
            var DataGrid3 = document.getElementById('DataGrid3');
            var CourseTable = document.getElementById('CourseTable');
            var CourseList = document.getElementById('CourseList');
            var LinkButton2 = document.getElementById('LinkButton2');
            //alert('');
            //debugger;
            if (DataGrid1 && DataGrid3 && CourseTable.style.display == cst_inline) {
                if (CourseList.style.display == cst_none) {
                    CourseList.style.display = cst_inline;
                    LinkButton2.innerHTML = '關閉目前各課程的排課時數';

                    DataGrid1.style.filter = 'alpha(opacity=100)';
                    TimerID1 = setInterval("highlightit(30)", 50)
                }
                else {
                    CourseList.style.display = cst_none;
                    LinkButton2.innerHTML = '檢視目前各課程的排課時數';
                    DataGrid1.style.filter = 'alpha(opacity=30)';
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

        function highlightit(num) {
            var DataGrid1 = document.getElementById('DataGrid1');
            if (num != 100) {			//表示要透明化
                if (DataGrid1.filters.alpha.opacity > num)
                    DataGrid1.filters.alpha.opacity -= 15;
                else
                    clearInterval(TimerID1);
            }
            else {
                if (DataGrid1.filters.alpha.opacity < 100)
                    DataGrid1.filters.alpha.opacity += 15;
                else
                    clearInterval(TimerID2);
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;課程管理&gt;&gt;單月排課作業</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tbody>
                <tr>
                    <td>
                        <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <tr>
                                <td align="center">(此作業需先將課程代碼和師資資料先設定完成，在執行此作業)</td>
                            </tr>
                        </table>
                        <table class="table_nw" id="DetailTable1" width="100%" runat="server" cellpadding="1" cellspacing="1">
                            <tr>
                                <td class="bluecol" style="width: 20%">訓練機構 </td>
                                <td colspan="3" class="whitecol">
                                    <asp:TextBox ID="center" runat="server" Width="40%" onfocus="this.blur()"></asp:TextBox>
                                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">班級 </td>
                                <td colspan="3" class="whitecol">
                                    <asp:TextBox ID="OCID1" runat="server" Width="30%" onfocus="this.blur()"></asp:TextBox>
                                    <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">訓練時段 </td>
                                <td class="whitecol"><font>
                                    <asp:Label ID="TPeriod" runat="server" CssClass="font"></asp:Label>
                                    <input id="TPeriodValue" type="hidden" name="TPeriodValue" runat="server"></font>
                                </td>
                                <td class="bluecol">訓練期間 </td>
                                <td class="whitecol">
                                    <asp:Label ID="labTDate" runat="server" CssClass="font"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">訓練時數 </td>
                                <td class="whitecol">
                                    <asp:Label ID="THours" runat="server" CssClass="font"></asp:Label></td>
                                <td class="bluecol">已排時數<br>
                                    剩餘時數 </td>
                                <td class="whitecol">
                                    <asp:Label ID="UseHour" runat="server" CssClass="font"></asp:Label><br>
                                    <asp:Label ID="LeftHour" runat="server" CssClass="font"></asp:Label>
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
                                <td colspan="2" class="whitecol">
                                    <asp:LinkButton ID="LinkButton2" runat="server" Width="60%" ForeColor="Blue">檢視目前各課程的排課時數</asp:LinkButton>
                                    <div id="CourseList" style="display: none; width: 80%; position: absolute; background-color: white">
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
                                <td colspan="4" class="whitecol" width="100%" align="center">課程選擇清單
                                    <input id="STDate" type="hidden" name="STDate" runat="server">
                                    <input id="FTDate" type="hidden" name="FTDate" runat="server">
                                    <input id="Button1" type="button" value="編輯選擇清單" name="Button1" runat="server" class="asp_button_M">
                                    <asp:Button ID="Button2" runat="server" Text="重載選擇清單" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="4" class="whitecol" width="100%">
                                    <div style="overflow-y: auto; height: 410px;">
                                        <asp:TreeView ID="TreeView1" runat="server" CssClass="fontMenu" ForeColor="#333300"></asp:TreeView>
                                        <br />
                                        <center>
                                            <asp:Label ID="msg2" runat="server" ForeColor="Red" CssClass="font"></asp:Label></center>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">課程名稱 </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="CourseID" Style="cursor: pointer" runat="server" ToolTip="輸入課程代碼可以自動轉換成課程名稱,點選兩下可以跳出視窗選擇課程名稱" EnableViewState="False" Columns="45"></asp:TextBox>
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
                                    <asp:TextBox ID="OLessonTeah1" runat="server" ToolTip="點選兩下可以跳出視窗選擇教師" Columns="22" onfocus="this.blur()"></asp:TextBox>
                                    <input id="OLessonTeah1Value" type="hidden" runat="server">
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">
                                    <asp:Label ID="labTechN2" runat="server" Text="助教1"></asp:Label></td>
                                <td class="whitecol">
                                    <asp:TextBox ID="OLessonTeah2" runat="server" ToolTip="點選兩下可以跳出視窗選擇助教" Columns="22" onfocus="this.blur()"></asp:TextBox>
                                    <input id="OLessonTeah2Value" type="hidden" runat="server">
                                </td>
                                <td class="bluecol">
                                    <asp:Label ID="labTechN3" runat="server" Text="助教2"></asp:Label></td>
                                <td class="whitecol">
                                    <asp:TextBox ID="OLessonTeah3" runat="server" ToolTip="點選兩下可以跳出視窗選擇助教" Columns="22" onfocus="this.blur()"></asp:TextBox>
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
                                        <asp:ListItem Value="0" Selected="True">一般</asp:ListItem>
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
                                    <asp:CheckBox ID="ClassSort2" runat="server" CssClass="font" Text="第1-4節"></asp:CheckBox>
                                    <asp:CheckBox ID="ClassSort3" runat="server" CssClass="font" Text="第5-8節"></asp:CheckBox>
                                    <asp:CheckBox ID="ClassSort4" runat="server" CssClass="font" Text="第1-8節"></asp:CheckBox>
                                    <asp:CheckBox ID="ClassSort5" runat="server" CssClass="font" Text="第9-12節"></asp:CheckBox>
                                </td>
                                <td colspan="2" class="whitecol">
                                    <asp:Button ID="Button10" runat="server" Text="刪除排課-依節數" CssClass="asp_button_M"></asp:Button></td>
                            </tr>
                            <tr>
                                <td colspan="4" class="whitecol">
                                    <div align="center">
                                        <asp:Button ID="Button9" runat="server" Text="新增排課" CssClass="asp_button_M"></asp:Button></div>
                                </td>
                            </tr>
                            <tr>
                                <td align="center" colspan="4" class="whitecol">
                                    <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label></td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </tbody>
        </table>
        <table id="CourseTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" PageSize="16" AllowPaging="True" CellPadding="8">
                        <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                        <ItemStyle BackColor="#FFFFFF"></ItemStyle>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:TemplateColumn HeaderText="">
                                <HeaderStyle Width="4%"></HeaderStyle>
                                <ItemTemplate>
                                    <input id="SelectClass1" type="checkbox" runat="server" name="SelectClass1">
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="日期(星期)">
                                <HeaderStyle Width="12%"></HeaderStyle>
                                <ItemTemplate>
                                    <asp:LinkButton ID="LinkButton1" runat="server" ForeColor="Blue"></asp:LinkButton>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn DataField="Class1" HeaderText="節次1">
                                <HeaderStyle Width="7%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Class2" HeaderText="節次2">
                                <HeaderStyle Width="7%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Class3" HeaderText="節次3">
                                <HeaderStyle Width="7%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Class4" HeaderText="節次4">
                                <HeaderStyle Width="7%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Class5" HeaderText="節次5">
                                <HeaderStyle Width="7%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Class6" HeaderText="節次6">
                                <HeaderStyle Width="7%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Class7" HeaderText="節次7">
                                <HeaderStyle Width="7%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Class8" HeaderText="節次8">
                                <HeaderStyle Width="7%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Class9" HeaderText="節次9">
                                <HeaderStyle Width="7%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Class10" HeaderText="節次10">
                                <HeaderStyle Width="7%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Class11" HeaderText="節次11">
                                <HeaderStyle Width="7%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Class12" HeaderText="節次12">
                                <HeaderStyle Width="7%"></HeaderStyle>
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
            <tr>
                <td align="center" colspan="4" class="whitecol">
                    <asp:Button ID="Button4" runat="server" Text="回排課列表" CssClass="asp_button_M"></asp:Button></td>
            </tr>
        </table>
        <asp:HiddenField ID="hid_TNum" runat="server" />
    </form>
    <iframe id="Iframe2" style="display: none" src="FindCourse.aspx"></iframe>
</body>
</html>
