<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_05_005.aspx.vb" Inherits="WDAIIP.SD_05_005" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>結訓成績登錄</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <style type="text/css">
        .FixedTitleRow { z-index: 10; position: relative; background-color: #e6ecf0; top: expression(this.offsetParent.scrollTop); }
        .FixedTitleColumn { position: relative; left: expression(this.parentElement.offsetParent.scrollLeft); }
        .FixedDataColumn { position: relative; left: expression(this.parentElement.offsetParent.parentElement.scrollLeft); }
        .DivWidth { position: static; width: 600px; display: inline; height: 350px; overflow: auto; cursor: default; }
        .DivHeight { position: static; overflow-x: hidden; overflow-y: scroll; display: inline; height: 350px; cursor: default; }
    </style>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        function GETvalue() {
            document.getElementById('Button11').click();
        }
        function SetOneOCID() {
            document.getElementById('Button12').click();
        }
        function choose_class() {
            //var RID = document.form1.RIDValue.value;
            var RIDValue = document.getElementById('RIDValue');
            var FunID = getParamValue('ID');
            var TMID1 = document.getElementById('TMID1');
            var OCID1 = document.getElementById('OCID1');
            var TMIDValue1 = document.getElementById('TMIDValue1');
            var OCIDValue1 = document.getElementById('OCIDValue1');
            var Button4 = document.getElementById('Button4');
            var Button6 = document.getElementById('Button6');
            var Button12 = document.getElementById('Button12');

            TMID1.value = '';
            OCID1.value = '';
            TMIDValue1.value = '';
            OCIDValue1.value = '';

            Button6.disabled = true;
            Button4.disabled = true;
            if (OCID1.value == '') { Button12.click(); }
            openClass('../02/SD_02_ch.aspx?BtnName=Button6&ID=' + FunID + '&RID=' + RIDValue.value);
        }

        function grade() {
            var Label1 = document.getElementById('Label1');
            var class_grade = document.getElementById('class_grade');
            var ChooseClass = document.getElementById('ChooseClass');
            var DBClass = document.getElementById('DBClass');
            var behavior = document.getElementById('behavior');
            var Button3 = document.getElementById('Button3');

            if (class_grade.checked) {
                Button3.disabled = false;
                ChooseClass.value = DBClass.value;
                if (DBClass.value != '')
                    Label1.innerHTML = '已選擇資料庫中的課程資料';
                else
                    Label1.innerHTML = '尚未選擇資料庫中的課程';
            }
            else if (behavior.checked) {
                ChooseClass.value = '';
                Label1.innerHTML = '';
                Button3.disabled = true;
            }
            else {
                Button3.disabled = true;
            }
        }

        function search() {
            var msg = '';
            var rbCPflag = false; //顯示是否取得結訓資格(defalut: false)
            var OCIDValue1 = document.getElementById('OCIDValue1');
            var rbCreditPoints = document.getElementById('rbCreditPoints');
            var behavior = document.getElementById('behavior');
            var class_grade = document.getElementById('class_grade');

            if (OCIDValue1.value == '') msg += '請選擇班級!\n';
            //debugger;
            if (rbCreditPoints) {
                if (rbCreditPoints.style.display != 'none') {
                    rbCPflag = true; //顯示是否取得結訓資格(true)
                }
            }

            if (rbCPflag) {
                if (behavior.checked == false && class_grade.checked == false && rbCreditPoints.checked == false) {
                    msg += '請選擇成績種類\n';
                }
            }
            else {
                //if(document.form1.class_grade.checked==true && document.form1.ChooseClass.value=='') msg+='請選擇主課程\n';
                if (behavior.checked == false && class_grade.checked == false) {
                    msg += '請選擇登入成績種類(課程或操行)\n';
                }
            }

            if (msg != '') {
                alert(msg);
                return false;
            }
            return true;
        }

        function getcourse() {
            var OCIDValue1 = document.getElementById('OCIDValue1');
            var Label1 = document.getElementById('Label1');
            var ChooseClass = document.getElementById('ChooseClass');

            if (OCIDValue1.value == '') {
                alert('請先選擇班級!!');
            }
            else {
                ChooseClass.value = '';
                Label1.innerHTML = '';
                window.open('SD_05_005c.aspx?OCID=' + OCIDValue1.value, '', 'width=700,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0')
            }
        }

        function chkdata(num) {
            //var cst_iDG2_COL_學號 = 0
            //var cst_iDG2_COL_姓名 = 1
            //var cst_iDG2_COL_出勤扣分 = 2
            //var cst_iDG2_COL_獎懲扣分 = 3
            //var cst_iDG2_COL_實際上課時數 = 4
            //var cst_iDG2_COL_上課比率 = 5
            var cst_iDG2_COL_導師加減分 = 6
            var cst_iDG2_COL_教務課加減分 = 7
            //var cst_iDG2_COL_操行成績 = 8
            //var cst_iDG2_COL_是否核發結訓證書 = 9

            var mytable = document.getElementById('DataGrid1');//dgInput
            var mytable2 = document.getElementById('DataGrid2');
            var msg = '';

            //debugger;
            if (num == 1) {
                //mytable = document.getElementById('DataGrid1');
                for (var i = 1; i < mytable.rows.length; i++) {
                    var mytext = mytable.rows[i].cells[2].children[0];
                    if (mytext.value == '') {
                        msg += '請輸入成績(第' + i + '列)\n';
                    }
                    else if (mytext.value != '' && mytext.value != '0') {
                        if (!isPositiveInt(mytext.value) && !isPositiveFloat(mytext.value)) msg += '成績必須輸入正數(第' + i + '列)\n'
                        else if (isInt(mytext.value) && (parseInt(mytext.value, 10) > 100 || parseInt(mytext.value, 10) < -100)) msg += '成績不能超過100分(第' + i + '列):' + mytext1.value + '\n';
                        else if (isFloat(mytext.value) && (parseFloat(mytext.value, 10) > 100 || parseFloat(mytext.value, 10) < -100)) msg += '成績最多100分(第' + i + '列):' + mytext1.value + '\n';
                    }
                }
            }
            else {
                //mytable = document.getElementById('DataGrid2');
                for (var i = 1; i < mytable2.rows.length; i++) {
                    var mytext1 = mytable2.rows[i].cells[cst_iDG2_COL_導師加減分].children[0];
                    var mytext2 = mytable2.rows[i].cells[cst_iDG2_COL_教務課加減分].children[0];
                    if (mytext1.value != '' && mytext1.value != undefined) {
                        if (!isFloat(mytext1.value) && !isInt(mytext1.value)) msg += '導師加減分必須輸入正數或負數(第' + i + '列):' + mytext1.value + '\n';
                        else if (isInt(mytext1.value) && (parseInt(mytext1.value, 10) > 100 || parseInt(mytext1.value, 10) < -100)) msg += '導師加減分至多100分(第' + i + '列):' + mytext1.value + '\n';
                        else if (isFloat(mytext1.value) && (parseFloat(mytext1.value, 10) > 100 || parseFloat(mytext1.value, 10) < -100)) msg += '導師加減分最多100分(第' + i + '列):' + mytext1.value + '\n';
                    }
                    if (mytext2.value != '' && mytext2.value != undefined) {
                        if (!isFloat(mytext2.value) && !isInt(mytext2.value)) msg += '教務課加減分必須輸入正數或負數(第' + i + '列):' + mytext2.value + '\n';
                        else if (isInt(mytext2.value) && (parseInt(mytext2.value, 10) > 100 || parseInt(mytext2.value, 10) < -100)) msg += '教務課加減分至多100分(第' + i + '列):' + mytext2.value + '\n';
                        else if (isFloat(mytext2.value) && (parseFloat(mytext2.value, 10) > 100 || parseFloat(mytext2.value, 10) < -100)) msg += '教務課加減分最多100分(第' + i + '列):' + mytext2.value + '\n';
                    }
                }
            }

            if (msg != '') {
                alert(msg);
                return false;
            }
        }
        function chall() {
            var mytable = document.getElementById('Table4')
            var mycheck1 = mytable.rows[0].cells[0].children[0];
            for (var i = 1; i < mytable.rows.length; i++) {
                var mycheck = mytable.rows[i].cells[0].children[0];
                mycheck.checked = mycheck1.checked
            }
            ChkSelSocid();    //勾選時更新學號 (表首第一列--全選 )
            ChkSelCourID();   //勾選時更新課程代碼
        }

        /**-----  090408 start----*/
        //取得 選擇學員學號  (table4--列)
        function ChkSelSocid() {
            var MyTable = document.getElementById('Table4');
            var SelSocid = document.getElementById('SelSocid');
            var SOCID = '';
            for (var i = 1; i < MyTable.rows.length; i++) {
                if (MyTable.rows[i].cells[0].children[0].checked) {
                    if (SOCID != '') { SOCID += ','; }
                    SOCID += MyTable.rows[i].cells[0].children[0].value;
                }
            }
            if (SOCID != '') {
                SelSocid.value = SOCID;
            }
        }

        //取得選擇課程代碼  (table4--欄)
        function ChkSelCourID() {
            var MyTable = document.getElementById('Table4');
            var SelCourID = document.getElementById('SelCourID');
            var CourID = '';

            for (var i = 3; i < MyTable.rows[0].cells.length; i++) {
                //預設抓畫面上有顯示的科目
                if (CourID != '') { CourID += ','; }
                CourID += "'" + MyTable.rows[0].cells[i].children[0].value + "'";
            }
            if (CourID != '') {
                SelCourID.value = CourID;
            }
        }

        /**-----  090408  end----*/
        function CheckPrint(BasicPoint, PrintDataTyp) {
            //var RID = document.form1.RIDValue.value;
            var RIDValue = document.getElementById('RIDValue');
            var MyTable = document.getElementById('Table4');
            var CCIDVALUE = document.getElementById('CCIDVALUE');
            var CCID = document.getElementById('CCID');
            /*報表以網頁方式呈現*/
            var OCIDValue1 = document.getElementById('OCIDValue1');
            var PercentSet = document.getElementById('PercentSet');
            var Print1 = document.getElementById('Print1');
            var SOCID = '';
            for (var i = 1; i < MyTable.rows.length; i++) {
                if (MyTable.rows[i].cells[0].children[0].checked) {
                    if (SOCID != '') { SOCID += ','; }
                    SOCID += MyTable.rows[i].cells[0].children[0].value;
                }
            }
            if (SOCID == '') {
                alert('請選擇學員');
                return false;
            }

            var checkedItems = getCheckBoxListItemsChecked();
            if (CCIDVALUE.value == '' && CCID.value == '') {
                alert('至少須選擇一個科目！！');
                return false;
            }

            //有設百分比
            if (PercentSet.value == 'Y') {
                //科目有選取時
                if (CCIDVALUE.value != '') {
                    window.open('SD_05_005_R.aspx?BasicPoint=' + BasicPoint + checkedItems + '&SOCID=' + SOCID + '&OCID=' + OCIDValue1.value + '&RID=' + RIDValue.value + '&ChooseClass=' + CCIDVALUE.value + '&SignType=' + Print1.value + '&PrintDataTyp=' + PrintDataTyp);
                }
                else {
                    window.open('SD_05_005_R.aspx?BasicPoint=' + BasicPoint + checkedItems + '&SOCID=' + SOCID + '&OCID=' + OCIDValue1.value + '&RID=' + RIDValue.value + '&ChooseClass=' + CCID.value + '&SignType=' + Print1.value + '&PrintDataTyp=' + PrintDataTyp);
                }
            }
            else {
                //無百分比設定
                if (CCIDVALUE.value != '') {
                    window.open('SD_05_005_R.aspx?BasicPoint=' + BasicPoint + checkedItems + '&SOCID=' + SOCID + '&OCID=' + OCIDValue1.value + '&RID=' + RIDValue.value + '&ChooseClass=' + CCIDVALUE.value + '&SignType=' + Print1.value + '&PrintDataTyp=' + PrintDataTyp);
                }
                else {
                    window.open('SD_05_005_R.aspx?BasicPoint=' + BasicPoint + checkedItems + '&SOCID=' + SOCID + '&OCID=' + OCIDValue1.value + '&RID=' + RIDValue.value + '&ChooseClass=' + CCID.value + '&SignType=' + Print1.value + '&PrintDataTyp=' + PrintDataTyp);
                }
            }
            return true;
        }
        function getCheckBoxListItemsChecked() {
            var checkedValues = '';
            var elementRef = document.getElementById('cb_signer');
            var checkBoxArray = elementRef.getElementsByTagName('input');
            for (var i = 0; i < checkBoxArray.length; i++) {
                var checkBoxRef = checkBoxArray[i];
                if (checkBoxRef.checked == true) {
                    if (checkedValues.length > 0) checkedValues += '';
                    checkedValues += '&chk' + String(i + 1) + '=1';
                }
                else {
                    if (checkedValues.length > 0) checkedValues += '';
                    checkedValues += '&chk' + String(i + 1) + '=0';
                }
            }
            var elementRef2 = document.getElementById('cb_signer2');
            var checkBoxArray2 = elementRef2.getElementsByTagName('input');
            for (var i = 0; i < checkBoxArray2.length; i++) {
                var checkBoxRef2 = checkBoxArray2[i];
                if (checkBoxRef2.checked == true) {
                    if (checkedValues.length > 0) checkedValues += '';
                    checkedValues += '&chk' + String(i + 10) + '=1';
                }
                else {
                    if (checkedValues.length > 0) checkedValues += '';
                    checkedValues += '&chk' + String(i + 10) + '=0';
                }
            }
            return checkedValues;
        }

        function readCheckBoxList() {
            var checkedItems = getCheckBoxListItemsChecked();
        }
        function checkAll() {
            var cb_signer = document.getElementById("cb_signer");
            for (var i = 0; i < cb_signer.getElementsByTagName("input").length; i++) { document.getElementById("cb_signer_" + i).checked = true; }
        }
        function deleteAll() {
            var cb_signer = document.getElementById("cb_signer");
            for (var i = 0; i < cb_signer.getElementsByTagName("input").length; i++) { document.getElementById("cb_signer_" + i).checked = false; }
        }
        function ReverseAll() {
            var cb_signer = document.getElementById("cb_signer");
            for (var i = 0; i < cb_signer.getElementsByTagName("input").length; i++) {
                var objCheck = document.getElementById("cb_signer_" + i);
                if (objCheck.checked)
                    objCheck.checked = false;
                else
                    objCheck.checked = true;
            }
        }
        function ChangeAll(j) {
            var cells_num2 = 9;
            var MyTable = document.getElementById('DataGrid2');
            if (!MyTable) return;
            for (i = 1; i < MyTable.rows.length; i++) {
                if (MyTable.rows[i].cells[cells_num2].children[0].selectedIndex == 0 && !MyTable.rows[i].cells[cells_num2].children[0].disabled) {
                    MyTable.rows[i].cells[cells_num2].children[0].selectedIndex = j;
                }
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;結訓成績登錄</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <%--<table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0"><tr><td>
                       <asp:Label ID="TitleLab1" runat="server"></asp:Label><asp:Label ID="TitleLab2" runat="server">
                       首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">結訓成績登錄</font></asp:Label></td></tr></table>--%>
                    <table class="table_sch" id="Table3" cellpadding="1" cellspacing="1">
                        <tr>
                            <td class="bluecol" style="width: 20%">訓練機構 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" runat="server">
                                <input id="Button10" type="button" value="..." name="Button10" runat="server" class="button_b_Mini">
                                <asp:Button ID="Button12" Style="display: none" runat="server"></asp:Button>
                                <asp:Button ID="Button11" Style="display: none" runat="server" Text="Button11"></asp:Button>
                                <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">職類/班別 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="button_b_Mini">
                                <asp:Button ID="Button6" runat="server" Text="查詢班級資料(隱藏)" CssClass="asp_button_M"></asp:Button>
                                <span id="HistoryList" style="position: absolute; display: none; left: 270px">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                <asp:Label ID="Label2" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">成績總類 </td>
                            <td class="whitecol">
                                <table class="font" id="Table5" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td>
                                            <asp:RadioButton ID="class_grade" runat="server" Text="課程" GroupName="gr"></asp:RadioButton>
                                            <input id="Button3" disabled onclick="getcourse();" type="button" value="搜尋主課程" runat="server" class="asp_button_M">
                                            <asp:Label ID="Label1" runat="server" CssClass="font"></asp:Label>
                                            <input id="ChooseClass" type="hidden" name="ChooseClass" runat="server">
                                            <input id="DBClass" type="hidden" name="DBClass" runat="server">
                                            <input id="SelClassCount" type="hidden" name="ChooseClass" runat="server">
                                            <input id="SelSocid" type="hidden" name="SelSocid" runat="server">
                                            <input id="SelCourID" type="hidden" name="SelCourID" runat="server">
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:RadioButton ID="behavior" runat="server" Text="操行" GroupName="gr"></asp:RadioButton>&nbsp;
										    <asp:RadioButton ID="rbCreditPoints" runat="server" Text="是否取得結訓資格" GroupName="gr"></asp:RadioButton>
                                            <asp:LinkButton ID="LinkButton1" runat="server"></asp:LinkButton>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td align="center" class="whitecol">
                                <p>
                                    <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>&nbsp;
                                <asp:Button ID="add_Result" runat="server" Text="登錄成績" Visible="False" CssClass="asp_button_M"></asp:Button>&nbsp;
                                <asp:Button ID="Button4" runat="server" Text="計算總成績" CssClass="asp_button_M"></asp:Button>
                                    <asp:Label ID="Msg2" runat="server" ForeColor="Red"></asp:Label>
                                </p>
                            </td>
                        </tr>
                    </table>
                    <table id="tb_EditResult" width="100%" runat="server">
                        <tr>
                            <td class="whitecol" align="center">學員
							<asp:DropDownList ID="dl_studName" runat="server" AutoPostBack="True">
                            </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td valign="top" align="center">
                                <asp:DataGrid ID="dgInput" Width="100%" CssClass="font" runat="server" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:BoundColumn DataField="CourseName" HeaderText="科目" ItemStyle-Width="200px">
                                            <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="成績">
                                            <HeaderStyle HorizontalAlign="Center" Width="80%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" CssClass="whitecol"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:TextBox ID="RESULTS" Width="80%" runat="server"></asp:TextBox>
                                                <input id="CourIDVal" type="hidden" runat="server" />
                                                <input id="SocidVal" type="hidden" runat="server">
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="bt_save" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button></td>
                        </tr>
                    </table>
                    <table class="font" id="GradeTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td align="center"></td>
                        </tr>
                        <tr>
                            <td align="center">
                                <div id="scrollDiv" runat="server" align="left">
                                    <asp:Table ID="Table4" runat="server" CssClass="font" BorderWidth="1px" CellPadding="1" CellSpacing="0">
                                    </asp:Table>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <asp:Label Style="z-index: 0" ID="Label3" runat="server" CssClass="font"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="Button7" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>&nbsp;&nbsp;&nbsp;&nbsp;
							<asp:Button ID="Button8" runat="server" Text="消除勾選科目" CssClass="asp_button_M"></asp:Button>
                                <input id="CCID" type="hidden" name="DBClass" runat="server">
                                <input id="CCIDVALUE" type="hidden" name="CCIDchked" runat="server">
                                <input id="PercentSet" type="hidden" name="PercentSet" runat="server">
                                <input id="Print1" type="hidden" name="Print1" runat="server">
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="PrintTB" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr id="TR_SingType" align="center" runat="server">
                            <td class="whitecol">
                                <asp:RadioButtonList ID="RB_SignType" AutoPostBack="True" runat="server" RepeatDirection="Horizontal" CssClass="font11">
                                    <asp:ListItem Value="1" Selected="True">簽核方式1</asp:ListItem>
                                    <asp:ListItem Value="2">簽核方式2</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="TR_Signer2" align="center" runat="server">
                            <td class="whitecol">
                                <asp:CheckBoxList ID="cb_signer2" runat="server" RepeatDirection="Horizontal" CssClass="font11">
                                    <asp:ListItem Value="1">承辦</asp:ListItem>
                                    <asp:ListItem Value="2">主管</asp:ListItem>
                                    <asp:ListItem Value="3">批示</asp:ListItem>
                                </asp:CheckBoxList>
                            </td>
                        </tr>
                        <tr id="TR_Signer" align="center" runat="server">
                            <td class="whitecol">
                                <asp:CheckBoxList ID="cb_signer" runat="server" CssClass="font11" RepeatDirection="Horizontal" RepeatColumns="3">
                                    <asp:ListItem Value="1">導師</asp:ListItem>
                                    <asp:ListItem Value="2">教學行政股長</asp:ListItem>
                                    <asp:ListItem Value="3">批示</asp:ListItem>
                                    <asp:ListItem Value="4">股長</asp:ListItem>
                                    <asp:ListItem Value="5">科長</asp:ListItem>
                                </asp:CheckBoxList>
                            </td>
                        </tr>
                        <%--<tr id="TR_Signer" align="center" runat="server"><td class="whitecol" >
                          <asp:CheckBoxList ID="cb_signer" runat="server" Width="448px" CssClass="font11" RepeatDirection="Horizontal" RepeatColumns="3" Height="20px">
                          <asp:ListItem Value="1">主訓老師</asp:ListItem><asp:ListItem Value="2">教學股長</asp:ListItem><asp:ListItem Value="3">秘書</asp:ListItem>
                          <asp:ListItem Value="4">導師</asp:ListItem><asp:ListItem Value="5">教務課長</asp:ListItem><asp:ListItem Value="6">簡任技正</asp:ListItem>
                          <asp:ListItem Value="7">股長</asp:ListItem><asp:ListItem Value="8">輔導課長</asp:ListItem>
                          <asp:ListItem Value="9">主任</asp:ListItem></asp:CheckBoxList></td></tr>--%>
                        <tr id="Bn_Print" runat="server">
                            <td align="center" class="whitecol">
                                <asp:Button ID="Button9" runat="server" Text="列印空白成績表" CssClass="asp_Export_M"></asp:Button>
                                &nbsp;&nbsp;
							<asp:Button ID="Button5" runat="server" Text="列印成績表" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <table id="BehaviorTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td colspan="3">
                                <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" Visible="False" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn DataField="StudentID" HeaderText="學號">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Name" HeaderText="姓名">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="MinusPoint" HeaderText="出勤扣分">
                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="total" HeaderText="獎懲扣分">
                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="實際上課時數">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" CssClass="whitecol"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="labACTUALHOURS" runat="server" Text=""></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="上課比率">
                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" CssClass="whitecol"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="labTRAINRATIO" runat="server" Text=""></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>

                                        <asp:TemplateColumn HeaderText="導師加減分">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" CssClass="whitecol"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:TextBox ID="TextBox2" runat="server" Width="66%"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="教務課加減分">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" CssClass="whitecol"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:TextBox ID="TextBox3" runat="server" Width="66%"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn HeaderText="操行成績">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="12%" />
                                            <ItemStyle HorizontalAlign="Center" CssClass="whitecol2" />
                                            <HeaderTemplate>
                                                是否核發結訓證書<br />
                                                <asp:DropDownList ID="SelectAll" runat="server" CssClass="whitecol2">
                                                    <asp:ListItem Value="請選擇">請選擇</asp:ListItem>
                                                    <asp:ListItem Value="是">是</asp:ListItem>
                                                    <asp:ListItem Value="否">否</asp:ListItem>
                                                </asp:DropDownList>
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:DropDownList ID="CreditPoints" runat="server" CssClass="whitecol2">
                                                    <asp:ListItem Value="請選擇">請選擇</asp:ListItem>
                                                    <asp:ListItem Value="是">是</asp:ListItem>
                                                    <asp:ListItem Value="否">否</asp:ListItem>
                                                </asp:DropDownList>
                                                <input id="hidDataKeys" type="hidden" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <%-- <asp:BoundColumn Visible="False" DataField="TechPoint" HeaderText="TechPoint"></asp:BoundColumn>
                                        <asp:BoundColumn Visible="False" DataField="RemedPoint" HeaderText="RemedPoint"></asp:BoundColumn>
                                        <asp:BoundColumn Visible="False" DataField="StudStatus" HeaderText="StudStatus"></asp:BoundColumn>
                                        <asp:BoundColumn Visible="False" DataField="OCID" HeaderText="OCID"></asp:BoundColumn>--%>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <asp:Label Style="z-index: 0" ID="Label4" runat="server" CssClass="font"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="Button2" runat="server" Text="計算儲存" Visible="False" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <p style="margin-top: 3px; margin-bottom: 3px" align="center">
                        <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                    </p>
                </td>
            </tr>
        </table>
        <%--Button4.計算總成績--%>
        <%--bt_save.儲存--%>
        <asp:HiddenField ID="Hid_OCID1" runat="server" />
        <asp:HiddenField ID="Hid_THOURS" runat="server" />
        <%--<asp:HiddenField ID="Hid_eMP" runat="server" />--%>
        <asp:HiddenField ID="Hid_lockBtn2" runat="server" />
        <asp:HiddenField ID="Hid_lockBtn2_MSG" runat="server" />
        <asp:HiddenField ID="Hid_lockBtn4" runat="server" />
        <asp:HiddenField ID="Hid_lockBtn4_MSG" runat="server" />

        <asp:HiddenField ID="hidSD05005CHK551" runat="server" />
        <asp:HiddenField ID="hidSD05005CHK572" runat="server" />
        <asp:HiddenField ID="hidSD05005CHK598" runat="server" />
        <asp:HiddenField ID="hidSD05005CHK618" runat="server" />
        <asp:HiddenField ID="hidSD05005CHK653" runat="server" />
    </form>
</body>
</html>
