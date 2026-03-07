<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_04_002_Course.aspx.vb" Inherits="WDAIIP.SD_04_002_Course" %>

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>SD_04_002_Course</title>
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
        /*
		網址上的參數
		TextField (要回傳的課程名稱欄位)
		HiddenField (要回傳的課程流水號欄位)
		*/
        function returnValue(CourseName, CourseValue, Tech1, TechName1, Tech2, TechName2, Tech3, TechName3, Room) {
            //document.form1.Type.value==1 'Edit','DataGrid中的課程欄位'
            if (document.form1.hid_Type1.value == '1') {
                opener.document.getElementById(getParamValue('CourseValue')).value = CourseValue;
                opener.document.getElementById(getParamValue('CourseName')).value = CourseName;
            }
            else {
                if (window.opener.document.form1.CourseID != null) window.opener.document.form1.CourseID.value = CourseName;
                if (window.opener.document.form1.CourseIDValue != null) window.opener.document.form1.CourseIDValue.value = CourseValue;
                if (window.opener.document.form1.OLessonTeah1 != null) window.opener.document.form1.OLessonTeah1.value = TechName1;
                if (window.opener.document.form1.OLessonTeah1Value != null) window.opener.document.form1.OLessonTeah1Value.value = Tech1;
                if (window.opener.document.form1.OLessonTeah2 != null) window.opener.document.form1.OLessonTeah2.value = TechName2;
                if (window.opener.document.form1.OLessonTeah2Value != null) window.opener.document.form1.OLessonTeah2Value.value = Tech2;
                if (window.opener.document.form1.OLessonTeah3 != null) window.opener.document.form1.OLessonTeah3.value = TechName3;
                if (window.opener.document.form1.OLessonTeah3Value != null) window.opener.document.form1.OLessonTeah3Value.value = Tech3;
                if (window.opener.document.form1.Room != null) window.opener.document.form1.Room.value = Room;
            }
            window.close();

            //if (getParamValue('TextField')!=''){
            //eval('parent.document.all.'+getParamValue('TextField')).value=CourseName;
            //eval('parent.document.all.'+getParamValue('HiddenField')).value=CourseValue;
            //eval('parent.document.all.'+getParamValue('Tech1Field')).value=Tech1;
            //eval('parent.document.all.'+getParamValue('TechName1Field')).value=TechName1;
            //eval('parent.document.all.'+getParamValue('Tech2Field')).value=Tech2;
            //eval('parent.document.all.'+getParamValue('TechName2Field')).value=TechName2;
            //eval('parent.document.all.'+getParamValue('RoomField')).value=Room;
            //	}
            //	else{
            //	parent.document.all.CourseID.value=CourseName;
            //	parent.document.all.CourseIDValue.value=CourseValue;
            //	parent.document.all.OLessonTeah1.value=TechName1;
            //	parent.document.all.OLessonTeah1Value.value=Tech1;
            //	parent.document.all.OLessonTeah2.value=TechName2;
            //	parent.document.all.OLessonTeah2Value.value=Tech2;
            //	parent.document.all.Room.value=Room;
            //	}
        }

        function returnValue68(CourseName, CourseValue, Tech1, TechName1, Tech2, TechName2, Room, Classification1) {
            //document.form1.Type.value==1 'Edit','DataGrid中的課程欄位'
            var opCourseValue = opener.document.getElementById(getParamValue('CourseValue'));
            var opCourseName = opener.document.getElementById(getParamValue('CourseName'));
            var opCourseID = window.opener.document.form1.CourseID;
            var opCourseIDValue = window.opener.document.form1.CourseIDValue;
            var opOLessonTeah1 = window.opener.document.form1.OLessonTeah1;
            var opOLessonTeah1Value = window.opener.document.form1.OLessonTeah1Value;
            var opOLessonTeah2 = window.opener.document.form1.OLessonTeah2;
            var opOLessonTeah2Value = window.opener.document.form1.OLessonTeah2Value;
            var opRoom = window.opener.document.form1.Room;
            var oplabTechN2 = opener.document.getElementById('labTechN2');
            var cst_inline = ''; //'inline';
            var cst_none = 'none';
            /*
			alert('document.form1.Type.value' + document.form1.Type.value);
			alert('Classification1'+Classification1);
			if (document.form1.Type.value == '1') {
			opCourseValue.value = CourseValue;
			opCourseName.value = CourseName;
			}
			*/
            if (opCourseID != null) opCourseID.value = CourseName;
            if (opCourseIDValue != null) opCourseIDValue.value = CourseValue;
            if (opOLessonTeah1 != null) opOLessonTeah1.value = TechName1;
            if (opOLessonTeah1Value != null) opOLessonTeah1Value.value = Tech1;
            //if (window.opener.document.form1.OLessonTeah3 != null) window.opener.document.form1.OLessonTeah3.value = TechName3;
            //if (window.opener.document.form1.OLessonTeah3Value != null) window.opener.document.form1.OLessonTeah3Value.value = Tech3;
            opOLessonTeah2.style.display = cst_inline;
            oplabTechN2.style.display = cst_inline;
            //CLASSIFICATION1 1:學科/2術科
            if (Classification1 == '1') {
                //1個
                if (opOLessonTeah2 != null) opOLessonTeah2.value = '';
                if (opOLessonTeah2Value != null) opOLessonTeah2Value.value = '';
                opOLessonTeah2.style.display = cst_none;
                oplabTechN2.style.display = cst_none;
            }
            if (Classification1 == '2') {
                //2個
                if (opOLessonTeah2 != null) opOLessonTeah2.value = TechName2;
                if (opOLessonTeah2Value != null) opOLessonTeah2Value.value = Tech2;
            }
            if (opRoom != null) opRoom.value = Room;
            window.close();
        }

        function GetCLSID() { wopen('../../TC/01/TC_01_005_Classid.aspx', 'CLSID', 800, 800, 1); }

        function ClearClassid() {
            var Classid_Hid = document.getElementById('Classid_Hid');
            var Classid = document.getElementById('Classid');
            Classid_Hid.value = '';
            Classid.value = '';
            //document.all.Classid_Hid.Value='';
            //document.all.Classid.Value='';
        }
    </script>
</head>
<body leftmargin="0" topmargin="0">
    <form id="form1" method="post" runat="server">
        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td width="100%">
                    <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="bluecol" width="20%"><font>課程代碼</font> </td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="CourseID" runat="server" MaxLength="50" Width="50%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%"><font>課程名稱</font> </td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="CourseName" runat="server" MaxLength="100" Width="50%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%"><font>學/術科</font> </td>
                            <td class="whitecol" width="80%">
                                <asp:DropDownList ID="Classification1" runat="server">
                                    <asp:ListItem Value="">===請選擇===</asp:ListItem>
                                    <asp:ListItem Value="1">學科</asp:ListItem>
                                    <asp:ListItem Value="2">術科</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%"><font>共同/一般/專業</font> </td>
                            <td class="whitecol" width="80%">
                                <asp:DropDownList ID="Classification2" runat="server">
                                    <asp:ListItem Value="">===請選擇===</asp:ListItem>
                                    <asp:ListItem Value="0">共同</asp:ListItem>
                                    <asp:ListItem Value="1">一般</asp:ListItem>
                                    <asp:ListItem Value="2">專業</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%"><font>隸屬班級</font> </td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="Classid" runat="server" Width="45%"></asp:TextBox>
                                <input id="Button1" type="button" value="選擇" name="Button1" runat="server" class="asp_button_S">
                                <input id="Classid_Hid" type="hidden" runat="server">
                                <input id="Button3" type="button" value="清除" name="Button3" runat="server" class="asp_button_S">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%"><font>訓練職類</font> </td>
                            <td class="whitecol" width="80%">
                                <asp:DropDownList ID="BusID" runat="server" AutoPostBack="True"></asp:DropDownList>
                                <br>
                                <asp:DropDownList ID="JobID" runat="server" AutoPostBack="True"></asp:DropDownList>
                                <br>
                                <asp:DropDownList ID="TrainID" runat="server"></asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="2" width="100%">
                                <input id="hid_Type1" type="hidden" runat="server">
                                <%--	
							<input id="hid_fieldname" type="hidden" name="fieldname" runat="server">
							<input id="TextField" type="hidden" name="TextField" runat="server">
							<input id="HiddenField" type="hidden" name="HiddenField" runat="server">
                                --%> &nbsp;
							<asp:Button ID="Button2" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>&nbsp;
							<asp:Button ID="btnSaveCheckBox" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                        <tr id="trTreeView1" runat="server">
                            <td colspan="2" width="100%" class="whitecol">
                                <div style="overflow-y: auto; height: 410px;">
                                    <%-- <iewc:treeview id="TreeView1" runat="server"></iewc:treeview> --%>
                                    <asp:TreeView ID="TreeView1" runat="server" CssClass="fontMenu"></asp:TreeView>
                                    <br />
                                    <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                                </div>
                            </td>
                        </tr>
                        <tr id="trCheckBoxList1" runat="server">
                            <td colspan="2" width="100%" class="whitecol">
                                <div style="overflow-y: auto; height: 410px;">
                                    <asp:CheckBoxList ID="CheckBoxList1" runat="server" CssClass="font" RepeatColumns="2" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                    <br />
                                    <asp:Label ID="msg2" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                                </div>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="HidrqType" runat="server" />
        <asp:HiddenField ID="HidReqRID" runat="server" />
    </form>
</body>
</html>
