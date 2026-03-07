<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="FindCourse.aspx.vb" Inherits="WDAIIP.FindCourse" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>FindCourse</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        //if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        function ReturnCourse(CourID, CourseName, Tech1, TechName1, Tech2, TechName2, Tech3, TechName3, Room) {
            eval('parent.document.forms[0].' + document.getElementById('TextField').value).value = CourseName;
            eval('parent.document.forms[0].' + document.getElementById('ValueField').value).value = CourID;
            eval('parent.document.forms[0].' + document.getElementById('Tech1Field').value).value = Tech1;
            eval('parent.document.forms[0].' + document.getElementById('TechName1Field').value).value = TechName1;
            eval('parent.document.forms[0].' + document.getElementById('Tech2Field').value).value = Tech2;
            eval('parent.document.forms[0].' + document.getElementById('TechName2Field').value).value = TechName2;
            var can_t3f_1 = false;
            var can_tn3f_1 = false;
            var Tech3Field = document.getElementById('Tech3Field');
            var TechName3Field = document.getElementById('TechName3Field');
            if (Tech3Field) { if (Tech3Field.value != "") { can_t3f_1 = true; } }
            if (TechName3Field) { if (TechName3Field.value != "") { can_tn3f_1 = true; } }
            if (can_t3f_1 && can_tn3f_1) {
                var objT3 = eval('parent.document.forms[0].' + Tech3Field.value);
                var objTN3 = eval('parent.document.forms[0].' + TechName3Field.value);
                if (objT3 && objTN3) {
                    eval('parent.document.forms[0].' + Tech3Field.value).value = Tech3;
                    eval('parent.document.forms[0].' + TechName3Field.value).value = TechName3;
                }
            }
            eval('parent.document.forms[0].' + document.getElementById('RoomField').value).value = Room;

            document.getElementById('CourseID').value = '';
            document.getElementById('TextField').value = '';
            document.getElementById('ValueField').value = '';
            document.getElementById('Tech1Field').value = '';
            document.getElementById('TechName1Field').value = '';
            document.getElementById('Tech2Field').value = '';
            document.getElementById('TechName2Field').value = '';
            if (Tech3Field) { Tech3Field.value = ''; }
            if (TechName3Field) { TechName3Field.value = ''; }
            document.getElementById('RoomField').value = '';
            document.getElementById('RID').value = '';
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <asp:Button ID="Button1" runat="server" Text="Search"></asp:Button>
        <asp:TextBox ID="CourseID" runat="server" Columns="1"></asp:TextBox>
        <asp:TextBox ID="TextField" runat="server" Columns="1"></asp:TextBox>
        <asp:TextBox ID="ValueField" runat="server" Columns="1"></asp:TextBox>
        <asp:TextBox ID="RID" runat="server" Columns="1"></asp:TextBox>
        <asp:TextBox ID="Tech1Field" runat="server" Columns="1"></asp:TextBox>
        <asp:TextBox ID="TechName1Field" runat="server" Columns="1"></asp:TextBox>
        <asp:TextBox ID="Tech2Field" runat="server" Columns="1"></asp:TextBox>
        <asp:TextBox ID="TechName2Field" runat="server" Columns="1"></asp:TextBox>
        <asp:TextBox ID="Tech3Field" runat="server" Columns="1"></asp:TextBox>
        <asp:TextBox ID="TechName3Field" runat="server" Columns="1"></asp:TextBox>
        <asp:TextBox ID="RoomField" runat="server" Columns="1"></asp:TextBox>
    </form>
</body>
</html>
