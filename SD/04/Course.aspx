<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Course.aspx.vb" Inherits="WDAIIP.Course" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>主-副課程</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <%--<link href="../../css/css.css" rel="stylesheet" type="text/css" />--%>
    <link href="../../css/style.css" rel="stylesheet" type="text/css" />
    <script language="javascript" src="../../js/openwin/openwin.js" type="text/javascript"></script>
    <script language="javascript" src="../../js/common.js" type="text/javascript"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        //if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script language="javascript" type="text/javascript">
        function returnValue(ICourse, NCourse) {
            if (document.form1.modifytype.value == 'Add') {
                if (getParamValue('fieldname') == '') {
                    opener.document.form1.OCourseID.value = NCourse;
                    opener.document.form1.OCourseIDValue.value = ICourse;
                }
                else {
                    //當網址列上fieldname 和 hiddenname帶有參數時
                    //則指定fieldname 和 hiddenname上的參數值來回傳
                    //ex.fieldname=CourseID
                    // 則回傳母視窗中的TextBox ID為CourseID值
                    eval('opener.document.form1.' + getParamValue('fieldname')).value = NCourse;
                    eval('opener.document.form1.' + getParamValue('hiddenname')).value = ICourse;
                }
                window.close();
            } else if (document.form1.modifytype.value == 'Edit') {
                opener.document.form1[document.form1.fieldname.value].value = NCourse;
                opener.document.form1.OCourseIDValue.value = ICourse;
                window.close();
            }
        }
    </script>
    <style type="text/css">
        body { margin: 20px; overflow-y: auto; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <input type="hidden" runat="server" id="modifytype" name="modifytype" />
        <input type="hidden" runat="server" id="fieldname" name="fieldname" />
        <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td width="20%" class="bluecol">學/術科</td>
                <td width="30%" class="whitecol">
                    <asp:DropDownList ID="Classification1" runat="server">
                        <asp:ListItem Value="">請選擇</asp:ListItem>
                        <asp:ListItem Value="1">學科</asp:ListItem>
                        <asp:ListItem Value="2">術科</asp:ListItem>
                    </asp:DropDownList>
                </td>
                <td width="20%" class="bluecol">共同/一般/專業</td>
                <td width="30%" class="whitecol">
                    <asp:DropDownList ID="Classification2" runat="server">
                        <asp:ListItem Value="">請選擇</asp:ListItem>
                        <asp:ListItem Value="0">共同</asp:ListItem>
                        <asp:ListItem Value="1">一般</asp:ListItem>
                        <asp:ListItem Value="2">專業</asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td width="20%" class="bluecol">課程名稱
                </td>
                <td class="whitecol" colspan="3">
                    <asp:TextBox ID="CourseName" runat="server" Columns="30" Width="222px"></asp:TextBox>
                    <%--<font color="#000000"><asp:CheckBox ID="CB_Main" runat="server" Text="上層查詢" /></font>--%>
                </td>
            </tr>
            <tr>
                <td width="20%" class="bluecol">隸屬班級
                </td>
                <td class="whitecol" colspan="3">
                    <asp:TextBox ID="Classid" runat="server" Columns="30" Width="222px"></asp:TextBox>
                    <input type="hidden" id="Classid_Hid" runat="server" name="Classid_Hid" />
                    <input class="asp_button_M" id="Button2" onclick="javascript: wopen('../../TC/01/TC_01_005_Classid.aspx', '班級代碼', 450, 450, 1)" type="button" value="選擇" name="choice_button" runat="server" />
                    <input class="asp_button_M" type="button" value="清除" onclick="document.getElementById('Classid').value = ''; document.getElementById('Classid_Hid').value = '';" />
                </td>
            </tr>
            <tr>
                <td width="20%" class="bluecol">訓練職類
                </td>
                <td class="whitecol" colspan="3">
                    <asp:DropDownList ID="bus" AutoPostBack="True" runat="server">
                    </asp:DropDownList>
                    <asp:DropDownList ID="job" AutoPostBack="True" runat="server">
                    </asp:DropDownList>
                    <asp:DropDownList ID="train" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td colspan="4" align="center" class="whitecol">
                    <asp:Button ID="But_Sub" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                    <asp:Button ID="btnClear" runat="server" Text="清除" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
            <tr>
                <td colspan="4" align="left" class="whitecol">
                    <%-- <iewc:TreeView ID="TreeView1" runat="server"></iewc:TreeView>--%>
                    <asp:TreeView ID="TreeView1" runat="server" CssClass="fontMenu">
                    </asp:TreeView>
                    <br />
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
