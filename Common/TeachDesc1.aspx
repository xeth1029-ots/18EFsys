<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TeachDesc1.aspx.vb" Inherits="WDAIIP.TeachDesc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href="../css/css.css" rel="stylesheet" type="text/css" />
    <%--<script language="javascript" src="../js/calendar/calendar.js"></script>--%>
    <script src="../js/OpenWin/openwin.js" type="text/javascript"></script>
    <script src="../js/common.js" type="text/javascript"></script>
    <script language="javascript" type="text/javascript">
        function setTDValue(Flag, MyValue) {
            var Hid_VALUE1 = document.getElementById('Hid_VALUE1');
            if (Flag) {
                if (Hid_VALUE1.value.indexOf(MyValue) == -1) {
                    if (Hid_VALUE1.value != '') { Hid_VALUE1.value += ","; }
                    Hid_VALUE1.value += MyValue;
                }
            }
            else {
                if (Hid_VALUE1.value.indexOf(MyValue) != -1) {
                    Hid_VALUE1.value = Hid_VALUE1.value.replace(',' + MyValue + ',', ',')
                    Hid_VALUE1.value = Hid_VALUE1.value.replace(',' + MyValue, '')
                    Hid_VALUE1.value = Hid_VALUE1.value.replace(MyValue + ',', '')
                    Hid_VALUE1.value = Hid_VALUE1.value.replace(MyValue, '')
                }
            }
            //alert(Hid_VALUE1.value);
        }

        //ReturnValue
        function ReturnValue() {
            var Hid_VALUE1 = document.getElementById('Hid_VALUE1');
            var txtSubject2 = document.getElementById('txtSubject2');
            //alert(Hid_VALUE1.value);
            //alert(txtSubject2.value);
            if (opener == undefined) { window.close(); return; }
            return true;
            var oDocu1 = opener.document;
            var oForm1 = oDocu1.form1;
            //oDocu1.getElementById(getParamValue('ValueField')).value = TechID;
            //oDocu1.getElementById(getParamValue('TextField')).value = TechName;
            //window.close();
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <table id="tbCJOB1" runat="server" class="font" border="0">
            <tr>
                <td><asp:Label ID="Label1" runat="server" Text="授課師資條件："></asp:Label></td>
                <td><asp:DropDownList ID="ddlTeachingCond1" runat="server" AutoPostBack="true"></asp:DropDownList></td>
            </tr>
            <tr>
                <td valign="top"><br /><asp:Label ID="Label2" runat="server" Text="助教條件："></asp:Label></td>
                <td>
                    <asp:CheckBoxList ID="CblTeachingCond2" runat="server" CssClass="font"></asp:CheckBoxList>
                    <asp:Label ID="LabSB3" runat="server" Text="其他(如為TTQS相關性課程，請填寫此欄位"></asp:Label>
                    <br /><asp:TextBox ID="txtSubject2" runat="server" Columns="30" Rows="8" TextMode="MultiLine"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td align="right" colspan="2">
                    <asp:Button ID="btnSend1" runat="server" Text="選擇" CssClass="asp_button_S" />
                    <%--<input id="btnSend1" onclick="javascript:ReturnValue();" type="button" value="選擇" runat="server" class="asp_button_S" />--%>
                </td>
            </tr>
        </table>
        <input id="Hid_VALUE1" type="hidden" runat="server" />
        <%--<input id="Hid_Text1" type="hidden" runat="server" />--%>
        <input id="Hid_OPTextBox1" type="hidden" runat="server" />
    </form>
</body>
</html>