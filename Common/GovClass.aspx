<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="GovClass.aspx.vb" Inherits="WDAIIP.GovClass" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>GovClass</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <%--<LINK href="../style.css" type="text/css" rel="stylesheet">--%>
    <link rel="stylesheet" type="text/css" href="../../css/style.css" />
    <script type="text/javascript" language="javascript">
        function ReturnValue() {
            if (document.form1.Radio1) {
                if (document.form1.Radio1.value == "") {
                    alert('請選擇層級!!!');
                    return false;
                }
            }
            if (document.form1.GCode1.value == "") {
                alert('請選擇類別!!!');
                return false;
            } else if (document.form1.GCode2.value == "") {
                alert('請選擇課程!!!');
                return false;
            }
            opener.document.form1.elements[document.form1.fieldname.value].value = form1.GCode2.options[form1.GCode2.selectedIndex].text;
            opener.document.form1.GCIDValue.value = document.form1.GCode2.value;
            opener.document.form1.GCID1Value.value = document.form1.GCode1.value;
        }

        function ClearValue() {
            //開放選擇訓練職類
            opener.document.form1.elements['btu_sel'].disabled = false;
            opener.document.form1.elements[document.form1.fieldname.value].value = '';
            opener.document.form1.GCIDValue.value = '';
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <input id="fieldname" type="hidden" name="fieldname" runat="server">
        <table id="errortable" runat="server" class="font" border="0" style="display: none" width="100%">
            <tr>
                <td colspan="2">
                    <input id="Button3" onclick="javascript: ClearValue();" type="button" value="清除" name="Button3" runat="server"></td>
            </tr>
            <tr>
                <td colspan="2">
                    <asp:Label ID="Label1" runat="server" ForeColor="Red">資料異常，請重新選擇訓練業別!!!</asp:Label></td>
            </tr>
        </table>
        <table id="normaltable" runat="server" class="font" border="0" width="100%">
            <tr id="trRadio1" runat="server">
                <td id="busTD" runat="server" class="bluecol" style="width: 20%">層級</td>
                <td class="whitecol" style="width: 80%">
                    <asp:RadioButtonList ID="Radio1" runat="server" AutoPostBack="True" RepeatLayout="Flow">
                        <asp:ListItem Value="1" Selected="True">院</asp:ListItem>
                        <asp:ListItem Value="2">局</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="bluecol" style="width: 20%">類別 </td>
                <td class="whitecol" style="width: 80%">
                    <asp:DropDownList ID="GCode1" runat="server" AutoPostBack="True"></asp:DropDownList></td>
            </tr>
            <tr>
                <td class="bluecol" style="width: 20%">課程 </td>
                <td class="whitecol" style="width: 80%">
                    <asp:DropDownList ID="GCode2" runat="server"></asp:DropDownList></td>
            </tr>
            <tr>
                <td colspan="2" align="center" class="whitecol">
                    <input id="Button1" onclick="javascript: ReturnValue();" type="button" value="選擇" name="but_sub" runat="server" class="button_b_M">
                    <input id="Button2" onclick="javascript: ClearValue();" type="button" value="清除" name="but_Cls" runat="server" class="button_b_M">
                </td>
            </tr>
        </table>
    </form>
</body>
</html>