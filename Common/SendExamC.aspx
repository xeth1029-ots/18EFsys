<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SendExamC.aspx.vb" Inherits="WDAIIP.SendExamC" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>請選擇 檢定職類與考試級別</title>
    <meta content="microsoft visual studio .net 7.1" name="generator" />
    <meta content="visual basic .net 7.1" name="code_language" />
    <meta content="javascript" name="vs_defaultclientscript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetschema" />
    <script type="text/javascript" src="../Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="../Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script language="javascript" type="text/javascript">
        function ReturnValue2() {
            var ddlEXAMGROUP = document.form1.ddlEXAMGROUP;
            var ddlEXAM3 = document.form1.ddlEXAM3;
            var ddlEXLEVEL = document.form1.ddlEXLEVEL;

            var fieldnameGP = document.form1.fieldnameGP;
            var fieldnameXM = document.form1.fieldnameXM;
            var fieldvalueXM = document.form1.fieldvalueXM;
            var fieldnameLV = document.form1.fieldnameLV;
            var fieldvalueLV = document.form1.fieldvalueLV;
            var fieldbtnN1 = document.form1.fieldbtnN1;

            var oform1 = opener.document.form1;

            $("#errmsg_GP1").hide();
            $("#errmsg_EX1").hide();
            $("#errmsg_LV1").hide();
            if (ddlEXAMGROUP && ddlEXAMGROUP.value == "") {
                $("#errmsg_GP1").show();
                return false;
            }
            if (ddlEXAM3 && ddlEXAM3.value == "") {
                $("#errmsg_EX1").show();
                return false;
            }
            if (ddlEXLEVEL && ddlEXLEVEL.value == "") {
                $("#errmsg_LV1").show();
                return false;
            }

            if (fieldnameGP.value != '') { oform1.elements[fieldnameGP.value].value = ddlEXAMGROUP.options[ddlEXAMGROUP.selectedIndex].text; }
            if (fieldnameLV.value != '') { oform1.elements[fieldnameLV.value].value = ddlEXLEVEL.options[ddlEXLEVEL.selectedIndex].text; }
            if (fieldvalueLV.value != '') { oform1.elements[fieldvalueLV.value].value = ddlEXLEVEL.value; }
            if (fieldnameXM.value != '') { oform1.elements[fieldnameXM.value].value = ddlEXAM3.options[ddlEXAM3.selectedIndex].text; }
            if (fieldvalueXM.value != '') { oform1.elements[fieldvalueXM.value].value = ddlEXAM3.value; }
            if (fieldbtnN1.value != '') { oform1.elements[fieldbtnN1.value].disabled = false; }
            window.close();
        }
        /**
		function CleaningValue() {
		opener.document.form1.elements[document.form1.fieldname.value].value = "";
		opener.document.form1.cjobValue.value = "";
		window.close();
		}
		**/
    </script>
    <link href="../css/style.css" type="text/css" rel="stylesheet">
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="tb_EXAM3" runat="server" class="table_nw" border="0" width="100%">
            <tr>
                <td class="bluecol" style="width: 20%">職類群</td>
                <td class="whitecol">
                    <asp:DropDownList ID="ddlEXAMGROUP" runat="server" AutoPostBack="true">
                    </asp:DropDownList><span id="errmsg_GP1" style="display: none; color: #FF0000">請選擇 職類群!!!</span>
                </td>
            </tr>
            <tr>
                <td class="bluecol">檢定職類</td>
                <td class="whitecol">
                    <asp:DropDownList ID="ddlEXAM3" runat="server" AutoPostBack="true">
                    </asp:DropDownList><span id="errmsg_EX1" style="display: none; color: #FF0000">請選擇 檢定職類!!!</span>
                </td>
            </tr>
            <tr>
                <td class="bluecol">級別</td>
                <td class="whitecol">
                    <asp:DropDownList ID="ddlEXLEVEL" runat="server">
                    </asp:DropDownList><span id="errmsg_LV1" style="display: none; color: #FF0000">請選擇 級別!!!</span>
                </td>
            </tr>
            <tr>
                <td align="center" colspan="2" class="whitecol">
                    <input id="button2" onclick="javascript: ReturnValue2();" type="button" value="選擇" name="button1" runat="server" class="asp_button_M" />
                </td>
            </tr>
        </table>
        <%--<input id="button2" onclick="javascript:CleaningValue();" type="button" value="清除" name="button2" runat="server" class="asp_button_S" />--%>
        <input id="fieldnameGP" type="hidden" runat="server" />
        <input id="fieldnameXM" type="hidden" runat="server" />
        <input id="fieldnameLV" type="hidden" runat="server" />
        <input id="fieldvalueXM" type="hidden" runat="server" />
        <input id="fieldvalueLV" type="hidden" runat="server" />
        <input id="fieldbtnN1" type="hidden" runat="server" />

    </form>
</body>
</html>