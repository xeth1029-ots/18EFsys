<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SendCJob.aspx.vb" Inherits="WDAIIP.SendCJob" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>請選擇通俗職類</title>
    <meta content="microsoft visual studio .net 7.1" name="generator" />
    <meta content="visual basic .net 7.1" name="code_language" />
    <meta content="javascript" name="vs_defaultclientscript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetschema" />
    <script language="javascript" type="text/javascript">
        function ReturnValue() {
            var CJOB_TYPE = document.form1.CJOB_TYPE;
            var CJOB_NO = document.form1.CJOB_NO;
            var fieldname = document.form1.fieldname;
            var oform1 = opener.document.form1;
            if (CJOB_TYPE && CJOB_TYPE.value == "") {
                alert('請選擇大類!!!');
                return false;
            }
            if (CJOB_NO && CJOB_NO.value == "") {
                alert('請選擇小類!!!');
                return false;
            }
            if (oform1 && CJOB_NO && fieldname) {
                oform1.elements[fieldname.value].value = CJOB_NO.options[CJOB_NO.selectedIndex].text;
                oform1.cjobValue.value = CJOB_NO.value;
                window.close();
            }
        }

        /*2016通俗*/
        function ReturnValue2() {
            var ddlCJOB16A1 = document.form1.ddlCJOB16A1;
            var ddlCJOB16A2 = document.form1.ddlCJOB16A2;
            var ddlCJOB16A3 = document.form1.ddlCJOB16A3;
            //var CJOB_NO = document.form1.CJOB_NO;
            var fieldname = document.form1.fieldname;
            var oform1 = opener.document.form1;
            if (ddlCJOB16A1 && ddlCJOB16A1.value == "") {
                alert('請選擇大類!!!');
                return false;
            }
            if (ddlCJOB16A2 && ddlCJOB16A2.value == "") {
                alert('請選擇中類!!!');
                return false;
            }
            if (ddlCJOB16A3 && ddlCJOB16A3.value == "") {
                alert('請選擇小類!!!');
                return false;
            }
            if (oform1 && ddlCJOB16A3 && fieldname) {
                oform1.elements[fieldname.value].value = ddlCJOB16A3.options[ddlCJOB16A3.selectedIndex].text;
                oform1.cjobValue.value = ddlCJOB16A3.value;
                window.close();
            }
        }
        /** function CleaningValue(){opener.document.form1.elements[document.form1.fieldname.value].value = "";	
        pener.document.form1.cjobValue.value = "";window.close();} **/
    </script>
    <link href="../css/style.css" type="text/css" rel="stylesheet">
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="tbCJOB1" runat="server" class="table_nw" border="0" width="100%">
            <tr>
                <td class="bluecol" style="width: 20%">大類
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="CJOB_TYPE" runat="server" AutoPostBack="true">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol" style="width: 20%">小類
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="CJOB_NO" runat="server" AutoPostBack="true">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td align="center" colspan="2" class="whitecol">
                    <input id="button1" onclick="javascript: ReturnValue();" type="button" value="選擇" name="button1" runat="server" class="asp_button_M" />
                </td>
            </tr>
        </table>
        <table id="tbCJOB2" runat="server" class="table_nw" border="0" width="100%">
            <tr>
                <td class="bluecol" style="width: 20%">大類
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="ddlCJOB16A1" runat="server" AutoPostBack="true">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">中類
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="ddlCJOB16A2" runat="server" AutoPostBack="true">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">小類
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="ddlCJOB16A3" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td align="center" colspan="2" class="whitecol">
                    <input id="button2" onclick="javascript: ReturnValue2();" type="button" value="選擇" name="button1" runat="server" class="asp_button_M" />
                </td>
            </tr>
        </table>
        <%--<input id="button2" onclick="javascript:CleaningValue();" type="button" value="清除" name="button2" runat="server" class="asp_button_S" />--%>
        <input id="fieldname" type="hidden" runat="server" />
    </form>
</body>
</html>
