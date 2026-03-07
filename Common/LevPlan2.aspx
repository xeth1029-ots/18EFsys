<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="LevPlan2.aspx.vb" Inherits="WDAIIP.LevPlan2" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>LevPlan2</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        //if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function ReturnLevPlan2() {
            var obj1 = document.getElementById('Years');
            var obj2 = document.getElementById('drpDist');
            var obj3 = document.getElementById('txtPlan');
            var obj4 = document.getElementById('drpGov');
            var obj5 = document.getElementById('drpOrg');
            document.getElementById('YearsValue').value = obj1.value;
            document.getElementById('DistValue').value = obj2.value;
            document.getElementById('PlanIDValue').value = obj3.value;
            document.getElementById('OrgIDValue').value = obj5.value;
            var msg1 = document.getElementById('YearsValue').value;
            var msg2 = document.getElementById('DistValue').value;
            var msg3 = document.getElementById('PlanIDValue').value;
            var msg4 = document.getElementById('RIDValue').value;
            var msg5 = document.getElementById('OrgIDValue').value;
            var id1 = obj1.selectedIndex;
            var id2 = obj2.selectedIndex;
            var id3 = obj3.selectedIndex;
            var id4 = obj4.selectedIndex;
            var id5 = obj5.selectedIndex;
            var txt1 = '', txt2 = '', txt3 = '', txt4 = '', txt5 = '';
            if (msg1 != '') txt1 = obj1.options[id1].text;
            if (msg2 != '') txt2 = obj2.options[id2].text;
            if (msg3 != '') txt3 = obj3.options[id3].text;
            if (msg4 != '') txt4 = '_' + obj4.options[id4].text;
            if (msg5 != '') {
                txt5 = '_' + obj5.options[id5].text;
                //alert(txt3+'_'+txt4+'_'+txt5);
                //alert(msg1+','+msg2+','+msg3+','+msg4+','+msg5);
                if (getParamValue('YearsField') != '') {
                    opener.document.getElementById(getParamValue('YearsField')).value = msg1;
                }
                if (getParamValue('DistField') != '') {
                    opener.document.getElementById(getParamValue('DistField')).value = msg2;
                }
                if (getParamValue('PlanIDField') != '') {
                    opener.document.getElementById(getParamValue('PlanIDField')).value = msg3;
                }
                if (getParamValue('RIDField') != '') {
                    opener.document.getElementById(getParamValue('RIDField')).value = msg4;
                }
                if (getParamValue('OrgIDField') != '') {
                    opener.document.getElementById(getParamValue('OrgIDField')).value = msg5;
                }
                if (getParamValue('TextField') != '') {
                    opener.document.getElementById(getParamValue('TextField')).value = txt3 + txt4 + txt5;
                }
                window.close();
            }
            else {
                alert('請選擇訓練機構!!');
            }
        }

        function ReturnLevPlan2_cls() {
            //var msg1='', msg2='', msg3='', msg4='', msg5='' ;
            var msg1 = "", msg2 = "", msg3 = "", msg4 = "", msg5 = "";
            if (getParamValue('YearsField') != '') {
                opener.document.getElementById(getParamValue('YearsField')).value = msg1;
            }
            if (getParamValue('DistField') != '') {
                opener.document.getElementById(getParamValue('DistField')).value = msg2;
            }
            if (getParamValue('PlanIDField') != '') {
                opener.document.getElementById(getParamValue('PlanIDField')).value = msg3;
            }
            if (getParamValue('RIDField') != '') {
                opener.document.getElementById(getParamValue('RIDField')).value = msg4;
            }
            if (getParamValue('OrgIDField') != '') {
                opener.document.getElementById(getParamValue('OrgIDField')).value = msg5;
            }
            if (getParamValue('TextField') != '') {
                opener.document.getElementById(getParamValue('TextField')).value = '';
            }
            window.close();
        }
    </script>
    <%--<LINK href="../style.css" type="text/css" rel="stylesheet">--%>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="TB_A" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tbody>
                <tr>
                    <td>
                        <table class="font" id="TB_1" width="100%" runat="server">
                            <tr>
                                <td class="bluecol" style="width: 20%">年度</td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="Years" runat="server" AutoPostBack="True"></asp:DropDownList>
                                    <input id="YearsValue" type="hidden" name="YearsValue" runat="server">
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">轄區</td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="drpDist" runat="server" AutoPostBack="True"></asp:DropDownList>
                                    <input id="DistValue" type="hidden" name="DistValue" runat="server">
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">計劃代碼</td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="txtPlan" runat="server" AutoPostBack="True"></asp:DropDownList>
                                    <input id="PlanIDValue" type="hidden" name="PlanIDValue" runat="server">
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">補助單位</td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="drpGov" runat="server" AutoPostBack="True"></asp:DropDownList>
                                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">訓練機構<font color="yellow">*</font></td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="drpOrg" runat="server"></asp:DropDownList>
                                    <input id="OrgIDValue" type="hidden" name="OrgIDValue" runat="server">
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </tbody>
        </table>
        <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td align="center" colspan="2" class="whitecol">
                    <input id="send" type="button" value="送出" name="back" runat="server" class="asp_button_M" onclick="ReturnLevPlan2();">
                    <input id="clear" type="button" value="清空" name="back" runat="server" class="asp_button_M" onclick="ReturnLevPlan2_cls();">
                </td>
            </tr>
        </table>
    </form>
</body>
</html>