<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_05_021_R.aspx.vb" Inherits="WDAIIP.TR_05_021_R" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>原住民職訓措施辦理情形</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        function SelectAll(obj, hidobj) {
            var num = getCheckBoxListValue(obj).length;
            var myallcheck = document.getElementById(obj + '_' + 0);
            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
        }

        function search() {
            var msg = '';
            var STDate1 = document.getElementById('STDate1');
            var STDate2 = document.getElementById('STDate2');
            var FTDate1 = document.getElementById('FTDate1');
            var FTDate2 = document.getElementById('FTDate2');
            /*if(document.form1.Syear.selectedIndex==0) msg+='請選擇年度\n';*/
            /*if(!isChecked(document.getElementsByName('DistID'))) msg+='請選擇轄區\n';
			if(!isChecked(document.getElementsByName('TPlanID'))) msg+='請選擇訓練計畫\n';*/
            if (STDate1.value != '') {
                if (!checkDate(STDate1.value)) msg += '開訓日期的起始日不是正確的日期格式\n';
            }
            if (STDate2.value != '') {
                if (!checkDate(STDate2.value)) msg += '開訓日期的結束日不是正確的日期格式\n';
            }
            if (FTDate1.value != '') {
                if (!checkDate(FTDate1.value)) msg += '結訓日期的起始日不是正確的日期格式\n';
            }
            if (FTDate2.value != '') {
                if (!checkDate(FTDate2.value)) msg += '結訓日期的結束日不是正確的日期格式\n';
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td><font face="新細明體">
                    <table id="Table1" class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label><asp:Label ID="TitleLab2" runat="server">
							首頁&gt;&gt;訓練與就業需求管理&gt;&gt;統計分析&gt;&gt;<FONT color="#990000">原住民職訓措施辦理情形</FONT>
                                </asp:Label>
                            </td>
                        </tr>
                    </table>
                </font>
                    <table id="Table2" class="table_sch" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol" width="100">年度 </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="Syear" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">轄區 </td>
                            <td class="whitecol">&nbsp;<asp:CheckBoxList ID="DistID" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                            </asp:CheckBoxList>
                                <input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練計畫 </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="TPlanID" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3">
                                </asp:CheckBoxList>
                                <input id="TPlanHidden" type="hidden" value="0" name="TPlanHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">開訓期間 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="STDate1" runat="server" Columns="10"></asp:TextBox><img onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" style="cursor: pointer" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30"><font color="#000000">～
								<asp:TextBox ID="STDate2" runat="server" Columns="10"></asp:TextBox><img onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" style="cursor: pointer" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30"></font> </td>
                        </tr>
                        <tr>
                            <td class="bluecol">結訓期間 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="FTDate1" runat="server" Columns="10"></asp:TextBox><img onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" style="cursor: pointer" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30"><font color="#000000">～
								<asp:TextBox ID="FTDate2" runat="server" Columns="10"></asp:TextBox><img onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" style="cursor: pointer" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30"></font> </td>
                        </tr>
                    </table>
                    <p align="center">
                        <asp:Button ID="Export1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button><font face="新細明體">&nbsp;</font>
                        <%--<asp:Button ID="Button1" runat="server" CssClass="asp_button_S" Text="列印"></asp:Button>--%>
                    </p>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
