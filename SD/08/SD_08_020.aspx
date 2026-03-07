<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_08_020.aspx.vb" Inherits="WDAIIP.SD_08_020" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>職業訓練生活津貼請領檢核表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../style.css" type="text/css" rel="stylesheet">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script>
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
            if (isEmpty(document.form1.STDate1) && isEmpty(document.form1.STDate2)) msg += '請選擇開訓日期範圍!\n';
            /*if(!isChecked(document.getElementsByName('TPlanID'))) msg+='請選擇訓練計畫\n';*/

            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <p>
            <table id="FrameTable" cellspacing="1" cellpadding="1" width="600" border="0">
                <tr>
                    <td>
                        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <tr>
                                <td>首頁&gt;&gt;學員動態管理&gt;&gt;職業訓練生活津貼&gt;&gt;<font color="#990000">職業訓練生活津貼請領檢核表</font></td>
                            </tr>
                        </table>
                        <table class="font" id="Table2" style="width: 696px; height: 285px" cellspacing="1" cellpadding="1"
                            width="696" border="0" runat="server">
                            <tr>
                                <td class="CM_TD1" style="height: 28px" width="80"><font face="新細明體">&nbsp;&nbsp;&nbsp; 
											開訓區間 </font><font color="red">*</font></td>
                                <td class="CM_TD2" style="height: 28px">
                                    <asp:TextBox ID="STDate1" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');"
                                        alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">~<asp:TextBox ID="STDate2" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');"
                                            alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></td>
                            </tr>
                            <%--<TR>
									<TD class="CM_TD1" width="80"><FONT face="新細明體">&nbsp;&nbsp;&nbsp; 結訓區間</FONT></TD>
									<TD class="CM_TD2"><asp:textbox id="FTDate1" runat="server" Columns="10"></asp:textbox><IMG style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');"
											alt="" src="../../images/show-calendar.gif" align="top">~<asp:textbox id="FTDate2" runat="server" Columns="10"></asp:textbox><IMG style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');"
											alt="" src="../../images/show-calendar.gif" align="top"></TD>
								</TR>
								<TR>
									<TD class="CM_TD1"><FONT face="新細明體">&nbsp;&nbsp;&nbsp; 轄區</FONT></TD>
									<TD class="CM_TD2"><asp:CheckBoxList id="DistID" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow"></asp:CheckBoxList>
										<INPUT id="DistHidden" type="hidden" value="0" runat="server" NAME="DistHidden">
									</TD>
								</TR>--%>
                            <tr>
                                <td class="CM_TD1">&nbsp;&nbsp;&nbsp; 訓練計畫</td>
                                <td class="CM_TD2">
                                    <asp:CheckBoxList ID="TPlanID" runat="server" RepeatDirection="Horizontal" CssClass="font" CellSpacing="0"
                                        CellPadding="0" RepeatColumns="3">
                                    </asp:CheckBoxList>
                                    <input id="TPlanHidden" type="hidden" value="0" name="TPlanHidden" runat="server" size="1">
                                </td>
                            </tr>
                            <tr>
                                <td align="center" colspan="2">
                                    <asp:Button ID="Print" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </p>
    </form>
</body>
</html>
