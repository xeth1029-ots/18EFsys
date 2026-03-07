<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_04_004_R.aspx.vb" Inherits="WDAIIP.TR_04_004_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>TR_04_004_R</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function chk() {
            var msg = '';
            //if (document.form1.Syear.value == '') {
            //msg += "請選擇列印年度!\n";
            //}
            if (document.form1.STDate1.value != '') {
                if (!IsDate(document.form1.STDate1.value)) msg += '開訓日期的起始日不是正確的日期格式\n';
            }
            if (document.form1.STDate2.value != '') {
                if (!IsDate(document.form1.STDate2.value)) msg += '開訓日期的迄日不是正確的日期格式\n';
            }
            if (document.form1.FTDate1.value != '') {
                if (!IsDate(document.form1.FTDate1.value)) msg += '結訓日期的起始日不是正確的日期格式\n';
            }
            if (document.form1.FTDate2.value != '') {
                if (!IsDate(document.form1.FTDate2.value)) msg += '結訓日期的迄日不是正確的日期格式\n';
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
        }
        function IsDate(MyDate) {
            if (MyDate != '') {
                if (!checkDate(MyDate))
                    return false;
            }
            return true;
        }
        function showHide(type) {
            if (type == 1) {
                document.all.nxlayer_01.style.visibility = 'visible'
                document.all.Syear.style.visibility = 'hidden'
            } else {
                document.all.nxlayer_01.style.visibility = 'hidden'
                document.all.Syear.style.visibility = 'visible'
            }
        }
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
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td title="點選可以查看說明" style="cursor: pointer" onclick="showHide(1)">
                                <asp:Label ID="Label1" runat="server"></asp:Label>
                                <asp:Label ID="Label2" runat="server">
									首頁&gt;&gt;訓練與就業需求管理&gt;&gt;就業追蹤報表&gt;&gt;<FONT color="#990000">職業訓練累計統計總表_依縣市別及性別</FONT>
                                </asp:Label>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table class="font" id="nxlayer_01" style="border-right: #9eb5cd 1px solid; border-top: #9eb5cd 1px solid; filter: progid: DXImageTransform.Microsoft.Shadow(Color=#919899, Strength=4, Direction=135); visibility: hidden; border-left: #9eb5cd 1px solid; border-bottom: #9eb5cd 1px solid; position: absolute" cellspacing="0" cellpadding="0" width="100%" border="0">
                        <tbody>
                            <tr>
                                <td width="90%" bgcolor="#ffffff" align="center">
                                    <a onclick="showHide(0)" href="#"><font color="black">關閉[X]</font></a>
                                </td>
                            </tr>
                            <tr>
                                <td class="dashline" style="height: 1px" height="1">
                                    <u></u>
                                </td>
                            </tr>
                            <tr>
                                <td style="padding-right: 8px; padding-left: 8px; padding-bottom: 6px; padding-top: 8px" width="100%" bgcolor="#f1faff" colspan="2">說明：<br>
                                    若輸入開訓期間:2006/01/01~2006/12/31 則會查出2006年開訓的班級 (含2006年開訓但結訓日跨2007年的班級)
                                <br>
                                    若輸入結訓期間:2006/01/01~2006/12/31 則會查出2006年結訓的班級 (含2005年開訓,但是結訓日期是在2006年的班級)
                                <br>
                                    若輸入開訓期間:2006/01/01~ 不填
                                <br>
                                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 結訓期間: &nbsp;&nbsp;&nbsp; 不填 &nbsp; &nbsp; ~ 2006/12/31 則會查出在2006年開訓及結訓的班級 (不含跨年度)
                                <br>
                                    以上情況的前提是:[年度] 需選擇 [====請選擇====]
                                </td>
                            </tr>
                        </tbody>
                    </table>
                    <table class="table_sch" id="Table2" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" width="80">年度
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="Syear" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">轄區
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="DistrictList" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                </asp:CheckBoxList>
                                <input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練計畫
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="TPlanID" TabIndex="3" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="3">
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">開訓期間
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="STDate1" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"><font color="#000000">～
                                <asp:TextBox ID="STDate2" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></font>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">結訓期間
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="FTDate1" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"><font color="#000000">～
                                <asp:TextBox ID="FTDate2" runat="server" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></font>
                            </td>
                        </tr>
                    </table>
                    <p align="center">
                        <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                    </p>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
