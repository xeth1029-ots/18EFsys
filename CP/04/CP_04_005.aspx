<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_04_005.aspx.vb" Inherits="WDAIIP.CP_04_005" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>CP_04_005</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <style type="text/css">
        .class_link A { color: #000000; }
            .class_link A:link { color: #0000ff; }
            .class_link A:hover { color: #0000ff; }
        A:visited { color: #0000ff; }
        A:active { color: #0000ff; }
    </style>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        //選擇全部
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


        //檢查日期格式
        function check_date() {
            if (!checkDate(form1.SSTDate.value) || !checkDate(form1.ESTDate.value)) {
                document.form1.SSTDate.value = '';
                document.form1.ESTDate.value = '';
                alert('請輸入正確的日期格式,YYYY/MM/DD!!\n');
            }

            if (!checkDate(form1.SFTDate.value) || !checkDate(form1.EFTDate.value)) {
                document.form1.SFTDate.value = '';
                document.form1.EFTDate.value = '';
                alert('請輸入正確的日期格式,YYYY/MM/DD!!\n');
            }

        }

        //檢查開訓及結訓日期為
        function search() {

            var msg = '';

            if (document.form1.SSTDate.value != '') {
                if ((document.form1.SFTDate.value == '') && (document.form1.EFTDate.value == '') && (document.form1.ESTDate.value == '')) msg += '請選擇開訓迄日!\n';

            }

            if (document.form1.ESTDate.value != '') {

                if ((document.form1.SFTDate.value == '') && (document.form1.EFTDate.value == '') && (document.form1.SSTDate.value == '')) msg += '請選擇開訓起日!\n';

            }

            if (document.form1.SFTDate.value != '') {

                if ((document.form1.SSTDate.value == '') && (document.form1.ESTDate.value == '') && (document.form1.EFTDate.value == '')) msg += '請選擇結訓訖日!\n';

            }

            if (document.form1.EFTDate.value != '') {
                if ((document.form1.SSTDate.value == '') && (document.form1.ESTDate.value == '') && (document.form1.SFTDate.value == '')) msg += '請選擇結訓起日!\n';
            }

            if ((document.form1.SSTDate.value == '') && (document.form1.ESTDate.value == '') && (document.form1.SFTDate.value == '') && (document.form1.EFTDate.value == '')) msg += '開訓日期及結訓日期請擇一輸入!\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }


    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <table id="Table9" class="font" width="100%">
                        <tr>
                            <td>
                                <font class="font" size="2">首頁&gt;&gt;訓練查核與績效管理&gt;&gt;訓練資料查詢&gt;&gt;<font color="#800000">上課地點統計表</font></font>
                            </td>
                        </tr>
                    </table>
                    <table id="Table1" class="table_sch" cellspacing="1" cellpadding="1">
                        <tbody>
                            <tr id="Dist" runat="server">
                                <td width="12%" class="bluecol">轄區
                                </td>
                                <td bgcolor="#ffecec" class="whitecol">
                                    <asp:CheckBoxList ID="DistID" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                                    </asp:CheckBoxList>
                                    <input id="DistHidden" value="0" type="hidden" name="DistHidden" runat="server">
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">縣市
                                </td>
                                <td bgcolor="#ffecec" class="whitecol">
                                    <asp:CheckBoxList ID="CityID" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="8">
                                    </asp:CheckBoxList>
                                    <input id="CityHidden" value="0" type="hidden" name="CityHidden" runat="server">
                                </td>
                            </tr>
                            <tr>
                                <td style="width: 12%" class="bluecol">訓練計畫
                                </td>
                                <td bgcolor="#ffecec" class="whitecol">
                                    <asp:CheckBoxList ID="TPlanID" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="3">
                                    </asp:CheckBoxList>
                                    <input id="TPlanHidden" value="0" type="hidden" name="TPlanHidden" runat="server">
                                </td>
                                <tr>
                                    <td style="width: 12%;" class="bluecol">開訓日期
                                    </td>
                                    <td bgcolor="#ffecec" class="whitecol">
                                        <asp:TextBox ID="SSTDate" runat="server" MaxLength="10" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= SSTDate.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">&nbsp;~
									<asp:TextBox ID="ESTDate" runat="server" MaxLength="10" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= ESTDate.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">
                                    </td>
                                </tr>
                            <tr>
                                <td style="width: 12%;" class="bluecol">結訓日期
                                </td>
                                <td bgcolor="#ffecec" class="whitecol">
                                    <asp:TextBox ID="SFTDate" runat="server" MaxLength="10" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= SFTDate.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">&nbsp;~
								<asp:TextBox ID="EFTDate" runat="server" MaxLength="10" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= EFTDate.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">
                                </td>
                            </tr>
                            <tr>
                                <td style="width: 12%" class="bluecol">報表類別
                                </td>
                                <td bgcolor="#ffecec" class="whitecol">
                                    <asp:RadioButtonList ID="PrintStaus" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                        <asp:ListItem Value="1" Selected="True">縣市</asp:ListItem>
                                        <asp:ListItem Value="2">鄉鎮市區</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </td>
            </tr>
        </table>
        <table id="Table6" class="font" border="0" cellspacing="0" cellpadding="0" width="740">
            <tr align="center">
                <td>
                    <asp:Button ID="print" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                    <asp:Button ID="bt_reset" runat="server" Text="重新設定" CssClass="asp_button_M"></asp:Button>
                    <asp:Button ID="bt_export" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
