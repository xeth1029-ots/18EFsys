<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_04_002.aspx.vb" Inherits="WDAIIP.CP_04_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>CP_04_002</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
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

            //假如是轄區物件
            if (obj == 'DistrictList') {
                var MyValue = getCheckBoxListValue(obj);
                for (i = 1; i < num; i++) {
                    if (MyValue.charAt(i) == '1') {
                        document.getElementById('Dist' + (i - 1)).style.display = 'inline';
                    }
                    else {
                        document.getElementById('Dist' + (i - 1)).style.display = 'none';
                    }
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
        }
        /* function ShowUnit(){
        if(document.getElementById('ShowTR').style.display=='none'){
        document.getElementById('ShowConUnit').innerHTML='關閉管控單位';
        document.getElementById('ShowTR').style.display='inline';
        }
        else{
        document.getElementById('ShowConUnit').innerHTML='展開管控單位';
        document.getElementById('ShowTR').style.display='none';
        }
        }*/
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <table class="font" width="600">
                        <tr>
                            <td>
                                <font class="font" size="2">首頁&gt;&gt;訓練查核與績效管理&gt;&gt;訓練資料查詢&gt;&gt;</font><font class="font" color="#800000" size="2">計畫資料</font>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch" id="Table1" cellpadding="1" cellspacing="1">
                        <tr>
                            <td width="100" class="bluecol_need">年度
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="yearlist" runat="server" AutoPostBack="True">
                                </asp:DropDownList>
                                <asp:RequiredFieldValidator ID="MustYear" runat="server" ErrorMessage="請選擇年度" Display="Dynamic" ControlToValidate="yearlist"></asp:RequiredFieldValidator></FONT></FONT>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">轄區
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="DistrictList" runat="server" CssClass="font" RepeatDirection="Horizontal" Width="512px" Height="11px" CellPadding="0" CellSpacing="0">
                                </asp:CheckBoxList>
                                <input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 10%" class="bluecol">縣市
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="CityList" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="8">
                                </asp:CheckBoxList>
                                <input id="CityHidden" type="hidden" value="0" runat="server" name="CityHidden">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練計畫
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="PlanList" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3" CellPadding="0" CellSpacing="0">
                                </asp:CheckBoxList>
                                <input id="TPlanHidden" type="hidden" value="0" name="TPlanHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">開訓日期
                            </td>
                            <td bgcolor="#ffecec" class="whitecol">
                                <asp:TextBox ID="SSTDate" runat="server" Columns="10" MaxLength="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= SSTDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">&nbsp;~
                            <asp:TextBox ID="ESTDate" runat="server" Columns="10" MaxLength="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= ESTDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">管控單位
                            </td>
                            <td bgcolor="#ffecec" class="whitecol">
                                <table class="font" id="Table3" cellspacing="0" cellpadding="0" width="100%" border="0">
                                    <tr id="Dist0" runat="server">
                                        <td>
                                            <asp:CheckBoxList ID="IsConUnit1" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="2" CellPadding="0" CellSpacing="0">
                                            </asp:CheckBoxList>
                                        </td>
                                    </tr>
                                    <tr id="Dist1" runat="server">
                                        <td>
                                            <asp:CheckBoxList ID="IsConUnit2" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="2" CellPadding="0" CellSpacing="0">
                                            </asp:CheckBoxList>
                                        </td>
                                    </tr>
                                    <tr id="Dist2" runat="server">
                                        <td>
                                            <asp:CheckBoxList ID="IsConUnit3" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="2" CellPadding="0" CellSpacing="0">
                                            </asp:CheckBoxList>
                                        </td>
                                    </tr>
                                    <tr id="Dist3" runat="server">
                                        <td>
                                            <asp:CheckBoxList ID="IsConUnit4" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="2" CellPadding="0" CellSpacing="0">
                                            </asp:CheckBoxList>
                                        </td>
                                    </tr>
                                    <tr id="Dist4" runat="server">
                                        <td>
                                            <asp:CheckBoxList ID="IsConUnit5" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="2" CellPadding="0" CellSpacing="0">
                                            </asp:CheckBoxList>
                                        </td>
                                    </tr>
                                    <tr id="Dist5" runat="server">
                                        <td>
                                            <asp:CheckBoxList ID="IsConUnit6" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="2" CellPadding="0" CellSpacing="0">
                                            </asp:CheckBoxList>
                                        </td>
                                    </tr>
                                    <tr id="Dist6" runat="server">
                                        <td>
                                            <asp:CheckBoxList ID="IsConUnit7" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="2" CellPadding="0" CellSpacing="0">
                                            </asp:CheckBoxList>
                                        </td>
                                    </tr>
                                </table>
                                <asp:Literal ID="msg" runat="server"></asp:Literal>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <table class="font" id="Table5" cellspacing="0" cellpadding="0" width="740" border="0">
            <tr>
                <td align="center">
                    <font face="新細明體">
                        <asp:Button ID="bt_search" runat="server" Text="明細查詢" Width="60px" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="bt_search1" runat="server" Width="60px" Text="統計查詢" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="bt_reset" runat="server" Text="重新設定" Width="60px" CssClass="asp_button_M"></asp:Button></font>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
