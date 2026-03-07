<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_11_007.aspx.vb" Inherits="WDAIIP.SD_11_007" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>勞保明細資料查詢</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="../../js/TIMS.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function GETvalue() {
            document.getElementById('Button12').click();
        }

        function ClearData() {
            document.getElementById('TMID1').value = '';
            document.getElementById('OCID1').value = '';
            document.getElementById('TMIDValue1').value = '';
            document.getElementById('OCIDValue1').value = '';
        }

        function CheckPrint() {
            var msg = '';
            if (getRadioValue(document.getElementsByName('searchMode')) == '1') {
                if (document.form1.RIDValue.value == '') { msg += '請選擇[訓練機構]\n'; }
                if (document.form1.OCIDValue1.value == '') { msg += '請選擇[班級]\n'; }
            }
            if (getRadioValue(document.getElementsByName('searchMode')) == '2') {
                if (document.form1.birthday.value != '' && !checkDate(document.form1.birthday.value)) { msg = msg + '[出生日期]時間格式不正確!\n'; }
                if (document.form1.birthday.value == '' && document.form1.IDNO.value == '') { msg += '[身分證號碼]及[出生日期]必須擇一輸入\n'; }
            }
            if (getRadioValue(document.getElementsByName('searchMode')) == '3') {
                if (document.form1.start_date.value != '') {
                    if (!checkDate(document.form1.start_date.value)) msg += '[開訓起日]不是正確的日期格式\n';
                }
                if (document.form1.end_date.value != '') {
                    if (!checkDate(document.form1.end_date.value)) msg += '[開訓迄日]不是正確的日期格式\n';
                }
                if (document.form1.end_date.value != '' && document.form1.start_date.value != '' && document.form1.end_date.value < document.form1.start_date.value) { msg += '[開訓迄日]必需大於[開訓起日]\n'; }
                if (document.form1.start_date.value == '' && document.form1.end_date.value == '') { msg += '請輸入[開訓起迄日期]\n'; }
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function choose_class() {
            openClass('../02/SD_02_ch.aspx?RID=' + document.getElementById('RIDValue').value);
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td align="center">
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;勞保明細資料查詢</asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch" id="Table3">
                        <tr>
                            <td class="bluecol" style="width: 20%">查詢類型</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="searchMode" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" AutoPostBack="True">
                                    <asp:ListItem Value="1" Selected="True">依班別查詢</asp:ListItem>
                                    <asp:ListItem Value="2">依個別學員查詢</asp:ListItem>
                                    <asp:ListItem Value="3">依開訓日期查詢</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="org_tr" runat="server">
                            <td class="bluecol_need">訓練機構</td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                <input id="Button8" type="button" value="..." name="Button5" runat="server">
                                <input id="DistValue" type="hidden" name="DistValue" runat="server">
                                <asp:Button ID="Button12" Style="display: none" runat="server" Text="Button12"></asp:Button>
                                <span id="HistoryList2" style="display: none; position: absolute" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr id="class_tr" runat="server">
                            <td class="bluecol_need">班級名稱</td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input id="Button5" onclick="choose_class();" type="button" value="..." name="Button5" runat="server">
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                <span id="HistoryList" style="display: none; left: 28%; position: absolute">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr id="idno_tr" runat="server">
                            <td class="bluecol">身分證號碼</td>
                            <td class="whitecol">
                                <asp:TextBox ID="IDNO" runat="server" Width="20%"></asp:TextBox></td>
                        </tr>
                        <tr id="birthday_tr" runat="server">
                            <td class="bluecol">出生日期</td>
                            <td class="whitecol">
                                <asp:TextBox ID="birthday" runat="server" Width="15%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= birthday.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr id="date_tr" runat="server">
                            <td class="bluecol_need">開訓起迄日期</td>
                            <td class="whitecol">
                                <asp:TextBox ID="start_date" runat="server" Width="15%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span> ~
                                <asp:TextBox ID="end_date" runat="server" Width="15%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr id="tr_ddl_INQUIRY_S" runat="server">
                            <td class="bluecol_need">查詢原因</td>
                            <td class="whitecol">
                                <asp:DropDownList ID="ddl_INQUIRY_Sch" runat="server"></asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="4" class="whitecol">
                                <asp:Label ID="labPageSize" runat="server" DESIGNTIMEDRAGDROP="30" ForeColor="SlateBlue">顯示列數</asp:Label>
                                <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                <asp:Button ID="Button2" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                <%--<asp:Button ID="Print" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>--%>
                            </td>
                        </tr>
                    </table>
                    <table id="Table4" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td align="center">
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AllowPaging="True" AutoGenerateColumns="False" CssClass="font" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn HeaderText="編號">
                                            <HeaderStyle HorizontalAlign="Center" Width="3%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Name" HeaderText="姓名">
                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="idno" HeaderText="身分證號碼">
                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="birthday" HeaderText="出生日期">
                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ActNo" HeaderText="保險證號">
                                            <HeaderStyle Width="8%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="COMNAME" HeaderText="投保單位名稱">
                                            <HeaderStyle Width="15%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="MDATE" HeaderText="勞保異動日期">
                                            <HeaderStyle Width="8%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="MODIFYDATE" HeaderText="系統勾稽<br>異動日期">
                                            <HeaderStyle Width="8%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <%--<asp:BoundColumn DataField="Actname" HeaderText="投保單位名稱">
                                            <HeaderStyle Width="20%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="MdateADD" HeaderText="生效日期" DataFormatString="{0:d}">
                                            <HeaderStyle Width="8%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="MdateL" HeaderText="退保日期" DataFormatString="{0:d}">
                                            <HeaderStyle Width="8%"></HeaderStyle>
                                        </asp:BoundColumn>--%>
                                        <asp:BoundColumn DataField="CLASSCNAME" HeaderText="參訓班別名稱">
                                            <HeaderStyle Width="20%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDATE" HeaderText="參訓班別開訓日">
                                            <HeaderStyle Width="8%"></HeaderStyle>
                                        </asp:BoundColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <uc1:PageControler ID="PageControler1" runat="server" Visible="False"></uc1:PageControler>
                            </td>
                        </tr>
                    </table>
                    <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
