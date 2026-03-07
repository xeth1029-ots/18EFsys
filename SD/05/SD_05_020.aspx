<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_020.aspx.vb" Inherits="WDAIIP.SD_05_020" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>學員重複參訓</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript">
        function closeDiv() {
            document.getElementById('eMeng').style.visibility = 'hidden';
        }

        function EXLINPUT_Chenk() {
            var msg = '';

            if (document.form1.start_date.value != '') {
                if (!checkDate(document.form1.start_date.value)) msg += '[重複參訓日期的起日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }
            if (document.form1.end_date.value != '') {
                if (!checkDate(document.form1.end_date.value)) msg += '[重複參訓日期的迄日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }
            if (document.form1.end_date.value != '' && document.form1.start_date.value != '' && document.form1.end_date.value < document.form1.start_date.value) {
                msg += '[重複參訓日期的迄日]必需大於[重複參訓日期的起日]\n';
            }

            if (document.form1.start_date.value == '') {
                msg += '請輸入[重複參訓日期起日] 或 輸入[重複參訓日期]之[起日]及[迄日]\n';
            }

            if (document.form1.File1.value == '') {
                msg += '請按[瀏覽]鍵,選取欲匯入的檔案\n';
            }

            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function GETvalue() { document.getElementById('Button7').click(); }

        function choose_class() {
            openClass('../02/SD_02_ch.aspx?RID=' + document.getElementById('RIDValue').value);
        }

        function CheckPrint() {
            var msg = '';

            if (getRadioValue(document.getElementsByName('searchMode')) == '1') {
                if (document.form1.RIDValue.value == '') { msg += '請選擇[訓練機構]\n'; }
                if (document.form1.OCIDValue1.value == '') { msg += '請選擇[班級]\n'; }
            }

            if (getRadioValue(document.getElementsByName('searchMode')) == '2') {
                if (document.form1.start_date.value != '') {
                    if (!checkDate(document.form1.start_date.value)) msg += '[重複參訓日期的起日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
                }
                if (document.form1.end_date.value != '') {
                    if (!checkDate(document.form1.end_date.value)) msg += '[重複參訓日期的迄日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
                }

                if (document.form1.end_date.value != '' && document.form1.start_date.value != '' && document.form1.end_date.value < document.form1.start_date.value) { msg += '[重複參訓日期的迄日]必需大於[重複參訓日期的起日]\n'; }

                if (document.form1.start_date.value == '') { msg += '請輸入[重複參訓日期起日] 或 輸入[重複參訓日期]之[起日]及[迄日]\n'; }

                if (document.form1.IDNO.value == '') { msg += '請輸入[身分證號碼]\n'; }
            }

            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">學員重複參訓</font>
                    </asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td align="center">

                    <table class="table_sch" id="Table3" cellpadding="1" cellspacing="1" width="100%" border="0">
                        <tr>
                            <td class="bluecol" width="15%">查詢類型 </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="searchMode" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" AutoPostBack="True">
                                    <asp:ListItem Value="1" Selected="True">依班別查詢</asp:ListItem>
                                    <asp:ListItem Value="2">依個別學員查詢</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="date_tr" runat="server">
                            <td class="bluecol">重複參訓日期 </td>
                            <td class="whitecol" runat="server" colspan="3">
                                <asp:TextBox ID="start_date" runat="server" Columns="13" Width="15%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                ~
                                <asp:TextBox ID="end_date" runat="server" Columns="13" Width="15%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            </td>
                        </tr>
                        <tr id="org_tr" runat="server">
                            <td class="bluecol">訓練機構 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" size="1" name="RIDValue" runat="server">
                                <input id="Button8" type="button" value="..." name="Button5" runat="server" class="button_b_Mini">
                                <input id="DistValue" type="hidden" size="1" name="DistValue" runat="server">
                                <asp:Button ID="Button7" Style="display: none" runat="server" Text="Button7"></asp:Button>
                                <span id="HistoryList2" style="display: none; position: absolute" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%">
                                    </asp:Table>
                                </span></td>
                        </tr>
                        <tr id="class_tr" runat="server">
                            <td class="bluecol">班級名稱 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input id="Button5" onclick="choose_class();" type="button" value="..." name="Button5" runat="server" class="button_b_Mini">
                                <input id="OCIDValue1" type="hidden" size="1" name="OCIDValue1" runat="server">
                                <input id="TMIDValue1" type="hidden" size="1" name="TMIDValue1" runat="server">
                                <span id="HistoryList" style="display: none; left: 270px; position: absolute">
                                    <asp:Table ID="HistoryTable" runat="server" Width="310px">
                                    </asp:Table>
                                </span>

                            </td>
                        </tr>
                        <tr id="idno_tr" runat="server">
                            <td class="bluecol">身分證號碼 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="IDNO" runat="server"></asp:TextBox>
                            </td>
                        </tr>
                        <tr id="trButton10" runat="server">
                            <td class="bluecol">匯入名冊 </td>
                            <td class="whitecol" colspan="3">
                                <input id="File1" type="file" name="File1" runat="server" size="40" accept=".xls,.ods" />
                                <asp:Button ID="Button10" runat="server" Text="匯入名冊" CssClass="asp_Export_M"></asp:Button>(必須為ods或xls格式)
								<asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="../../Doc/StudIDNO_v21.zip" CssClass="font" ForeColor="#8080FF">下載整批上載格式檔</asp:HyperLink>
                            </td>
                        </tr>
                        <tr id="tr_ddl_INQUIRY_S" runat="server">
                            <td class="bluecol_need">查詢原因</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="ddl_INQUIRY_Sch" runat="server"></asp:DropDownList>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>&nbsp;
                                <asp:Button ID="Print" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                                <input id="Idnew2" type="hidden" size="1" name="Idnew2" runat="server">
                            </td>
                        </tr>
                    </table>
                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                </td>
            </tr>
        </table>

        <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AllowPaging="True" AutoGenerateColumns="False">
                        <HeaderStyle CssClass="head_navy" />
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <Columns>
                            <asp:TemplateColumn HeaderText="序號"></asp:TemplateColumn>
                            <asp:BoundColumn DataField="Name" HeaderText="姓名">
                                <HeaderStyle Width="40px"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="IDNO" HeaderText="身分證號碼"></asp:BoundColumn>
                            <asp:BoundColumn DataField="DISTNAME" HeaderText="轄區中心"></asp:BoundColumn>
                            <asp:BoundColumn DataField="planname" HeaderText="訓練計畫"></asp:BoundColumn>
                            <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構"></asp:BoundColumn>
                            <asp:BoundColumn DataField="Classname" HeaderText="班級名稱"></asp:BoundColumn>
                            <asp:BoundColumn DataField="Sfdate" HeaderText="受訓期間"></asp:BoundColumn>
                            <asp:BoundColumn DataField="WEEKS" HeaderText="上課時間"></asp:BoundColumn>
                            <asp:BoundColumn DataField="DISTNAME2" HeaderText="重複參訓-轄區中心"></asp:BoundColumn>
                            <asp:BoundColumn DataField="planname2" HeaderText="重複參訓-訓練計畫"></asp:BoundColumn>
                            <asp:BoundColumn DataField="OrgName2" HeaderText="重複參訓-訓練機構"></asp:BoundColumn>
                            <asp:BoundColumn DataField="Classname2" HeaderText="重複參訓-班級名稱"></asp:BoundColumn>
                            <asp:BoundColumn DataField="Sfdate2" HeaderText="重複參訓-受訓期間"></asp:BoundColumn>
                            <asp:BoundColumn DataField="WEEKS2" HeaderText="重複參訓-上課時間"></asp:BoundColumn>
                            <asp:BoundColumn DataField="STUDSTATUS_N1" HeaderText="訓練狀態"></asp:BoundColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                </td>
            </tr>
        </table>
        <table class="font" id="DataGridTable2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AllowPaging="True" AutoGenerateColumns="False">
                        <HeaderStyle CssClass="head_navy" />
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <Columns>
                            <asp:TemplateColumn HeaderText="序號"></asp:TemplateColumn>
                            <asp:BoundColumn DataField="Name3" HeaderText="姓名">
                                <HeaderStyle Width="40px"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="IDNO3" HeaderText="身分證號碼"></asp:BoundColumn>
                            <asp:BoundColumn DataField="DISTNAME3" HeaderText="轄區中心"></asp:BoundColumn>
                            <asp:BoundColumn DataField="planname3" HeaderText="訓練計畫"></asp:BoundColumn>
                            <asp:BoundColumn DataField="OrgName3" HeaderText="訓練機構"></asp:BoundColumn>
                            <asp:BoundColumn DataField="Classname3" HeaderText="班級名稱"></asp:BoundColumn>
                            <asp:BoundColumn DataField="Sfdate3" HeaderText="受訓期間"></asp:BoundColumn>
                            <asp:BoundColumn DataField="ExamName" HeaderText="技能檢定"></asp:BoundColumn>
                            <asp:BoundColumn DataField="WEEKS" HeaderText="上課時間"></asp:BoundColumn>
                            <asp:BoundColumn DataField="StudStatus2" HeaderText="訓練&lt;BR&gt;狀態"></asp:BoundColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <uc1:PageControler ID="PageControler2" runat="server"></uc1:PageControler>
                </td>
            </tr>
        </table>
        <table class="font" id="DataGridTable4" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td>
                    <asp:DataGrid ID="Datagrid4" runat="server" Width="100%" CssClass="font" AllowPaging="True" AutoGenerateColumns="False">
                        <HeaderStyle CssClass="head_navy" />
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <Columns>
                            <asp:TemplateColumn HeaderText="序號"></asp:TemplateColumn>
                            <asp:BoundColumn DataField="Name4" HeaderText="姓名">
                                <HeaderStyle Width="40px"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="IDNO4" HeaderText="身分證號碼"></asp:BoundColumn>
                            <asp:BoundColumn DataField="DISTNAME4" HeaderText="轄區中心"></asp:BoundColumn>
                            <asp:BoundColumn DataField="planname4" HeaderText="訓練計畫"></asp:BoundColumn>
                            <asp:BoundColumn DataField="OrgName4" HeaderText="訓練機構"></asp:BoundColumn>
                            <asp:BoundColumn DataField="Classname4" HeaderText="班級名稱"></asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="受訓期間"></asp:TemplateColumn>
                            <asp:BoundColumn DataField="ExamName" HeaderText="技能檢定"></asp:BoundColumn>
                            <asp:BoundColumn DataField="WEEKS" HeaderText="上課時間"></asp:BoundColumn>
                            <asp:BoundColumn DataField="StudStatus2" HeaderText="訓練&lt;BR&gt;狀態"></asp:BoundColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <uc1:PageControler ID="PageControler4" runat="server"></uc1:PageControler>
                </td>
            </tr>
        </table>
        <table class="font" id="eMeng" style="border-right: #455690 1px solid; border-top: #a6b4cf 1px solid; visibility: visible; border-left: #a6b4cf 1px solid; width: 376px; border-bottom: #455690 1px solid; height: 248px; background-color: #c9d3f3" cellspacing="1" cellpadding="1" width="376" border="0" runat="server">
            <tr>
                <td background="../../images/MSNTitle.gif">
                    <table class="font" id="Table4" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td><strong><font color="#0000ff">問題轉入資料訊息：</font></strong> </td>
                            <td style="cursor: hand" onclick="closeDiv();" align="center" width="15">
                                <img src="../../images/CloseMsn.gif" width="13" height="13" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td style="border-right: #b9c9ef 1px solid; padding-right: 10px; border-top: #728eb8 1px solid; padding-left: 10px; font-size: 12px; padding-bottom: 10px; border-left: #728eb8 1px solid; width: 100%; color: #1f336b; padding-top: 15px; border-bottom: #b9c9ef 1px solid; height: 100%" align="center" background="../../images/MsnBack.gif" colspan="1" height="100">
                    <asp:DataGrid ID="Datagrid3" runat="server" Width="100%" CssClass="font" AllowPaging="True" AutoGenerateColumns="False" Height="208px">
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <ItemStyle BackColor="#FFFFFF" />
                        <Columns>
                            <asp:BoundColumn DataField="Index" HeaderText=" 第幾筆錯誤"></asp:BoundColumn>
                            <asp:BoundColumn DataField="IDNO" HeaderText="身分證號碼"></asp:BoundColumn>
                            <asp:BoundColumn DataField="Reason" HeaderText="原因"></asp:BoundColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
