<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_01_001_add.aspx.vb" Inherits="WDAIIP.CP_01_001_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>實地訪查紀錄表</title>
    <meta content="True" name="vs_snapToGrid">
    <meta content="True" name="vs_showGrid">
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function chkdata() {
            var msg = '';
            if (document.form1.OCIDValue1.value == '') msg += '請選擇職類\n';
            if (document.form1.ApplyDate.value == '') msg += '請輸入訪查日期\n';
            if (document.form1.ApplyDate.value != '' && !checkDate(document.form1.ApplyDate.value)) msg += '訪查日期的時間格式不正確\n';
            if (document.form1.AuthCount.value != '' && !isUnsignedInt(document.form1.AuthCount.value)) msg += '核定人數必須為數字\n';
            if (document.form1.TurthCount.value != '' && !isUnsignedInt(document.form1.TurthCount.value)) msg += '實到人數必須為數字\n';
            if (document.form1.TurnoutCount.value != '' && !isUnsignedInt(document.form1.TurnoutCount.value)) msg += '請假人數必須為數字\n';
            if (document.form1.TruancyCount.value != '' && !isUnsignedInt(document.form1.TruancyCount.value)) msg += '缺(曠)課人數必須為數字\n';
            if (document.form1.RejectCount.value != '' && !isUnsignedInt(document.form1.RejectCount.value)) msg += '退訓人數必須為數字\n';
            if (!isChecked(document.form1.Data1)) msg += '請選擇訓練日誌1的選項\n';
            if (!isChecked(document.form1.Data2)) msg += '請選擇學員名冊1的選項\n';
            if (!isChecked(document.form1.Data3)) msg += '請選擇學員簽到(退)表1的選項\n';
            if (!isChecked(document.form1.Data4)) msg += '請選擇課程表1的選項\n';
            if (!isChecked(document.form1.Data5)) msg += '請選擇請假單1的選項\n';
            if (!isChecked(document.form1.Data6)) msg += '請選擇退訓申請單1的選項\n';
            if (!isChecked(document.form1.Data7)) msg += '請選擇勞保投保資料1的選項\n';
            if (!isChecked(document.form1.Item1_1)) msg += '請回答有無週(月)課程表?\n';
            if (!isChecked(document.form1.Item1_2)) msg += '請回答是否依課程表授課?\n';
            if (document.form1.Item1_3.value == '') msg += '請輸入課目或課題為何?\n';
            if (!isChecked(document.form1.Item2_1)) msg += '請回答教學(訓練)日誌是否確實填寫?\n';
            if (!isChecked(document.form1.Item2_2)) msg += '請回答有否按時呈主管核閱?\n';
            if (!isChecked(document.form1.Item3_1)) msg += '請回答教師(職業訓練師)與助教姓名?是否與計畫相符?\n';
            if (document.form1.Item3_1Tech.value == '') msg += '請輸入講師姓名?\n';
            if (!isChecked(document.form1.Item3_2)) msg += '請回答學員學習情況是否良好?\n';
            if (!isChecked(document.form1.Item4_1)) msg += '請回答教學環境(教室或訓練工場)是否整齊?\n';
            if (!isChecked(document.form1.Item5_1)) msg += '請回答有無具體輔導活動或其他事項?\n';
            if (!isChecked(document.form1.Item6_1)) msg += '請回答職業訓練生活津貼是否依規定申請並發放?\n';
            if (document.form1.CurseName.value == '') msg += '請輸入培訓單位人員姓名?\n';
            if (document.form1.VisitorName.value == '') msg += '請輸入訪視人員姓名?\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function Turnoutchk() {
            if (!isUnsignedInt(document.form1.TurnoutCount.value)) {
                alert('請假人數必須為數字\n');
                document.form1.TurnoutCount.value = '';
                return false;
            }
            if (document.form1.TurnoutCount.value == '0') {
                //document.form1.Data5_0.checked=true;
                document.form1.Data5[0].checked = true;
                document.form1.Data5Note.value = document.getElementById('Data5Note').value.replace('無人請假', '') + '無人請假';
            }
        }

        function Rejectchk() {
            if (!isUnsignedInt(document.form1.RejectCount.value)) {
                alert('退訓人數必須為數字\n');
                document.form1.RejectCount.value = '';
                return false;
            }
            if (document.form1.RejectCount.value == '0') {
                //document.form1.Data6_0.checked=true;
                document.form1.Data6[0].checked = true;
                document.form1.Data6Note.value = document.getElementById('Data6Note').value.replace('無人退訓', '') + '無人退訓';
            }
        }

        function callCalendar() {
            /* if (document.form1.ApplyDate.value!=''){
            //alert('1 '+ document.form1.ApplyDate.value);
            openCalendar('ApplyDate',document.form1.ApplyDate.value,document.form1.ApplyDate.value,document.form1.ApplyDate.value,'','Button5');
            }
            else if (document.form1.Enddate.value!=''){
            //alert('2 '+ document.form1.end_date.value);
            openCalendar('ApplyDate',document.form1.end_date.value,document.form1.end_date.value,document.form1.end_date.value,'','Button5');
            }
            else{
            show_calendar('ApplyDate','','','CY/MM/DD');
            }   */
            //openCalendar('ApplyDate', document.form1.StartDate.value, document.form1.EndDate.value, document.form1.NowDate.value, '', 'Button5');
            openCalendar('ApplyDate', $('#StartDate').val(), $('#EndDate').val(), $('#NowDate').val(), '', 'Button5');
        }
    </script>
    <style type="text/css">
        .style1 { font-size: 12px; color: Black; line-height: 22px; background-color: #e9f1fe; padding: 2px; height: 31px; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <%--
                    <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>首頁&gt;&gt;訓練查核與績效管理&gt;&gt;統計分析&gt;&gt;<font color="#990000">實地訪查紀錄表</font></td>
                        </tr>
                    </table>
                    --%>
                    <table class="table_sch" id="Table3" cellspacing="1" cellpadding="1" border="0" width="100%">
                        <tr>
                            <td class="bluecol" width="20%">機構</td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini"><input id="RIDValue" type="hidden" name="Hidden1" runat="server">
                                <input id="EndDate" type="hidden" name="EndDate" runat="server"><input id="StartDate" type="hidden" name="StartDate" runat="server"><input id="NowDate" type="hidden" name="NowDate" runat="server"><br>
                                <span id="HistoryList2" style="display: none; position: absolute">
                                    <asp:Table ID="HistoryRID" runat="server" Width="30%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">職類/班別</td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input id="Button3" onclick="javascript: window.open('../CP_01_ch.aspx?RID=' + document.form1.RIDValue.value, '', 'width=540,height=520,location=0,status=0,menubar=0,scrollbars=1,resizable=0');" type="button" value="..." name="Button3" runat="server" class="button_b_Mini"><input id="TMIDValue1" type="hidden" name="Hidden1" runat="server"><input id="OCIDValue1" type="hidden" name="Hidden2" runat="server"><br>
                                <span id="HistoryList" style="display: none; left: 270px; position: absolute">
                                    <asp:Table ID="HistoryTable" runat="server" Width="30%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訪查時間</td>
                            <td class="whitecol">
                                <asp:TextBox ID="ApplyDate" runat="server" Width="15%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="callCalendar();" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                <asp:Button ID="Button5" runat="server" Text="查詢出勤" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td width="100%" colspan="2">
                                <table class="font" id="Table4" cellspacing="1" cellpadding="1" border="0" width="100%">
                                    <tr>
                                        <td colspan="2" class="bluecol">
                                            <div align="center">學員出缺勤狀況</div>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">核定人數:
                                            <asp:TextBox ID="AuthCount" runat="server" Width="20%"></asp:TextBox>人
                                            <div>實到人數:<asp:TextBox ID="TurthCount" runat="server" Width="20%"></asp:TextBox>人</div>
                                            <div>請假人數:<asp:TextBox ID="TurnoutCount" runat="server" Width="20%"></asp:TextBox>人</div>
                                        </td>
                                        <td class="whitecol">缺(曠)課人數:
                                            <asp:TextBox ID="TruancyCount" runat="server" Width="20%"></asp:TextBox>人
                                            <div>退訓人數:<asp:TextBox ID="RejectCount" runat="server" Width="20%"></asp:TextBox>人</div>
                                            <div>點名未到課學員，另以電話抽訪</div>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <table class="table_sch" id="Table6" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td class="bluecol">
                                            <div align="center">書面資料</div>
                                        </td>
                                        <td class="bluecol"></td>
                                        <td class="bluecol">攜回影本</td>
                                        <td class="bluecol">說明</td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">一、訓練日誌</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Data1" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                                <asp:ListItem Value="1">備齊</asp:ListItem>
                                                <asp:ListItem Value="2">未備</asp:ListItem>
                                                <asp:ListItem Value="3">部份備有</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="DataCopy1" runat="server" MaxLength="50" Width="50%"></asp:TextBox></td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="Data1Note" runat="server" MaxLength="100" TextMode="MultiLine"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">二<font>、學員名冊</font></td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Data2" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                                <asp:ListItem Value="1">備齊</asp:ListItem>
                                                <asp:ListItem Value="2">未備</asp:ListItem>
                                                <asp:ListItem Value="3">部份備有</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="DataCopy2" runat="server" MaxLength="50" Width="50%"></asp:TextBox></td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="Data2Note" runat="server" MaxLength="100" TextMode="MultiLine"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">三<font>、學員簽到(退)表</font></td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Data3" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                                <asp:ListItem Value="1">備齊</asp:ListItem>
                                                <asp:ListItem Value="2">未備</asp:ListItem>
                                                <asp:ListItem Value="3">部份備有</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="DataCopy3" runat="server" MaxLength="50" Width="50%"></asp:TextBox></td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="Data3Note" runat="server" MaxLength="100" TextMode="MultiLine"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">四<font>、課程表</font></td>
                                        <td class="whitecol">
                                            <font>
                                                <asp:RadioButtonList ID="Data4" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                                    <asp:ListItem Value="1">備齊</asp:ListItem>
                                                    <asp:ListItem Value="2">未備</asp:ListItem>
                                                    <asp:ListItem Value="3">部份備有</asp:ListItem>
                                                </asp:RadioButtonList>
                                            </font>
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="DataCopy4" runat="server" MaxLength="50" Width="50%"></asp:TextBox></td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="Data4Note" runat="server" MaxLength="100" TextMode="MultiLine"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">五<font>、請假單</font></td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Data5" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                                <asp:ListItem Value="1">備齊</asp:ListItem>
                                                <asp:ListItem Value="2">未備</asp:ListItem>
                                                <asp:ListItem Value="3">部份備有</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="DataCopy5" runat="server" MaxLength="50" Width="50%"></asp:TextBox></td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="Data5Note" runat="server" MaxLength="100" TextMode="MultiLine"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">六<font>、退訓申請單</font></td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Data6" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                                <asp:ListItem Value="1">備齊</asp:ListItem>
                                                <asp:ListItem Value="2">未備</asp:ListItem>
                                                <asp:ListItem Value="3">部份備有</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="DataCopy6" runat="server" MaxLength="50" Width="50%"></asp:TextBox></td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="Data6Note" runat="server" MaxLength="100" TextMode="MultiLine"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">七<font>、勞保投保資料</font></td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Data7" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                                <asp:ListItem Value="1">備齊</asp:ListItem>
                                                <asp:ListItem Value="2">未備</asp:ListItem>
                                                <asp:ListItem Value="3">部份備有</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="DataCopy7" runat="server" MaxLength="50" Width="50%"></asp:TextBox></td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="Data7Note" runat="server" MaxLength="100" TextMode="MultiLine"></asp:TextBox></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <table class="table_sch" id="Table5" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td align="center" class="bluecol">項次</td>
                                        <td align="center" class="bluecol">訪查項目</td>
                                        <td align="center" class="bluecol">訪查實況</td>
                                        <td align="center" class="bluecol">處理情形</td>
                                        <td align="center" class="bluecol">備註</td>
                                    </tr>
                                    <tr>
                                        <td rowspan="3" class="whitecol" align="center">一</td>
                                        <td class="whitecol">1.有無週(月)課程表?</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Item1_1" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                                <asp:ListItem Value="Y">有</asp:ListItem>
                                                <asp:ListItem Value="N">無</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td rowspan="3" class="whitecol">
                                            <asp:TextBox ID="Item1Pros" runat="server" TextMode="MultiLine" Width="60%"></asp:TextBox></td>
                                        <td rowspan="3" class="whitecol">
                                            <asp:TextBox ID="Item1Note" runat="server" TextMode="MultiLine" Width="60%"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">2.是否依課程表授課?</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Item1_2" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                                <asp:ListItem Value="Y">是</asp:ListItem>
                                                <asp:ListItem Value="N">否</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">3.課目或課題為何?</td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="Item1_3" runat="server" Width="60%"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td rowspan="2" class="whitecol" align="center">二</td>
                                        <td class="whitecol">1.教學(訓練)日誌是否確實填寫?</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Item2_1" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                                <asp:ListItem Value="Y">是</asp:ListItem>
                                                <asp:ListItem Value="N">否</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td rowspan="2" class="td_light">
                                            <asp:TextBox ID="Item2Pros" runat="server" TextMode="MultiLine" Width="60%"></asp:TextBox></td>
                                        <td rowspan="2" class="td_light">
                                            <asp:TextBox ID="Item2Note" runat="server" TextMode="MultiLine" Width="60%"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">2.有否按時呈主管核閱?</td>
                                        <td class="whitecol">
                                            <font>
                                                <asp:RadioButtonList ID="Item2_2" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                                    <asp:ListItem Value="Y">有</asp:ListItem>
                                                    <asp:ListItem Value="N">否</asp:ListItem>
                                                </asp:RadioButtonList>
                                            </font>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td rowspan="2" class="whitecol" align="center">三</td>
                                        <td class="whitecol">1.教師(職業訓練師)與助教姓名?是否與計畫相符?</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Item3_1" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                                <asp:ListItem Value="Y">是</asp:ListItem>
                                                <asp:ListItem Value="N">否</asp:ListItem>
                                            </asp:RadioButtonList>
                                            <div>講師:<asp:TextBox ID="Item3_1Tech" runat="server" Width="50%"></asp:TextBox></div>
                                            <div>助教:<asp:TextBox ID="Item3_1Tutor" runat="server" Width="50%"></asp:TextBox></div>
                                        </td>
                                        <td rowspan="2" class="whitecol">
                                            <asp:TextBox ID="Item3Pros" runat="server" TextMode="MultiLine"></asp:TextBox></td>
                                        <td rowspan="2" class="whitecol">
                                            <asp:TextBox ID="Item3Note" runat="server" TextMode="MultiLine"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">2.學員學習情況是否良好?</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Item3_2" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                                <asp:ListItem Value="Y">是</asp:ListItem>
                                                <asp:ListItem Value="N">否</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol" align="center">四</td>
                                        <td class="whitecol">教學環境(教室或訓練工場)是否整齊?</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Item4_1" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                                <asp:ListItem Value="Y">是</asp:ListItem>
                                                <asp:ListItem Value="N">否</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="Item4Pros" runat="server" TextMode="MultiLine" Height="60px"></asp:TextBox></td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="Item4Note" runat="server" TextMode="MultiLine" Height="60px"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol" align="center">五</td>
                                        <td class="whitecol">有無具體輔導活動或其他事項?</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Item5_1" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                                <asp:ListItem Value="Y">是</asp:ListItem>
                                                <asp:ListItem Value="N">否</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="Item5Pros" runat="server" TextMode="MultiLine" Height="60px"></asp:TextBox></td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="Item5Note" runat="server" TextMode="MultiLine" Height="60px"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol" align="center">六</td>
                                        <td class="whitecol">職業訓練生活津貼是否依規定申請並發放?</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Item6_1" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                                <asp:ListItem Value="Y">是</asp:ListItem>
                                                <asp:ListItem Value="N">否</asp:ListItem>
                                                <asp:ListItem Value="D">不符合發放</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td class="whitecol">
                                            <div style="margin-top: 3px; margin-bottom: 3px">請領職業訓練生活津貼</div>
                                            <div style="margin-top: 3px; margin-bottom: 3px">共<asp:TextBox ID="Item6Count1" runat="server" Width="30%"></asp:TextBox>人。</div>
                                            <div style="margin-top: 3px; margin-bottom: 3px"></div>
                                            <div style="margin-top: 3px; margin-bottom: 3px">到勤狀況及身分查驗:</div>
                                            <div style="margin-top: 3px; margin-bottom: 3px">無異狀<asp:TextBox ID="Item6Count2" runat="server" Width="30%"></asp:TextBox>人。</div>
                                            <div style="margin-top: 3px; margin-bottom: 3px"></div>
                                            <div style="margin-top: 3px; margin-bottom: 3px">需繼續追蹤<asp:TextBox ID="Item6Count3" runat="server" Width="30%"></asp:TextBox>人、</div>
                                            <div style="margin-top: 3px; margin-bottom: 3px">姓名:<asp:TextBox ID="Item6Names" runat="server" Width="50%"></asp:TextBox></div>
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="Item6Note" runat="server" TextMode="MultiLine" Height="150px"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol" align="center">七</td>
                                        <td class="whitecol">參訓學員反映意見及問題</td>
                                        <td colspan="3" class="whitecol">
                                            <asp:TextBox ID="Item7Note" runat="server" TextMode="MultiLine" Width="60%"></asp:TextBox></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <table class="font" id="table7" cellspacing="1" width="100%" border="0">
                                    <tr>
                                        <td class="bluecol" width="20%">培訓姓名</td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="CurseName" runat="server" MaxLength="10" Width="40%"></asp:TextBox></td>
                                        <td class="bluecol" width="20%">訪視姓名</td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="VisitorName" runat="server" MaxLength="10" Width="40%"></asp:TextBox></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                    <div align="center" class="whitecol">
                        <asp:Button ID="Button1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                        <input id="Button4" type="button" value="回查詢頁面" name="Button4" runat="server" class="button_b_M">
                    </div>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
