<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_01_001_add_99.aspx.vb" Inherits="WDAIIP.CP_01_001_add_99" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>實地訪查紀錄表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" type="text/javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script language="javascript" type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function check(obj, obj2) {
            if (document.getElementById(obj).checked) {
                document.getElementById(obj2).checked = false;
            }
        }

        function chkdata() {
            var msg = '';
            //            debugger;
            //            alert('debuggeor');
            if (document.form1.OCIDValue1.value == '') msg += '請選擇職類\n';
            if (document.form1.ApplyDate.value == '') msg += '請輸入訪查日期\n';
            if (document.form1.ApplyDate.value != '' && !checkDate(document.form1.ApplyDate.value)) msg += '訪查日期的時間格式不正確\n';
            if (document.form1.Tplanid.value != '15') {
                if (document.form1.AuthCount.value != '' && !isUnsignedInt(document.form1.AuthCount.value)) msg += '核定人數必須為數字\n';
                if (document.form1.TurthCount.value != '' && !isUnsignedInt(document.form1.TurthCount.value)) msg += '實到人數必須為數字\n';
                if (document.form1.TurnoutCount.value != '' && !isUnsignedInt(document.form1.TurnoutCount.value)) msg += '請假人數必須為數字\n';
                if (document.form1.TruancyCount.value != '' && !isUnsignedInt(document.form1.TruancyCount.value)) msg += '缺(曠)課人數必須為數字\n';
                if (document.form1.RejectCount.value != '' && !isUnsignedInt(document.form1.RejectCount.value)) msg += '退訓人數必須為數字\n';
                if (document.form1.AheadjobCount.value != '' && !isUnsignedInt(document.form1.AheadjobCount.value)) msg += '提前就業人數必須為數字\n';
            }
            //else
            //	{
            //if(document.form1.StudyticketCount.value!='' && !isUnsignedInt(document.form1.StudyticketCount.value)) msg+='學習券人數必須為數字\n';
            //	}
            if (!isChecked(document.form1.Data1)) msg += '請選擇訓練日誌的選項\n';
            if (!isChecked(document.form1.Data3)) msg += '請選擇學員簽到(退)表的選項\n';
            if (!isChecked(document.form1.Data5)) msg += '請選擇請假單的選項\n';
            if (!isChecked(document.form1.Data6)) msg += '請選擇退訓/提前就業申請表的選項\n';
            //if(!isChecked(document.form1.Data7)) msg+='請選擇勞保加/退保明細表的選項\n';
            //if(!isChecked(document.form1.Data9)) msg+='請選擇學員書籍(講義)、材料領用表的選項\n';
            if (!isChecked(document.form1.Data10)) msg += '請選擇職訓生活津貼補助印領清冊的選項\n';
            //if(!isChecked(document.form1.Data11)) msg+='請選擇學員服務手冊或權利義務公告相關文件的選項\n';
            //if(isEmpty(document.form1.Data11)) msg+='請選擇出缺勤狀況\n';
            if (!isChecked(document.form1.Item1_1)) msg += '請回答有無週(月)課程表?\n';
            if (!isChecked(document.form1.Item1_2)) msg += '請回答是否依課程表授課?\n';
            if (document.form1.Item1_3.value == '') msg += '請輸入課目或課題為何?\n';
            if (!isChecked(document.form1.Item3_1)) msg += '請回答教師與助教是否與計畫相符?\n';
            if (document.form1.Item3_1Tech.value == '') msg += '請輸入講師姓名?\n';
            if (!isChecked(document.form1.Item19)) msg += '請回答有無書籍(講義)領用表?\n';
            if (!isChecked(document.form1.Item20)) msg += '請回答有無材料領用表?\n';
            if (!isChecked(document.form1.Item21)) msg += '請回答訓練設施設備是否依契約提供學員使用?\n';
            if (!isChecked(document.form1.Item2_1)) msg += '請回答教學(訓練)日誌是否確實填寫?\n';
            if (!isChecked(document.form1.Item2_2)) msg += '請回答有否按時呈主管核閱?\n';
            if (!isChecked(document.form1.Item23)) msg += '請回答學員生活、就業輔導與管理機制是否依契約規範辦理?\n';
            if (!isChecked(document.form1.Item24)) msg += '請回答是否依契約規範提供學員問題反應申訴管道?\n';
            if (!isChecked(document.form1.Item25)) msg += '請回答是否為參訓學員辦理勞工保險加退保?\n';
            if (!isChecked(document.form1.Item26)) msg += '請回答是否依契約規範公告學員權益義務或編製參訓學員服務手冊?\n';
            //if(!isChecked(document.form1.Item27)) msg+='請回答是否對已結訓學員是否依規定發給研習證書?\n';
            //if(isEmpty(document.form1.Item28)) msg+='請回答有無自費參訓學員?\n';
            //if(!isChecked(document.form1.Item28_2)) msg+='請回答訓練費用是否繳交主辦單位?\n';
            if (!isChecked(document.form1.Item6_1)) msg += '請回答職業訓練生活津貼是否依規定申請並核發?\n';
            //if(!isChecked(document.form1.Item29)) msg+='請回答培訓單位無巧立名目強制收取費用?\n';
            //if(!isChecked(document.form1.Item30)) msg+='請回答職業訓練機構是否依規定懸掛設立許可證書?\n';
            if (document.form1.CurseName.value == '') msg += '請輸入培訓單位人員姓名?\n';
            if (document.form1.VisitorName.value == '') msg += '請輸入訪視人員姓名?\n';
            if (document.getElementById('D1c').checked) { if (document.form1.DataCopy1.value == '') msg += '請輸入第1題的附件內容?\n'; }
            if (document.getElementById('D2c').checked) { if (document.form1.DataCopy3.value == '') msg += '請輸入第2題的附件內容?\n'; }
            if (document.getElementById('D3c').checked) { if (document.form1.DataCopy5.value == '') msg += '請輸入第3題的附件內容?\n'; }
            if (document.getElementById('D4c').checked) { if (document.form1.DataCopy6.value == '') msg += '請輸入第4題的附件內容?\n'; }
            //if(document.getElementById('D5c').checked){if(document.form1.DataCopy7.value=='') msg+='請輸入第5題的附件內容?\n';}
            //if(document.getElementById('D6c').checked){if(document.form1.DataCopy9.value=='') msg+='請輸入第6題的附件內容?\n';}
            if (document.getElementById('D7c').checked) { if (document.form1.DataCopy10.value == '') msg += '請輸入第5題的附件內容?\n'; }
            //if(document.getElementById('D8c').checked){if(document.form1.DataCopy11.value=='') msg+='請輸入第8題的附件內容?\n';}
            var cst_dc3_1 = 0;
            var cst_dc3_2 = 1;
            var cst_dc3_3 = 2;
            if (document.getElementsByName('D3c3').length > 3) {
                cst_dc3_1 = 1; //cst_dc3_(D3c3,D4c3) 3項
                cst_dc3_2 = 2;
                cst_dc3_3 = 3;
            }
            var cst_dc4_1 = 0;
            var cst_dc4_2 = 1;
            var cst_dc4_3 = 2;
            var cst_dc4_4 = 3;
            if (document.getElementsByName('D7c3').length > 4) {
                cst_dc4_1 = 1; //cst_dc4_ 4項
                cst_dc4_2 = 2;
                cst_dc4_3 = 3;
                cst_dc4_4 = 4;
            }
            // alert(document.getElementsByName('D3c3').item(2).checked);
            // alert(document.getElementsByName('D7c3').item(3).checked);
            if (document.getElementsByName('D3c3').item(cst_dc3_3).checked) { if (document.form1.Data5Note.value == '') msg += '請輸入第3題的其他說明?\n'; }
            if (document.getElementsByName('D4c3').item(cst_dc3_3).checked) { if (document.form1.Data6Note.value == '') msg += '請輸入第4題的其他說明?\n'; }
            //if(document.getElementsByName('D5c3').item(3).checked){if(document.form1.Data7Note.value=='') msg+='請輸入第5題的其他說明?\n';}				
            //if(document.getElementsByName('D6c3').item(3).checked){if(document.form1.Data9Note.value=='') msg+='請輸入第6題的其他說明?\n';}				
            if (document.getElementsByName('D7c3').item(cst_dc4_4).checked) { if (document.form1.Data10Note.value == '') msg += '請輸入第5題的其他說明?\n'; }
            //if(document.getElementsByName('D8c3').item(3).checked){if(document.form1.Data11Note.value=='') msg+='請輸入第8題的其他說明?\n';}											
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function callCalendar() {
            openCalendar('ApplyDate', document.form1.StartDate.value, document.form1.EndDate.value, document.form1.NowDate.value, '', 'Button5');
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <%--
                    <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>首頁&gt;&gt;訓練查核與績效管理&gt;&gt;統計分析&gt;&gt;實地訪查紀錄表</td>
                        </tr>
                    </table>
                    --%>
                    <table class="font" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td width="20%" class="bluecol">機構</td>
                            <td width="80%" class="whitecol">
                                <asp:TextBox ID="center" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox>
                                <input id="Button2" type="button" value="..." name="Button2" runat="server">
                                <input id="RIDValue" type="hidden" name="Hidden1" runat="server">
                                <input id="EndDate" type="hidden" name="EndDate" runat="server">
                                <input id="StartDate" type="hidden" name="StartDate" runat="server">
                                <input id="NowDate" type="hidden" name="NowDate" runat="server">
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">職類/班別</td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input id="Button3" onclick="javascript: window.open('../CP_01_ch.aspx?RID=' + document.form1.RIDValue.value, '', 'width=540,height=520,location=0,status=0,menubar=0,scrollbars=1,resizable=0');" type="button" value="..." name="Button3" runat="server">
                                <input id="TMIDValue1" type="hidden" name="Hidden1" runat="server"><input id="OCIDValue1" type="hidden" name="Hidden2" runat="server">
                                <span id="HistoryList" style="position: absolute; display: none; left: 28%">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">訪查時間</td>
                            <td class="whitecol"><span id="span01" runat="server">
                                <asp:TextBox ID="ApplyDate" runat="server" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= ApplyDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span></td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <table class="font" id="Table6" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                    <tr align="center">
                                        <td class="bluecol" width="20%">書面資料</td>
                                        <td class="bluecol" width="25%"></td>
                                        <td class="bluecol" width="20%">佐證資料</td>
                                        <td class="bluecol" width="35%">說明</td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">1.教學(訓練)日誌</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Data1" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">備齊</asp:ListItem>
                                                <asp:ListItem Value="3">部份備有</asp:ListItem>
                                                <asp:ListItem Value="2">未備</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td class="whitecol">
                                            <div>
                                                <asp:CheckBox ID="D1c" runat="server"></asp:CheckBox>如附件&nbsp;&nbsp;
                                                <asp:TextBox ID="DataCopy1" runat="server" Width="50%"></asp:TextBox>
                                            </div>
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="Data1Note" runat="server" Width="90%" MaxLength="100" TextMode="MultiLine"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">2.學員簽到(退)表</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Data3" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">備齊</asp:ListItem>
                                                <asp:ListItem Value="3">部份備有</asp:ListItem>
                                                <asp:ListItem Value="2">未備</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td class="whitecol">
                                            <asp:CheckBox ID="D2c" runat="server"></asp:CheckBox>如附件 &nbsp;
                                            <asp:TextBox ID="DataCopy3" runat="server" Width="50%"></asp:TextBox>
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="Data3Note" runat="server" Width="90%" MaxLength="100" TextMode="MultiLine"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">3.請假單</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Data5" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">備齊</asp:ListItem>
                                                <asp:ListItem Value="3">部份備有</asp:ListItem>
                                                <asp:ListItem Value="2">未備</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td class="whitecol">
                                            <asp:RadioButton ID="D3c" runat="server"></asp:RadioButton>如附件
                                            <asp:TextBox ID="DataCopy5" runat="server" Width="50%"></asp:TextBox><br>
                                            <asp:RadioButton ID="D3c2" runat="server"></asp:RadioButton>免提供
                                        </td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="D3c3" runat="server">
                                                <asp:ListItem Value="1">攜回影本</asp:ListItem>
                                                <asp:ListItem Value="2">無學員請假情形，故免提供</asp:ListItem>
                                                <asp:ListItem Value="3">其他(請說明)：</asp:ListItem>
                                            </asp:RadioButtonList>
                                            <asp:TextBox ID="Data5Note" runat="server" Width="90%"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">4.退訓／提前就業申請表</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Data6" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">備齊</asp:ListItem>
                                                <asp:ListItem Value="3">部份備有</asp:ListItem>
                                                <asp:ListItem Value="2">未備</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td class="whitecol">
                                            <asp:RadioButton ID="D4c" runat="server"></asp:RadioButton>如附件
                                            <asp:TextBox ID="DataCopy6" runat="server" Width="50%"></asp:TextBox><br>
                                            <asp:RadioButton ID="D4c2" runat="server"></asp:RadioButton>免提供
                                        </td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="D4c3" runat="server">
                                                <asp:ListItem Value="1">攜回影本</asp:ListItem>
                                                <asp:ListItem Value="2">前次訪查已提供過，故免提供</asp:ListItem>
                                                <asp:ListItem Value="3">其他(請說明)：</asp:ListItem>
                                            </asp:RadioButtonList>
                                            <asp:TextBox ID="Data6Note" runat="server" Width="90%"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">5.職業訓練生活津貼補助印領清冊</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Data10" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">備齊</asp:ListItem>
                                                <asp:ListItem Value="3">部份備有</asp:ListItem>
                                                <asp:ListItem Value="2">未備</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td class="whitecol">
                                            <asp:RadioButton ID="D7c" runat="server"></asp:RadioButton>如附件
                                            <asp:TextBox ID="DataCopy10" runat="server" Width="50%"></asp:TextBox><br>
                                            <asp:RadioButton ID="D7c2" runat="server"></asp:RadioButton>免提供
                                        </td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="D7c3" runat="server">
                                                <asp:ListItem Value="1">攜回影本</asp:ListItem>
                                                <asp:ListItem Value="2">前次訪查已提供過，故免提供</asp:ListItem>
                                                <asp:ListItem Value="3">學習券計畫或無申請補助，故免提供</asp:ListItem>
                                                <asp:ListItem Value="4">其他(請說明)：</asp:ListItem>
                                            </asp:RadioButtonList>
                                            <asp:TextBox ID="Data10Note" runat="server" Width="90%"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">6.出缺勤狀況</td>
                                        <td id="notstudy3" colspan="3" runat="server" class="whitecol">核定：<asp:TextBox ID="AuthCount" runat="server" Width="10%"></asp:TextBox>人； 實到：<asp:TextBox ID="TurthCount" runat="server" Width="10%"></asp:TextBox>人； 請假：<asp:TextBox ID="TurnoutCount" runat="server" Width="10%"></asp:TextBox>人； 缺(曠)課：<asp:TextBox ID="TruancyCount" runat="server" Width="10%"></asp:TextBox>人；<br>
                                            退訓：<asp:TextBox ID="RejectCount" runat="server" Width="10%"></asp:TextBox>人； (含提前就業：<asp:TextBox ID="AheadjobCount" runat="server" Width="10%"></asp:TextBox>人)。 ※點名未到課學員，另以電話抽訪。
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <table class="font" id="Table5" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td width="36%" align="center" colspan="2" class="bluecol">訪查項目</td>
                                        <td width="20%" align="center" class="bluecol">訪查實況</td>
                                        <td width="22%" align="center" class="bluecol">處理情形</td>
                                        <td width="22%" align="center" class="bluecol">備註</td>
                                    </tr>
                                    <tr>
                                        <td rowspan="4" width="14%" class="bluecol">課程(師資)實施狀況</td>
                                        <td class="whitecol">1.有無週(月)課程表?</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Item1_1" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">有</asp:ListItem>
                                                <asp:ListItem Value="2">無</asp:ListItem>
                                                <asp:ListItem Value="3">免填</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td rowspan="4" class="whitecol">
                                            <asp:TextBox ID="Item1Pros" runat="server" Width="90%" TextMode="MultiLine" Height="260px"></asp:TextBox></td>
                                        <td rowspan="4" class="whitecol">
                                            <asp:TextBox ID="Item1Note" runat="server" Width="90%" TextMode="MultiLine" Height="260px"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">2.是否依課程表授課?</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Item1_2" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">是</asp:ListItem>
                                                <asp:ListItem Value="2">否</asp:ListItem>
                                                <asp:ListItem Value="3">免填</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">3.課目或課題為何?</td>
                                        <td class="whitecol">課目：<asp:TextBox ID="Item1_3" runat="server" Width="70%"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">教師與助教是否與計畫相符?</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Item3_1" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">是</asp:ListItem>
                                                <asp:ListItem Value="2">否</asp:ListItem>
                                                <asp:ListItem Value="3">免填</asp:ListItem>
                                            </asp:RadioButtonList>
                                            教師：<asp:TextBox ID="Item3_1Tech" runat="server" Width="70%"></asp:TextBox><br>
                                            助教：<asp:TextBox ID="Item3_1Tutor" runat="server" Width="70%"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td rowspan="3" class="bluecol">教材設施運用狀況</td>
                                        <td class="whitecol">1.有無書籍(講義)領用表?</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Item19" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">有</asp:ListItem>
                                                <asp:ListItem Value="2">無</asp:ListItem>
                                                <asp:ListItem Value="3">免填</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td rowspan="3" class="whitecol">
                                            <asp:TextBox ID="Item2Pros" runat="server" Width="90%" TextMode="MultiLine" Height="100px"></asp:TextBox></td>
                                        <td rowspan="3" class="whitecol">
                                            <asp:TextBox ID="Item2Note" runat="server" Width="90%" TextMode="MultiLine" Height="100px"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">2.有無材料領用表?</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Item20" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">有</asp:ListItem>
                                                <asp:ListItem Value="2">無</asp:ListItem>
                                                <asp:ListItem Value="3">免填</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">3.訓練設施設備是否依契約提供學員使用?</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Item21" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">是</asp:ListItem>
                                                <asp:ListItem Value="2">否</asp:ListItem>
                                                <asp:ListItem Value="3">免填</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td rowspan="6" class="bluecol">教務管理狀況</td>
                                        <td class="whitecol">1.教學(訓練)日誌是否確實填寫?</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Item2_1" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">是</asp:ListItem>
                                                <asp:ListItem Value="2">否</asp:ListItem>
                                                <asp:ListItem Value="3">免填</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td rowspan="6" class="whitecol">
                                            <asp:TextBox ID="Item3Pros" runat="server" Width="90%" TextMode="MultiLine" Height="200px"></asp:TextBox></td>
                                        <td rowspan="6" class="whitecol">
                                            <asp:TextBox ID="Item3Note" runat="server" Width="90%" TextMode="MultiLine" Height="200px"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">2.有否按時呈主管核閱?</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Item2_2" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">是</asp:ListItem>
                                                <asp:ListItem Value="2">否</asp:ListItem>
                                                <asp:ListItem Value="3">免填</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">3.學員生活、就業輔導與管理機制是否依契約挸範辦理?</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Item23" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">是</asp:ListItem>
                                                <asp:ListItem Value="2">否</asp:ListItem>
                                                <asp:ListItem Value="3">免填</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">4.是否依契約規範提供學員問題反應申訴管道?</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Item24" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">是</asp:ListItem>
                                                <asp:ListItem Value="2">否</asp:ListItem>
                                                <asp:ListItem Value="3">免填</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">5.是否為參訓學員辦理勞工保險加退保?</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Item25" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">是</asp:ListItem>
                                                <asp:ListItem Value="2">否</asp:ListItem>
                                                <asp:ListItem Value="3">免填</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">6.是否依契約規範公告學員權益義務管理狀況義務或編製參訓學員服務手冊?</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Item26" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">是</asp:ListItem>
                                                <asp:ListItem Value="2">否</asp:ListItem>
                                                <asp:ListItem Value="3">免填</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">費用(津貼)收核狀況</td>
                                        <td class="whitecol">1.職業訓練生活津貼是否依規定申請並核發?</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Item6_1" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">是</asp:ListItem>
                                                <asp:ListItem Value="2">否</asp:ListItem>
                                                <asp:ListItem Value="3">免填</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="Item4Pros" runat="server" Width="90%" TextMode="MultiLine" Height="100px"></asp:TextBox></td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="Item4Note" runat="server" Width="90%" TextMode="MultiLine" Height="100px"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">參訓學員反映意見及問題</td>
                                        <td colspan="4" class="whitecol">
                                            <asp:TextBox ID="Item7Note" runat="server" Width="90%" TextMode="MultiLine" Height="60px"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">綜合建議</td>
                                        <td colspan="2" class="whitecol">
                                            <asp:TextBox ID="Item31Note" runat="server" Width="90%" TextMode="MultiLine" Height="80px"></asp:TextBox></td>
                                        <td class="bluecol">缺失處理</td>
                                        <td class="whitecol">
                                            <input id="Item32_4" type="radio" value="4" name="Item32" runat="server">無缺失<br>
                                            <input id="Item32_1" type="radio" value="1" name="Item32" runat="server">限期改善，研提檢討報告<br>
                                            <input id="Item32_2" type="radio" value="2" name="Item32" runat="server">擇期進行訪查<br>
                                            <input id="Item32_3" type="radio" value="3" name="Item32" runat="server">其他(請說明)：<asp:TextBox ID="Item32Note" runat="server" Width="100px"></asp:TextBox>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <table class="font" id="table7" cellspacing="1" width="100%" border="0">
                                    <tr>
                                        <td class="bluecol" width="14%">培訓姓名</td>
                                        <td class="whitecol" width="42%">
                                            <asp:TextBox ID="CurseName" runat="server" MaxLength="10"></asp:TextBox></td>
                                        <td class="bluecol" width="22%">訪視姓名</td>
                                        <td class="whitecol" width="22%">
                                            <asp:TextBox ID="VisitorName" runat="server" MaxLength="10"></asp:TextBox></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                    <div align="center">
                        <asp:Button ID="Button1" runat="server" Text="儲存" CssClass="button_b_M"></asp:Button>&nbsp;
                        <input id="Button4" type="button" value="回查詢頁面" name="Button4" runat="server" class="button_b_M" />
                        <input id="Tplanid" type="hidden" name="Tplanid" runat="server" />
                    </div>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
