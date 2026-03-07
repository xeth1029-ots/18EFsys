<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_11_004_add17.aspx.vb" Inherits="WDAIIP.SD_11_004_add17" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>受訓學員意見調查表</title>
    <meta http-equiv="Content-Type" content="text/html; charset=big5">
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function ChkData1() {
            var msg = '';
            var cst_msgA2 = "勾選「其他」選項 必須填寫文字說明\n";
            var cst_t1Q = "學員基本資料>>(一)參加產投方案動機：";
            var cst_t2Q = "第一部份：(一)您獲得本次課程的訊息來源：";
            var cst_t3Q = "第一部份：(二)參加本次課程的主要原因：";
            var cst_t4Q = "第一部份：(三)選擇本訓練單位的主要原因：";
            var S1chk_CBLVAL = getCBLValue('S1chk');
            var A1chk_CBLVAL = getCBLValue('A1chk');
            var A2_RBLVAL = getRBLValue('A2');
            var A3_RBLVAL = getRBLValue('A3');
            var S16_NOTE = document.getElementById('S16_NOTE');
            var A1_10_NOTE = document.getElementById('A1_10_NOTE');
            var A2_7_NOTE = document.getElementById('A2_7_NOTE');
            var A3_5_NOTE = document.getElementById('A3_5_NOTE');
            //alert("test");
            //alert(S1chk_CBLVAL.indexOf("'5'"));
            //alert(A2_RBLVAL);alert(A3_RBLVAL);
            if (isEmpty('S1chk')) msg += '學員基本資料：(一)參加產投方案動機（可複選）：\n';
            if (S1chk_CBLVAL.indexOf("'5'") != -1 && S16_NOTE.value == "") {
                msg += cst_t1Q + cst_msgA2;
            };
            if (isEmpty('S2')) msg += '學員基本資料：(二)是否為第1次參加產業人才投資方案課程？\n';
            if (isEmpty('S3')) msg += '學員基本資料：(三)服務單位員工人數：\n';
            if (isEmpty('A1chk')) msg += '第一部份：(一)您獲得本次課程的訊息來源（可複選）：\n';
            if (A1chk_CBLVAL.indexOf("'9'") != -1 && A1_10_NOTE.value == "") {
                msg += cst_t2Q + cst_msgA2;
            };
            if (isEmpty('A2')) msg += '第一部份：(二)參加本次課程的主要原因：\n';
            if (A2_RBLVAL.indexOf("7") != -1 && A2_7_NOTE.value == "") {
                msg += cst_t3Q + cst_msgA2;
            };
            if (isEmpty('A3')) msg += '第一部份：(三)選擇本訓練單位的主要原因：\n';
            if (A3_RBLVAL.indexOf("5") != -1 && A3_5_NOTE.value == "") {
                msg += cst_t4Q + cst_msgA2;
            };
            if (isEmpty('A4')) msg += '第一部份：(四)沒有參加本方案訓練之前，每年參加訓練支出的費用？\n';
            if (isEmpty('A5')) msg += '第一部份：(五)如果沒有補助訓練費用，你每年願意自費參加訓練課程的金額？\n';
            if (isEmpty('A6')) msg += '第一部份：(六)您認為本次課程的訓練費用是否合理？\n';
            if (isEmpty('A7')) msg += '第一部份：(七)結訓後對於工作的規劃？\n';
            if (isEmpty('B11')) msg += '第二部份：(一)訓練課程1.課程內容符合期望\n';
            if (isEmpty('B12')) msg += '第二部份：(一)訓練課程2.課程難易安排適當\n';
            if (isEmpty('B13')) msg += '第二部份：(一)訓練課程3.課程總時數適當\n';
            if (isEmpty('B14')) msg += '第二部份：(一)訓練課程4.課程符合實務需求\n';
            if (isEmpty('B15')) msg += '第二部份：(一)訓練課程5.課程符合產業發展趨勢\n';
            if (isEmpty('B21')) msg += '第二部份：(二)講師1.滿意講師的教學態度\n';
            if (isEmpty('B22')) msg += '第二部份：(二)講師2.滿意講師的教學方法\n';
            if (isEmpty('B23')) msg += '第二部份：(二)講師3.滿意講師的課程專業度\n';
            if (isEmpty('B31')) msg += '第二部份：(三)教材1.對於訓練教材感到滿意\n';
            if (isEmpty('B32')) msg += '第二部份：(三)教材2.訓練教材能夠輔助課程學習\n';
            if (isEmpty('B41')) msg += '第二部份：(四)訓練環境1.您對於訓練場地感到滿意\n';
            if (isEmpty('B42')) msg += '第二部份：(四)訓練環境2.您對於訓練設備感到滿意\n';
            if (isEmpty('B43')) msg += '第二部份：(四)訓練環境3.您認為實作設備的數量適當\n';
            if (isEmpty('B44')) msg += '第二部份：(四)訓練環境4.您認為實作設備新穎\n';
            if (isEmpty('B51')) msg += '第二部份：(五)訓練評量:訓練評量（如：訓後測驗、專題報告、作品展示等）能促進學習效果\n';
            if (isEmpty('B61')) msg += '第二部份：(六)立即學習效果1.您認為在訓練課程中，課程內容能讓您專注\n';
            if (isEmpty('B62')) msg += '第二部份：(六)立即學習效果2.您在完成訓練後，已充份學習訓練課程所教授知識或技能\n';
            if (isEmpty('B63')) msg += '第二部份：(六)立即學習效果3.您在完成訓練後，有學習到新的知識或技能\n';
            if (isEmpty('B71')) msg += '第二部份：(七)整體意見1.您對於訓練單位的課程安排與授課情形感到滿意\n';
            if (isEmpty('B72')) msg += '第二部份：(七)整體意見2.您對於訓練單位的行政服務感到滿意\n';
            if (isEmpty('B73')) msg += '第二部份：(七)整體意見3.您對於產業人才投資方案感到滿意\n';
            if (isEmpty('B74')) msg += '第二部份：(七)整體意見4.您認為完成本訓練課程對於目前或未來工作有幫助\n';
            if (isEmpty('C11')) msg += '第三部份：(一)若本訓練課程沒有補助，是否會全額自費參訓？\n';
            if (msg != '') {
                msg = '請確認下列答案：\n' + msg;
                alert(msg);
                return false;
            }
            return true;
        }

        function insert_next() {
            var mainpageUrl = 'SD_11_004.aspx';
            var pageUrl = 'SD_11_004_add17.aspx';
            var Re_OCID = document.getElementById("Re_OCID");
            var Re_SOCID = document.getElementById("Re_SOCID");
            var Re_ID = document.getElementById("Re_ID");
            var sHref = mainpageUrl + '?ProcessType=Back&ocid=' + Re_OCID.value + '&ID=' + Re_ID.value;
            if (window.confirm("儲存成功!!是否繼續新增下一筆?")) {
                sHref = pageUrl + '?ProcessType=Next&ocid=' + Re_OCID.value + '&SOCID=' + Re_SOCID.value + '&ID=' + Re_ID.value;
            }
            location.href = sHref;
        }
    </script>
    <%--.font { font-size: 12px; line-height: 24px; color: #696969; }--%>
    <style type="text/css">
        .BBstyle1 { color: #000000; background-color: #ecf7ff; }
        .BBstyle2 { color: #000000; font-size: 12px; line-height: 22px; text-align: center; background-color: #CCD8EE; padding: 2px; }
        .BBstyle_t1 { color: #000000; background-color: #ecf7ff; font-weight: bold; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <input id="Re_OCID" type="hidden" name="Re_OCID" runat="server" size="1">
        <input id="Re_SOCID" type="hidden" name="Re_SOCID" runat="server" size="1">
        <input id="ProcessType" type="hidden" name="ProcessType" runat="server">
        <input id="Re_ID" type="hidden" name="Re_ID" runat="server">
        <asp:CustomValidator ID="CustomValidator1" runat="server" ClientValidationFunction="CheckSurvey" ErrorMessage="CustomValidator" Display="None"></asp:CustomValidator>
        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                  <%--<tr><td><asp:Label ID="TitleLab1" runat="server"></asp:Label><asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<FONT color="#990000">受訓學員意見調查表</FONT></asp:Label></td></tr>--%>
                        <tr>
                            <td>
                                <asp:Label ID="Label_Stud" runat="server" CssClass="font"></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                <asp:Label ID="Label_Name" runat="server" CssClass="font"></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                <asp:Label ID="Label_Status" runat="server"></asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="TableName" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr id="StdTr" runat="server">
                            <td class="bluecol" width="20%">學員 </td>
                            <td colspan="3" class="whitecol">
                                <asp:DropDownList ID="SOCID" runat="server" AutoPostBack="True"></asp:DropDownList></td>
                        </tr>
                    </table>
                    <table class="font" id="tb3_Datalist" runat="server" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="BBstyle_t1">
                                <asp:Literal ID="Lit_msg_1" runat="server"></asp:Literal></td>
                        </tr>
                        <tr>
                            <td class="table_title"><span><strong>學員基本資料</strong></span> </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1"><span><strong>(一) 參加產投方案動機（可複選）：</strong></span> </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1">
                                <asp:CheckBoxList ID="S1chk" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3" Width="90%">
                                    <asp:ListItem Value="1">1.加強原專長相關技能</asp:ListItem>
                                    <asp:ListItem Value="2">2.培育第二專長或轉換其他行職業所需技能</asp:ListItem>
                                    <asp:ListItem Value="3">3.考取相關檢定或證照</asp:ListItem>
                                    <asp:ListItem Value="4">4.拓展人脈</asp:ListItem>
                                    <asp:ListItem Value="5">5.使用政府提供之訓練費用補助</asp:ListItem>
                                    <asp:ListItem Value="6">6.其他（請說明）</asp:ListItem>
                                </asp:CheckBoxList>
                            </td>
                        </tr>
                        <tr>
                            <td style="background-color: #ecf7ff" class="whitecol">
                                <div align="right">
                                    其他：<asp:TextBox ID="S16_NOTE" runat="server" MaxLength="100" Columns="40" Width="33%"></asp:TextBox>
                                    &nbsp;&nbsp;&nbsp;&nbsp;
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1"><span><strong>(二) 是否為第1次參加產業人才投資方案課程？</strong></span>
                                <asp:RadioButtonList ID="S2" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3" RepeatLayout="Flow">
                                    <asp:ListItem Value="1">1.是</asp:ListItem>
                                    <asp:ListItem Value="2">2.否</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1"><span><strong>(三) 服務單位員工人數：</strong></span> </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1">
                                <asp:RadioButtonList ID="S3" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="4" Width="90%">
                                    <asp:ListItem Value="1">1.9人以下</asp:ListItem>
                                    <asp:ListItem Value="2">2.10人-29人</asp:ListItem>
                                    <asp:ListItem Value="3">3.30人-49人</asp:ListItem>
                                    <asp:ListItem Value="4">4.50人-99人</asp:ListItem>
                                    <asp:ListItem Value="5">5.100人-199人</asp:ListItem>
                                    <asp:ListItem Value="6">6.200人-499人</asp:ListItem>
                                    <asp:ListItem Value="7">7.500人以上</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="table_title"><span><strong>第一部份：參加產投方案考量因素</strong></span> </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1"><span><strong>(一) 您獲得本次課程的訊息來源（可複選）：</strong></span> </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1">
                                <asp:CheckBoxList ID="A1chk" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="4" Width="90%">
                                    <asp:ListItem Value="1">1.本署或分署網站</asp:ListItem>
                                    <asp:ListItem Value="2">2.就業服務中心</asp:ListItem>
                                    <asp:ListItem Value="3">3.訓練單位</asp:ListItem>
                                    <asp:ListItem Value="4">4.搜尋網站</asp:ListItem>
                                    <asp:ListItem Value="5">5.報紙</asp:ListItem>
                                    <asp:ListItem Value="6">6.廣播</asp:ListItem>
                                    <asp:ListItem Value="7">7.電視</asp:ListItem>
                                    <asp:ListItem Value="8">8.親友介紹</asp:ListItem>
                                    <asp:ListItem Value="9">9.社群媒體(ex：臉書、LINE)</asp:ListItem>
                                    <asp:ListItem Value="10">10.其他</asp:ListItem>
                                </asp:CheckBoxList>
                            </td>
                        </tr>
                        <tr>
                            <td style="background-color: #ecf7ff" class="whitecol">
                                <div align="right">
                                    其他：<asp:TextBox ID="A1_10_NOTE" runat="server" MaxLength="100" Columns="50" Width="33%"></asp:TextBox>
                                    &nbsp;&nbsp;&nbsp;&nbsp;
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1"><span><strong>(二) 參加本次課程的主要原因：</strong></span> </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1">
                                <asp:RadioButtonList ID="A2" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="4" Width="90%">
                                    <asp:ListItem Value="1">1.課程符合就業市場需求</asp:ListItem>
                                    <asp:ListItem Value="2">2.課程符合目前工作需求</asp:ListItem>
                                    <asp:ListItem Value="3">3.課程符合個人興趣</asp:ListItem>
                                    <asp:ListItem Value="4">4.可取得課程相關證照或證書</asp:ListItem>
                                    <asp:ListItem Value="5">5.學習第二專長</asp:ListItem>
                                    <asp:ListItem Value="6">6.師資具知名度或專業性</asp:ListItem>
                                    <asp:ListItem Value="7">7.其他</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td style="background-color: #ecf7ff" class="whitecol">
                                <div align="right">
                                    其他：<asp:TextBox ID="A2_7_NOTE" runat="server" MaxLength="100" Columns="50" Width="33%"></asp:TextBox>
                                    &nbsp;&nbsp;&nbsp;&nbsp;
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1"><span><strong>(三) 選擇本訓練單位的主要原因：</strong></span> </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1">
                                <asp:RadioButtonList ID="A3" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3" Width="90%">
                                    <asp:ListItem Value="1">1.環境、設備良好</asp:ListItem>
                                    <asp:ListItem Value="2">2.具課程專業度</asp:ListItem>
                                    <asp:ListItem Value="3">3.行政人員服務良好</asp:ListItem>
                                    <asp:ListItem Value="4">4.為訓練單位之會員</asp:ListItem>
                                    <asp:ListItem Value="5">5.其他</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td style="background-color: #ecf7ff" class="whitecol">
                                <div align="right">
                                    其他：<asp:TextBox ID="A3_5_NOTE" runat="server" MaxLength="100" Columns="50" Width="33%"></asp:TextBox>
                                    &nbsp;&nbsp;&nbsp;&nbsp;
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1"><span><strong>(四) 沒有參加本方案訓練之前，每年參加訓練支出的費用？</strong></span> </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1">
                                <asp:RadioButtonList ID="A4" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="4" Width="90%">
                                    <asp:ListItem Value="1">1.0元</asp:ListItem>
                                    <asp:ListItem Value="2">2.999元以下</asp:ListItem>
                                    <asp:ListItem Value="3">3.1,000元-3,999元</asp:ListItem>
                                    <asp:ListItem Value="4">4.4,000元-6,999元</asp:ListItem>
                                    <asp:ListItem Value="5">5.7,000元-9,999元</asp:ListItem>
                                    <asp:ListItem Value="6">6.10,000元-19,999元</asp:ListItem>
                                    <asp:ListItem Value="7">7.20,000元-29,999元</asp:ListItem>
                                    <asp:ListItem Value="8">8.30,000元-39,999元</asp:ListItem>
                                    <%--<asp:ListItem Value="9">9.4,000元-6,999元</asp:ListItem>--%>
                                    <asp:ListItem Value="9">9.40,000元以上</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1"><span><strong>(五) 如果沒有補助訓練費用，你每年願意自費參加訓練課程的金額？</strong></span> </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1">
                                <asp:RadioButtonList ID="A5" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="4" Width="90%">
                                    <asp:ListItem Value="1">1.0元</asp:ListItem>
                                    <asp:ListItem Value="2">2.999元以下</asp:ListItem>
                                    <asp:ListItem Value="3">3.1,000元-3,999元</asp:ListItem>
                                    <asp:ListItem Value="4">4.4,000元-6,999元</asp:ListItem>
                                    <asp:ListItem Value="5">5.7,000元-9,999元</asp:ListItem>
                                    <asp:ListItem Value="6">6.10,000元-19,999元</asp:ListItem>
                                    <asp:ListItem Value="7">7.20,000元-29,999元</asp:ListItem>
                                    <asp:ListItem Value="8">8.30,000元-39,999元</asp:ListItem>
                                    <%--<asp:ListItem Value="9">9.4,000元-6,999元</asp:ListItem>--%>
                                    <asp:ListItem Value="9">9.40,000元以上</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1"><span><strong>(六) 您認為本次課程的訓練費用是否合理？</strong></span>
                                <asp:RadioButtonList ID="A6" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="4" RepeatLayout="Flow">
                                    <asp:ListItem Value="1">1.偏高</asp:ListItem>
                                    <asp:ListItem Value="2">2.合理</asp:ListItem>
                                    <asp:ListItem Value="3">3.偏低</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1"><span><strong>(七) 結訓後對於工作的規劃？</strong></span> </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1">
                                <asp:RadioButtonList ID="A7" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="4" Width="90%">
                                    <asp:ListItem Value="1">1.留在原職位</asp:ListItem>
                                    <asp:ListItem Value="2">2.轉調較能發揮潛能部門</asp:ListItem>
                                    <asp:ListItem Value="3">3.轉換至同業的其他公司</asp:ListItem>
                                    <asp:ListItem Value="4">4.轉換至不同行業的公司</asp:ListItem>
                                    <asp:ListItem Value="5">5.希望自己創業</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="table_title"><span><strong>第二部份：訓練課程設計與執行過程意見調查</strong></span> </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1" align="center">
                                <table class="font" id="Table4" cellspacing="0" cellpadding="0" width="100%" border="0">
                                    <tr>
                                        <td width="50%" align="center">題目 </td>
                                        <td width="10%" align="center">非常同意 </td>
                                        <td width="10%" align="center">同意 </td>
                                        <td width="10%" align="center">普通 </td>
                                        <td width="10%" align="center">不同意 </td>
                                        <td width="10%" align="center">非常不同意 </td>
                                    </tr>
                                    <tr>
                                        <td class="BBstyle1" colspan="6"><span><strong>(一) 訓練課程</strong></span> </td>
                                    </tr>
                                    <tr>
                                        <td>1.課程內容符合期望 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="B11" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>2.課程難易安排適當 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="B12" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>3.課程總時數適當 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="B13" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>4.課程符合實務需求 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="B14" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>5.課程符合產業發展趨勢 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="B15" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="BBstyle1" colspan="6"><span><strong>(二)講師</strong></span> </td>
                                    </tr>
                                    <tr>
                                        <td>1.滿意講師的教學態度 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="B21" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>2.滿意講師的教學方法 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="B22" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>3.滿意講師的課程專業度 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="B23" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="BBstyle1" colspan="6"><span><strong>(三)教材</strong></span> </td>
                                    </tr>
                                    <tr>
                                        <td>1.對於訓練教材感到滿意 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="B31" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>2.訓練教材能夠輔助課程學習 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="B32" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="BBstyle1" colspan="6"><span><strong>(四)訓練環境</strong></span> </td>
                                    </tr>
                                    <tr>
                                        <td>1.您對於訓練場地感到滿意 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="B41" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>2.您對於訓練設備感到滿意 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="B42" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>3.您認為實作設備的數量適當 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="B43" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>4.您認為實作設備新穎 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="B44" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="BBstyle1" colspan="6"><span><strong>(五)訓練評量</strong></span> </td>
                                    </tr>
                                    <tr>
                                        <td>訓練評量（如：訓後測驗、專題報告、作品展示等）<br />
                                            能促進學習效果 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="B51" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="BBstyle1" colspan="6"><span><strong>(六)立即學習效果</strong></span> </td>
                                    </tr>
                                    <tr>
                                        <td>1.您認為在訓練課程中，課程內容能讓您專注 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="B61" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>2.您在完成訓練後，已充份學習訓練課程所教授知識或技能 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="B62" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>3.您在完成訓練後，有學習到新的知識或技能 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="B63" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="BBstyle1" colspan="6"><span><strong>(七)整體意見</strong></span> </td>
                                    </tr>
                                    <tr>
                                        <td>1.您對於訓練單位的課程安排與授課情形感到滿意 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="B71" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>2.您對於訓練單位的行政服務感到滿意 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="B72" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>3.您對於產業人才投資方案感到滿意 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="B73" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>4.您認為完成本訓練課程對於目前或未來工作有幫助 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="B74" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td class="table_title"><span><strong>第三部份：其他建議</strong></span> </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1"><span><strong>(一)若本訓練課程沒有補助，是否會全額自費參訓？</strong></span> </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1">
                                <asp:RadioButtonList ID="C11" runat="server" CssClass="font" RepeatDirection="Horizontal" Width="90%">
                                    <asp:ListItem Value="1">1.一定會</asp:ListItem>
                                    <asp:ListItem Value="2">2.應該會</asp:ListItem>
                                    <asp:ListItem Value="3">3.普通</asp:ListItem>
                                    <asp:ListItem Value="4">4.應該不會</asp:ListItem>
                                    <asp:ListItem Value="5">5.一定不會</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1"><span><strong>(二)其他建議：</strong></span> </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1" align="left">
                                <asp:TextBox ID="C21_NOTE" runat="server" MaxLength="500" Rows="8" TextMode="MultiLine" Width="77%"></asp:TextBox><br />
                            </td>
                        </tr>
                    </table>
                    <div align="center" class="whitecol">
                        <asp:Button ID="Button1" runat="server" Text="儲存" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="Button2" runat="server" Text="回上一頁" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="next_but" runat="server" Text="不儲存填寫下一位" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                    </div>
                </td>
            </tr>
        </table>
        <input id="HidDASOURCE" type="hidden" name="HidDASOURCE" runat="server">
    </form>
</body>
</html>
