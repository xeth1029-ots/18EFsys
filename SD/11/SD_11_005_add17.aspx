<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_11_005_add17.aspx.vb" Inherits="WDAIIP.SD_11_005_add17" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>受訓學員訓後動態調查表</title>
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
        function Clk_Q1x4() {
            var Q1abc = document.getElementById("Q1abc");
            Q1abc.disabled = true;
            if (!isEmpty('rblQ1_4')) {
                Q1abc.disabled = false;
            }
        }

        function ChkData() {
            var msg = '';
            if (isEmpty('rblQ1_1') && isEmpty('rblQ1_2') && isEmpty('rblQ1_3')
				&& isEmpty('rblQ1_4') && isEmpty('rblQ1_5') && isEmpty('rblQ1_6')) {
                msg += '請選擇(一)請問您目前的就業狀況為何?\n';
            }
            if (!isEmpty('rblQ1_4')) {
                if (isEmpty('Q1abc')) { msg += '選擇(一)目前的就業狀況為4.創業，請選擇?(a.b.或c.)\n'; }
            }
            if (isEmpty('Q2')) { msg += '請選擇(二)請問您的薪資於結訓後有提升嗎?\n'; }
            if (isEmpty('Q3')) { msg += '請選擇(三)請問您擔任的職位有變化嗎?\n'; }
            if (isEmpty('Q4')) { msg += '請選擇(四)請問您對目前工作的滿意度是否有變化?\n'; }
            if (isEmpty('Q5')) { msg += '請選擇(五)請問您目前工作內容是否與本次參訓課程有相關?\n'; }
            if (isEmpty('Q6')) { msg += '請選擇(六)請問您是否有繼續參與本方案的意願?\n'; }
            if (isEmpty('Q7')) { msg += '請選擇(七)結訓後是否有與下列人員保持聯絡?(可複選)?\n'; }
            //alert(isEmpty('Q7'));
            var Q7x1 = "";
            if (getCheckBoxListValue("Q7").charAt(0) == '1') { Q7x1 = "1"; }
            if (getCheckBoxListValue("Q7").charAt(1) == '1') { Q7x1 = "1"; }
            if (getCheckBoxListValue("Q7").charAt(2) == '1') { Q7x1 = "1"; }
            if (getCheckBoxListValue("Q7").charAt(3) == '1' && Q7x1 == "1") {
                msg += '(七)結訓後是否有與下列人員保持聯絡，已選擇無，但又選擇其它答案，邏輯異常\n';
            }
            if (isEmpty('Q211')) { msg += '請選擇 二、訓練成效 (一)1.答案?\n'; }
            if (isEmpty('Q212')) { msg += '請選擇 二、訓練成效 (一)2.答案?\n'; }
            if (isEmpty('Q213')) { msg += '請選擇 二、訓練成效 (一)3.答案?\n'; }
            if (isEmpty('Q214')) { msg += '請選擇 二、訓練成效 (一)4.答案?\n'; }
            if (isEmpty('Q215')) { msg += '請選擇 二、訓練成效 (一)5.答案?\n'; }
            if (isEmpty('Q216')) { msg += '請選擇 二、訓練成效 (一)6.答案?\n'; }
            if (isEmpty('Q217')) { msg += '請選擇 二、訓練成效 (一)7.答案?\n'; }
            if (isEmpty('Q218')) { msg += '請選擇 二、訓練成效 (一)8.答案?\n'; }
            if (isEmpty('Q221')) { msg += '請選擇 二、訓練成效 (二)1.答案?\n'; }
            if (isEmpty('Q222')) { msg += '請選擇 二、訓練成效 (二)2.答案?\n'; }
            if (isEmpty('Q223')) { msg += '請選擇 二、訓練成效 (二)3.答案?\n'; }
            if (isEmpty('Q224')) { msg += '請選擇 二、訓練成效 (二)4.答案?\n'; }
            if (isEmpty('Q225')) { msg += '請選擇 二、訓練成效 (二)5.答案?\n'; }
            if (isEmpty('Q226')) { msg += '請選擇 二、訓練成效 (二)6.答案?\n'; }
            if (msg != '') {
                alert(msg);
                return false;
            }
            return true;
        }

        function insert_next(v_socid) {
            var aspxSDform1 = "SD_11_005_add17.aspx";
            var aspxSDForm2 = "SD_11_005.aspx";
            var vHref = "";
            var Re_ID = document.getElementById("Re_ID");
            var Re_OCID = document.getElementById("Re_OCID");
            var Re_SOCID = document.getElementById("Re_SOCID");
            vHref = aspxSDForm2 + '?ProcessType=Back&OCID=' + Re_OCID.value + '&ID=' + Re_ID.value;
            if (window.confirm("是否繼續新增下一筆?")) {
                vHref = aspxSDform1 + '?ProcessType=next&OCID=' + Re_OCID.value + '&SOCID=' + v_socid + '&ID=' + Re_ID.value;
            }
            location.href = vHref;
        }

    </script>
    <%--
	<style type="text/css">
		.style1 { color: #000000; }
		.style2 { color: #000000; }
	</style>--%>
    <style type="text/css">
        .BBstyle1 { color: #000000; background-color: #ecf7ff; }
        .BBstyle2 { color: #000000; font-size: 12px; line-height: 22px; text-align: center; background-color: #CCD8EE; padding: 2px; }
        .BBstyle_t1 { color: #000000; background-color: #ecf7ff; font-weight: bold; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <input id="ProcessType" type="hidden" name="ProcessType" runat="server">
        <input id="Re_OCID" type="hidden" name="Re_OCID" runat="server">
        <input id="Re_SOCID" type="hidden" name="Re_SOCID" runat="server">
        <input id="Re_ID" type="hidden" name="Re_ID" runat="server">
        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <%--<tr><td><asp:Label ID="TitleLab1" runat="server"></asp:Label><asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<FONT color="#990000">受訓學員訓後動態調查表</FONT></asp:Label></td></tr>--%>
                        <tr>
                            <td class="whitecol">
                                <asp:Label ID="Label_Stud" runat="server" CssClass="font"></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                <asp:Label ID="Label_Name" runat="server" CssClass="font"></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;
                                <asp:Label ID="Label_Status" runat="server"></asp:Label>
                                <asp:Button ID="btnClear1" runat="server" Text="Clear1" Visible="False" CssClass="asp_button_M" />
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="TableName" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr id="StdTr" runat="server">
                            <td width="20%" class="bluecol">學員</td>
                            <td colspan="3">
                                <asp:DropDownList ID="ddlSOCID" runat="server" AutoPostBack="True" ForeColor="Black"></asp:DropDownList></td>
                        </tr>
                    </table>
                    <table class="font" id="tb3_Datalist" runat="server" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="BBstyle_t1">本問卷係本署為瞭解學員參加本訓練後的就業狀況與未來動向，請協助調查並以打ˇ方式填答每一題項，如有其他意見請於<br />
                                &nbsp;&nbsp; 三、其他建議欄以文字敘述，供本署改進之參考。謝謝！ </td>
                        </tr>
                        <tr>
                            <td class="table_title"><span><strong>一、學員部分</strong></span> </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1"><span><strong>(一)請問您目前的就業狀況為何?</strong></span> </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1">
                                <table class="font" id="tbQ1" cellspacing="0" cellpadding="0" width="90%" border="0">
                                    <tr>
                                        <td>
                                            <asp:RadioButton ID="rblQ1_1" runat="server" GroupName="Q1" Text="1.留任原公司" /></td>
                                        <td>
                                            <asp:RadioButton ID="rblQ1_2" runat="server" GroupName="Q1" Text="2.轉換至同產業的公司" /></td>
                                        <td>
                                            <asp:RadioButton ID="rblQ1_3" runat="server" GroupName="Q1" Text="3.轉換至不同產業的公司" /></td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:RadioButton ID="rblQ1_4" runat="server" GroupName="Q1" Text="4.創業" />
                                            (<asp:RadioButtonList ID="Q1abc" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3" RepeatLayout="Flow">
                                                <asp:ListItem Value="a">a.實體</asp:ListItem>
                                                <asp:ListItem Value="b">b.網路</asp:ListItem>
                                                <asp:ListItem Value="c">c.兩者皆有</asp:ListItem>
                                            </asp:RadioButtonList>
                                            ) </td>
                                        <td>
                                            <asp:RadioButton ID="rblQ1_5" runat="server" GroupName="Q1" Text="5.已離職，待業中" /></td>
                                        <td>
                                            <asp:RadioButton ID="rblQ1_6" runat="server" GroupName="Q1" Text="6.其他" /></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1"><span><strong>(二)請問您的薪資於結訓後有提升嗎?</strong></span> </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1">
                                <asp:RadioButtonList ID="Q2" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3" Width="80%">
                                    <asp:ListItem Value="1">1.大幅提升</asp:ListItem>
                                    <asp:ListItem Value="2">2.小幅提升</asp:ListItem>
                                    <asp:ListItem Value="3">3.沒有變化</asp:ListItem>
                                    <asp:ListItem Value="4">4.小幅減少</asp:ListItem>
                                    <asp:ListItem Value="5">5.大幅減少</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1"><span><strong>(三)請問您擔任的職位有變化嗎?</strong></span> </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1">
                                <asp:RadioButtonList ID="Q3" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3" Width="80%">
                                    <asp:ListItem Value="1">1.升職</asp:ListItem>
                                    <asp:ListItem Value="2">2.調職</asp:ListItem>
                                    <asp:ListItem Value="3">3.沒有變化</asp:ListItem>
                                    <asp:ListItem Value="4">4.降職</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1"><span><strong>(四)請問您對目前工作的滿意度是否有變化?</strong></span> </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1">
                                <asp:RadioButtonList ID="Q4" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3" Width="80%">
                                    <asp:ListItem Value="1">1.大幅提升</asp:ListItem>
                                    <asp:ListItem Value="2">2.小幅提升</asp:ListItem>
                                    <asp:ListItem Value="3">3.沒有變化</asp:ListItem>
                                    <asp:ListItem Value="4">4.小幅減少</asp:ListItem>
                                    <asp:ListItem Value="5">5.大幅減少</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1"><span><strong>(五)請問您目前工作內容是否與本次參訓課程有相關?</strong></span> </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1">
                                <asp:RadioButtonList ID="Q5" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3" Width="80%">
                                    <asp:ListItem Value="1">1.非常相關</asp:ListItem>
                                    <asp:ListItem Value="2">2.相關</asp:ListItem>
                                    <asp:ListItem Value="3">3.尚可</asp:ListItem>
                                    <asp:ListItem Value="4">4.不相關</asp:ListItem>
                                    <asp:ListItem Value="5">5.非常不相關</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1"><span><strong>(六)請問您是否有繼續參與本方案的意願?</strong></span> </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1">
                                <asp:RadioButtonList ID="Q6" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3" Width="80%">
                                    <asp:ListItem Value="1">1.非常想參與</asp:ListItem>
                                    <asp:ListItem Value="2">2.想參與</asp:ListItem>
                                    <asp:ListItem Value="3">3.尚可</asp:ListItem>
                                    <asp:ListItem Value="4">4.不想參與</asp:ListItem>
                                    <asp:ListItem Value="5">5.非常不想參與</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1"><span><strong>(七)結訓後是否有與下列人員保持聯絡?(可複選)</strong></span> </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1">
                                <asp:CheckBoxList ID="Q7" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3" Width="80%">
                                    <asp:ListItem Value="1">1.教師</asp:ListItem>
                                    <asp:ListItem Value="2">2.學員</asp:ListItem>
                                    <asp:ListItem Value="3">3.工作人員</asp:ListItem>
                                    <asp:ListItem Value="4">4.無</asp:ListItem>
                                </asp:CheckBoxList>
                            </td>
                        </tr>
                        <tr>
                            <td class="table_title"><span><strong>二、訓練成效</strong></span> </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1">
                                <table class="font" cellspacing="0" cellpadding="0" width="98%" border="0">
                                    <tr>
                                        <td width="50%" align="center">題目 </td>
                                        <td width="10%" align="center">非常同意 </td>
                                        <td width="10%" align="center">同意 </td>
                                        <td width="10%" align="center">普通 </td>
                                        <td width="10%" align="center">不同意 </td>
                                        <td width="10%" align="center">非常不同意 </td>
                                    </tr>
                                    <tr>
                                        <td colspan="6"><span><strong>(一)訓練技能運用</strong></span> </td>
                                    </tr>
                                    <tr>
                                        <td>1.參加訓練後，對工作能力更有信心 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="Q211" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>2.參加訓練後，有助於提升工作技能 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="Q212" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>3.參加訓練後，有助於提升工作效率 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="Q213" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>4.參加訓練後，能增進我的問題解決能力 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="Q214" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>5.參加訓練後，能將所學應用到工作上 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="Q215" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>6.參加訓練後，能將所學應用於日常生活中 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="Q216" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>7.是否同意參加訓練對第二專長有幫助 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="Q217" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>8.是否同意參加訓練對目前工作表現有幫助 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="Q218" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td colspan="6"><span><strong>(二)訓練成果</strong></span> </td>
                                    </tr>
                                    <tr>
                                        <td>1.參加訓練後，有助於提升我的績效考核 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="Q221" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>2.參加訓練後，有助於薪資的調升 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="Q222" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>3.參加訓練後，有助於職位的升遷 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="Q223" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>4.參加訓練後，有助於獲得證照 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="Q224" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>5.參加訓練後，有助於發展職涯 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="Q225" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
                                                <asp:ListItem Value="1">非常同意</asp:ListItem>
                                                <asp:ListItem Value="2">同意</asp:ListItem>
                                                <asp:ListItem Value="3">普通</asp:ListItem>
                                                <asp:ListItem Value="4">不同意</asp:ListItem>
                                                <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>6.參加訓練後，有助於強化個人職場競爭力 </td>
                                        <td colspan="5">
                                            <asp:RadioButtonList ID="Q226" runat="server" RepeatDirection="Horizontal" CssClass="font" Width="100%">
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
                            <td class="table_title"><span class="style2"><strong>三、其他建議：</strong></span> </td>
                        </tr>
                        <tr>
                            <td class="BBstyle1">
                                <div align="center">
                                    <asp:TextBox ID="Q3_Note" runat="server" MaxLength="500" Rows="5" TextMode="MultiLine" Width="70%"></asp:TextBox>
                                </div>
                                <br />
                            </td>
                        </tr>
                    </table>
                    <div align="center" class="whitecol">
                        <asp:Button ID="btnSave1" runat="server" Text="儲存" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="btnBack1" runat="server" Text="回上一頁" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="btnNext1" runat="server" Text="不儲存填寫下一位" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                    </div>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="Hid_DASOURCE" runat="server" />
    </form>
</body>
</html>
