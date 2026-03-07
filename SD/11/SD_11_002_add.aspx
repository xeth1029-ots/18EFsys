<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_11_002_add.aspx.vb" Inherits="WDAIIP.SD_11_002_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_11_002_add</title>
    <meta http-equiv="Content-Type" content="text/html; charset=big5">
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script language="javascript">
        function disable_radio1() {
            if (getValue("RadioButtonList1_2") == '4' && getValue("RadioButtonList1_3") == '2') {
                document.getElementById("RadioButtonList1_4").disabled = true;
                document.getElementById("RadioButtonList1_5").disabled = true;
                document.getElementById("Q1_5Other").disabled = true;
                document.getElementById("RadioButtonList1_6").disabled = true;
                document.getElementById("Q1_6Other").disabled = true;
                document.getElementById("RadioButtonList1_7").disabled = true;
            } else {
                document.getElementById("RadioButtonList1_4").disabled = false;
                document.getElementById("RadioButtonList1_5").disabled = false;
                document.getElementById("Q1_5Other").disabled = false;
                document.getElementById("RadioButtonList1_6").disabled = false;
                document.getElementById("Q1_6Other").disabled = false;
                document.getElementById("RadioButtonList1_7").disabled = false;
            }
            if (getValue("RadioButtonList1_2") == '4' && getValue("RadioButtonList1_3") == '1') {
                document.getElementById("RadioButtonList1_8").disabled = true;
                document.getElementById("Q1_8Other").disabled = true;
            } else {
                document.getElementById("RadioButtonList1_8").disabled = false;
                document.getElementById("Q1_8Other").disabled = false;
            }
            if (getValue("RadioButtonList1_2") == '4') {
                document.getElementById("RadioButtonList2_7").disabled = true;
                document.getElementById("RadioButtonList2_8").disabled = true;
                document.getElementById("RadioButtonList2_9").disabled = true;
                document.getElementById("RadioButtonList2_10").disabled = true;
            } else {
                document.getElementById("RadioButtonList2_7").disabled = false;
                document.getElementById("RadioButtonList2_8").disabled = false;
                document.getElementById("RadioButtonList2_9").disabled = false;
                document.getElementById("RadioButtonList2_10").disabled = false;
            }
            if (getValue("RadioButtonList2_10") == '4' || getValue("RadioButtonList2_10") == '5') {
                document.getElementById("RadioButtonList2_11").disabled = false;
                document.getElementById("Q2_11Other").disabled = false;
            } else {
                document.getElementById("RadioButtonList2_11").disabled = true;
                document.getElementById("Q2_11Other").disabled = true;
            }
        }

        function insert_next() {
            if (window.confirm("是否繼續新增下一筆?")) {
                location.href = 'SD_11_002_add.aspx?ProcessType=Next&ocid=' + document.getElementById("Re_OCID").value + '&Stuedntid=' + document.getElementById("Re_Studentid").value + '&ID=' + document.getElementById("Re_ID").value;
            }
            else {
                location.href = 'SD_11_002.aspx?ProcessType=Back&ocid=' + document.getElementById("Re_OCID").value + '&ID=' + document.getElementById("Re_ID").value;
            }
        }

        function CheckSurvey(source, args) {
            args.IsValid = true;
            source.errormessage = "";
            if (getValue("RadioButtonList1_2") == '1' || getValue("RadioButtonList1_2") == '2' || getValue("RadioButtonList1_2") == '3') {
                if (getValue("RadioButtonList1_3") == '1') {
                    if (getValue("RadioButtonList1_4") == '' && document.ggetElementById("RadioButtonList1_4").disabled == false) {
                        args.IsValid = false;
                        source.errormessage += '請選擇第一部分的問題四<br>';
                    }
                    if (getValue("RadioButtonList1_5") == '' && document.getElementById("RadioButtonList1_5").disabled == false) {
                        args.IsValid = false;
                        source.errormessage += '請選擇第一部分的問題五<br>';
                    }
                    if (getValue("RadioButtonList1_6") == '' && document.getElementById("RadioButtonList1_6").disabled == false) {
                        args.IsValid = false;
                        source.errormessage += '請選擇第一部分的問題六<br>';
                    }
                    if (getValue("RadioButtonList1_7") == '' && document.getElementById("RadioButtonList1_7").disabled == false) {
                        args.IsValid = false;
                        source.errormessage += '請選擇第一部分的問題七<br>';
                    }
                }
            }
            if (getValue("RadioButtonList1_2") == '4' && RadioButtonList1_3 == '1') {
                if (getValue("RadioButtonList1_8") == '' && document.getElementById("RadioButtonList1_8").disabled == false) {
                    args.IsValid = false;
                    source.errormessage += '請選擇第一部分的問題八<br>';
                }
            }
            if (getValue("RadioButtonList1_5") == '97' && getValue("Q1_5Other") == '' && document.getElementById("RadioButtonList1_5").disabled == false) {
                args.IsValid = false;
                source.errormessage += '請輸入第一部分問題五中的其他選項<br>';
            }
            if (getValue("RadioButtonList1_6") == '97' && getValue("Q1_6Other") == '' && document.getElementById("RadioButtonList1_6").disabled == false) {
                args.IsValid = false;
                source.errormessage += '請輸入第一部分問題六中的其他選項<br>';
            }
            if (document.getElementById("RadioButtonList1_8").disabled == false && getValue("RadioButtonList1_8") == '97' && getValue("Q1_8Other") == '') {
                args.IsValid = false;
                source.errormessage += '請輸入第一部分問題八中的其他選項<br>';
            }
            if (getValue("RadioButtonList1_2") == '1' || getValue("RadioButtonList1_2") == '2' || getValue("RadioButtonList1_2") == '3') {
                if (getValue("RadioButtonList2_7") == '' && document.getElementById("RadioButtonList2_7").disabled == false) {
                    args.IsValid = false;
                    source.errormessage += '請選擇第二部分的問題七<br>';
                }
                if (getValue("RadioButtonList2_8") == '' && document.getElementById("RadioButtonList2_8").disabled == false) {
                    args.IsValid = false;
                    source.errormessage += '請選擇第二部分的問題八<br>';
                }
                if (getValue("RadioButtonList2_9") == '' && document.getElementById("RadioButtonList2_9").disabled == false) {
                    args.IsValid = false;
                    source.errormessage += '請選擇第二部分的問題九<br>';
                }
                //alert(getValue("RadioButtonList2_10"));
                if (getValue("RadioButtonList2_10") == '' && document.getElementById("RadioButtonList2_10").disabled == false) {
                    args.IsValid = false;
                    source.errormessage += '請選擇第二部分的問題十<br>';
                }
            }
            if (getCheckBoxListValue("CheckBoxList2_3").charAt(7) == '1' && getValue("Q2_3Other") == '') {
                args.IsValid = false;
                source.errormessage += '請輸入第二部分問題三中的其他選項<br>';
            }
            if (getValue("RadioButtonList2_11") == '97' && getValue("Q2_11Other") == '' && document.getElementById("RadioButtonList2_11").disabled == false) {
                args.IsValid = false;
                source.errormessage += '請輸入第二部分問題十一中的其他選項<br>';
            }
            if (getCheckBoxListValue("CheckBoxList3_4").charAt(8) == '1' && getValue("Q3_4Other") == '') {
                args.IsValid = false;
                source.errormessage += '請輸入第三部分問題四中的其他選項<br>';
            }
            if (parseInt(getCheckBoxListValue("CheckBoxList2_3"), 10) == 0) {
                args.IsValid = false;
                source.errormessage += '請選擇第二部分的問題三<br>';
            }
            if (parseInt(getCheckBoxListValue("CheckBoxList3_4"), 10) == 0) {
                args.IsValid = false;
                source.errormessage += '請選擇第三部分的問題四<br>';
            }
        }

        function window_onload() {
            disable_radio1();
        }

        function printDoc() {
            window.print();
            //debugger;
            //if (!factory.object) {
            //    return false;
            //} else {
            //    factory.printing.header = '';
            //    factory.printing.footer = '';
            //    factory.printing.portrait = true;
            //    factory.printing.Print(true);
            //}
        }
    </script>
    <%--<style type="text/css">
        .style1 { color: #000000; }
        .style2 { color: #ffffff; }
        .style3 { color: #ffffff; }
    </style>--%>
</head>
<%--<body onload="window_onload()">--%>
<body>
    <!-- MeadCo ScriptX -->
    <%--<object style="display: none" id="factory" codebase="../../scriptx/smsx.cab#Version=6,6,440,26" classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" viewastext></object>--%>
    <form id="form1" method="post" runat="server">
        <div id="divPage" runat="server" style="overflow-y: auto;">
            <input id="Re_OCID" type="hidden" name="Re_OCID" runat="server">
            <input id="Re_Studentid" type="hidden" name="Re_Studentid" runat="server">
            <input id="ProcessType" type="hidden" name="ProcessType" runat="server">
            <input id="Re_ID" style="width: 56px; height: 22px" type="hidden" size="4" name="Re_ID" runat="server">
            <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                <tr>
                    <td>
                        <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <tr>
                                <td>
                                    <asp:Label ID="Label_Name" runat="server" CssClass="font"></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;
                                    <asp:Label ID="Label_Stud" runat="server" CssClass="font"></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;
                                    <asp:Label ID="Label_Status" runat="server"></asp:Label>
                                </td>
                            </tr>
                        </table>
                        <table class="table_sch" id="Table3" width="100%">
                            <tr>
                                <td class="bluecol">【第一部份：目前工作背景-8】</td>
                            </tr>
                            <tr>
                                <td class="whitecol">
                                    <div>
                                        <asp:Label ID="Label1" runat="server" CssClass="FONT" BackColor="Transparent" BorderColor="Transparent" ForeColor="Black">[調查對象: 全體受訪者]</asp:Label></div>
                                    <div>1.※請問您結訓後等待就業的期間 ?</div>
                                    <div>
                                        <asp:RadioButtonList ID="RadioButtonList1_1" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="1">結訓後至一個月內</asp:ListItem>
                                            <asp:ListItem Value="2">1(含)-3個月</asp:ListItem>
                                            <asp:ListItem Value="3">3個月以上</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </div>
                                    <asp:RequiredFieldValidator ID="Re_R1_1" runat="server" ControlToValidate="RadioButtonList1_1" Display="None" ErrorMessage="請選擇第一部分的問題一"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol">
                                    <div>2.※請問您您目前的工作已從事多久?</div>
                                    <div>
                                        <asp:RadioButtonList ID="RadioButtonList1_2" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="1">3個月以上</asp:ListItem>
                                            <asp:ListItem Value="2">一個月以上至三個月內</asp:ListItem>
                                            <asp:ListItem Value="3">結訓後至一個月內</asp:ListItem>
                                            <asp:ListItem Value="4">尚未有工作</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </div>
                                    <asp:RequiredFieldValidator ID="Re_R1_2" runat="server" ControlToValidate="RadioButtonList1_2" Display="None" ErrorMessage="請選擇第一部分的問題二"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol">
                                    <div>3.請問您參加職訓後有沒有換工作？</div>
                                    <div>
                                        <asp:RadioButtonList ID="RadioButtonList1_3" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="1">有</asp:ListItem>
                                            <asp:ListItem Value="2">沒有</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </div>
                                    <asp:RequiredFieldValidator ID="Re_R1_3" runat="server" ControlToValidate="RadioButtonList1_3" Display="None" ErrorMessage="請選擇第一部分的問題三"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol">
                                    <div>
                                        <asp:Label ID="Label2" runat="server" CssClass="FONT" BackColor="Transparent" BorderColor="Transparent" ForeColor="Black">[調查對象: 目前有工作者且有換工作]</asp:Label></div>
                                    <div>4.整體來說滿不滿意目前的工作？</div>
                                    <div>
                                        <asp:RadioButtonList ID="RadioButtonList1_4" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                            <asp:ListItem Value="2">滿意</asp:ListItem>
                                            <asp:ListItem Value="3">普通</asp:ListItem>
                                            <asp:ListItem Value="4">不滿意</asp:ListItem>
                                            <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </div>
                                    <asp:CustomValidator ID="CustomValidator1" runat="server" Display="None" ErrorMessage="CustomValidator" ClientValidationFunction="CheckSurvey"></asp:CustomValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol">
                                    <div>5.請問您目前是從事哪一種行業？</div>
                                    <div>
                                        <asp:RadioButtonList ID="RadioButtonList1_5" runat="server" CssClass="font" RepeatColumns="2">
                                            <asp:ListItem Value="1">農林漁牧業</asp:ListItem>
                                            <asp:ListItem Value="2">製造業</asp:ListItem>
                                            <asp:ListItem Value="3">礦業及土石採取業</asp:ListItem>
                                            <asp:ListItem Value="4">水電燃氣業</asp:ListItem>
                                            <asp:ListItem Value="5">營造業</asp:ListItem>
                                            <asp:ListItem Value="6">批發及零售業及餐飲</asp:ListItem>
                                            <asp:ListItem Value="77">運輸、倉儲業及通信業</asp:ListItem>
                                            <asp:ListItem Value="8">金融保險及不動產</asp:ListItem>
                                            <asp:ListItem Value="9">工商服務業</asp:ListItem>
                                            <asp:ListItem Value="10">社會服務及個人服務業</asp:ListItem>
                                            <asp:ListItem Value="11">公共行政業</asp:ListItem>
                                            <asp:ListItem Value="97">其他（請說明-如國防事業）</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </div>
                                    <div>
                                        <asp:TextBox ID="Q1_5Other" runat="server" Width="55%"></asp:TextBox>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol">
                                    <div>6.那工作內容為何？</div>
                                    <div>
                                        <asp:RadioButtonList ID="RadioButtonList1_6" runat="server" CssClass="font" RepeatColumns="2">
                                            <asp:ListItem Value="1">主管及監督人員</asp:ListItem>
                                            <asp:ListItem Value="2">專業人員及技術人員</asp:ListItem>
                                            <asp:ListItem Value="3">工程師</asp:ListItem>
                                            <asp:ListItem Value="4">技術工</asp:ListItem>
                                            <asp:ListItem Value="5">事務工作人員</asp:ListItem>
                                            <asp:ListItem Value="6">業務員及售貨員</asp:ListItem>
                                            <asp:ListItem Value="7">生產及機械操作員</asp:ListItem>
                                            <asp:ListItem Value="8">作業員</asp:ListItem>
                                            <asp:ListItem Value="9">體力工</asp:ListItem>
                                            <asp:ListItem Value="97">其他(請說明)</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </div>
                                    <div>
                                        <asp:TextBox ID="Q1_6Other" runat="server" Width="55%"></asp:TextBox></div>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol">
                                    <div>7.目前工作每月的收入大約在哪一個範圍內？</div>
                                    <div>
                                        <asp:RadioButtonList ID="RadioButtonList1_7" runat="server" CssClass="font" RepeatColumns="2">
                                            <asp:ListItem Value="1">6萬元以上</asp:ListItem>
                                            <asp:ListItem Value="2">5-6(不含)萬元</asp:ListItem>
                                            <asp:ListItem Value="3">4-5(不含)萬元</asp:ListItem>
                                            <asp:ListItem Value="4">3-4(不含)萬元</asp:ListItem>
                                            <asp:ListItem Value="5">3(不含)萬元以下</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol">
                                    <div>
                                        <asp:Label ID="Label3" runat="server" CssClass="FONT">[調查對象: 目前&#27809;有工作但有換工作者]</asp:Label></div>
                                    <div>8.請問您離開職訓前那一份工作的原因為何？</div>
                                    <div>
                                        <asp:RadioButtonList ID="RadioButtonList1_8" runat="server" CssClass="font" RepeatColumns="2">
                                            <asp:ListItem Value="1">工作場所歇業或業務緊縮</asp:ListItem>
                                            <asp:ListItem Value="2">對原有的工作不滿意</asp:ListItem>
                                            <asp:ListItem Value="3">健康不良</asp:ListItem>
                                            <asp:ListItem Value="4">季節性或臨時性工作結束</asp:ListItem>
                                            <asp:ListItem Value="5">女性結婚或生育</asp:ListItem>
                                            <asp:ListItem Value="6">退休</asp:ListItem>
                                            <asp:ListItem Value="7">家務太忙</asp:ListItem>
                                            <asp:ListItem Value="97">其他（請說明）</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </div>
                                    <div>
                                        <asp:TextBox ID="Q1_8Other" runat="server" Width="55%"></asp:TextBox></div>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">【第二部份：職訓與工作-11】</td>
                            </tr>
                            <tr>
                                <td class="whitecol">
                                    <div>
                                        <asp:Label ID="Label4" runat="server" CssClass="FONT" BackColor="Transparent" BorderColor="Transparent" ForeColor="Black">[調查對象: 全體受訪者]</asp:Label></div>
                                    <div><span class="style1">1.※整體而言：您認為參加職訓課程對再就業或轉業是否有幫助？ </span><font color="#ffffff">&nbsp;&nbsp;&nbsp;</font></div>
                                    <div>
                                        <asp:RadioButtonList ID="RadioButtonList2_1" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="1">非常有幫助</asp:ListItem>
                                            <asp:ListItem Value="2">有幫助</asp:ListItem>
                                            <asp:ListItem Value="3">普通</asp:ListItem>
                                            <asp:ListItem Value="4">沒幫助</asp:ListItem>
                                            <asp:ListItem Value="5">完全沒幫助</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </div>
                                    <asp:RequiredFieldValidator ID="Re_R2_1" runat="server" ControlToValidate="RadioButtonList2_1" Display="None" ErrorMessage="請選擇第二部分的問題一"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol">
                                    <div>2.請問您認為參加職訓對您工作能力提升的幫助大不大？</div>
                                    <div>
                                        <asp:RadioButtonList ID="RadioButtonList2_2" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="1">非常有幫助</asp:ListItem>
                                            <asp:ListItem Value="2">有幫助</asp:ListItem>
                                            <asp:ListItem Value="3">普通</asp:ListItem>
                                            <asp:ListItem Value="4">沒幫助</asp:ListItem>
                                            <asp:ListItem Value="5">完全沒幫助</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </div>
                                    <asp:RequiredFieldValidator ID="Re_R2_2" runat="server" ControlToValidate="RadioButtonList2_2" Display="None" ErrorMessage="請選擇第二部分的問題二"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol">
                                    <div>3.請問您認為參加職訓，對您個人最大的收獲為何？ (複選題)</div>
                                    <div>
                                        <asp:CheckBoxList ID="CheckBoxList2_3" runat="server" CssClass="font" RepeatColumns="2">
                                            <asp:ListItem Value="1">收入增加</asp:ListItem>
                                            <asp:ListItem Value="2">升遷</asp:ListItem>
                                            <asp:ListItem Value="3">工作能力提升</asp:ListItem>
                                            <asp:ListItem Value="4">工作考績</asp:ListItem>
                                            <asp:ListItem Value="5">培養第二專長</asp:ListItem>
                                            <asp:ListItem Value="6">培養個人興趣</asp:ListItem>
                                            <asp:ListItem Value="7">增廣見聞</asp:ListItem>
                                            <asp:ListItem Value="97">其他(請註明)</asp:ListItem>
                                        </asp:CheckBoxList>
                                    </div>
                                    <div>
                                        <asp:TextBox ID="Q2_3Other" runat="server" Width="55%"></asp:TextBox></div>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol">
                                    <div>4.請問您認為受完訓後，對自信心提升上的幫助？</div>
                                    <div>
                                        <asp:RadioButtonList ID="RadioButtonList2_4" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="1">非常有幫助</asp:ListItem>
                                            <asp:ListItem Value="2">有幫助</asp:ListItem>
                                            <asp:ListItem Value="3">普通</asp:ListItem>
                                            <asp:ListItem Value="4">沒幫助</asp:ListItem>
                                            <asp:ListItem Value="5">完全沒幫助</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </div>
                                    <asp:RequiredFieldValidator ID="Re_R2_4" runat="server" ControlToValidate="RadioButtonList2_4" Display="None" ErrorMessage="請選擇第二部分的問題四"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol">
                                    <div>5.請問您認為受完訓後，對升遷上的幫助？</div>
                                    <div>
                                        <asp:RadioButtonList ID="RadioButtonList2_5" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="1">非常有幫助</asp:ListItem>
                                            <asp:ListItem Value="2">有幫助</asp:ListItem>
                                            <asp:ListItem Value="3">普通</asp:ListItem>
                                            <asp:ListItem Value="4">沒幫助</asp:ListItem>
                                            <asp:ListItem Value="5">完全沒幫助</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </div>
                                    <asp:RequiredFieldValidator ID="Re_R2_5" runat="server" ControlToValidate="RadioButtonList2_5" Display="None" ErrorMessage="請選擇第二部分的問題五"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol">
                                    <div>6.請問您認為受完訓後，對加薪上的幫助？</div>
                                    <div>
                                        <asp:RadioButtonList ID="RadioButtonList2_6" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="1">非常有幫助</asp:ListItem>
                                            <asp:ListItem Value="2">有幫助</asp:ListItem>
                                            <asp:ListItem Value="3">普通</asp:ListItem>
                                            <asp:ListItem Value="4">沒幫助</asp:ListItem>
                                            <asp:ListItem Value="5">完全沒幫助</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </div>
                                    <asp:RequiredFieldValidator ID="Re_R2_6" runat="server" ControlToValidate="RadioButtonList2_6" Display="None" ErrorMessage="請選擇第二部分的問題六"></asp:RequiredFieldValidator>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol">
                                    <div>
                                        <asp:Label ID="Label5" runat="server" CssClass="FONT">[以下各題調查對象: 目前有工作者]</asp:Label></div>
                                    <div>7.請問您目前工作收入是比職訓前增加？減少？還是一樣？</div>
                                    <div>
                                        <asp:RadioButtonList ID="RadioButtonList2_7" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="1">增加非常多</asp:ListItem>
                                            <asp:ListItem Value="2">增加</asp:ListItem>
                                            <asp:ListItem Value="3">一樣</asp:ListItem>
                                            <asp:ListItem Value="4">減少</asp:ListItem>
                                            <asp:ListItem Value="5">減少非常多</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol">
                                    <div>8.請問您目前的職等是比職訓前提升？下降？還是一樣？</div>
                                    <div>
                                        <asp:RadioButtonList ID="RadioButtonList2_8" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="1">職等高跳升</asp:ListItem>
                                            <asp:ListItem Value="2">職等微升</asp:ListItem>
                                            <asp:ListItem Value="3">一樣</asp:ListItem>
                                            <asp:ListItem Value="4">職等下降</asp:ListItem>
                                            <asp:ListItem Value="5">無法比較</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol">
                                    <div>9.請問目前公司有支持您將訓練獲得的知識技能運用在工作上嗎？</div>
                                    <div>
                                        <asp:RadioButtonList ID="RadioButtonList2_9" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="1">有</asp:ListItem>
                                            <asp:ListItem Value="2">沒有</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol">
                                    <div>10.請問您目前您有將訓練所獲得的知識技能運用在工作上嗎？</div>
                                    <div>
                                        <asp:RadioButtonList ID="RadioButtonList2_10" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="1">非常有幫助</asp:ListItem>
                                            <asp:ListItem Value="2">有幫助</asp:ListItem>
                                            <asp:ListItem Value="3">普通</asp:ListItem>
                                            <asp:ListItem Value="4">沒幫助</asp:ListItem>
                                            <asp:ListItem Value="5">完全沒幫助</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol">
                                    <div>
                                        <asp:Label ID="Label6" runat="server" CssClass="FONT">[11題調查對象:上一題回答沒有者]</asp:Label></div>
                                    <div>11.為何沒有？ (上一題回答 沒幫助 或 完全沒幫忙 者)</div>
                                    <div>
                                        <asp:RadioButtonList ID="RadioButtonList2_11" runat="server" CssClass="font">
                                            <asp:ListItem Value="1">訓練課程與工作不相關</asp:ListItem>
                                            <asp:ListItem Value="2">已轉換與職業訓練課程無關之工作</asp:ListItem>
                                            <asp:ListItem Value="97">其他</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </div>
                                    <div>
                                        <asp:TextBox ID="Q2_11Other" runat="server" Width="55%"></asp:TextBox></div>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">【第三部份：未來與建議-5】</td>
                            </tr>
                            <tr>
                                <td class="whitecol">
                                    <div>
                                        <asp:Label ID="Label7" runat="server" CssClass="FONT" BackColor="Transparent" BorderColor="Transparent" ForeColor="Black">[調查對象: 全體受訪者]</asp:Label></div>
                                    <div>1.※整體而言：您認為職訓師資的教學態度與方法，您是否滿意?</div>
                                    <div>
                                        <asp:RadioButtonList ID="RadioButtonList3_1" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                            <asp:ListItem Value="2">滿意</asp:ListItem>
                                            <asp:ListItem Value="3">普通</asp:ListItem>
                                            <asp:ListItem Value="4">不滿意</asp:ListItem>
                                            <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </div>
                                    <div>
                                        <asp:RequiredFieldValidator ID="Re_R3_1" runat="server" ControlToValidate="RadioButtonList3_1" Display="None" ErrorMessage="請選擇第三部分的問題一"></asp:RequiredFieldValidator></div>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol">
                                    <div>2.※整體而言：您對受訓單位之訓練場地的環境與設備及行政服務，您是否滿意?</div>
                                    <div>
                                        <asp:RadioButtonList ID="RadioButtonList3_2" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                            <asp:ListItem Value="2">滿意</asp:ListItem>
                                            <asp:ListItem Value="3">普通</asp:ListItem>
                                            <asp:ListItem Value="4">不滿意</asp:ListItem>
                                            <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </div>
                                    <div>
                                        <asp:RequiredFieldValidator ID="Re_R3_2" runat="server" ControlToValidate="RadioButtonList3_2" Display="None" ErrorMessage="請選擇第三部分的問題二"></asp:RequiredFieldValidator></div>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol">
                                    <div>3.若有機會，您會參加勞動力發展署委託或公訓機構辦理的更進一步的訓練嗎？</div>
                                    <div>
                                        <asp:RadioButtonList ID="RadioButtonList3_3" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="1">一定會</asp:ListItem>
                                            <asp:ListItem Value="2">會</asp:ListItem>
                                            <asp:ListItem Value="3">視況狀會</asp:ListItem>
                                            <asp:ListItem Value="4">不會</asp:ListItem>
                                            <asp:ListItem Value="5">完全不會</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </div>
                                    <div>
                                        <asp:RequiredFieldValidator ID="Re_R3_3" runat="server" ControlToValidate="RadioButtonList3_3" Display="None" ErrorMessage="請選擇第三部分的問題三"></asp:RequiredFieldValidator></div>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol">
                                    <div>4.您參加訓練後覺得訓練單位需要再加強的地方為何？（複選）</div>
                                    <div>
                                        <asp:CheckBoxList ID="CheckBoxList3_4" runat="server" CssClass="font">
                                            <asp:ListItem Value="1">課程內容與教材</asp:ListItem>
                                            <asp:ListItem Value="2">訓練師資</asp:ListItem>
                                            <asp:ListItem Value="3">訓練設備</asp:ListItem>
                                            <asp:ListItem Value="4">訓練場地</asp:ListItem>
                                            <asp:ListItem Value="5">訓練時數</asp:ListItem>
                                            <asp:ListItem Value="6">生活管理</asp:ListItem>
                                            <asp:ListItem Value="7">就業輔導</asp:ListItem>
                                            <asp:ListItem Value="8">已經都很好</asp:ListItem>
                                            <asp:ListItem Value="97">其他</asp:ListItem>
                                        </asp:CheckBoxList>
                                    </div>
                                    <div>
                                        <asp:TextBox ID="Q3_4Other" runat="server" Width="55%"></asp:TextBox></div>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol">
                                    <%--<div>5.※請問您覺得職業訓練局需要再加強的地方為何？</div>--%>
                                    <div>5.※請問您覺得勞動部勞動力發展署需要再加強的地方為何？</div>
                                    <div>
                                        <asp:TextBox ID="Q3_5" runat="server" Width="55%" TextMode="MultiLine" Rows="5" Columns="20"></asp:TextBox>
                                        <asp:RequiredFieldValidator ID="Requiredfieldvalidator1" runat="server" ControlToValidate="Q3_5" Display="None" ErrorMessage="請輸入第三部分的問題五"></asp:RequiredFieldValidator>
                                    </div>
                                </td>
                            </tr>
                        </table>
                        <div align="center" class="whitecol">
                            <asp:Button ID="Button1" runat="server" Text="儲存" CssClass="asp_Export_M"></asp:Button>
                            <%--<asp:button id="BtnBack1" runat="server" Text="回上一頁" CausesValidation="False" CssClass="asp_Export_M"></asp:button>--%>
                            <asp:button id="BtnBack1" runat="server" Text="回上一頁" CausesValidation="False" CssClass="asp_Export_M"></asp:button>
                            <%--<input id="Button2" type="button" value="回上一頁" name="Button2" runat="server" class="asp_Export_M">--%>
                            <asp:Button ID="next_but" runat="server" Text="不儲存填寫下一位" CausesValidation="False" CssClass="asp_Export_M"></asp:Button>
                        </div>
                        <div align="center" class="whitecol">
                            <asp:ValidationSummary ID="Summary" runat="server" ShowMessageBox="True" ShowSummary="False" DisplayMode="List"></asp:ValidationSummary>
                        </div>
                    </td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
