<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_11_001_add.aspx.vb" Inherits="TIMS.SD_11_001_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_11_001_add</title>
    <meta http-equiv="Content-Type" content="text/html; charset=big5">
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script src="../../js/common.js"></script>
    <script language="javascript">
        function disable_radio3() {
            if (getValue("RadioButtonList3_4") == '2') {
                document.getElementById("RadioButtonList3_5").disabled = true;
                document.getElementById("RadioButtonList3_6").disabled = true;
                document.getElementById("RadioButtonList3_7").disabled = true;
            } else {
                document.getElementById("RadioButtonList3_5").disabled = false;
                document.getElementById("RadioButtonList3_6").disabled = false;
                document.getElementById("RadioButtonList3_7").disabled = false;
            }
        }

        function disable_radio5() {
            if (getValue("RadioButtonList5_3") == '2') {
                document.getElementById("RadioButtonList5_4").disabled = true;
            } else {
                document.getElementById("RadioButtonList5_4").disabled = false;
            }
        }

        function CheckSurvey(source, args) {
            args.IsValid = true;
            source.errormessage = "";
            if (getValue("RadioButtonList3_4") == '1') {
                if (getValue("RadioButtonList3_5") == '') {
                    args.IsValid = false;
                    source.errormessage += '請選擇第三部分的問題五<br>';
                }
                if (getValue("RadioButtonList3_6") == '') {
                    args.IsValid = false;
                    source.errormessage += '請選擇第三部分的問題六<br>';
                }
                if (getValue("RadioButtonList3_7") == '') {
                    args.IsValid = false;
                    source.errormessage += '請選擇第三部分的問題七<br>';
                }
            }
            if (getValue("RadioButtonList5_3") == '1') {
                if (getValue("RadioButtonList5_4") == '') {
                    args.IsValid = false;
                    source.errormessage += '請選擇第五部分的問題四-test2<br>';
                }
            }
        }

        function insert_next() {
            if (window.confirm("是否繼續新增下一筆?")) {
                location.href = 'SD_11_001_add.aspx?ProcessType=Next&ocid=' + document.getElementById("Re_OCID").value + '&Stuedntid=' + document.getElementById("Re_Studentid").value + '&ID=' + document.getElementById("Re_ID").value;
            } else {
                location.href = 'SD_11_001.aspx?ProcessType=Back&ocid=' + document.getElementById("Re_OCID").value + '&ID=' + document.getElementById("Re_ID").value;
            }
        }

        function BAK() {
            location.href = 'SD_11_001.aspx?ProcessType=Back&ocid=' + document.getElementById("Re_OCID").value + '&ID=' + document.getElementById("Re_ID").value;

        }

        function window_onload() {
            if (document.getElementById("Qtype_Value").value == 'B') { }
            else {
                disable_radio5();
                disable_radio3();
            }
        }

        /*** 20080724 andy ***/
        function ChgFont(QType) {
            if (QType == 'B') {
                document.getElementById("TD_R5").innerText = "【第五部份：職訓與工作-5】";
            }
            else {
                document.getElementById("TD_R5").innerText = "【第五部份：證照與工作-4】";
            }
        }

        function printDoc() {
            if (!factory.object) {
                return
            } else {
                factory.printing.header = '';
                factory.printing.footer = '';
                factory.printing.portrait = true;
                factory.printing.Print(true);
            }
        }
    </script>
    <style type="text/css">
        .style1 {
            color: #000000;
        }
        .style2 {
            color: #ffffff;
        }
        P.text {
            color: #ffffff;
            font-style: italic;
        }
    </style>
</head>
<%--<body onload="window_onload()" ms_positioning="GridLayout">--%>
    <body>
    <!-- MeadCo ScriptX -->
    <object id="factory" style="display: none" classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="../../scriptx/smsx.cab#Version=6,6,440,26" viewastext></object>
    <form id="form1" method="post" runat="server">
        <input id="Re_OCID" type="hidden" name="Re_OCID" runat="server">
        <input id="Re_Studentid" type="hidden" name="Re_Studentid" runat="server">
        <asp:CustomValidator ID="CustomValidator1" runat="server" Display="None" ErrorMessage="CustomValidator" ClientValidationFunction="CheckSurvey"></asp:CustomValidator>
        <input id="ProcessType" style="width: 40px; height: 22px" type="hidden" size="1" name="ProcessType" runat="server">
        <input id="Re_ID" style="width: 40px; height: 22px" type="hidden" size="1" name="Re_ID" runat="server">
        <input id="Qtype_Value" style="width: 80px; height: 22px" type="hidden" size="8" runat="server">
        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="Label_Stud" runat="server" CssClass="font"></asp:Label>
                                <asp:Label ID="Label_Name" runat="server" CssClass="font"></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;
                                <asp:Label ID="Label_Status" runat="server"></asp:Label>
                                <asp:Label ID="Label1" runat="server" Visible="False" ForeColor="Red">問卷類型尚未設定，請聯絡系統管理員！！</asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch" id="TableName">
                        <tr id="StdTr" runat="server">
                            <td width="20%" class="bluecol">學員</td>
                            <td id="TD_Stud" colspan="3" class="whitecol" width="80%"><asp:DropDownList ID="SOCID" runat="server" AutoPostBack="True"></asp:DropDownList></td>
                        </tr>
                    </table>
                    <table class="table_sch" id="Table3" runat="server">
                        <tr>
                            <td class="bluecol" id="TD_R1">【第一部份：課程與教材-3】</td>
                        </tr>
                        <tr bgcolor="#ecf7ff">
                            <td id="TD_R1_1" runat="server" class="whitecol">
                                <div><span class="style1"><asp:Label ID="Label_R1_1" runat="server">1.請問您這次參加的職訓課程，對課程內容安排及銜接是否滿意？</asp:Label></span></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList1_1" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                        <asp:ListItem Value="2">滿意</asp:ListItem>
                                        <asp:ListItem Value="3">普通</asp:ListItem>
                                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <asp:RequiredFieldValidator ID="Re_R1_1" runat="server" Display="None" ErrorMessage="請選擇第一部分的問題一" ControlToValidate="RadioButtonList1_1"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr>
                            <td id="TD_R1_2" runat="server" class="whitecol">
                                <div><span class="style1"><asp:Label ID="Label_R1_2" runat="server">2.請問您這次參加的職訓課程，對課程時數安排是否滿意？ </asp:Label></span></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList1_2" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                        <asp:ListItem Value="2">滿意</asp:ListItem>
                                        <asp:ListItem Value="3">普通</asp:ListItem>
                                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <asp:RequiredFieldValidator ID="Re_R1_2" runat="server" Display="None" ErrorMessage="請選擇第一部分的問題二" ControlToValidate="RadioButtonList1_2"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr bgcolor="#ecf7ff">
                            <td id="TD_R1_3" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R1_3" runat="server">3.請問您這次參加的職訓課程，對使用的上課教材與訓練設施（如工具 /材料）是否滿意？</asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList1_3" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                        <asp:ListItem Value="2">滿意</asp:ListItem>
                                        <asp:ListItem Value="3">普通</asp:ListItem>
                                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <div><asp:RequiredFieldValidator ID="Re_R1_3" runat="server" Display="None" ErrorMessage="請選擇第一部分的問題三" ControlToValidate="RadioButtonList1_3"></asp:RequiredFieldValidator></div>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" id="TD_R2">【第二部份：師資與教學-5】</td>
                        </tr>
                        <tr>
                            <td id="TD_R2_1" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R2_1" runat="server"> 1.請問您滿不滿意老師專業知識？ </asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList2_1" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                        <asp:ListItem Value="2">滿意</asp:ListItem>
                                        <asp:ListItem Value="3">普通</asp:ListItem>
                                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <div><asp:RequiredFieldValidator ID="Re_R2_1" runat="server" Display="None" ErrorMessage="請選擇第二部分的問題一" ControlToValidate="RadioButtonList2_1"></asp:RequiredFieldValidator></div>
                            </td>
                        </tr>
                        <tr>
                            <td id="TD_R2_2" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R2_2" runat="server">2.請問您滿不滿意老師教學態度及教學耐心？</asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList2_2" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                        <asp:ListItem Value="2">滿意</asp:ListItem>
                                        <asp:ListItem Value="3">普通</asp:ListItem>
                                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <div><asp:RequiredFieldValidator ID="Re_R2_2" runat="server" Display="None" ErrorMessage="請選擇第二部分的問題二" ControlToValidate="RadioButtonList2_2"></asp:RequiredFieldValidator></div>
                            </td>
                        </tr>
                        <tr>
                            <td id="TD_R2_3" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R2_3" runat="server">3.請問您滿不滿意老師實務操作之教導能力？</asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList2_3" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                        <asp:ListItem Value="2">滿意</asp:ListItem>
                                        <asp:ListItem Value="3">普通</asp:ListItem>
                                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <div><asp:RequiredFieldValidator ID="Re_R2_3" runat="server" Display="None" ErrorMessage="請選擇第二部分的問題三" ControlToValidate="RadioButtonList2_3"></asp:RequiredFieldValidator></div>
                            </td>
                        </tr>
                        <tr>
                            <td id="TD_R2_4" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R2_4" runat="server">4.請問您滿不滿意老師與學員間之互動？</asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList2_4" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                        <asp:ListItem Value="2">滿意</asp:ListItem>
                                        <asp:ListItem Value="3">普通</asp:ListItem>
                                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <div><asp:RequiredFieldValidator ID="Re_R2_4" runat="server" Display="None" ErrorMessage="請選擇第二部分的問題四" ControlToValidate="RadioButtonList2_4"></asp:RequiredFieldValidator></div>
                            </td>
                        </tr>
                        <tr>
                            <td id="TD_R2_5" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R2_5" runat="server">5.請問您對老師教學準備工作是否滿意？</asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList2_5" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                        <asp:ListItem Value="2">滿意</asp:ListItem>
                                        <asp:ListItem Value="3">普通</asp:ListItem>
                                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <div><asp:RequiredFieldValidator ID="Re_R2_5" runat="server" Display="None" ErrorMessage="請選擇第二部分的問題五" ControlToValidate="RadioButtonList2_5"></asp:RequiredFieldValidator></div>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" id="TD_R3">【第三部份：學習環境與行政支援-8】</td>
                        </tr>
                        <tr>
                            <td id="TD_R3_1" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R3_1" runat="server">1.請問您滿不滿意訓練單位的上課環境？</asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList3_1" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                        <asp:ListItem Value="2">滿意</asp:ListItem>
                                        <asp:ListItem Value="3">普通</asp:ListItem>
                                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <div><asp:RequiredFieldValidator ID="Re_R3_1" runat="server" Display="None" ErrorMessage="請選擇第三部分的問題一" ControlToValidate="RadioButtonList3_1"></asp:RequiredFieldValidator></div>
                            </td>
                        </tr>
                        <tr>
                            <td id="TD_R3_2" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R3_2" runat="server">2.請問您這次參加的職訓課程，對訓練單位公共安全(如無障礙設施）是否滿意？</asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList3_2" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                        <asp:ListItem Value="2">滿意</asp:ListItem>
                                        <asp:ListItem Value="3">普通</asp:ListItem>
                                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <div><asp:RequiredFieldValidator ID="Re_R3_2" runat="server" Display="None" ErrorMessage="請選擇第三部分的問題二" ControlToValidate="RadioButtonList3_2"></asp:RequiredFieldValidator></div>
                            </td>
                        </tr>
                        <tr>
                            <td id="TD_R3_3" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R3_3" runat="server">3.請問您對訓練單位行政支援（如求助導師及申訴管道）是否滿意？</asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList3_3" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                        <asp:ListItem Value="2">滿意</asp:ListItem>
                                        <asp:ListItem Value="3">普通</asp:ListItem>
                                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <div><asp:RequiredFieldValidator ID="Re_R3_3" runat="server" Display="None" ErrorMessage="請選擇第三部分的問題三" ControlToValidate="RadioButtonList3_3"></asp:RequiredFieldValidator></div>
                            </td>
                        </tr>
                        <tr>
                            <td id="TD_R3_4" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R3_4" runat="server">4.請問您參加職訓的機構有提供就業輔導嗎？</asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList3_4" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">有</asp:ListItem>
                                        <asp:ListItem Value="2">沒有</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <div><asp:RequiredFieldValidator ID="Re_R3_4" runat="server" Display="None" ErrorMessage="請選擇第三部分的問題四" ControlToValidate="RadioButtonList3_4"></asp:RequiredFieldValidator></div>
                            </td>
                        </tr>
                        <tr>
                            <td id="TD_R3_5" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R3_5" runat="server">5.那您滿不滿意其提供就業輔導服務？ (針對第四題回答有者) </asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList3_5" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                        <asp:ListItem Value="2">滿意</asp:ListItem>
                                        <asp:ListItem Value="3">普通</asp:ListItem>
                                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td id="TD_R3_6" runat="server" class="whitecol">
                                <div>6.<asp:Label ID="Label_R3_6" runat="server">6.請問您這次參加的職訓課程，對訓練單位提供就業資訊是否滿意？ (針對第四題回答有者) </asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList3_6" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                        <asp:ListItem Value="2">滿意</asp:ListItem>
                                        <asp:ListItem Value="3">普通</asp:ListItem>
                                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td id="TD_R3_7" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R3_7" runat="server">7.請問您這次參加的職訓課程，對訓練單位提供就業推介服務是否滿意？ (針對第四題回答有者)</asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList3_7" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                        <asp:ListItem Value="2">滿意</asp:ListItem>
                                        <asp:ListItem Value="3">普通</asp:ListItem>
                                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" id="TD_R4">【第四部份：學習效果-6】</td>
                        </tr>
                        <tr>
                            <td id="TD_R4_1" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R4_1" runat="server">1.請問您覺得自己上課內容吸收程度如何？</asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList4_1" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                        <asp:ListItem Value="2">滿意</asp:ListItem>
                                        <asp:ListItem Value="3">普通</asp:ListItem>
                                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <div><asp:RequiredFieldValidator ID="Re_R4_1" runat="server" Display="None" ErrorMessage="請選擇第四部分的問題一" ControlToValidate="RadioButtonList4_1"></asp:RequiredFieldValidator></div>
                            </td>
                        </tr>
                        <tr>
                            <td id="TD_R4_2" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R4_2" runat="server">2.您對於自己職訓這段期間表現打幾分？</asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList4_2" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">100-90</asp:ListItem>
                                        <asp:ListItem Value="2">89-80</asp:ListItem>
                                        <asp:ListItem Value="3">79-70</asp:ListItem>
                                        <asp:ListItem Value="4">69-60</asp:ListItem>
                                        <asp:ListItem Value="5">60分以下</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <div><asp:RequiredFieldValidator ID="Re_R4_2" runat="server" Display="None" ErrorMessage="請選擇第四部分的問題二" ControlToValidate="RadioButtonList4_2"></asp:RequiredFieldValidator></div>
                            </td>
                        </tr>
                        <tr>
                            <td id="TD_R4_3" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R4_3" runat="server">3.請問若用考試評估您的學習效果，您的學習效果如何？</asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList4_3" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                        <asp:ListItem Value="2">滿意</asp:ListItem>
                                        <asp:ListItem Value="3">普通</asp:ListItem>
                                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <div><asp:RequiredFieldValidator ID="Re_R4_3" runat="server" Display="None" ErrorMessage="請選擇第四部分的問題三" ControlToValidate="RadioButtonList4_3"></asp:RequiredFieldValidator></div>
                            </td>
                        </tr>
                        <tr>
                            <td id="TD_R4_4" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R4_4" runat="server">4.請問若用交作業評估您的學習效果，您的學習效果如何？</asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList4_4" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                        <asp:ListItem Value="2">滿意</asp:ListItem>
                                        <asp:ListItem Value="3">普通</asp:ListItem>
                                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <div><asp:RequiredFieldValidator ID="Re_R4_4" runat="server" Display="None" ErrorMessage="請選擇第四部分的問題四" ControlToValidate="RadioButtonList4_4"></asp:RequiredFieldValidator></div>
                            </td>
                        </tr>
                        <tr>
                            <td id="TD_R4_5" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R4_5" runat="server">5.請問若用實習評估您的學習效果，您的學習效果如何？</asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList4_5" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                        <asp:ListItem Value="2">滿意</asp:ListItem>
                                        <asp:ListItem Value="3">普通</asp:ListItem>
                                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <div><asp:RequiredFieldValidator ID="Re_R4_5" runat="server" Display="None" ErrorMessage="請選擇第四部分的問題五" ControlToValidate="RadioButtonList4_5"></asp:RequiredFieldValidator></div>
                            </td>
                        </tr>
                        <tr>
                            <td id="TD_R4_6" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R4_6" runat="server">6.請問您受訓前的期待與受訓後的感受有沒有落差？感受如何？</asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList4_6" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                        <asp:ListItem Value="2">滿意</asp:ListItem>
                                        <asp:ListItem Value="3">普通</asp:ListItem>
                                        <asp:ListItem Value="4">不滿意</asp:ListItem>
                                        <asp:ListItem Value="5">非常不滿意</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <div><asp:RequiredFieldValidator ID="Re_R4_6" runat="server" Display="None" ErrorMessage="請選擇第四部分的問題六" ControlToValidate="RadioButtonList4_6"></asp:RequiredFieldValidator></div>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" id="TD_R5">【第五部份：證照與工作-4】</td>
                        </tr>
                        <tr>
                            <td id="TD_R5_1" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R5_1" runat="server">1.請問您參加此職訓的目的之一是不是要考證照？</asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList5_1" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">是</asp:ListItem>
                                        <asp:ListItem Value="2">不是</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <div><asp:RequiredFieldValidator ID="Re_R5_1" runat="server" Display="None" ErrorMessage="請選擇第五部分的問題一" ControlToValidate="RadioButtonList5_1"></asp:RequiredFieldValidator></div>
                            </td>
                        </tr>
                        <tr>
                            <td id="TD_R5_2" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R5_2" runat="server">2.請問您這次參加的職訓課程，對您的證照考試幫助大不大？</asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList5_2" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常有幫助</asp:ListItem>
                                        <asp:ListItem Value="2">有幫助</asp:ListItem>
                                        <asp:ListItem Value="3">普通</asp:ListItem>
                                        <asp:ListItem Value="4">沒幫助</asp:ListItem>
                                        <asp:ListItem Value="5">完全沒幫助</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <asp:RequiredFieldValidator ID="Re_R5_2" runat="server" Display="None" ErrorMessage="請選擇第五部分的問題二" ControlToValidate="RadioButtonList5_2"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr>
                            <td id="TD_R5_3" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R5_3" runat="server">3.請問您這次參加的職訓課程，有沒有考到證照？</asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList5_3" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">有</asp:ListItem>
                                        <asp:ListItem Value="2">沒有</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <asp:RequiredFieldValidator ID="Re_R5_3" runat="server" Display="None" ErrorMessage="請選擇第五部分的問題三" ControlToValidate="RadioButtonList5_3"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr>
                            <td id="TD_R5_4" style="height: 97px" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R5_4" runat="server">4.考到證照後，對找工作有沒有幫助？ (針對上一題回答有者)</asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList5_4" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常有幫助</asp:ListItem>
                                        <asp:ListItem Value="2">有幫助</asp:ListItem>
                                        <asp:ListItem Value="3">普通</asp:ListItem>
                                        <asp:ListItem Value="4">沒幫助</asp:ListItem>
                                        <asp:ListItem Value="5">完全沒幫助</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <div><asp:RequiredFieldValidator ID="Re_R5_4" runat="server" Display="None" ErrorMessage="請選擇第五部分的問題四" ControlToValidate="RadioButtonList5_4"></asp:RequiredFieldValidator></div>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" id="TD_R6" runat="server">【第六部份：職訓與工作-5】</td>
                        </tr>
                        <tr>
                            <td id="TD_R6_1" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R6_1" runat="server">1.受訓後，有沒有找工作或換工作？</asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList6_1" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">有</asp:ListItem>
                                        <asp:ListItem Value="2">沒有</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <div><asp:RequiredFieldValidator ID="Re_R6_1" runat="server" Display="None" ErrorMessage="請選擇第六部分的問題一" ControlToValidate="RadioButtonList6_1"></asp:RequiredFieldValidator></div>
                            </td>
                        </tr>
                        <tr>
                            <td id="TD_R6_2" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R6_2" runat="server">2.受訓所學知識技能，對找工作有沒有幫助？</asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList6_2" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常有幫助</asp:ListItem>
                                        <asp:ListItem Value="2">有幫助</asp:ListItem>
                                        <asp:ListItem Value="3">普通</asp:ListItem>
                                        <asp:ListItem Value="4">沒幫助</asp:ListItem>
                                        <asp:ListItem Value="5">完全沒幫助</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <div>
                                    <asp:RequiredFieldValidator ID="Re_R6_2" runat="server" Display="None" ErrorMessage="請選擇第六部分的問題二" ControlToValidate="RadioButtonList6_2"></asp:RequiredFieldValidator>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td id="TD_R6_3" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R6_3" runat="server">3.受訓頒發的結業證書，對找工作有沒有幫助？ </asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList6_3" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常有幫助</asp:ListItem>
                                        <asp:ListItem Value="2">有幫助</asp:ListItem>
                                        <asp:ListItem Value="3">普通</asp:ListItem>
                                        <asp:ListItem Value="4">沒幫助</asp:ListItem>
                                        <asp:ListItem Value="5">完全沒幫助</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <div><asp:RequiredFieldValidator ID="Re_R6_3" runat="server" Display="None" ErrorMessage="請選擇第六部分的問題三" ControlToValidate="RadioButtonList6_3"></asp:RequiredFieldValidator></div>
                            </td>
                        </tr>
                        <tr>
                            <td id="TD_R6_4" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R6_4" runat="server">4.老師名氣，對找工作有沒有幫助？</asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList6_4" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常有幫助</asp:ListItem>
                                        <asp:ListItem Value="2">有幫助</asp:ListItem>
                                        <asp:ListItem Value="3">普通</asp:ListItem>
                                        <asp:ListItem Value="4">沒幫助</asp:ListItem>
                                        <asp:ListItem Value="5">完全沒幫助</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <div><asp:RequiredFieldValidator ID="Re_R6_4" runat="server" Display="None" ErrorMessage="請選擇第六部分的問題四" ControlToValidate="RadioButtonList6_4"></asp:RequiredFieldValidator></div>
                            </td>
                        </tr>
                        <tr>
                            <td id="TD_R6_5" runat="server" class="whitecol">
                                <div><asp:Label ID="Label_R6_5" runat="server">5.職訓機構名聲，對找工作有沒有幫助？</asp:Label></div>
                                <div>
                                    <asp:RadioButtonList ID="RadioButtonList6_5" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常有幫助</asp:ListItem>
                                        <asp:ListItem Value="2">有幫助</asp:ListItem>
                                        <asp:ListItem Value="3">普通</asp:ListItem>
                                        <asp:ListItem Value="4">沒幫助</asp:ListItem>
                                        <asp:ListItem Value="5">完全沒幫助</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                                <div><asp:RequiredFieldValidator ID="Re_R6_5" runat="server" Display="None" ErrorMessage="請選擇第六部分的問題五" ControlToValidate="RadioButtonList6_5"></asp:RequiredFieldValidator></div>
                            </td>
                        </tr>
                    </table>
                    <div align="center">
                        <asp:Button ID="Button1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="BtnBak" runat="server" Text="回上一頁" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                        <input id="Button2" onclick="history.back( );location.reload();" type="button" value="回上一頁" name="Button2" runat="server" class="asp_button_M">
                        <asp:Button ID="next_but" runat="server" Text="不儲存填寫下一位" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                    </div>
                    <div align="center"><asp:ValidationSummary ID="Summary" runat="server" DisplayMode="List" ShowSummary="False" ShowMessageBox="True"></asp:ValidationSummary></div>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>