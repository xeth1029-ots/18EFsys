<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_11_004_add12.aspx.vb" Inherits="WDAIIP.SD_11_004_add12" %>

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
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function ChkData() {
            var msg = '';
            if (isEmpty('Q1_1')) { msg += '1-1.課程內容與工作性質是否相關\n'; }
            if (isEmpty('Q1_2')) { msg += '1-2.課程名稱是否適當\n'; }
            if (isEmpty('Q1_3')) { msg += '1-3.教材內容是否適當\n'; }
            if (isEmpty('Q1_4')) { msg += '1-4.本項訓練發給教材情形\n'; }
            if (isEmpty('Q1_5')) { msg += '1-5.發給方式\n'; }
            //debugger;
            if (isEmpty('Q1_61') && isEmpty('Q1_62') && isEmpty('Q1_63')) {
                msg += '1-6.訓練時數是否適當\n';
            }
            if (!isEmpty('Q1_62') && (isEmpty('Q1_6_CCourName') || isEmpty('Q1_6_CHour'))) {
                msg += '1-6A.訓練時數是否適當，應增加(課程名稱)與時數\n';
            }
            if (!isEmpty('Q1_63') && (isEmpty('Q1_6_MCourName') || isEmpty('Q1_6_MHour'))) {
                msg += '1-6B.訓練時數是否適當，應減少(課程名稱)與時數\n';
            }
            /*
			//無術科免填
			if (isEmpty('Q2_11') && isEmpty('Q2_12') && isEmpty('Q2_13')) {
			msg += '【第二部份】術科時數是否適當\n'; 
			}
			if (!isEmpty('Q2_12') && (isEmpty('Q2_1_CCourName') || isEmpty('Q2_1_CHour'))) {
			msg += '【第二部份】術科時數是否適當，應增加(課程名稱)與時數\n'; 
			}
			if (!isEmpty('Q2_13') && (isEmpty('Q2_1_MCourName') || isEmpty('Q2_1_MHour'))) {
			msg += '【第二部份】術科時數是否適當，應減少(課程名稱)與時數\n'; 
			}
			if (isEmpty('Q2_2')) { msg += '【第二部份】術科內容是否適當\n'; }
			if (isEmpty('Q2_3')) { msg += '【第二部份】術科操作解說是否充分\n'; }
			if (isEmpty('Q2_4')) { msg += '【第二部份】訓練設備是否充足\n'; }
			if (isEmpty('Q2_5')) { msg += '【第二部份】訓練設備現狀\n'; }
			*/
            if (isEmpty('Q3_1')) { msg += '3-1.請選擇教師的教學態度\n'; }
            if (isEmpty('Q3_2')) { msg += '3-2.請選擇教師師資的教學方法或技巧\n'; }
            if (isEmpty('Q3_3')) { msg += '3-3.請選擇講授課程時間控制是否適當\n'; }
            if (isEmpty('Q4')) { msg += '4.請選擇你瞭解整體課程內容\n'; }
            if (isEmpty('Q5_1') && isEmpty('Q5_2') && isEmpty('Q5_3')
					 && isEmpty('Q5_4') && isEmpty('Q5_5') && isEmpty('Q5_6')) {
                msg += '5.你獲得招訓消息的來源為\n';
            }
            if (!isEmpty('Q5_4') && isEmpty('Q5_Note_News')) {
                msg += '5A.你獲得招訓消息的來源，報紙;(報紙名稱) \n';
            }
            if (!isEmpty('Q5_6') && isEmpty('Q5_Note_Other')) {
                msg += '5B.你獲得招訓消息的來源，其他 \n';
            }
            if (isEmpty('Q6_1') && isEmpty('Q6_2') && isEmpty('Q6_3')) {
                msg += '6.訓練費用\n';
            }
            if (!isEmpty('Q6_1') && isEmpty('Q6_Note1')) {
                msg += '6A.訓練費用，自行負擔 \n';
            }
            if (!isEmpty('Q6_2') && isEmpty('Q6_Note2')) {
                msg += '6B.訓練費用，服務單位負擔 \n';
            }
            if (isEmpty('Q7')) { msg += '7.參加本項訓練後，你能掌握訓練課程所教授知識或技能\n'; }
            if (isEmpty('Q7_8')) { msg += '8.參加本項訓練後，你有把握自已能所學的知識應用到工作上\n'; }
            if (isEmpty('Q7_9')) { msg += '9.完成訓練後，你願意找機會將所學8的知識／技能應用在工作中\n'; }
            if (isEmpty('Q8')) { msg += '10.你對訓練單位的行政服務滿意度\n'; }
            if (isEmpty('Q9_1')) { msg += '11-1.瞭解補助對象\n'; }
            if (isEmpty('Q9_2')) { msg += '11-2.瞭解補助經費標準\n'; }
            if (isEmpty('Q9_3')) { msg += '11-3.瞭解補助流程\n'; }
            if (isEmpty('Q10')) { msg += '12.整體而言，你對於參加本計畫訓練的滿意度\n'; }
            if (msg == '') {
                return true;
            } else {
                msg = '請確認下列答案：\n' + msg;
                alert(msg);
                return false;
            }
        }

        function insert_next() {
            var mainpageUrl = 'SD_11_004.aspx'
            var pageUrl = 'SD_11_004_add12.aspx'
            if (window.confirm("是否繼續新增下一筆?")) {
                location.href = pageUrl + '?ProcessType=Next&ocid=' + document.getElementById("Re_OCID").value + '&SOCID=' + document.getElementById("Re_SOCID").value + '&ID=' + document.getElementById("Re_ID").value;
            }
            else {
                location.href = mainpageUrl + '?ProcessType=Back&ocid=' + document.getElementById("Re_OCID").value + '&ID=' + document.getElementById("Re_ID").value;
            }
        }
    </script>
    <style type="text/css">
        .style1 { color: #000000; }
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
                        <%--<tr>
						<td>
							<asp:Label ID="TitleLab1" runat="server"></asp:Label>
							<asp:Label ID="TitleLab2" runat="server">
									首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<FONT color="#990000">受訓學員意見調查表</FONT>
							</asp:Label>
						</td>
					    </tr>--%>
                        <tr>
                            <td>
                                <asp:Label ID="Label_Stud" runat="server" CssClass="font"></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                <asp:Label ID="Label_Name" runat="server" CssClass="font"></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;
                                <asp:Label ID="Label_Status" runat="server"></asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="TableName" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr id="StdTr" runat="server">
                            <td class="bluecol" width="20%">學員</td>
                            <td colspan="3" class="whitecol">
                                <asp:DropDownList ID="SOCID" runat="server" AutoPostBack="True"></asp:DropDownList></td>
                        </tr>
                    </table>
                    <table class="font" id="Table3_Datalist" runat="server" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="table_title">【第一部份：講授內容-6】</td>
                        </tr>
                        <tr bgcolor="#ecf7ff">
                            <td class="td_light">
                                <div><font color="#ffffff">&nbsp;&nbsp;&nbsp;</font><span class="style1"><strong>1. 課程內容與工作性質是否相關？</strong></span></div>
                                <div>
                                    <asp:RadioButtonList ID="Q1_1" runat="server" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常相關</asp:ListItem>
                                        <asp:ListItem Value="2">相關</asp:ListItem>
                                        <asp:ListItem Value="3">尚可</asp:ListItem>
                                        <asp:ListItem Value="4">不相關</asp:ListItem>
                                        <asp:ListItem Value="5">非常不相關</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div><font color="#ffffff">&nbsp;&nbsp;&nbsp;</font><span class="style1"><strong>2. 課程名稱是否適當？</strong></span></div>
                                <div>
                                    <asp:RadioButtonList ID="Q1_2" runat="server" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常適當</asp:ListItem>
                                        <asp:ListItem Value="2">適當</asp:ListItem>
                                        <asp:ListItem Value="3">尚可</asp:ListItem>
                                        <asp:ListItem Value="4">不適當</asp:ListItem>
                                        <asp:ListItem Value="5">非常不適當</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div><span class="style1">&nbsp;<strong> 3. 教材內容是否適當？</strong></span></div>
                                <div>
                                    <asp:RadioButtonList ID="Q1_3" runat="server" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常適當</asp:ListItem>
                                        <asp:ListItem Value="2">適當</asp:ListItem>
                                        <asp:ListItem Value="3">尚可</asp:ListItem>
                                        <asp:ListItem Value="4">不適當</asp:ListItem>
                                        <asp:ListItem Value="5">非常不適當</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div><span class="style1">&nbsp;<strong> 4. 本項訓練發給教材情形：</strong></span></div>
                                <div>
                                    <asp:RadioButtonList ID="Q1_4" runat="server" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">僅發給教科書</asp:ListItem>
                                        <asp:ListItem Value="2">僅發給講義</asp:ListItem>
                                        <asp:ListItem Value="3">發給教科書與講義</asp:ListItem>
                                        <asp:ListItem Value="4">兩者均無</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div><span class="style1">&nbsp;<strong> 5. 發給方式：</strong></span></div>
                                <div>
                                    <asp:RadioButtonList ID="Q1_5" runat="server" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">訓練之前發給整套教材</asp:ListItem>
                                        <asp:ListItem Value="2">隨課程進度給單頁式講義</asp:ListItem>
                                        <asp:ListItem Value="3">其他</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td style="background-color: #e9f1fe" class="whitecol">
                                <div><span class="style1">&nbsp;<strong> 6. 訓練時數是否適當？</strong></span></div>
                                <div>
                                    <asp:RadioButton ID="Q1_61" runat="server" GroupName="Q1" Text="適當"></asp:RadioButton>
                                </div>
                                <div>
                                    <asp:RadioButton ID="Q1_62" runat="server" GroupName="Q1" Text="應增加(課程名稱)"></asp:RadioButton>
                                    <asp:TextBox ID="Q1_6_CCourName" runat="server" Width="25%" MaxLength="100"></asp:TextBox>&nbsp;增加
                                    <asp:TextBox ID="Q1_6_CHour" runat="server" Width="10%" MaxLength="5"></asp:TextBox>小時
                                </div>
                                <div>
                                    <asp:RadioButton ID="Q1_63" runat="server" GroupName="Q1" Text="應減少(課程名稱)"></asp:RadioButton>
                                    <asp:TextBox ID="Q1_6_MCourName" runat="server" Width="25%" MaxLength="100"></asp:TextBox>&nbsp;&nbsp;減少
                                    <asp:TextBox ID="Q1_6_MHour" runat="server" Width="10%" MaxLength="5"></asp:TextBox>小時
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="table_title">【第二部份：術科操作-5 (如無術科免填) 】</td>
                        </tr>
                        <tr>
                            <td style="background-color: #e9f1fe" class="whitecol">
                                <div><span class="style1">&nbsp;<strong> 1. 術科時數是否適當？</strong></span></div>
                                <div>
                                    <asp:RadioButton ID="Q2_11" runat="server" GroupName="Q2" Text="適當"></asp:RadioButton>
                                </div>
                                <div>
                                    <asp:RadioButton ID="Q2_12" runat="server" GroupName="Q2" Text="應增加(課程名稱)"></asp:RadioButton>
                                    <asp:TextBox ID="Q2_1_CCourName" runat="server" Width="248px" MaxLength="100"></asp:TextBox>&nbsp;增加
                                    <asp:TextBox ID="Q2_1_CHour" runat="server" Width="10%" MaxLength="5"></asp:TextBox>小時
                                </div>
                                <div>
                                    <asp:RadioButton ID="Q2_13" runat="server" GroupName="Q2" Text="應減少(課程名稱)"></asp:RadioButton>
                                    <asp:TextBox ID="Q2_1_MCourName" runat="server" Width="248px" MaxLength="100"></asp:TextBox>&nbsp;&nbsp;減少
                                    <asp:TextBox ID="Q2_1_MHour" runat="server" Width="10%" MaxLength="5"></asp:TextBox>小時
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div><font color="#ffffff">&nbsp;&nbsp;&nbsp;</font><span class="style1"><strong>2. 術科內容是否適當？</strong></span></div>
                                <div>
                                    <asp:RadioButtonList ID="Q2_2" runat="server" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">太多</asp:ListItem>
                                        <asp:ListItem Value="2">適當</asp:ListItem>
                                        <asp:ListItem Value="3">太少</asp:ListItem>
                                        <asp:ListItem Value="4">其他</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div><font color="#ffffff">&nbsp;&nbsp;&nbsp;</font><span class="style1"><strong>3. 術科操作解說是否充分？</strong></span></div>
                                <div>
                                    <asp:RadioButtonList ID="Q2_3" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">很充分</asp:ListItem>
                                        <asp:ListItem Value="2">尚可</asp:ListItem>
                                        <asp:ListItem Value="3">需改善</asp:ListItem>
                                        <asp:ListItem Value="4">其他</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div><font color="#ffffff">&nbsp;&nbsp;&nbsp;</font><span class="style1"><strong>4. 訓練設備是否充足？</strong></span></div>
                                <div>
                                    <asp:RadioButtonList ID="Q2_4" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">很充足</asp:ListItem>
                                        <asp:ListItem Value="2">尚可</asp:ListItem>
                                        <asp:ListItem Value="3">應再充實</asp:ListItem>
                                        <asp:ListItem Value="4">其他</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div><font color="#ffffff">&nbsp;&nbsp;&nbsp;</font><span class="style1"><strong>5. 訓練設備現狀 :</strong></span></div>
                                <div>
                                    <asp:RadioButtonList ID="Q2_5" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">新穎</asp:ListItem>
                                        <asp:ListItem Value="2">尚可</asp:ListItem>
                                        <asp:ListItem Value="3">陳舊</asp:ListItem>
                                        <asp:ListItem Value="4">應淘汰</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="table_title">【第三部份：講授人員-3】</td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div>&nbsp;&nbsp; <strong>1. 教師的教學態度</strong></div>
                                <div>
                                    <asp:RadioButtonList ID="Q3_1" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                        <asp:ListItem Value="2">滿意</asp:ListItem>
                                        <asp:ListItem Value="3">尚可</asp:ListItem>
                                        <asp:ListItem Value="4">不太滿意</asp:ListItem>
                                        <asp:ListItem Value="5">很不滿意</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div>&nbsp;&nbsp; <strong>2. 教師師資的教學方法或技巧</strong></div>
                                <div>
                                    <asp:RadioButtonList ID="Q3_2" runat="server" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                        <asp:ListItem Value="2">滿意</asp:ListItem>
                                        <asp:ListItem Value="3">尚可</asp:ListItem>
                                        <asp:ListItem Value="4">不太滿意</asp:ListItem>
                                        <asp:ListItem Value="5">很不滿意</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div>&nbsp;&nbsp; <strong>3. 講授課程時間控制是否適當？</strong></div>
                                <div>
                                    <asp:RadioButtonList ID="Q3_3" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常適當</asp:ListItem>
                                        <asp:ListItem Value="2">適當</asp:ListItem>
                                        <asp:ListItem Value="3">尚可</asp:ListItem>
                                        <asp:ListItem Value="4">不適當</asp:ListItem>
                                        <asp:ListItem Value="5">非常不適當</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="table_title">【第四部份：】</td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div>&nbsp;&nbsp;&nbsp;<strong>你瞭解整體課程內容</strong></div>
                                <div>
                                    <asp:RadioButtonList ID="Q4" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常同意</asp:ListItem>
                                        <asp:ListItem Value="2">同意</asp:ListItem>
                                        <asp:ListItem Value="3">無意見</asp:ListItem>
                                        <asp:ListItem Value="4">不同意</asp:ListItem>
                                        <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="table_title">【第五部份：】</td>
                        </tr>
                        <tr>
                            <td style="background-color: #e9f1fe" class="whitecol">
                                <div><span class="style1">&nbsp;&nbsp; <strong>你獲得招訓消息的來源為？</strong></span></div>
                                <div>
                                    <asp:RadioButton ID="Q5_1" runat="server" GroupName="Q5" Text="電視"></asp:RadioButton>
                                    <asp:RadioButton ID="Q5_2" runat="server" GroupName="Q5" Text="廣播"></asp:RadioButton>
                                    <asp:RadioButton ID="Q5_3" runat="server" GroupName="Q5" Text="網路"></asp:RadioButton>
                                    <asp:RadioButton ID="Q5_4" runat="server" GroupName="Q5" Text="報紙;(報紙名稱)"></asp:RadioButton>
                                    <asp:TextBox ID="Q5_Note_News" runat="server" Width="20%" MaxLength="100"></asp:TextBox>
                                </div>
                                <div>
                                    <asp:RadioButton ID="Q5_5" runat="server" GroupName="Q5" Text="服務單位"></asp:RadioButton>
                                    <asp:RadioButton ID="Q5_6" runat="server" GroupName="Q5" Text="其他"></asp:RadioButton>
                                    <asp:TextBox ID="Q5_Note_Other" runat="server" Width="20%" MaxLength="100"></asp:TextBox>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="table_title">【第六部份：】</td>
                        </tr>
                        <tr>
                            <td style="background-color: #e9f1fe" class="whitecol">
                                <div>
                                    &nbsp;&nbsp;&nbsp;<strong>訓練費用總共 </strong>
                                    <asp:Label ID="Label1" runat="server" Width="40px"></asp:Label><strong>元，其中自行繳納</strong>&nbsp;
                                    <asp:Label ID="Label2" runat="server" Width="40px"></asp:Label><strong>元。 除政府補助外，自行繳納費用負擔方式(可複選)</strong>
                                </div>
                                <div>
                                    <asp:CheckBox ID="Q6_1" runat="server" Text="自行負擔"></asp:CheckBox>
                                    <asp:TextBox ID="Q6_Note1" runat="server" Width="10%" MaxLength="100"></asp:TextBox>元&nbsp; &nbsp; &nbsp;&nbsp;
                                    <asp:CheckBox ID="Q6_2" runat="server" Text="服務單位負擔"></asp:CheckBox>
                                    <asp:TextBox ID="Q6_Note2" runat="server" Width="10%" MaxLength="100"></asp:TextBox>元&nbsp;&nbsp; &nbsp; &nbsp;&nbsp;
                                    <asp:CheckBox ID="Q6_3" runat="server" Text="兩者皆有"></asp:CheckBox>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="table_title">【第七部份：】</td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div>&nbsp;&nbsp;&nbsp;<strong>參加本項訓練後，你能掌握訓練課程所教授知識或技能</strong></div>
                                <div>
                                    <asp:RadioButtonList ID="Q7" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常同意</asp:ListItem>
                                        <asp:ListItem Value="2">同意</asp:ListItem>
                                        <asp:ListItem Value="3">無意見</asp:ListItem>
                                        <asp:ListItem Value="4">不同意</asp:ListItem>
                                        <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="table_title">【第八部份：】</td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div>&nbsp;&nbsp; <strong>參加本項訓練後，你有把握自已能所學的知識應用到工作上</strong></div>
                                <div>
                                    <asp:RadioButtonList ID="Q7_8" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常有信心</asp:ListItem>
                                        <asp:ListItem Value="2">有信心</asp:ListItem>
                                        <asp:ListItem Value="3">尚可</asp:ListItem>
                                        <asp:ListItem Value="4">沒信心</asp:ListItem>
                                        <asp:ListItem Value="5">完全沒信心</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="table_title">【第九部份：】</td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div>&nbsp;&nbsp; <strong>完成訓練後，你願意找機會將所學的知識／技能應用在工作中</strong></div>
                                <div>
                                    <asp:RadioButtonList ID="Q7_9" runat="server" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常願意</asp:ListItem>
                                        <asp:ListItem Value="2">願意</asp:ListItem>
                                        <asp:ListItem Value="3">無意見</asp:ListItem>
                                        <asp:ListItem Value="4">不願意</asp:ListItem>
                                        <asp:ListItem Value="5">非常不願意</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="table_title">【第十部份：】</td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div>&nbsp;&nbsp; <strong>你對訓練單位的行政服務滿意度</strong></div>
                                <div>
                                    <asp:RadioButtonList ID="Q8" runat="server" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                        <asp:ListItem Value="2">滿意</asp:ListItem>
                                        <asp:ListItem Value="3">尚可</asp:ListItem>
                                        <asp:ListItem Value="4">不太滿意</asp:ListItem>
                                        <asp:ListItem Value="5">很不滿意</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="table_title">【第11部份：你瞭解<asp:Label ID="Label4" runat="server"></asp:Label>規劃設計】</td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div>&nbsp;&nbsp;&nbsp;<strong>瞭解補助對象</strong></div>
                                <div>
                                    <asp:RadioButtonList ID="Q9_1" runat="server" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常同意</asp:ListItem>
                                        <asp:ListItem Value="2">同意</asp:ListItem>
                                        <asp:ListItem Value="3">無意見</asp:ListItem>
                                        <asp:ListItem Value="4">不同意</asp:ListItem>
                                        <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div>&nbsp;&nbsp; <strong>瞭解補助經費標準</strong></div>
                                <div>
                                    <strong>
                                        <asp:RadioButtonList ID="Q9_2" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="1">非常同意</asp:ListItem>
                                            <asp:ListItem Value="2">同意</asp:ListItem>
                                            <asp:ListItem Value="3">無意見</asp:ListItem>
                                            <asp:ListItem Value="4">不同意</asp:ListItem>
                                            <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </strong>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div>&nbsp;&nbsp; <strong>瞭解補助流程</strong></div>
                                <div>
                                    <strong>
                                        <asp:RadioButtonList ID="Q9_3" runat="server" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="1">非常同意</asp:ListItem>
                                            <asp:ListItem Value="2">同意</asp:ListItem>
                                            <asp:ListItem Value="3">無意見</asp:ListItem>
                                            <asp:ListItem Value="4">不同意</asp:ListItem>
                                            <asp:ListItem Value="5">非常不同意</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </strong>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="table_title">【第12部份：】</td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div>&nbsp;&nbsp; <strong>整體而言，你對於<asp:Label ID="Label5" runat="server"></asp:Label>是否滿意</strong></div>
                                <div>
                                    <strong>
                                        <asp:RadioButtonList ID="Q10" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                            <asp:ListItem Value="2">滿意</asp:ListItem>
                                            <asp:ListItem Value="3">尚可</asp:ListItem>
                                            <asp:ListItem Value="4">不太滿意</asp:ListItem>
                                            <asp:ListItem Value="5">很不滿意</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </strong>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="table_title">【第13部份：】</td>
                        </tr>
                        <tr>
                            <td style="background-color: #e9f1fe" class="whitecol">&nbsp;&nbsp; <strong>若無補助訓練計畫經費，你每年願意以自費參加相關訓練課程之金額
								<asp:Label ID="Label3" runat="server"></asp:Label>
                                <asp:TextBox ID="Q11" runat="server" Width="20%" MaxLength="10"></asp:TextBox>元</strong>
                            </td>
                        </tr>
                        <tr>
                            <td class="table_title">【第14部份：】</td>
                        </tr>
                        <tr>
                            <td class="td_light">&nbsp;&nbsp; <strong>你對訓練單位提供之無障礙訓練環境是否滿意？（領有身心障礙手冊之學員填寫，惠請填寫） </strong><strong>
                                <asp:RadioButtonList ID="Q14" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="1">非常滿意</asp:ListItem>
                                    <asp:ListItem Value="2">滿意</asp:ListItem>
                                    <asp:ListItem Value="3">尚可</asp:ListItem>
                                    <asp:ListItem Value="4">不太滿意</asp:ListItem>
                                    <asp:ListItem Value="5">很不滿意</asp:ListItem>
                                </asp:RadioButtonList></strong>
                            </td>
                        </tr>
                        <tr>
                            <td class="table_title">【第15部份：】</td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <strong>其他建議</strong>
                                <div>
                                    <asp:TextBox ID="Q12" runat="server" Width="584px" Height="80px" TextMode="MultiLine" MaxLength="200"></asp:TextBox>
                                </div>
                            </td>
                        </tr>
                    </table>
                    <div align="center" class="whitecol">
                        <asp:Button ID="Button1" runat="server" Text="儲存" CausesValidation="False" CssClass="asp_button_M"></asp:Button>&nbsp;
                        <asp:Button ID="Button2" runat="server" Text="回上一頁" CausesValidation="False" CssClass="asp_button_M"></asp:Button>&nbsp;
                        <asp:Button ID="next_but" runat="server" Text="不儲存填寫下一位" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                    </div>
                    <div align="center">
                        <asp:ValidationSummary ID="Summary" runat="server" Width="577px" DisplayMode="List" ShowSummary="False" ShowMessageBox="True"></asp:ValidationSummary>
                    </div>
                </td>
            </tr>
        </table>
        <input id="HidDASOURCE" type="hidden" name="HidDASOURCE" runat="server" size="1">
    </form>
</body>
</html>
