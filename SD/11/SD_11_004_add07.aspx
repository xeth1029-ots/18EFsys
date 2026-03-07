<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_11_004_add07.aspx.vb" Inherits="TIMS.SD_11_004_add07" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_11_004_add07</title>
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
        function ChkData() {
            var msg = '';
            if (isEmpty('Q1_1')) { msg += '請選擇課程內容與工作性質是否相關\n'; }
            if (isEmpty('Q1_2')) { msg += '請選擇課程名稱是否適當\n'; }
            if (isEmpty('Q1_3')) { msg += '請選擇教材內容是否適當\n'; }
            if (isEmpty('Q1_4')) { msg += '請選擇本項訓練發給教材情形\n'; }
            if (isEmpty('Q1_5')) { msg += '請選擇發給方式\n'; }
            if (isEmpty('Q3_1')) { msg += '請選擇教師的教學態度\n'; }
            if (isEmpty('Q3_2')) { msg += '請選擇教師師資的教學方法或技巧\n'; }
            if (isEmpty('Q3_3')) { msg += '請選擇講授課程時間控制是否適當\n'; }
            if (isEmpty('Q4')) { msg += '請選擇你對整體課程瞭解的程度\n'; }
            if (isEmpty('Q7')) { msg += '請選擇參加本項訓練後，對就業安定工作是否有幫助\n'; }
            if (isEmpty('Q8')) { msg += '請選擇你對訓練單位的行政服務滿意度\n'; }
            if (isEmpty('Q9')) { msg += '請選擇整體而言，你對於參加本計畫訓練的滿意度為：\n'; }
            if (msg == '') {
                return true;
            } else {
                alert(msg);
                return false;
            }
        }

        function insert_next() {
            if (window.confirm("是否繼續新增下一筆?")) {
                location.href = 'SD_11_004_add07.aspx?ProcessType=next&ocid=' + document.getElementById("Re_OCID").value + '&SOCID=' + document.getElementById("Re_SOCID").value + '&ID=' + document.getElementById("Re_ID").value;
            } else {
                location.href = 'SD_11_004.aspx?ProcessType=Back&ocid=' + document.getElementById("Re_OCID").value + '&ID=' + document.getElementById("Re_ID").value;
            }
        }
    </script>
    <style type="text/css">
        .style1 { COLOR: #000000; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <input id="Re_OCID" type="hidden" name="Re_OCID" runat="server"><input id="Re_SOCID" type="hidden" name="Re_SOCID" runat="server">
        <asp:CustomValidator ID="CustomValidator1" runat="server" Display="None" ErrorMessage="CustomValidator" ClientValidationFunction="CheckSurvey"></asp:CustomValidator>
        <input id="ProcessType" style="width: 40px; height: 22px" type="hidden" size="1" name="ProcessType" runat="server">
        <input id="Re_ID" style="width: 40px; height: 22px" type="hidden" size="1" name="Re_ID" runat="server">
        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<FONT color="#990000">受訓學員意見調查表</FONT></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:Label ID="Label_Stud" runat="server" CssClass="font"></asp:Label>
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                <asp:Label ID="Label_Name" runat="server" CssClass="font"></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;
                                <asp:Label ID="Label_Status" runat="server"></asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="TableName" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr id="StdTr" runat="server">
                            <td width="100" class="bluecol">學員</td>
                            <td colspan="3" class="td_light"><asp:DropDownList ID="SOCID" runat="server" AutoPostBack="True" Width="112px"></asp:DropDownList></td>
                        </tr>
                    </table>
                    <table class="font" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="bluecol"><font>【第一部份：講授內容-6】</font></td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div>
                                    <font color="#ffffff">&nbsp;&nbsp;&nbsp;</font><span class="style1"><strong>1. 課程內容與工作性質是否相關？		課程內容與工作性質是否相關？</strong></span>
                                </div>
                                <div>
                                    <asp:RadioButtonList ID="Q1_1" runat="server" CssClass="font" Width="584px" RepeatDirection="Horizontal">
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
                                <div>
                                    <font color="#ffffff">&nbsp;&nbsp;&nbsp;</font><span class="style1"><strong>2. 課程名稱是否適當？</strong></span>
                                </div>
                                <div>
                                    <asp:RadioButtonList ID="Q1_2" runat="server" CssClass="font" Width="584px" RepeatDirection="Horizontal">
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
                                    <asp:RadioButtonList ID="Q1_3" runat="server" CssClass="font" Width="584px" RepeatDirection="Horizontal">
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
                                    <asp:RadioButtonList ID="Q1_4" runat="server" CssClass="font" Width="584px" RepeatDirection="Horizontal">
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
                                    <asp:RadioButtonList ID="Q1_5" runat="server" CssClass="font" Width="584px" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">訓練之前發給整套教材</asp:ListItem>
                                        <asp:ListItem Value="2">隨課程進度給單頁式講義</asp:ListItem>
                                        <asp:ListItem Value="3">其他</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div><span class="style1">&nbsp;<strong> 6. 訓練時數是否適當？</strong></span></div>
                                <div><asp:RadioButton ID="Q1_61" runat="server" Width="104px" Text="適當" Height="12px" GroupName="Q1"></asp:RadioButton></div>
                                <div>
                                    <div style="float:left;"><asp:RadioButton ID="Q1_62" runat="server" Width="160px" Text="應增加(課程名稱)" GroupName="Q1"></asp:RadioButton></div>
                                    <div style="float:left;">
                                        <asp:TextBox ID="Q1_6_CCourName" runat="server" Width="248px"></asp:TextBox>&nbsp;增加
                                        <asp:TextBox ID="Q1_6_CHour" runat="server" Width="24px"></asp:TextBox>小時
                                    </div>
                                </div>
                                <div>
                                    <div style="float:left;"><asp:RadioButton ID="Q1_63" runat="server" Width="160px" Text="應減少(課程名稱)" GroupName="Q1"></asp:RadioButton></div>
                                    <div style="float:left;">
                                        <asp:TextBox ID="Q1_6_MCourName" runat="server" Width="248px"></asp:TextBox>&nbsp;&nbsp;減少
										<asp:TextBox ID="Q1_6_MHour" runat="server" Width="24px"></asp:TextBox>小時
                                    </div>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">【第二部份：術科操作-5 (如無術科免填) 】</td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div><span class="style1">&nbsp;<strong> 1. 術科時數是否適當？</strong></span></div>
                                <div><asp:RadioButton ID="Q2_11" runat="server" Width="104px" Text="適當" Height="12px" GroupName="Q2"></asp:RadioButton></div>
                                <div>
                                    <div style="float:left;"><asp:RadioButton ID="Q2_12" runat="server" Width="160px" Text="應增加(課程名稱)" GroupName="Q2"></asp:RadioButton></div>
                                    <div style="float:left;">
                                        <asp:TextBox ID="Q2_1_CCourName" runat="server" Width="248px"></asp:TextBox>&nbsp; 增加
										<asp:TextBox ID="Q2_1_CHour" runat="server" Width="24px"></asp:TextBox>小時
                                    </div>
                                </div>
                                <div>
                                    <div style="float:left;"><asp:RadioButton ID="Q2_13" runat="server" Width="160px" Text="應減少(課程名稱)" GroupName="Q2"></asp:RadioButton></div>
                                    <div style="float:left;">
                                        <asp:TextBox ID="Q2_1_MCourName" runat="server" Width="248px"></asp:TextBox>&nbsp;&nbsp;減少
										<asp:TextBox ID="Q2_1_MHour" runat="server" Width="24px"></asp:TextBox>小時
                                    </div>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div><font color="#ffffff">&nbsp;&nbsp;&nbsp;</font><span class="style1"><strong>2. 術科內容是否適當？</strong></span></div>
                                <div>
                                    <asp:RadioButtonList ID="Q2_2" runat="server" CssClass="font" Width="584px" RepeatDirection="Horizontal">
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
                                    <asp:RadioButtonList ID="Q2_3" runat="server" CssClass="font" Width="584px" RepeatDirection="Horizontal">
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
                                    <asp:RadioButtonList ID="Q2_4" runat="server" CssClass="font" Width="584px" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">很充分</asp:ListItem>
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
                                    <asp:RadioButtonList ID="Q2_5" runat="server" CssClass="font" Width="584px" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">新穎</asp:ListItem>
                                        <asp:ListItem Value="2">尚可</asp:ListItem>
                                        <asp:ListItem Value="3">陳舊</asp:ListItem>
                                        <asp:ListItem Value="4">應淘汰</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">【第三部份：講授人員-3】</td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div>&nbsp;&nbsp; <strong>1. 教師的教學態度</strong></div>
                                <div>
                                    <asp:RadioButtonList ID="Q3_1" runat="server" CssClass="font" Width="584px" RepeatDirection="Horizontal">
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
                                    <asp:RadioButtonList ID="Q3_2" runat="server" CssClass="font" Width="584px" RepeatDirection="Horizontal">
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
                                    <asp:RadioButtonList ID="Q3_3" runat="server" CssClass="font" Width="584px" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常適當</asp:ListItem>
                                        <asp:ListItem Value="2">適當</asp:ListItem>
                                        <asp:ListItem Value="3">尚可</asp:ListItem>
                                        <asp:ListItem Value="4">不適當</asp:ListItem>
                                        <asp:ListItem Value="5">很不適當</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">【第四部份：】</td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div>&nbsp;&nbsp;&nbsp;<strong>你對整體課程瞭解的程度</strong></div>
                                <div>
                                    <asp:RadioButtonList ID="Q4" runat="server" CssClass="font" Width="584px" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常瞭解</asp:ListItem>
                                        <asp:ListItem Value="2">大致瞭解</asp:ListItem>
                                        <asp:ListItem Value="3">尚可瞭解</asp:ListItem>
                                        <asp:ListItem Value="4">不太瞭解</asp:ListItem>
                                        <asp:ListItem Value="5">幾乎不瞭解</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">【第五部份：】</td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div><span class="style1">&nbsp;&nbsp; <strong>你獲得招訓消息的來源為？</strong></span></div>
                                <div>
                                    <div style="float:left;"><asp:RadioButton ID="Q5_1" runat="server" Width="160px" Text="電視" GroupName="Q5"></asp:RadioButton></div>
                                    <div style="float:left;"><asp:RadioButton ID="Q5_2" runat="server" Width="160px" Text="廣播" GroupName="Q5"></asp:RadioButton></div>
                                    <div style="float:left;"><asp:RadioButton ID="Q5_3" runat="server" Width="160px" Text="網路" GroupName="Q5"></asp:RadioButton></div>
                                    <div style="float:left;"><asp:RadioButton ID="Q5_4" runat="server" Width="160px" Text="報紙;(報紙名稱)" GroupName="Q5"></asp:RadioButton></div>
                                    <div style="float:left;"><asp:TextBox ID="Q5_Note_News" runat="server" Width="140px"></asp:TextBox></div>
                                </div>
                                <div>
                                    <asp:RadioButton ID="Q5_5" runat="server" Width="112px" Text="服務單位" Height="20px" GroupName="Q5"></asp:RadioButton>
                                    <asp:RadioButton ID="Q5_6" runat="server" Width="56px" Text="其他" Height="20px" GroupName="Q5"></asp:RadioButton>
                                    <asp:TextBox ID="Q5_Note_Other" runat="server" Width="168px"></asp:TextBox>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">【第六部份：】</td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div>
                                    &nbsp;&nbsp;&nbsp;<strong>訓練費用總共 </strong>
                                    <asp:Label ID="Label1" runat="server" Width="40px"></asp:Label><strong>元，其中自行繳納</strong>&nbsp;
                                    <asp:Label ID="Label2" runat="server" Width="40px"></asp:Label><strong>元。 
                                    除政府補助外，自行繳納費用負擔方式(可複選)</strong>
                                </div>
                                <div>
                                    <div style="float:left;"><asp:CheckBox ID="Q6_3" runat="server" Width="120px" Text="兩者皆有"></asp:CheckBox></div>
                                    <div style="float:left;"><asp:CheckBox ID="Q6_1" runat="server" Width="120px" Text="自行負擔"></asp:CheckBox>&nbsp;</div>
                                    <div style="float:left;"><asp:TextBox ID="Q6_Note1" runat="server" Width="120px"></asp:TextBox>&nbsp;</div>
                                    <div style="float:left;"><asp:CheckBox ID="Q6_2" runat="server" Width="120px" Text="服務單位負擔"></asp:CheckBox>&nbsp;</div>
                                    <div style="float:left;"><asp:TextBox ID="Q6_Note2" runat="server" Width="120px"></asp:TextBox></div>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">【第七部份：】</td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div>&nbsp;&nbsp;&nbsp;<strong>參加本項訓練後，對就業安定工作是否有幫助？</strong></div>
                                <div>
                                    <asp:RadioButtonList ID="Q7" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">對初次工作有幫助</asp:ListItem>
                                        <asp:ListItem Value="2">對目前工作效率有幫助</asp:ListItem>
                                        <asp:ListItem Value="3">對轉換工作有幫助</asp:ListItem>
                                        <asp:ListItem Value="4">對調昇職位有幫助</asp:ListItem>
                                        <asp:ListItem Value="5">其他</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">【第八部份：】</td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div>&nbsp;&nbsp; <strong>你對訓練單位的行政服務滿意度</strong></div>
                                <div>
                                    <asp:RadioButtonList ID="Q8" runat="server" CssClass="font" Width="584px" RepeatDirection="Horizontal">
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
                            <td class="bluecol">【第九部份：】</td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <div>&nbsp;&nbsp; <strong>你整體而言，你對於參加本計畫訓練的滿意度為：</strong></div>
                                <div>
                                    <asp:RadioButtonList ID="Q9" runat="server" CssClass="font" Width="584px" RepeatDirection="Horizontal">
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
                            <td class="bluecol">【第十部份：】</td>
                        </tr>
                        <tr>
                            <td class="td_light">
                                <strong>其他建議</strong>
                                <div><asp:TextBox ID="Q9_Note" runat="server" Width="584px" Height="76px" TextMode="MultiLine"></asp:TextBox></div>
                            </td>
                        </tr>
                    </table>
                    <div align="center">
                        <asp:Button ID="Button1" runat="server" Text="儲存" CausesValidation="False" CssClass="asp_button_M"></asp:Button>&nbsp;
                        <asp:Button ID="Button2" runat="server" Text="回上一頁" CausesValidation="False" CssClass="asp_button_M"></asp:Button>&nbsp;
                        <asp:Button ID="next_but" runat="server" Text="不儲存填寫下一位" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                    </div>
                    <div align="center"><asp:ValidationSummary ID="Summary" runat="server" Width="577px" DisplayMode="List" ShowSummary="False" ShowMessageBox="True"></asp:ValidationSummary></div>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>