<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_11_005_add08.aspx.vb" Inherits="TIMS.SD_11_005_add08" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_11_005_add08</title>
    <meta http-equiv="Content-Type" content="text/html; charset=big5">
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script src="../../js/common.js"></script>
    <script language="javascript">
        function ChkData() {
            var msg = '';
            if (isEmpty('Q1')) { msg += '請選擇 1.學員近況?\n'; }
            if (isEmpty('Q2')) { msg += '請選擇 2.結訓後薪資是否提升?\n'; }
            if (isEmpty('Q3')) { msg += '請選擇 3.職位是否變化?\n'; }
            if (isEmpty('Q4')) { msg += '請選擇 4.工作情況是否滿意?\n'; }
            if (isEmpty('Q5')) { msg += '請選擇 5.工作內容是否與參訓課程內容相關?\n'; }
            if (isEmpty('Q6')) { msg += '請選擇 6.參加訓練對目前的工作是否有幫助?\n'; }
            if (isEmpty('Q7')) { msg += '請選擇 7.承上題，參加本項訓練對學員的幫助是在哪方面 ?\n'; }
            if (isEmpty('Q8')) { msg += '請選擇 8.是否有繼續參與進修訓練的意願?\n'; }
            if (msg == '') {
                return true;
            } else {
                alert(msg);
                return false;
            }
        }

        function insert_next() {
            if (window.confirm("是否繼續新增下一筆?")) {
                location.href = 'SD_11_005_add08.aspx?ProcessType=next&ocid=' + document.getElementById("Re_OCID").value + '&SOCID=' + document.getElementById("Re_SOCID").value + '&ID=' + document.getElementById("Re_ID").value;
            } else {
                location.href = 'SD_11_005.aspx?ProcessType=Back&ocid=' + document.getElementById("Re_OCID").value + '&ID=' + document.getElementById("Re_ID").value;
            }
        }
    </script>
    <style type="text/css">
        .style1 {
            COLOR: #000000;
        }
        .style2 {
            COLOR: #ffffff;
        }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <input id="Re_OCID" type="hidden" name="Re_OCID" runat="server">
        <input id="Re_SOCID" type="hidden" name="Re_SOCID" runat="server">
        <asp:CustomValidator ID="CustomValidator1" runat="server" ClientValidationFunction="CheckSurvey" ErrorMessage="CustomValidator" Display="None"></asp:CustomValidator><input id="ProcessType" style="width: 40px; height: 22px" type="hidden" size="1" name="ProcessType" runat="server">
        <input id="Re_ID" style="width: 40px; height: 22px" type="hidden" size="1" name="Re_ID" runat="server">
        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="600" border="0">
            <tr>
                <td>
                    <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <%--
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">受訓學員意見調查表</font></asp:Label>
                            </td>
                        </tr>
                        --%>
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
                            <td width="100" bgcolor="#2aafc0"><font color="#ffffff">&nbsp;學員</font></td>
                            <td bgcolor="#ecf7ff" colspan="3"><asp:DropDownList ID="SOCID" runat="server" AutoPostBack="True"></asp:DropDownList></td>
                        </tr>
                    </table>
                    <table class="font" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td bgcolor="#2aafc0"><font color="#ffffff"><strong>【一、學員部份:】</strong></font></td>
                        </tr>
                        <tr bgcolor="#ecf7ff">
                            <td bgcolor="#ecf7ff">
                                <div><span class="style1"><strong>1.學員目前的近況為何?</strong></span></div>
                                <div>
                                    <asp:RadioButtonList ID="Q1" runat="server" CssClass="font" Width="584px" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">留任原公司</asp:ListItem>
                                        <asp:ListItem Value="2">轉換至同產業公司</asp:ListItem>
                                        <asp:ListItem Value="3">轉換至不同產業的公司</asp:ListItem>
                                        <asp:ListItem Value="4">待業中</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td bgcolor="#ecf7ff">
                                <div><span class="style1"><strong>2.學員於結訓後薪資有提升嗎 ?</strong></span></div>
                                <div>
                                    <asp:RadioButtonList ID="Q2" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">大幅提升</asp:ListItem>
                                        <asp:ListItem Value="2">小幅提升</asp:ListItem>
                                        <asp:ListItem Value="3">沒有變化</asp:ListItem>
                                        <asp:ListItem Value="4">小幅減少</asp:ListItem>
                                        <asp:ListItem Value="5">大幅減少</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr bgcolor="#ecf7ff">
                            <td bgcolor="#ecf7ff">
                                <div><span class="style1"><strong>3.學員的職位有變化嗎 ?</strong></span></div>
                                <div>
                                    <asp:RadioButtonList ID="Q3" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">升遷</asp:ListItem>
                                        <asp:ListItem Value="2">調職</asp:ListItem>
                                        <asp:ListItem Value="3">沒有變化</asp:ListItem>
                                        <asp:ListItem Value="4">降職</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr bgcolor="#ecf7ff">
                            <td bgcolor="#ecf7ff">
                                <div><span class="style1"><strong>4.學員的工作滿意度是否提升?</strong></span></div>
                                <div>
                                    <asp:RadioButtonList ID="Q4" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">大幅提升</asp:ListItem>
                                        <asp:ListItem Value="2">小幅提升</asp:ListItem>
                                        <asp:ListItem Value="3">沒有變化</asp:ListItem>
                                        <asp:ListItem Value="4">小幅降低</asp:ListItem>
                                        <asp:ListItem Value="5">大幅降低</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr bgcolor="#ecf7ff">
                            <td bgcolor="#ecf7ff">
                                <div><span class="style1"><strong>5.學員目前的工作內容是否與參訓課程內容相關 ?</strong></span></div>
                                <div>
                                    <asp:RadioButtonList ID="Q5" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常相關</asp:ListItem>
                                        <asp:ListItem Value="2">相關</asp:ListItem>
                                        <asp:ListItem Value="3">尚可</asp:ListItem>
                                        <asp:ListItem Value="4">不相關</asp:ListItem>
                                        <asp:ListItem Value="5">非常不相關</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr bgcolor="#ecf7ff">
                            <td bgcolor="#ecf7ff">
                                <div><span class="style1"><strong>6.學員認為參加訓練對工作表現是否有幫助 ?</strong></span></div>
                                <asp:RadioButtonList ID="Q6" runat="server" CssClass="font" Width="576px" RepeatDirection="Horizontal" Font-Size="X-Small">
                                    <asp:ListItem Value="1">幫助非常大</asp:ListItem>
                                    <asp:ListItem Value="2">幫助頗多</asp:ListItem>
                                    <asp:ListItem Value="3">有幫助</asp:ListItem>
                                    <asp:ListItem Value="4">幫助有限</asp:ListItem>
                                    <asp:ListItem Value="5">完全沒幫助</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td bgcolor="#ecf7ff">
                                <div><span class="style1"><strong>7.承上題，參加本項訓練對學員的幫助是在哪方面 ?</strong></span></div>
                                <div>
                                    <asp:RadioButtonList ID="Q7" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">對適應工作環境有幫助</asp:ListItem>
                                        <asp:ListItem Value="2">對目前工作績效有幫助</asp:ListItem>
                                        <asp:ListItem Value="3">對轉換工作跑道有幫助</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td bgcolor="#ecf7ff">
                                <div><strong>8<span class="style1">8.學員是否有繼續參與進修訓練的意願 ?</span></strong></div>
                                <div>
                                    <asp:RadioButtonList ID="Q8" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">非常想參與</asp:ListItem>
                                        <asp:ListItem Value="2">想參與</asp:ListItem>
                                        <asp:ListItem Value="3">尚無想法</asp:ListItem>
                                        <asp:ListItem Value="4">不想參與</asp:ListItem>
                                        <asp:ListItem Value="5">非常不想參與</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td bgcolor="#ecf7ff">
                                <div><strong>9.學員認為還需要加強哪方面的專業知識使工作進行得更順利 ?</strong></div>
                                <div>
                                    1.<asp:TextBox ID="Q9_1_Note" runat="server" MaxLength="100"></asp:TextBox>
                                    2.<asp:TextBox ID="Q9_2_Note" runat="server" MaxLength="100"></asp:TextBox>
                                    3.<asp:TextBox ID="Q9_3_Note" runat="server" MaxLength="100"></asp:TextBox>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td bgcolor="#ecf7ff">
                                <div><strong>10.學員常和本課程的哪些學員、教師或職員聯絡 ?</strong></div>
                                <div>
                                    1.<asp:TextBox ID="Q10_1_Note" runat="server" MaxLength="100"></asp:TextBox>
                                    2.<asp:TextBox ID="Q10_2_Note" runat="server" MaxLength="100"></asp:TextBox>
                                    3.<asp:TextBox ID="Q10_3_Note" runat="server" MaxLength="100"></asp:TextBox>
                                </div>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="Table4" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td bgcolor="#2aafc0"><font color="#ffffff"><strong>【二、企業部份:針對企業專班調查】</strong></font></td>
                        </tr>
                    </table>
                    <table class="font" id="Table5" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td bgcolor="#2aafc0"><font color="#ffffff">&nbsp;企業名稱</font></td>
                            <td bgcolor="#ecf7ff" colspan="3"><asp:TextBox ID="BusName" runat="server" Width="208px"></asp:TextBox></td>
                        </tr>
                    </table>
                    <table class="font" id="Table6" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td bgcolor="#ecf7ff">
                                <div><span class="style1"><strong>1.學員受訓後工作態度是否有改善 ?</strong></span></div>
                                <div>
                                    <asp:RadioButtonList ID="Q11" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">明顯改善</asp:ListItem>
                                        <asp:ListItem Value="2">略有改善</asp:ListItem>
                                        <asp:ListItem Value="3">沒有變化</asp:ListItem>
                                        <asp:ListItem Value="4">略有惡化</asp:ListItem>
                                        <asp:ListItem Value="5">明顯惡化</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td bgcolor="#ecf7ff">
                                <div><span class="style1"><strong>2.學員受訓後知識技術是否有提升 ?</strong></span></div>
                                <div>
                                    <asp:RadioButtonList ID="Q12" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">明顯提升</asp:ListItem>
                                        <asp:ListItem Value="2">略有提升</asp:ListItem>
                                        <asp:ListItem Value="3">沒有提升</asp:ListItem>
                                        <asp:ListItem Value="4">略有下降</asp:ListItem>
                                        <asp:ListItem Value="5">明顯下降</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td bgcolor="#ecf7ff">
                                <div><span class="style1"><strong>3.參訓學員工作能力是否有提升 ?</strong></span></div>
                                <div>
                                    <asp:RadioButtonList ID="Q13" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">明顯提升</asp:ListItem>
                                        <asp:ListItem Value="2">略有提升</asp:ListItem>
                                        <asp:ListItem Value="3">沒有提升</asp:ListItem>
                                        <asp:ListItem Value="4">略有下降</asp:ListItem>
                                        <asp:ListItem Value="5">明顯下降</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td bgcolor="#ecf7ff">
                                <div><span class="style1"><strong>4.企業欲藉由訓練改善之具體營業績效(例如顧客滿意,產品良率等等)是否有改變 ?</strong></span></div>
                                <div>
                                    <asp:RadioButtonList ID="Q14" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">明顯提升</asp:ListItem>
                                        <asp:ListItem Value="2">略有提升</asp:ListItem>
                                        <asp:ListItem Value="3">沒有提升</asp:ListItem>
                                        <asp:ListItem Value="4">略有下降</asp:ListItem>
                                        <asp:ListItem Value="5">明顯下降</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td bgcolor="#ecf7ff">
                                <div><span class="style1"><strong>5.企業之離職率是否有改變 ?</strong></span></div>
                                <div>
                                    <asp:RadioButtonList ID="Q15" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">明顯提升</asp:ListItem>
                                        <asp:ListItem Value="2">略有提升</asp:ListItem>
                                        <asp:ListItem Value="3">沒有提升</asp:ListItem>
                                        <asp:ListItem Value="4">略有下降</asp:ListItem>
                                        <asp:ListItem Value="5">明顯下降</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td bgcolor="#ecf7ff">
                                <div><span class="style1"><strong>6.企業繼續辦理專班的計畫 ?</strong></span></div>
                                <div>
                                    <asp:RadioButtonList ID="Q16" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">擴大辦理</asp:ListItem>
                                        <asp:ListItem Value="2">維持現有規模</asp:ListItem>
                                        <asp:ListItem Value="3">縮小辦理規模</asp:ListItem>
                                        <asp:ListItem Value="4">不再辦理</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                    </table>
                    <div align="center">
                        <asp:Button ID="Button1" runat="server" CausesValidation="False" Text="儲存"></asp:Button>
                        <asp:Button ID="Button2" runat="server" CausesValidation="False" Text="回上一頁"></asp:Button>
                        <asp:Button ID="next_but" runat="server" CausesValidation="False" Text="不儲存填寫下一位"></asp:Button>
                    </div>
                    <div align="center"><asp:ValidationSummary ID="Summary" runat="server" Width="577px" ShowMessageBox="True" ShowSummary="False" DisplayMode="List"></asp:ValidationSummary></div>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>