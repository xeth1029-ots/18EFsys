<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_11_005_add12.aspx.vb" Inherits="WDAIIP.SD_11_005_add12" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_11_005_add12</title>
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
        function Q1Is4() {
            if (getValue("Q1") == '4') {
                document.getElementById("Q2").disabled = true;
                document.getElementById("Q3").disabled = true;
                document.getElementById("Q4").disabled = true;
                document.getElementById("Q5").disabled = true;
            }
            else {
                document.getElementById("Q2").disabled = false;
                document.getElementById("Q3").disabled = false;
                document.getElementById("Q4").disabled = false;
                document.getElementById("Q5").disabled = false;
            }
        }

        function ChkData() {
            var msg = '';
            if (isEmpty('Q1')) { msg += '請選擇 1. 學員目前的近況為何 ?\n'; }
            if (getValue("Q1") != '4') {
                if (isEmpty('Q2')) { msg += '請選擇 2. 學員於結訓後薪資有提升嗎 ?\n'; }
                if (isEmpty('Q3')) { msg += '請選擇 3. 學員的職位有變化嗎 ?\n'; }
                if (isEmpty('Q4')) { msg += '請選擇 4. 學員對目前工作的滿意度是否有變化 ?\n'; }
                if (isEmpty('Q5')) { msg += '請選擇 5. 學員目前的工作內容是否與參訓課程內容相關 ?\n'; }
            }
            //if (isEmpty('Q6')) { msg += '請選擇 6. 學員是否同意參加訓練對目前工作表現或第二專長培育有幫助?(6-1、6-2可擇一選答)\n'; }
            if (isEmpty('Q6_7') && isEmpty('Q6_8')) { msg += '請選擇 (6-1、6-2可擇一選答)\n'; }
            //if (isEmpty('Q6_7')) { msg += '請選擇 6-1.學員是否同意參加訓練對目前工作表現有幫助 ?\n'; }
            //if (isEmpty('Q6_8')) { msg += '請選擇 6-2.學員是否同意參加訓練對第二專長培育有幫助 ?\n'; }
            //if (isEmpty('Q7')) { msg += '請選擇 9.承上題，參加本項訓練對學員的幫助是在哪方面 ?\n'; }
            if (isEmpty('Q8')) { msg += '請選擇 7. 學員是否有繼續參與進修訓練的意願 ?\n'; }
            if (msg == '') {
                return true;
            } else {
                alert(msg);
                return false;
            }
        }

        function insert_next() {
            var aspxSDform1 = "SD_11_005_add12.aspx";
            var aspxSDForm2 = "SD_11_005.aspx";
            if (window.confirm("是否繼續新增下一筆?")) {
                location.href = aspxSDform1 + '?ProcessType=next&ocid=' + document.getElementById("Re_OCID").value + '&SOCID=' + document.getElementById("Re_SOCID").value + '&ID=' + document.getElementById("Re_ID").value;
            }
            else {
                location.href = aspxSDForm2 + '?ProcessType=Back&ocid=' + document.getElementById("Re_OCID").value + '&ID=' + document.getElementById("Re_ID").value;
            }
        }
    </script>
    <style type="text/css">
        .style1 { color: #000000; }
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
                        <%--
                        <tr>
						    <td>
							    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
							    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;受訓學員訓後動態調查表</asp:Label>
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
                            <td width="20%" class="bluecol"><font color="#ffffff">&nbsp;&nbsp;&nbsp; 學員</font> </td>
                            <td bgcolor="#ecf7ff" colspan="3">
                                <asp:DropDownList ID="SOCID" runat="server" AutoPostBack="True"></asp:DropDownList></td>
                        </tr>
                    </table>
                    <table class="font" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr bgcolor="#ecf7ff">
                            <td bgcolor="#ecf7ff">
                                <div>本問卷係本署為瞭解學員參加本計畫訓練課程後，近況與未來動向，請協助調查並將表中每一項打ˇ表示，如有其他意見請於其他欄以文字敘述，供本署改進之參考。謝謝！ </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="table_title"><font color="#ffffff"><strong>【一、學員部份:】</strong></font></td>
                        </tr>
                        <tr bgcolor="#ecf7ff">
                            <td bgcolor="#ecf7ff">
                                <div><span class="style1"><strong>1.學員目前的近況為何 ?</strong></span></div>
                                <div>
                                    <asp:RadioButtonList ID="Q1" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">留任原公司</asp:ListItem>
                                        <asp:ListItem Value="2">轉換至同產業的公司</asp:ListItem>
                                        <asp:ListItem Value="3">轉換至不同產業的公司</asp:ListItem>
                                        <asp:ListItem Value="4">已離職，待業中(若選此項直接跳答第6題以後之題目)</asp:ListItem>
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
                                <div><span class="style1"><strong>4.學員對目前工作的滿意度是否有變化 ?</strong></span></div>
                                <div>
                                    <asp:RadioButtonList ID="Q4" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1">大幅提升</asp:ListItem>
                                        <asp:ListItem Value="2">小幅提升</asp:ListItem>
                                        <asp:ListItem Value="3">沒有變化</asp:ListItem>
                                        <asp:ListItem Value="4">小幅降低</asp:ListItem>
                                        <asp:ListItem Value="5">大幅減少</asp:ListItem>
                                    </asp:RadioButtonList>
                                </div>
                            </td>
                        </tr>
                        <tr bgcolor="#ecf7ff">
                            <td bgcolor="#ecf7ff">
                                <div><span class="style1">&nbsp;<strong>5. 學員目前的工作內容是否與參訓課程內容相關 ?</strong></span></div>
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
                                <div><span class="style1"><strong>6.學員是否同意參加訓練對目前工作表現或第二專長培育有幫助?(6-1、6-2可擇一選答)</strong></span></div>
                                <%--
                                <asp:RadioButtonList ID="Q6" runat="server" CssClass="font" Width="576px" RepeatDirection="Horizontal" Font-Size="X-Small">
								    <asp:ListItem Value="1">幫助非常大</asp:ListItem>
								    <asp:ListItem Value="2">幫助頗多</asp:ListItem>
								    <asp:ListItem Value="3">有幫助</asp:ListItem>
								    <asp:ListItem Value="4">幫助有限</asp:ListItem>
								    <asp:ListItem Value="5">完全沒幫助</asp:ListItem>
							    </asp:RadioButtonList>
                                --%>
                            </td>
                        </tr>
                        <tr bgcolor="#ecf7ff">
                            <td style="height: 59px" bgcolor="#ecf7ff">
                                <div><span class="style1"><strong>6-1.學員是否同意參加訓練對目前工作表現有幫助 ?</strong></span></div>
                                <asp:RadioButtonList ID="Q6_7" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="1">幫助非常大</asp:ListItem>
                                    <asp:ListItem Value="2">幫助頗多</asp:ListItem>
                                    <asp:ListItem Value="3">有幫助</asp:ListItem>
                                    <asp:ListItem Value="4">幫助有限</asp:ListItem>
                                    <asp:ListItem Value="5">完全沒幫助</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr bgcolor="#ecf7ff">
                            <td bgcolor="#ecf7ff">
                                <div><span class="style1"><strong>6-2.學員是否同意參加訓練對第二專長培育有幫助 ?</strong></span></div>
                                <asp:RadioButtonList ID="Q6_8" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="1">幫助非常大</asp:ListItem>
                                    <asp:ListItem Value="2">幫助頗多</asp:ListItem>
                                    <asp:ListItem Value="3">有幫助</asp:ListItem>
                                    <asp:ListItem Value="4">幫助有限</asp:ListItem>
                                    <asp:ListItem Value="5">完全沒幫助</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <%--
                        <tr>
						    <td bgcolor="#ecf7ff">
							    <div><span class="style1"><strong>9.承上題，參加本項訓練對學員的幫助是在哪方面 ?</strong></span></div>
							    <div>
								    <asp:RadioButtonList ID="Q7" runat="server" CssClass="font" Width="584px" RepeatDirection="Horizontal">
									    <asp:ListItem Value="1">對適應工作環境有幫助</asp:ListItem>
									    <asp:ListItem Value="2">對目前工作績效有幫助</asp:ListItem>
									    <asp:ListItem Value="3">對轉換工作跑道有幫助</asp:ListItem>
								    </asp:RadioButtonList>
							    </div>
						    </td>
					    </tr>
                        --%>
                        <tr>
                            <td bgcolor="#ecf7ff">
                                <div><span class="style1"><strong>7.學員是否有繼續參與進修訓練的意願 ?</strong></span></div>
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
                            <td bgcolor="#ecf7ff" class="whitecol"><span class="style1"><strong>8.學員認為還需要加強哪方面的專業知識使工作進行得更順利 ?</strong></span>
                                <div>
                                    1.<asp:TextBox ID="Q9_1_Note" runat="server" MaxLength="100"></asp:TextBox>
                                    2.<asp:TextBox ID="Q9_2_Note" runat="server" MaxLength="100"></asp:TextBox>
                                    3.<asp:TextBox ID="Q9_3_Note" runat="server" MaxLength="100"></asp:TextBox>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td bgcolor="#ecf7ff" class="whitecol"><span class="style1"><strong>9.學員常和本課程的哪些學員、教師或職員聯絡 ?</strong></span>
                                <div>
                                    1.<asp:TextBox ID="Q10_1_Note" runat="server" MaxLength="100"></asp:TextBox>
                                    2.<asp:TextBox ID="Q10_2_Note" runat="server" MaxLength="100"></asp:TextBox>
                                    3.<asp:TextBox ID="Q10_3_Note" runat="server" MaxLength="100"></asp:TextBox>
                                </div>
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
    </form>
</body>
</html>
