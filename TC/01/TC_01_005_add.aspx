<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_005_add.aspx.vb" Inherits="WDAIIP.TC_01_005_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>課程資料設定</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/OpenWin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function ChangeCourseSort(MyValue) {
            //alert(MyValue);//1.學科 2.術科
            //var Classification1_List = document.getElementById('Classification1_List');
            var Classification2_List = document.getElementById('Classification2_List');
            Classification2_List.value = '0'; //MyValue;
            //Classification2_List.disabled = false;
            if (MyValue == '2') {
                Classification2_List.value = '2'; //MyValue;
                //Classification2_List.disabled = true;
            }

            var cst_inline = ""; // "inline";
            var cst_none = "none";
            var LabelTeah1_2 = document.getElementById('LabelTeah1_2');
            var OLessonTeah1_2 = document.getElementById('OLessonTeah1_2');
            var TPlanIDValue = document.getElementById('TPlanIDValue');
            if (TPlanIDValue.value == '68') {
                LabelTeah1_2.style.display = cst_inline;
                OLessonTeah1_2.style.display = cst_inline;
                OLessonTeah1_2.value = "";
                if (MyValue == '1') {
                    LabelTeah1_2.style.display = cst_none;
                    OLessonTeah1_2.style.display = cst_none;
                }
            }
        }

        function CheckTMID(source, args) {
            args.IsValid = true;
            var trainValue = document.getElementById('trainValue');
            var Classification2_List = document.getElementById('Classification2_List');
            //0~2 0:共同1:一般2:專業
            if (Classification2_List.selectedIndex > 1) {
                if (trainValue.value == '') args.IsValid = false;
            }
        }

        function LessonTeah3(opentype, st, fieldname, hiddenname) {
            //alert("!!!");
            var RIDValue1 = document.getElementById('RIDValue1');
            if (st == '1') {//老師
                wopen('../../SD/04/LessonTeah1.aspx?RID=' + RIDValue1.value + '&type=' + opentype + '&fieldname=' + fieldname + '&hiddenname=' + hiddenname, 'LessonTeah1', 400, 300, 1);
            }
            if (st == '12') {//老師
                wopen('../../SD/04/LessonTeah1.aspx?RID=' + RIDValue1.value + '&type=' + opentype + '&fieldname=' + fieldname + '&hiddenname=' + hiddenname, 'LessonTeah12', 400, 300, 1);
            }
            if (st == '13') {//老師
                wopen('../../SD/04/LessonTeah1.aspx?RID=' + RIDValue1.value + '&type=' + opentype + '&fieldname=' + fieldname + '&hiddenname=' + hiddenname, 'LessonTeah13', 400, 300, 1);
            }
            if (st == '2') {//助教
                wopen('../../SD/04/LessonTeah2.aspx?RID=' + RIDValue1.value + '&type=' + opentype + '&fieldname=' + fieldname + '&hiddenname=' + hiddenname, 'LessonTeah2', 400, 300, 1);
            }
            if (st == '3') {
                //hiddenname//助教
                wopen('../../SD/04/LessonTeah2.aspx?RID=' + RIDValue1.value + '&type=' + opentype + '&fieldname=' + fieldname + '&hiddenname=' + hiddenname, 'LessonTeah3', 400, 300, 1);
            }
            if (st == '2b') {//助教
                wopen('../../SD/04/LessonTeah2.aspx?RID=' + RIDValue1.value + '&type=' + opentype + '&fieldname=' + fieldname + '&hiddenname=' + hiddenname, 'LessonTeah2b', 400, 300, 1);
            }
        }

        function Check_CourseName(source, args) {
            args.IsValid = true;
            if (!checkFileName(args.Value)) {
                args.IsValid = false;
            }
        }
    </script>
</head>
<body onload="">
    <form id="form1" method="post" runat="server">
        <table id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">

            <tr>
                <td>
                    <asp:Label ID="lblProecessType" runat="server" Visible="false"></asp:Label>
                    <table class="font" cellspacing="1" cellpadding="1" border="0">
                        <tr>
                            <td class="font">
                                <%--<asp:Literal ID="clientscript" runat="server"></asp:Literal>--%>
                                <asp:CustomValidator ID="CustomValidator1" runat="server" Display="None"></asp:CustomValidator>
                            </td>
                        </tr>
                    </table>
                    <table class="table_nw" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td id="td2" runat="server" class="bluecol_need" style="width: 20%">計畫階段 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="TBplan" runat="server" Width="70%" onfocus="this.blur()"></asp:TextBox>
                                <input id="choice_button" onclick="javascript: wopen('../../Common/LevPlan.aspx?winreload=1', '計畫階段', 850, 570, 1)" type="button" value="選擇" name="choice_button" runat="server" class="asp_button_M">
                                <asp:RequiredFieldValidator ID="plan" runat="server" Display="None" ControlToValidate="TBplan" ErrorMessage="請選擇計畫階段"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" style="width: 20%">課程代碼 </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="TB_CourseID" runat="server" Width="40%" MaxLength="8"></asp:TextBox>(最多8碼)
							<asp:RequiredFieldValidator ID="Re_CourseID" runat="server" Display="None" ControlToValidate="TB_CourseID" ErrorMessage="請輸入課程代碼"></asp:RequiredFieldValidator><asp:RegularExpressionValidator ID="Re_CourseID2" runat="server" Display="None" ControlToValidate="TB_CourseID" ErrorMessage="請填寫英數字且最多8碼" ValidationExpression="[0-9A-Za-z]{1,8}"></asp:RegularExpressionValidator>
                            </td>
                            <td width="100" class="bluecol_need" style="width: 20%">課程名稱 </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="TB_CourseName" runat="server" MaxLength="50" Width="70%"></asp:TextBox>
                                <asp:RequiredFieldValidator ID="Re_CourseName" runat="server" Display="None" ControlToValidate="TB_CourseName" ErrorMessage="請輸入課程名稱"></asp:RequiredFieldValidator><asp:CustomValidator ID="Custom" runat="server" Display="None" ControlToValidate="TB_CourseName" ErrorMessage="請勿輸入特殊字元" ClientValidationFunction="Check_CourseName"></asp:CustomValidator>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">學/術科 </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="Classification1_List" runat="server">
                                    <asp:ListItem Value="">請選擇</asp:ListItem>
                                    <asp:ListItem Value="1">學科</asp:ListItem>
                                    <asp:ListItem Value="2">術科</asp:ListItem>
                                </asp:DropDownList>
                                <asp:RequiredFieldValidator ID="Re_Classification1" runat="server" Display="None" ControlToValidate="Classification1_List" ErrorMessage="請選擇學/術科"></asp:RequiredFieldValidator>
                            </td>
                            <td class="bluecol_need">共同/一般/專業 </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="Classification2_List" runat="server">
                                    <asp:ListItem Value="">請選擇</asp:ListItem>
                                    <asp:ListItem Value="0">共同</asp:ListItem>
                                    <asp:ListItem Value="1">一般</asp:ListItem>
                                    <asp:ListItem Value="2">專業</asp:ListItem>
                                </asp:DropDownList>
                                <asp:RequiredFieldValidator ID="Re_Classification2" runat="server" Display="None" ControlToValidate="Classification2_List" ErrorMessage="請選擇共同/一般/專業"></asp:RequiredFieldValidator>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">小時數 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TB_Hours" runat="server" Width="30%"></asp:TextBox>
                                <asp:RegularExpressionValidator ID="Re_hours" runat="server" Display="None" ControlToValidate="TB_Hours" ErrorMessage="請輸入數字" ValidationExpression="[0-9]+"></asp:RegularExpressionValidator>
                            </td>
                            <td id="Td3" runat="server" class="bluecol">是否有效 </td>
                            <td class="whitecol">
                                <asp:CheckBox ID="CB_Valid" runat="server" Checked="True" CssClass="font"></asp:CheckBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">主課程 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TB_CourName" runat="server" onfocus="this.blur()" Width="80%"></asp:TextBox><br>
                                <input id="choice_MainCource" onclick="javascript: wopen('TC_01_005_MainCourse.aspx?rid=' + document.getElementById('RIDValue').value + '&amp;orgid=' + document.getElementById('orgid_value').value + '&amp;classid=' + document.getElementById('TB_CourseID').value, '主課程', 900, 500, 1)" type="button" value="挑選" runat="server" class="asp_button_M">
                                <input id="orgid_value" type="hidden" runat="server">
                                <asp:Button ID="Button1" runat="server" Text="清除" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                            </td>
                            <td class="bluecol_need">訓練職類 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="career" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." name="career" runat="server" class="button_b_Mini">
                                <input id="trainValue" type="hidden" name="trainValue" runat="server">
                                <asp:Button ID="Button4" runat="server" Text="清除" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                                <asp:CustomValidator ID="CustomValidator3" runat="server" Display="None" ErrorMessage="請輸入訓練職類" ClientValidationFunction="CheckTMID"></asp:CustomValidator>
                            </td>
                        </tr>
                        <tr>
                            <td id="Td4" runat="server" class="bluecol">隸屬班級 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="Classid" runat="server" onfocus="this.blur()" Columns="46" Width="40%"></asp:TextBox>
                                <%--<input id="Button2" title="含計畫年度屬性" onclick="javascript:wopen('TC_01_005_Classid.aspx?ProcessType='+document.getElementById('Type_str').value+'&amp;tplanid='+document.getElementById('TPlanIDValue').value,'班級代碼',1000,630,1)" type="button" value="選擇" name="Button2" runat="server" class="asp_button_M">--%>
                                <input id="Button2" title="含計畫年度屬性" type="button" value="選擇" name="Button2" runat="server" class="asp_button_M">
                                <asp:Button ID="Button3" runat="server" Text="清除" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                                <input id="Classid_Hid" type="hidden" name="Classid_Hid" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                <asp:Label ID="LabelTeah1" runat="server">教師1</asp:Label>
                            </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="OLessonTeah1" Style="cursor: pointer" Width="15%" MaxLength="8" runat="server" ToolTip="點選兩下可以跳出視窗選擇教師"></asp:TextBox>
                            </td>
                        </tr>
                        <tr id="trTeahList1" runat="server">
                            <td class="bluecol">
                                <asp:Label ID="LabelTeah1_2" runat="server">教師2</asp:Label>
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="OLessonTeah1_2" Style="cursor: pointer" Width="40%" MaxLength="8" runat="server" ToolTip="點選兩下可以跳出視窗選擇教師"></asp:TextBox>
                            </td>
                            <td class="bluecol">
                                <asp:Label ID="LabelTeah1_3" runat="server">教師3</asp:Label>
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="OLessonTeah1_3" Style="cursor: pointer" Width="40%" MaxLength="8" runat="server" ToolTip="點選兩下可以跳出視窗選擇教師"></asp:TextBox>
                            </td>
                        </tr>
                        <tr id="trTeahList2" runat="server">
                            <td class="bluecol">
                                <asp:Label ID="LabelTeah2" runat="server">助教1</asp:Label>
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="OLessonTeah2" Style="cursor: pointer" Width="40%" MaxLength="8" runat="server" ToolTip="點選兩下可以跳出視窗選擇教師"></asp:TextBox>
                            </td>
                            <td class="bluecol">
                                <asp:Label ID="LabelTeah3" runat="server">助教2</asp:Label>
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="OLessonTeah3" Style="cursor: pointer" Width="40%" MaxLength="8" runat="server" ToolTip="點選兩下可以跳出視窗選擇教師"></asp:TextBox>
                            </td>
                        </tr>
                        <tr id="trTeahList3" runat="server">
                            <td class="bluecol">
                                <asp:Label ID="LabelTeah2b" runat="server">助教1</asp:Label>
                            </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="OLessonTeah2b" Style="cursor: pointer" Width="15%" MaxLength="8" runat="server" ToolTip="點選兩下可以跳出視窗選擇教師"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">教室 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="Room" runat="server" MaxLength="30" Width="15%"></asp:TextBox>
                            </td>
                        </tr>
                        <%---
						<TR>
							<TD style="WIDTH: 79px; HEIGHT: 22px" align="left" bgColor="#ffcccc">&nbsp;&nbsp;&nbsp;是否計算<br>
								&nbsp;&nbsp; 排課時數
							</TD>
							<TD colSpan="3">
								<asp:RadioButtonList id="IsCountHours" runat="server" RepeatDirection="Horizontal" Font-Size="X-Small"
									Height="24px">
									<asp:ListItem Value="Y" Selected="True">是</asp:ListItem>
									<asp:ListItem Value="N">否</asp:ListItem>
								</asp:RadioButtonList><br>
								<font color="red">若選【否】<br>
									則此課程的排課將不列入排課時數的計算，其他相關計算排課時數功能，都會將此課程排除</font>
							</TD>
						</TR>
                        --%>
                    </table>
                    <table width="100%">
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="bt_save" runat="server" Text="儲存" Enabled="False" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="Button5" runat="server" Text="回上一頁" CausesValidation="False" class="asp_button_M"></asp:Button>
                                <asp:ValidationSummary ID="Summary" runat="server" ShowMessageBox="True" DisplayMode="List"></asp:ValidationSummary>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <input id="OLessonTeah1Value" type="hidden" name="OLessonTeah1Value" runat="server">
        <input id="OLessonTeah1_2Value" type="hidden" name="OLessonTeah1_2Value" runat="server">
        <input id="OLessonTeah1_3Value" type="hidden" name="OLessonTeah1_3Value" runat="server">
        <input id="OLessonTeah2Value" type="hidden" name="OLessonTeah2Value" runat="server">
        <input id="OLessonTeah3Value" type="hidden" name="OLessonTeah3Value" runat="server">
        <input id="OLessonTeah2bValue" type="hidden" name="OLessonTeah2bValue" runat="server">
        <input id="courid" type="hidden" runat="server">
        <input id="Re_ID" type="hidden" runat="server">
        <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
        <input id="PlanIDValue" type="hidden" name="PlanIDValue" runat="server">
        <input id="TPlanIDValue" type="hidden" name="TPlanIDValue" runat="server">
        <input id="Type_str" type="hidden" name="Type_str" runat="server">
        <input id="RIDValue1" type="hidden" name="RIDValue1" runat="server">
        <%--<input id="HidOCID1" type="hidden" name="HidOCID1" runat="server">--%>
        <asp:HiddenField ID="Hidsave1" runat="server" />
    </form>
</body>
</html>
