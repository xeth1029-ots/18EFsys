<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_009.aspx.vb" Inherits="WDAIIP.TC_01_009" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>開班日期修正</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script>
        function SelectAll() {
            var MyTable = document.getElementById('DataGrid1');
            var Flag = MyTable.rows[0].cells[0].children[0].checked;
            for (i = 1; i < MyTable.rows.length; i++) {
                MyTable.rows[i].cells[0].children[0].checked = Flag;
                SelectMyItem(Flag, i)
            }
        }

        function CheckData() {
            var msg = '';
            if (document.getElementById('CyclType').value != '' && !isUnsignedInt(document.getElementById('CyclType').value)) msg += '期別請輸入數字\n';
            if (!CheckMyDate(document.getElementById('start_date').value)) msg += '開訓日期必須為正確的時間格式\n';
            if (!CheckMyDate(document.getElementById('end_date').value)) msg += '結訓日期必須為正確的時間格式\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function SelectMyItem(Flag, num) {
            var MyTable = document.getElementById('DataGrid1')
            if (MyTable.rows[num].cells[3].children.length > 1) {
                MyTable.rows[num].cells[3].children[0].disabled = !Flag;
                if (Flag)
                    MyTable.rows[num].cells[3].children[1].style.display = 'inline'
                else
                    MyTable.rows[num].cells[3].children[1].style.display = 'none'
            }
            if (MyTable.rows[num].cells[4].children.length > 1) {
                MyTable.rows[num].cells[4].children[0].disabled = !Flag;
                if (Flag)
                    MyTable.rows[num].cells[4].children[1].style.display = 'inline'
                else
                    MyTable.rows[num].cells[4].children[1].style.display = 'none'
            }
            MyTable.rows[num].cells[5].children[0].disabled = !Flag;
            if (Flag)
                MyTable.rows[num].cells[5].children[1].style.display = 'inline'
            else
                MyTable.rows[num].cells[5].children[1].style.display = 'none'
            MyTable.rows[num].cells[6].children[0].disabled = !Flag;
            if (Flag)
                MyTable.rows[num].cells[6].children[1].style.display = 'inline'
            else
                MyTable.rows[num].cells[6].children[1].style.display = 'none'
            MyTable.rows[num].cells[7].children[0].disabled = !Flag;
            if (Flag)
                MyTable.rows[num].cells[7].children[1].style.display = 'inline'
            else
                MyTable.rows[num].cells[7].children[1].style.display = 'none'
        }

        var DateCount = 0;
        function SaveData() {
            var msg = '';
            DateCount = 0;
            var MyTable = document.getElementById('DataGrid1');
            for (i = 1; i < MyTable.rows.length; i++) {
                var Flag = MyTable.rows[i].cells[0].children[0].checked;
                if (Flag) {
                    var STDate = MyTable.rows[i].cells[3].children[0].value;
                    var FTDate = MyTable.rows[i].cells[4].children[0].value;
                    var FTDate2 = MyTable.rows[i].cells[4].children[2].value;
                    var SEnterDate = MyTable.rows[i].cells[5].children[0].value;
                    var FEnterDate = MyTable.rows[i].cells[6].children[0].value;
                    var CheckInDate = MyTable.rows[i].cells[7].children[0].value;
                    if (!CheckMyDate(STDate)) msg += '開訓日期必須為正確的時間格式[' + MyTable.rows[i].cells[1].innerHTML + ']\n';
                    if (!CheckMyDate(FTDate)) msg += '結訓日期必須為正確的時間格式[' + MyTable.rows[i].cells[1].innerHTML + ']\n';
                    else {
                        if (compareDate(FTDate, FTDate2) == -1 && MyTable.rows[i].cells[3].children.length <= 1) msg += '結訓日期只能延後不能提前[' + MyTable.rows[i].cells[1].innerHTML + ']\n';
                    }
                    if (CheckMyDate(STDate) && CheckMyDate(FTDate)) {
                        if (compareDate(STDate, FTDate) == 1) msg += '開訓日期不能比結訓日期晚[' + MyTable.rows[i].cells[1].innerHTML + ']\n'
                    }
                    if (!CheckMyDate(SEnterDate)) msg += '開始報名日期必須為正確的時間格式[' + MyTable.rows[i].cells[1].innerHTML + ']\n';
                    if (!CheckMyDate(FEnterDate)) msg += '結束報名日期必須為正確的時間格式[' + MyTable.rows[i].cells[1].innerHTML + ']\n';
                    if (!CheckMyDate(CheckInDate)) msg += '報到日期必須為正確的時間格式[' + MyTable.rows[i].cells[1].innerHTML + ']\n';
                    DateCount = 1;
                }
            }
            if (DateCount == 0) msg += '至少要選擇一項要變動的日期\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function CheckMyDate(MyDate) {
            if (MyDate == '') {
                DateCount++;
                return true;
            }
            else {
                return checkDate(MyDate)
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;開班資料設定&gt;&gt;開班日期修正</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="FrameTable1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td align="center">
                    <%--
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
					    <tr>
						    <td>首頁&gt;&gt;訓練機構管理&gt;&gt;開班資料設定&gt;&gt;<font color="#990000"><font color="#990000">開班日期修正</font></font></td>
					    </tr>
				    </table>
                    --%>
                    <table class="table_sch" id="Table2" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" width="20%">班級名稱</td>
                            <td class="whitecol" width="30%">
                                <asp:TextBox ID="ClassCName" runat="server" Columns="30" Width="40%"></asp:TextBox></td>
                            <td class="bluecol" width="20%">期別</td>
                            <td class="whitecol" width="30%">
                                <asp:TextBox ID="CyclType" runat="server" Columns="5" MaxLength="3" Width="40%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練職類</td>
                            <td class="whitecol">
                                <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Columns="30" Width="40%"></asp:TextBox>
                                <input id="career" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." name="career" runat="server" class="button_b_Mini">
                                <input id="trainValue" type="hidden" name="trainValue" runat="server">
                            </td>
                            <td class="bluecol">&nbsp;&nbsp;&nbsp;&nbsp;訓練時段</td>
                            <td class="whitecol">
                                <asp:DropDownList ID="TPeriod" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label></td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="40%"></asp:TextBox>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="button_b_Mini">
                                <input id="cjobValue" type="hidden" name="cjobValue" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練性質</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="TPropertyID" runat="server" Width="100%" RepeatDirection="Horizontal" RepeatLayout="Flow" CssClass="font">
                                    <asp:ListItem Value="不區分" Selected="True">不區分</asp:ListItem>
                                    <%--<asp:ListItem Value="0">職前</asp:ListItem>--%>
                                    <asp:ListItem Value="1">在職</asp:ListItem>
                                    <asp:ListItem Value="2">接受企業委託</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                            <td class="bluecol">&nbsp;&nbsp;&nbsp; 開班狀態</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="NotOpen" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="N" Selected="True">開班</asp:ListItem>
                                    <asp:ListItem Value="Y">不開班</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">開訓日期</td>
                            <td class="whitecol">
                                <asp:TextBox ID="start_date" Width="35%" runat="server"></asp:TextBox>
                                <span id="date1" runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>～
                                <asp:TextBox ID="end_date" Width="35%" runat="server"></asp:TextBox>
                                <span id="date2" runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                            <td class="bluecol">開訓狀態</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="ClassState" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="不區分" Selected="True">不區分</asp:ListItem>
                                    <asp:ListItem Value="已開訓">已開訓</asp:ListItem>
                                    <asp:ListItem Value="未開訓">未開訓</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button></td>
                        </tr>
                    </table>
                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td>
                    <font color="red">
                        系統提示：<br>
                        1.班級狀態如為「結訓」，則無法修改日期。(欲修改請先解除班級結訓狀態)<br>
                        2.已排課的班級不可修改開訓日期，且只可延長結訓日期。
                    </font>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False">
                        <AlternatingItemStyle />
                        <HeaderStyle CssClass="head_navy" />
                        <Columns>
                            <asp:TemplateColumn HeaderText="選取">
                                <HeaderStyle HorizontalAlign="Center" Width="3%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <HeaderTemplate>
                                    <input id="Checkbox1" type="checkbox" onclick="SelectAll();">
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <input id="OCID" type="checkbox" runat="server">
                                    <input id="hidExamDate" type="hidden" runat="server" />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn DataField="ClassCName" HeaderText="班級名稱">
                                <HeaderStyle Width="12%"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn HeaderText="班別代碼" HeaderStyle-Width="10%"></asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="開訓日期">
                                <HeaderStyle Width="15%"></HeaderStyle>
                                <ItemStyle CssClass="whitecol" />
                                <ItemTemplate>
                                    <asp:TextBox ID="STDate" runat="server" Columns="6" Width="84%"></asp:TextBox>
                                    <img id="IMG1" style="cursor: pointer" onclick="javascript:show_calendar('<%= STDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="20" height="22">
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="結訓日期">
                                <HeaderStyle Width="15%"></HeaderStyle>
                                <ItemStyle CssClass="whitecol" />
                                <ItemTemplate>
                                    <asp:TextBox ID="FTDate" runat="server" Columns="6" Width="84%"></asp:TextBox>
                                    <img id="IMG2" style="cursor: pointer" onclick="javascript:show_calendar('<%= STDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="20" height="22">
                                    <input id="FTDate2" type="hidden" runat="server">
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="報名開始">
                                <HeaderStyle Width="15%"></HeaderStyle>
                                <ItemStyle CssClass="whitecol" />
                                <ItemTemplate>
                                    <asp:TextBox ID="SEnterDate" runat="server" Columns="6" Width="84%"></asp:TextBox>
                                    <img id="IMG3" style="cursor: pointer" onclick="javascript:show_calendar('<%= STDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="20" height="22">
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="報名結束">
                                <HeaderStyle Width="15%"></HeaderStyle>
                                <ItemStyle CssClass="whitecol" />
                                <ItemTemplate>
                                    <asp:TextBox ID="FEnterDate" runat="server" Columns="6" Width="84%"></asp:TextBox>
                                    <img id="IMG4" style="cursor: pointer" onclick="javascript:show_calendar('<%= STDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="20" height="22">
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="報到日期">
                                <HeaderStyle Width="15%"></HeaderStyle>
                                <ItemStyle CssClass="whitecol" />
                                <ItemTemplate>
                                    <asp:TextBox ID="CheckInDate" runat="server" Columns="6" Width="84%"></asp:TextBox>
                                    <img id="IMG5" style="cursor: pointer" onclick="javascript:show_calendar('<%= STDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="20" height="22">
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn HeaderText="已排課">
                                <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                        </Columns>
                    </asp:DataGrid>
                </td>
            </tr>
            <tr>
                <td align="center" class="whitecol">
                    <asp:Button ID="Button2" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button></td>
            </tr>
        </table>
        <asp:HiddenField ID="Hid_ComIDNO" runat="server" />
        <asp:HiddenField ID="Hid_RID1" runat="server" />
    </form>
</body>
</html>
