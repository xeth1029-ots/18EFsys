<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_01_002_R.aspx.vb" Inherits="WDAIIP.SD_01_002_R" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>報名資料一覽表</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function choose_class() {
            var RID = document.form1.RIDValue.value;
            openClass('../02/SD_02_ch.aspx?special=2&RID=' + RID);
        }

        function GETvalue() {
            document.getElementById('Button13').click();
        }

        function chk() {
            var value1 = '';
            var TMIDValue1 = document.getElementById("<%=TMIDValue1.ClientID%>");
            var OCIDValue1 = document.getElementById("<%=OCIDValue1.ClientID%>");
            var cobValue = document.getElementById("<%=cjobValue.ClientID%>");
            var start_date = document.getElementById("<%=start_date.ClientID%>");
            var end_date = document.getElementById("<%=end_date.ClientID%>");

            if (TMIDValue1 && TMIDValue1.value != '') {
                value1 += '1';
            }
            if (OCIDValue1 && OCIDValue1.value != '') {
                value1 += '1';
            }
            if (cobValue && cobValue.value != '') {
                value1 += '1';
            }
            if ((start_date && start_date.value == '') || (end_date && end_date.value == '')) {
                value1 += '';
            }
            else {
                value1 += '1';
            }
            if (value1 == '') {
                window.alert('查詢條件請擇一挑選');
                return false;
            }
            else {
                return true;
            }
        }

        function chall() {
            var mytable = document.getElementById('DataGrid2')
            for (var i = 1; i < mytable.rows.length; i++) {
                var mycheck = mytable.rows[i].cells[0].children[0];
                if (mycheck.disabled == false) {
                    mycheck.checked = document.form1.Choose1.checked
                }
            }
        }

        function CheckPrint() {
            var MyTable = document.getElementById('DataGrid2');
            var hidExamID = document.getElementById('hidExamID');
            var hidIDNOVALUE = document.getElementById('hidIDNOVALUE');
            var ExamID = '';
            var IDNOVALUE = '';
            for (var i = 1; i < MyTable.rows.length; i++) {
                if (MyTable.rows[i].cells[0].children[0].checked) {
                    if (ExamID != '') { ExamID += ',' }
                    ExamID += MyTable.rows(i).cells(0).children(1).value;
                    //if (IDNOVALUE != '') { IDNOVALUE += ',' }
                    //IDNOVALUE += MyTable.rows(i).cells(0).children(2).value;
                }
            }
            if (ExamID == '') {
                alert('請選擇學員');
                return false;
            }
            else {
                hidExamID.value = ExamID;
                //hidIDNOVALUE.value = IDNOVALUE;
                //openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=member&filename=member_list&&path=TIMS&ExamID='+ExamID+'&IDNO='+IDNOVALUE+'&DistID='+document.getElementById('DistID').value+'&OCID1='+document.getElementById('OCIDValue1').value+'&CJOB_UNKEY='+document.getElementById('cjobValue').value+'&RID='+document.getElementById('RID').value+'&TPlanID='+document.getElementById('TPlanID').value+'&STDate1='+document.getElementById('start_date').value+'&STDate2='+document.getElementById('end_date').value+'&RelEnterDate1='+document.getElementById('EnterDate_start').value+'&RelEnterDate2='+document.getElementById('EnterDate_end').value);
                //openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=member&filename=member_list&ExamID=' + ExamID + '&DistID=' + document.getElementById('DistID').value + '&OCID1=' + document.getElementById('OCIDValue1').value + '&CJOB_UNKEY=' + document.getElementById('cjobValue').value + '&RID=' + document.getElementById('RID').value + '&TPlanID=' + document.getElementById('TPlanID').value + '&STDate1=' + document.getElementById('start_date').value + '&STDate2=' + document.getElementById('end_date').value + '&RelEnterDate1=' + document.getElementById('EnterDate_start').value + '&RelEnterDate2=' + document.getElementById('EnterDate_end').value);
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;報名資料一覽表</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch" id="Table3" cellpadding="1" cellspacing="1" width="100%">
                        <tr>
                            <td class="bluecol" width="20%">訓練機構</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="center" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="Button8" type="button" value="..." name="Button8" runat="server" class="button_b_Mini" />
                                <asp:Button ID="Button13" Style="display: none" runat="server" Text="Button13"></asp:Button>
                                <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">職類/班別</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="button_b_Mini" />
                                <input id="OCIDValue1" type="hidden" name="Hidden2" runat="server" />
                                <input id="TMIDValue1" type="hidden" name="Hidden1" runat="server" />
                                <span id="HistoryList" style="position: absolute; display: none; left: 270px">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">通俗職類</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="30%"></asp:TextBox>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="button_b_Mini" />
                                <input id="cjobValue" type="hidden" name="cjobValue" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">開訓日期</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="start_date" runat="server" Width="15%" MaxLength="10"></asp:TextBox>
                                <span id="span1" runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span> &nbsp;～&nbsp;
                                <asp:TextBox ID="end_date" runat="server" Width="15%" MaxLength="10"></asp:TextBox>
                                <span id="span2" runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span> &nbsp;(查詢條件請擇一挑選)
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">報名日期</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="EnterDate_start" runat="server" Width="15%" MaxLength="10"></asp:TextBox>
                                <span id="span3" runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= EnterDate_start.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span> &nbsp;～&nbsp;
                                <asp:TextBox ID="EnterDate_end" runat="server" Width="15%" MaxLength="10"></asp:TextBox>
                                <span id="span4" runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= EnterDate_end.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span> &nbsp;&nbsp;&nbsp;
                                <%--<input id="DistID" style="width: 6%; height: 5%" type="hidden" name="DistID" runat="server" />--%>
                                <%--<input id="RID" style="width: 6%; height: 5%" type="hidden" name="DistID" runat="server" />--%>
                                <%--<input id="TPlanID" style="width: 6%; height: 5%" type="hidden" name="DistID" runat="server" />--%>
                            </td>
                        </tr>
                        <tr id="Trwork2013a" runat="server">
                            <td class="bluecol">報名管道</td>
                            <td class="whitecol">
                                <%--
							    id="Trwork2013a" runat="server"
							    就服單位協助報名
							    <asp:RadioButtonList Style="z-index: 0" ID="rblEnterPathW" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
								   <asp:ListItem Value="A" Selected="True">不區分</asp:ListItem>
								   <asp:ListItem Value="Y">是</asp:ListItem>
								   <asp:ListItem Value="N">否</asp:ListItem>
							    </asp:RadioButtonList>
                                --%>
                                <asp:RadioButtonList Style="z-index: 0" ID="rblEnterPathW2" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="A" Selected="True">不區分</asp:ListItem>
                                    <asp:ListItem Value="CH4">一般推介單</asp:ListItem>
                                    <asp:ListItem Value="EPW">免試推介單</asp:ListItem>
                                    <asp:ListItem Value="EP2P">專案核定報名</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">匯出檔案格式</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                                    <asp:ListItem Value="ODS">ODS</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr runat="server">
                            <td class="whitecol" colspan="2" style="text-align: center;">
                                <asp:Button ID="Button3" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                &nbsp;<asp:Button ID="Button1" runat="server" Text="直接列印" CssClass="asp_Export_M"></asp:Button>
                                &nbsp;<asp:Button ID="Button2" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>&nbsp;
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td class="font" style="text-align: center;">
                    <asp:Label ID="labmsg1" runat="server" ForeColor="Red" Text="以「*」星號標註在訓中之學員，惟不可以本項標註作為甄試資格不符或不予錄訓之依據。"></asp:Label>
                </td>
            </tr>
        </table>
        <br />
        <table class="font" id="Table5" style="width: 100%" cellspacing="1" cellpadding="1" border="0" runat="server">
            <tr>
                <td align="center" class="whitecol">
                    <asp:Button ID="Button4" runat="server" Text="群組列印" Visible="False" CssClass="asp_Export_M"></asp:Button>
                    <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" AllowSorting="True" PageSize="100" AllowPaging="True" CellPadding="8">
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <HeaderStyle CssClass="head_navy" />
                        <Columns>
                            <asp:TemplateColumn>
                                <HeaderStyle HorizontalAlign="Center" Width="6%"></HeaderStyle>
                                <HeaderTemplate>
                                    <input onclick="chall();" type="checkbox" checked="checked" name="Choose1" />
                                </HeaderTemplate>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <input id="Checkbox1" type="checkbox" checked name="student" runat="server" />
                                    <input id="EXAMNO" type="hidden" runat="server" />
                                    <input id="HIDNO" type="hidden" runat="server" />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn SortExpression="Name">
                                <HeaderStyle HorizontalAlign="Center" Width="14%"></HeaderStyle>
                                <HeaderTemplate>姓名</HeaderTemplate>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="labStar1" runat="server" ForeColor="Red">*</asp:Label>
                                    <asp:Label ID="labName" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <%-- <asp:BoundColumn DataField="IDNO" SortExpression="IDNO" HeaderText="身分證號碼">
                                <HeaderStyle HorizontalAlign="Center" ForeColor="#B0E2FF" Width="14%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>--%>
                            <asp:BoundColumn DataField="EXAMNO" HeaderText="准考證號碼">
                                <HeaderStyle HorizontalAlign="Center" Width="14%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="OrgName" HeaderText="報名機構">
                                <HeaderStyle HorizontalAlign="Center" Width="25%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="ClassCName" HeaderText="報名班級">
                                <HeaderStyle HorizontalAlign="Center" Width="25%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn SortExpression="RelEnterDate">
                                <HeaderStyle HorizontalAlign="Center" Width="14%"></HeaderStyle>
                                <HeaderTemplate>報名日期</HeaderTemplate>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="labRelEnterDate" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid><uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                </td>
            </tr>
        </table>
        <div class="whitecol" align="center" style="width: 100%">
            <asp:Button ID="Button5" runat="server" Text="群組列印" Visible="False" CssClass="asp_Export_M"></asp:Button>
            <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
        </div>
        <input id="hidExamID" runat="server" type="hidden" />
        <input id="hidIDNOVALUE" runat="server" type="hidden" />
    </form>
</body>
</html>
