<%@ Page AspCompat="true" Language="vb" AutoEventWireup="true" CodeBehind="SD_01_001.aspx.vb" Inherits="WDAIIP.SD_01_001" EnableEventValidation="true" %>

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>報名登錄</title>
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
        function GETvalue() {
            document.getElementById('Button12').click();
        }

        function chkblank(num) {
            var msg = '';
            if (num == 1) {
                if (document.form1.IDNO.value == '') msg += '請輸入身分證號碼!';
            }
            else {
                if (document.form1.FExamNo.value != '' && document.form1.SExamNo.value == '') { msg += '請輸入准證號碼起始值!\n'; }
                if (document.form1.SExamNo.value != '' && document.form1.FExamNo.value == '') { msg += '請輸入准證號碼終值!\n'; }
                if (document.form1.start_date.value != '' && !checkDate(document.form1.start_date.value)) msg += '起始時間格式不正確\n';
                if (document.form1.end_date.value != '' && !checkDate(document.form1.end_date.value)) msg += '終至時間格式不正確\n';
                if (document.form1.transDate1.value != '' && !checkDate(document.form1.transDate1.value)) msg += 'e網轉入起始時間格式不正確\n';
                if (document.form1.transDate2.value != '' && !checkDate(document.form1.transDate2.value)) msg += 'e網轉入終至時間格式不正確\n';
                if ((document.form1.start_date.value == '' && document.form1.end_date.value == '') && document.form1.IDNO.value == '') {
                    msg += '請輸入報名日期起、迄或身分證號碼!\n';
                }
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function choose_class(num) {
            var RID = document.form1.RIDValue.value;
            document.form1.TMID1.value = '';
            document.form1.TMIDValue1.value = '';
            document.form1.OCID1.value = '';
            document.form1.OCIDValue1.value = '';
            openClass('../02/SD_02_ch.aspx?RWClass=1&RID=' + RID);
        }

        function chall() {
            var Mytable = document.getElementById('DataGrid2');
            for (var i = 1; i < Mytable.rows.length; i++) {
                var mycheck = Mytable.rows[i].cells[0].children[0];
                if (mycheck.disabled == false)
                    mycheck.checked = document.form1.Choose1.checked;
            }
        }

        function CheckPrint(SMpath, UserID) {
            var Mytable = document.getElementById('DataGrid2');
            var ExamNo = '';
            var SETID = '';
            var SerNum = '';
            for (var i = 1; i < Mytable.rows.length; i++) {
                var value1 = Mytable.rows[i].cells[0].children[1].value;
                var value2 = Mytable.rows[i].cells[0].children[2].value;
                var value3 = Mytable.rows[i].cells[0].children[3].value;
                if (Mytable.rows[i].cells[0].children[0].checked && value1 != '' && value2 != '' && value3 != '') {
                    if (ExamNo != '') ExamNo += ',';
                    ExamNo += '\'' + value1 + '\'';
                    if (SETID != '') SETID += ',';
                    SETID += '\'' + value2 + '\'';
                    if (SerNum != '') SerNum += ',';
                    SerNum += '\'' + value3 + '\'';
                }
            }
            if (ExamNo == '') {
                alert('請選擇學員');
                return false;
            } else {
                var pageVal = "";
                pageVal += '../../SQControl.aspx?SQ_AutoLogout=true&sys=Member&filename=SD_01_001';
                pageVal += '&path=' + SMpath;
                pageVal += '&ExamNo=' + ExamNo + '&SETID=' + SETID + '&SerNum=' + SerNum;
                pageVal += '&PrintShow=' + document.getElementById('PrintShow').value;
                pageVal += '&UserID=' + UserID;
                //debugger;
                openPrint(pageVal);
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;報名登錄</asp:Label>
                </td>
            </tr>
        </table>
        <table class="table_nw" id="table3" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td class="bluecol" width="20%">訓練機構</td>
                <td class="whitecol" width="80%">
                    <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                    <input id="RIDValue" type="hidden" runat="server" />
                    <input id="Button8" type="button" value="..." name="Button5" runat="server" class="asp_button_Mini" />
                    <asp:Button ID="Button12" Style="display: none" runat="server" Text="Button12" CssClass="asp_button_S"></asp:Button>
                    <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                        <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                    </span>
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">班級名稱 </td>
                <td class="whitecol" width="80%">
                    <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                    <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                    <input id="Button5" type="button" value="..." name="Button5" runat="server" class="asp_button_Mini" />
                    <%--<asp:CheckBox ID="SearchMore" runat="server" Text="查詢二、三志願"></asp:CheckBox>--%>
                    <input id="OCIDValue1" type="hidden" runat="server" />
                    <input id="TMIDValue1" type="hidden" runat="server" />
                    <span id="HistoryList" style="position: absolute; display: none; left: 30%">
                        <asp:Table ID="Historytable" runat="server" Width="100%"></asp:Table>
                    </span>
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">通俗職類 </td>
                <td class="whitecol" width="80%">
                    <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="30%"></asp:TextBox>
                    <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="asp_button_Mini" />
                    <input id="cjobValue" type="hidden" name="cjobValue" runat="server" />
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">
                    <asp:Label ID="labtIDNO" runat="server" Text="身分證號碼" Font-Bold="True"></asp:Label></td>
                <td class="whitecol" width="80%">
                    <asp:TextBox ID="IDNO" runat="server" Width="20%"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">開訓日期 </td>
                <td class="whitecol" width="80%">
                    <asp:TextBox ID="stdate1" runat="server" MaxLength="10" Width="15%"></asp:TextBox>
                    <span id="span1" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= stdate1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span> ~
                    <asp:TextBox ID="stdate2" runat="server" MaxLength="10" Width="15%"></asp:TextBox>
                    <span id="span2" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= stdate2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span>
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">報名日期 </td>
                <td class="whitecol" width="80%">
                    <asp:TextBox ID="start_date" runat="server" Width="15%"></asp:TextBox>
                    <span id="span3" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span> ~
                    <asp:TextBox ID="end_date" runat="server" Width="15%"></asp:TextBox>
                    <span id="span4" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span>
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">e網轉入日期 </td>
                <td class="whitecol" width="80%">
                    <asp:TextBox ID="transDate1" runat="server" Width="15%"></asp:TextBox>
                    <span id="span5" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= transDate1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span> ~
                    <asp:TextBox ID="transDate2" runat="server" Width="15%"></asp:TextBox>
                    <span id="span6" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= transDate2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span>
                </td>
            </tr>
            <tr id="trImport1" runat="server">
                <td class="bluecol" width="20%">匯入報名名冊 </td>
                <td class="whitecol" width="80%">
                    <input id="File1" type="file" size="50" name="File1" runat="server" accept=".xls,.ods" />
                    <asp:Button ID="btnIMPORT07" runat="server" Text="匯入名冊" CssClass="asp_button_M"></asp:Button>(必須為ods或xls格式)
                    <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="../../Doc/Stud_Temp_v14.zip" ForeColor="#8080FF" CssClass="font">下載整批上載格式檔</asp:HyperLink>
                    <asp:Button ID="btnPrintOCID1" runat="server" Text="列印匯入學員報名名冊用的班級代碼" CssClass="asp_Export_M"></asp:Button>
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">准考證號碼 </td>
                <td class="whitecol" width="80%">
                    <asp:TextBox ID="SExamNo" runat="server" Width="15%"></asp:TextBox>&nbsp;~&nbsp;
                    <asp:TextBox ID="FExamNo" runat="server" Width="15%"></asp:TextBox>
                </td>
            </tr>
            <%-- <tr id="Trwork2013a" runat="server">
					<td class="bluecol">就服單位協助報名</td>
				    <td class="whitecol">
						<asp:RadioButtonList Style="z-index: 0" ID="rblEnterPathW" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
							<asp:ListItem Value="A" Selected="True">不區分</asp:ListItem>
							<asp:ListItem Value="Y">是</asp:ListItem>
							<asp:ListItem Value="N">否</asp:ListItem>
						</asp:RadioButtonList>
					</td>
				</tr> --%>
            <tr id="Trwork2013a" runat="server">
                <td class="bluecol" width="20%">報名管道</td>
                <td class="whitecol" width="80%">
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
        </table>
        <table width="100%">
            <tr>
                <td class="whitecol">
                    <div align="center">
                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                        <asp:Button ID="add_but" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="Button2" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="Button3" runat="server" Text="查詢三合一資料" CssClass="asp_button_M" Visible="false"></asp:Button>
                    </div>
                </td>
            </tr>
        </table>
        <table id="table4" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td align="center">
                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <HeaderStyle CssClass="head_navy" />
                        <Columns>
                            <asp:BoundColumn HeaderText="編號">
                                <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Name" HeaderText="姓名">
                                <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="IDNO_MK" HeaderText="身分證號碼">
                                <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Birthday" HeaderText="出生日期">
                                <%--DataFormatString="{0:d}"--%>
                                <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="功能">
                                <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                                <ItemTemplate>
                                    <asp:LinkButton ID="Button1" runat="server" Text="新增" CommandName="add" CssClass="linkbutton"></asp:LinkButton>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                </td>
            </tr>
            <%-- <tr><td align="center"><uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler></td></tr> --%>
        </table>
        <div align="center">
            <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
        </div>
        <table class="font" id="table5" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td>
                    <asp:Label ID="Label1" runat="server" ForeColor="Green">* 表該學員有報名其他班級仍在訓中,請查詢學員參訓歷史</asp:Label>
                    <%--<br /><asp:Label ID="Label2" runat="server">滑鼠移至班級可以觀看第二、三志願的班級</asp:Label>--%>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <div id="Div2" runat="server">
                        <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="true" AllowSorting="true" CellPadding="8">
                            <AlternatingItemStyle BackColor="#EEEEEE" />
                            <HeaderStyle CssClass="head_navy" />
                            <Columns>
                                <asp:TemplateColumn>
                                    <HeaderTemplate>
                                        <input onclick="chall();" type="checkbox" checked name="Choose1">
                                    </HeaderTemplate>
                                    <ItemTemplate>
                                        <input id="Checkbox1" type="checkbox" checked name="student" runat="server" />
                                        <input id="ExamNO" type="hidden" name="ExamNO" runat="server" />
                                        <input id="SETID" type="hidden" name="SETID" runat="server" />
                                        <input id="SerNum" type="hidden" name="SerNum" runat="server" />
                                        <input id="Hid_IDNO_MK" type="hidden" name="Hid_IDNO_MK" runat="server" />
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:TemplateColumn HeaderText="編號">
                                    <ItemTemplate>
                                        <asp:Label ID="LNO" runat="server"></asp:Label><br>
                                        <asp:Label ID="star3" runat="server" ForeColor="Green" CssClass="font">*</asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:BoundColumn DataField="Name" SortExpression="NAME" HeaderText="姓名">
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="IDNO_MK" SortExpression="IDNO_MK" HeaderText="身分證號碼">
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="OrgName" SortExpression="ORGNAME" HeaderText="報名機構">
                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="CLASSCNAME1B" SortExpression="CLASSCNAME1B" HeaderText="報名班級">
                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="STDate" HeaderText="開課日期"></asp:BoundColumn>
                                <asp:BoundColumn DataField="FTDate" HeaderText="結訓日期"></asp:BoundColumn>
                                <asp:BoundColumn DataField="RelEnterDate" SortExpression="RELENTERDATE" HeaderText="報名日期"></asp:BoundColumn>
                                <asp:BoundColumn DataField="ExamNO" SortExpression="EXAMNO" HeaderText="准考證號碼">
                                    <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn HeaderText="報名管道"></asp:BoundColumn>
                                <asp:BoundColumn HeaderText="是否試算"></asp:BoundColumn>
                                <asp:BoundColumn HeaderText="錄取結果"></asp:BoundColumn>
                                <%--<asp:TemplateColumn HeaderText="協助基金">
                                    <ItemTemplate><asp:Label ID="LBudgetID97" runat="server"></asp:Label></ItemTemplate>
                                </asp:TemplateColumn>--%>
                                <asp:BoundColumn HeaderText="結訓情形"></asp:BoundColumn>
                                <%--<asp:BoundColumn HeaderText="就業情形"></asp:BoundColumn>--%>
                                <asp:TemplateColumn HeaderText="功能">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                                    <ItemTemplate>
                                        <asp:LinkButton ID="btnEditView6" runat="server" Text="修改" CommandName="edit" CssClass="linkbutton"></asp:LinkButton><br />
                                        <asp:LinkButton ID="Button4" runat="server" Text="刪除" CommandName="del" CssClass="linkbutton"></asp:LinkButton>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                            </Columns>
                            <PagerStyle Visible="False"></PagerStyle>
                        </asp:DataGrid>
                    </div>
                    <div class="whitecol">
                        <asp:Button ID="Button9" runat="server" Text="查詢參訓歷史" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="Button13" runat="server" Text="近兩年參訓資料" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="Button7" Text="列印報名表" runat="server" CssClass="asp_Export_M" Visible="false"></asp:Button>
                        <asp:Button ID="btnExport1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                    </div>
                </td>
            </tr>
            <tr>
                <td>
                    <div align="center">
                        <input id="PrintShow" type="hidden" name="PrintShow" runat="server" />
                        <input id="Years" type="hidden" name="Years" runat="server" />
                        <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                    </div>
                </td>
            </tr>
        </table>
        <input id="isBlack" type="hidden" name="isBlack" runat="server" />
        <input id="Blackorgname" type="hidden" name="Blackorgname" runat="server" />
        <input id="HidOCID1" type="hidden" runat="server" />
        <asp:HiddenField ID="Hid_PreUseLimited18a" runat="server" />
        <asp:Literal ID="JAVASCRIPT_LITERAL" runat="server"></asp:Literal>
    </form>
</body>
</html>
