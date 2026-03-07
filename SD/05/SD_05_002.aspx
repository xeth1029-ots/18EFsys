<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_002.aspx.vb" Inherits="WDAIIP.SD_05_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>學員出缺勤作業</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/TIMS.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        //TIMS.js: FrameLoad();
        function GETvalue() {
            document.getElementById('Button8').click();
        }

        function SetOneOCID() {
            document.getElementById('Button9').click();
        }

        function search() {
            var msg = '';
            var start_date = document.getElementById('start_date');
            var end_date = document.getElementById('end_date');
            //if(document.form1.OCIDValue1.value=='') msg+='請選擇職類班別!\n';
            if (start_date.value != '' && !checkDate(start_date.value)) msg += '請假起始日日期格式不正確\n';
            if (end_date.value != '' && !checkDate(end_date.value)) msg += '請假終至日日期格式不正確\n';

            if (msg != '') {
                alert(msg)
                return false;
            }
        }

        function schocid() {
            var msg = '';
            var OCIDValue1 = document.getElementById('OCIDValue1');
            if (OCIDValue1.value == '') msg += '請選擇職類班別!\n';
            if (msg != '') {
                alert(msg)
                return false;
            }
        }

        function choose_class() {
            var RIDValue = document.getElementById('RIDValue');
            var OCID1 = document.getElementById('OCID1');
            var Button9 = document.getElementById('Button9');

            if (OCID1.value == '') { Button9.click(); }
            openClass('../02/SD_02_ch.aspx?RID=' + RIDValue.value, 'Class');
        }
        function ShowFrame() {
            var FrameObj = document.getElementById('FrameObj');
            var HistoryTable = document.getElementById('HistoryTable');
            var HistoryList = document.getElementById('HistoryList');
            FrameObj.height = HistoryTable.rows.length * 20;
            if (FrameObj.height > 67) { FrameObj.height = 67; }
            FrameObj.style.display = HistoryList.style.display;
        }
    </script>
</head>
<body onload="FrameLoad();">
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;學員出缺勤作業</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td align="center">
                    <%--<table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">學員出缺勤作業</font> </td>
					</tr>
				</table>--%>
                    <table class="table_sch" id="Table3" cellpadding="1" cellspacing="1" width="100%">
                        <tr>
                            <td class="bluecol" width="20%">訓練機構 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" runat="server">
                                <input id="Button6" type="button" value="..." name="Button6" runat="server" class="button_b_Mini">
                                <asp:Button ID="Button9" Style="display: none" runat="server"></asp:Button>
                                <asp:Button ID="Button8" Style="display: none" runat="server" Text="Button8"></asp:Button>
                                <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%">
                                    </asp:Table>
                                </span></td>
                        </tr>
                        <tr>
                            <td class="bluecol">類別/班別 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="button_b_Mini">
                                <input id="TMIDValue1" type="hidden" name="Hidden1" runat="server">
                                <input id="OCIDValue1" type="hidden" name="Hidden3" runat="server">
                                <span id="HistoryList" style="z-index: 1; position: absolute; display: none; left: 270px">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                                <iframe id="FrameObj" style="position: absolute; display: none; left: 270px" frameborder="0" width="310" height="0"></iframe>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">通俗職類 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="30%"></asp:TextBox>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="button_b_Mini">
                                <input id="cjobValue" type="hidden" name="cjobValue" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">學員姓名 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="StdName" runat="server" Width="20%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">請假期間 </td>
                            <td class="whitecol" width="30%">
                                <span id="span1" runat="server">
                                    <asp:TextBox ID="start_date" runat="server" Width="33%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">～
							    <asp:TextBox ID="end_date" runat="server" Width="33%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                </span>
                            </td>
                            <td class="bluecol" width="20%">假別 </td>
                            <td class="whitecol" width="30%">
                                <asp:DropDownList ID="LeaveID" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr id="file_tr" runat="server">
                            <td class="bluecol">匯入刷卡紀錄 </td>
                            <td class="whitecol" colspan="3">
                                <input id="File1" type="file" name="File1" runat="server" size="40" accept=".xls,.ods" />
                                <asp:Button ID="btnImport" runat="server" Text="紀錄匯入" CssClass="asp_button_M"></asp:Button>(必須為ods或xls格式)
							<asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="../../Doc/Stud_CardRecord2_v21.zip" CssClass="font" ForeColor="#8080FF">下載整批上載格式檔</asp:HyperLink>
                            </td>
                        </tr>
                    </table>
                    <p align="center" class="whitecol">
                        <asp:Label ID="labPageSize" runat="server" font-family="Arial, Helvetica, sans-serif" Font-Size="9pt" DESIGNTIMEDRAGDROP="30" ForeColor="SlateBlue">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                        <asp:Button ID="Button1" runat="server" Text="查詢學員" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="Button2" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="Button4" runat="server" Text="查詢班級記錄" ToolTip="學員姓名、假別不列入查詢條件" CssClass="asp_button_M"></asp:Button>
                    </p>
                    <%--<p style="margin-top: 3px; margin-bottom: 3px" align="center">--%>
                    <div align="center">
                        <table id="Table4" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td>
                                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" AllowPaging="True" CellPadding="8">
                                        <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                        <Columns>
                                            <asp:BoundColumn HeaderText="序號">
                                                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="3%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="LastDate" HeaderText="最近登錄日期">
                                                <HeaderStyle HorizontalAlign="Center" Width="6%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="CLASSCNAME2" HeaderText="班別"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="STUDID2" HeaderText="學號">
                                                <HeaderStyle HorizontalAlign="Center" Width="4%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="Name" HeaderText="學員姓名">
                                                <HeaderStyle HorizontalAlign="Center" Width="6%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="Sick" HeaderText="病假"><HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle><ItemStyle HorizontalAlign="Center"></ItemStyle></asp:BoundColumn>
                                            <asp:BoundColumn DataField="Reason" HeaderText="事假"><HeaderStyle HorizontalAlign="Center" ></HeaderStyle><ItemStyle HorizontalAlign="Center"></ItemStyle></asp:BoundColumn>
                                            <asp:BoundColumn DataField="Publics" HeaderText="公假"><HeaderStyle HorizontalAlign="Center"></HeaderStyle><ItemStyle HorizontalAlign="Center"></ItemStyle></asp:BoundColumn>
                                            <asp:BoundColumn DataField="Skips" HeaderText="曠課"><HeaderStyle HorizontalAlign="Center"></HeaderStyle><ItemStyle HorizontalAlign="Center"></ItemStyle></asp:BoundColumn>
                                            <asp:BoundColumn DataField="Dead" HeaderText="喪假"><HeaderStyle HorizontalAlign="Center"></HeaderStyle><ItemStyle HorizontalAlign="Center"></ItemStyle></asp:BoundColumn>
                                            <asp:BoundColumn DataField="Late" HeaderText="遲到"><HeaderStyle HorizontalAlign="Center"></HeaderStyle><ItemStyle HorizontalAlign="Center"></ItemStyle></asp:BoundColumn>
                                            <%--<asp:BoundColumn DataField="Marry" HeaderText="婚假"><HeaderStyle HorizontalAlign="Center"></HeaderStyle><ItemStyle HorizontalAlign="Center"></ItemStyle></asp:BoundColumn>
                                            <asp:BoundColumn DataField="Birth" HeaderText="陪產假"><HeaderStyle HorizontalAlign="Center"></HeaderStyle><ItemStyle HorizontalAlign="Center"></ItemStyle></asp:BoundColumn>--%>
                                            <asp:BoundColumn DataField="notPunch" HeaderText="未打卡"><HeaderStyle HorizontalAlign="Center"></HeaderStyle><ItemStyle HorizontalAlign="Center"></ItemStyle></asp:BoundColumn>
                                            <asp:BoundColumn DataField="isolation" HeaderText="防疫隔離假"><HeaderStyle HorizontalAlign="Center"></HeaderStyle><ItemStyle HorizontalAlign="Center"></ItemStyle></asp:BoundColumn>
                                            <asp:BoundColumn DataField="detect" HeaderText="核酸檢測假"><HeaderStyle HorizontalAlign="Center"></HeaderStyle><ItemStyle HorizontalAlign="Center"></ItemStyle></asp:BoundColumn>
                                            <asp:BoundColumn DataField="vaccine" HeaderText="疫苗接種假"><HeaderStyle HorizontalAlign="Center"></HeaderStyle><ItemStyle HorizontalAlign="Center"></ItemStyle></asp:BoundColumn>                                            
                                            <asp:BoundColumn DataField="DaSource" HeaderText="資料來源">
                                                <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="功能">
                                                <HeaderStyle HorizontalAlign="Center" Width="6%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Button ID="Button3" runat="server" Text="修改" CssClass="asp_button_M"></asp:Button>
                                                    <asp:Button ID="Button5" runat="server" Text="查看" CssClass="asp_button_M"></asp:Button>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                        <PagerStyle Visible="False"></PagerStyle>
                                    </asp:DataGrid>
                                    <%--<asp:DataGrid ID="DataGrid1B" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" AllowPaging="True" CellPadding="8">
                                        <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                        <Columns>
                                            <asp:BoundColumn HeaderText="序號">
                                                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" Width="3%"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="LastDate" HeaderText="最近登錄日期" DataFormatString="{0:d}">
                                                <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="CLASSCNAME2" HeaderText="班別"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="StudentID" HeaderText="學號">
                                                <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="Name" HeaderText="學員姓名">
                                                <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="Publics" HeaderText="公假">
                                                <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="Dead" HeaderText="喪假">
                                                <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="Health" HeaderText="生理假">
                                                <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="Sick" HeaderText="病假">
                                                <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="Reason" HeaderText="事假">
                                                <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="Skips" HeaderText="曠課">
                                                <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="Late" HeaderText="遲到">
                                                <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="notPunch" HeaderText="未打卡">
                                                <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="DaSource" HeaderText="資料來源">
                                                <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="功能">
                                                <HeaderStyle HorizontalAlign="Center" Width="6%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Button ID="Button3" runat="server" Text="修改" CssClass="asp_button_M"></asp:Button>
                                                    <asp:Button ID="Button5" runat="server" Text="查看" CssClass="asp_button_M"></asp:Button>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                        <PagerStyle Visible="False"></PagerStyle>
                                    </asp:DataGrid>--%>
                                    <%--<p align="center">&nbsp;</p>--%>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <div align="center">
                                        <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                    </div>
                                </td>
                            </tr>
                        </table>
                        <asp:DataGrid ID="Datagrid2" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" Visible="False" CellPadding="8">
                            <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                            <Columns>
                                <asp:BoundColumn DataField="YearMon" HeaderText="年月">
                                    <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="Recorded" HeaderText="已紀錄(已填缺曠課)">
                                    <HeaderStyle HorizontalAlign="Center" Width="25%"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="NoRecord" HeaderText="紀錄狀況(無缺課、尚未記錄之日期)">
                                    <HeaderStyle HorizontalAlign="Center" Width="60%"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                </asp:BoundColumn>
                            </Columns>
                            <PagerStyle Visible="False"></PagerStyle>
                        </asp:DataGrid>
                    </div>
                    <div class="whitecol">
                        <asp:Button ID="Button7" runat="server" Text="回上一頁" Visible="False" CssClass="asp_button_M"></asp:Button>
                    </div>
                    <div>
                        <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
                    </div>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
