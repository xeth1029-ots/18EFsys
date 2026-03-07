<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_010_pop.aspx.vb" Inherits="WDAIIP.SD_05_010_pop" EnableEventValidation="false" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>學員參訓歷史</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" type="text/javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script language="javascript" type="text/javascript" src="../../js/common.js"></script>
    <script language="javascript" type="text/javascript" src="../../js/common.js"></script>
    <script language="javascript" type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;

        function printDoc() {
            if (_isIE) { window.print(); window.close(); }
            else { window.print(); }
            //if (!factory.object) {
            //    window.print();
            //    return false;
            //} else {
            //    factory.printing.header = '';
            //    factory.printing.footer = '';
            //    factory.printing.portrait = true;
            //    factory.printing.Print(true);
            //}
        }

        function chkdata() {
            var msg = '';
            if (document.form1.IDNO.value == '') msg += '請輸入身分證號碼\n';
            //else if(checkId(document.form1.IDNO.value)) msg+='身分證號碼錯誤!\n';

            if (msg != '') {
                //alert(msg)
                //return false;
            }
        }
        function GetMode() {
            document.form1.center.value = '';
            document.form1.RIDValue.value = '';
            document.form1.OCIDValue.value = '';
            document.form1.PlanID.value = '';
            for (var i = document.form1.OCID.options.length - 1; i >= 0; i--) {
                document.form1.OCID.options[i] = null;
            }
            document.form1.OCID.options[0] = new Option('請選擇機構');

            if (document.form1.DistID.selectedIndex != 0 && document.form1.TPlanID.selectedIndex != 0) {
                document.form1.Button3.disabled = false;
            }
            else {
                document.form1.Button3.disabled = true;
            }
        }
        function ShowPersonData(obj) {
            //var cst_inline = 'inline';
            var cst_inline = '';
            document.getElementById(obj).style.display = cst_inline; //'inline';
        }
        function HidPersonData(obj) {
            document.getElementById(obj).style.display = 'none';
        }
    </script>
    <style type="text/css">
        .style1 { color: #FF0000; }
    </style>
</head>
<body>
    <!-- MeadCo ScriptX -->
    <%--<object id="factory" style="display: none" codebase="../../scriptx/smsx.cab#Version=6,6,440,26" classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" viewastext></object>--%>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable2" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Panel ID="SearchTable" runat="server">
                        <table class="table_nw" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol" style="width: 20%">身分證號碼 </td>
                                <td class="whitecol" style="width: 30%">
                                    <asp:TextBox ID="IDNO" runat="server" Width="40%"></asp:TextBox>
                                </td>
                                <td class="bluecol" style="width: 20%">姓名 </td>
                                <td class="whitecol" style="width: 30%">
                                    <asp:TextBox ID="Name" runat="server" Width="40%"></asp:TextBox>
                                </td>
                            </tr>
                            <tr id="tr01a" runat="server">
                                <%--<td class="bluecol">轄區中心 </td>--%>
                                <td class="bluecol">轄區分署 </td>
                                <td colspan="3" class="whitecol">
                                    <asp:DropDownList ID="DistID" runat="server">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr id="tr01b" runat="server">
                                <td class="bluecol">訓練計畫 </td>
                                <td colspan="3" class="whitecol">
                                    <asp:DropDownList ID="TPlanID" runat="server">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr id="tr01c" runat="server">
                                <td class="bluecol">訓練機構 </td>
                                <td class="whitecol" colspan="3">
                                    <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                    <input id="RIDValue" type="hidden" runat="server"><input id="Button3" onclick="javascript: wopen('../../Common/MainOrg.aspx?DistID=' + document.form1.DistID.value + '&amp;TPlanID=' + document.form1.TPlanID.value, '訓練機構', 400, 400, 1)" type="button" value="..." name="Button3" runat="server" class="button_b_Mini">
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">通俗職類 </td>
                                <td class="whitecol" colspan="3">
                                    <asp:TextBox ID="txtCJOB_NAME" runat="server" Columns="30" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                    <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="button_b_Mini">
                                    <input id="cjobValue" type="hidden" name="cjobValue" runat="server">
                                </td>
                            </tr>
                            <tr id="tr01d" runat="server">
                                <td class="bluecol">班別 </td>
                                <td class="whitecol" colspan="3">
                                    <asp:DropDownList ID="OCID" runat="server">
                                    </asp:DropDownList>
                                    <input id="OCIDValue" type="hidden" runat="server"><input id="PlanID" type="hidden" runat="server">
                                    <asp:Button ID="Button2" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                            <tr id="tr01e" runat="server">
                                <td class="bluecol">訓練區間 </td>
                                <td class="whitecol" colspan="3"><font color="#ffffff">
                                    <asp:TextBox ID="STDate" runat="server" Columns="10" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top"><font color="#000000">～</font>
                                    <asp:TextBox ID="FTDate" runat="server" Columns="10" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top">
                                </font>&nbsp;&nbsp;&nbsp;&nbsp; </td>
                            </tr>
                            <tr id="tr02d" runat="server">
                                <td class="bluecol">匯入查詢資料 </td>
                                <td class="whitecol" colspan="3">
                                    <input id="File1" type="file" name="File1" runat="server" size="50" accept=".csv" />
                                    <asp:Button ID="Button13" runat="server" Text="查詢匯入資料" CssClass="asp_button_M"></asp:Button><br>
                                    (必須為csv格式)
								<%--
									<asp:hyperlink id="HyperLink1" runat="server" ForeColor="#8080FF" NavigateUrl="../../Doc/Stud_searchhistory.zip"
										CssClass="font">下載整批上載格式檔</asp:hyperlink>
                                --%>
                                </td>
                            </tr>
                        </table>
                        <table width="100%">
                            <tr>
                                <td align="center" class="whitecol">
                                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                    <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                    <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                            <tr>
                                <td align="center">
                                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                </td>
            </tr>
        </table>
        <table class="font" id="ShowDataTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td align="right">
                    <table class="font" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:Label ID="TPlanName" runat="server"></asp:Label>
                            </td>
                            <td align="right">(為避免消耗主機效能，最大搜尋筆數為2000筆)共計：
							<asp:Label ID="RecordCount" runat="server" />&nbsp;筆資料 </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td align="center" class="whitecol">
                    <input id="Button5" type="button" value="回上一頁" name="Button4" runat="server" class="asp_button_M" />
                </td>
            </tr>
            <tr id="Lab_TR" runat="server">
                <td>滑鼠移到姓名上可以觀看個人資料 </td>
            </tr>
            <tr>
                <td>
                    <div style="overflow-y: auto; height: 600px;">
                        <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" AllowSorting="True" AllowPaging="True" PageSize="20" AutoGenerateColumns="False" CssClass="font" CellPadding="8">
                            <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                            <Columns>
                                <asp:BoundColumn HeaderText="序號">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:TemplateColumn SortExpression="Name" HeaderText="姓名">
                                    <HeaderStyle ForeColor="#00ffff"></HeaderStyle>
                                    <ItemTemplate>
                                        <asp:Label ID="Label1" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.Name") %>'>
                                        </asp:Label>
                                        <table class="table_nw" id="Table4" style="display: none; position: absolute; border-collapse: collapse" cellspacing="1" cellpadding="1" width="450" bgcolor="white" border="0" bordercolor="#81ADE4" runat="server">
                                            <tr>
                                                <td class="bluecol_sub">姓名 </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="LName" runat="server"></asp:Label>
                                                </td>
                                                <td class="bluecol_sub">身分證號碼 </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="LIDNO" runat="server"></asp:Label>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td class="bluecol_sub">生日 </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="LBirthday" runat="server"></asp:Label>
                                                </td>
                                                <td class="bluecol_sub">性別 </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="LSex" runat="server"></asp:Label>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td class="bluecol_sub">身分別 </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="LIdent" runat="server"></asp:Label>
                                                </td>
                                                <td class="bluecol_sub">聯絡電話 </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="LTel" runat="server"></asp:Label>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td class="bluecol_sub">地址 </td>
                                                <td colspan="3" class="whitecol">
                                                    <asp:Label ID="LAddress" runat="server"></asp:Label>
                                                </td>
                                            </tr>
                                        </table>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:BoundColumn DataField="IDNO" SortExpression="IDNO" HeaderText="身分證號碼">
                                    <HeaderStyle ForeColor="#00ffff"></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="Birthday" SortExpression="Birthday" HeaderText="出生日期">
                                    <HeaderStyle ForeColor="#00ffff"></HeaderStyle>
                                </asp:BoundColumn>
                                <%--<asp:BoundColumn DataField="DistName" SortExpression="DistName" HeaderText="轄區&lt;BR&gt;中心">--%>
                                <asp:BoundColumn DataField="DistName" SortExpression="DistName" HeaderText="轄區&lt;BR&gt;分署">
                                    <HeaderStyle HorizontalAlign="Center" ForeColor="#00ffff"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="Years" HeaderText="年度">
                                    <HeaderStyle></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="PlanName" HeaderText="訓練計畫">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="OrgName" SortExpression="OrgName" HeaderText="訓練機構">
                                    <HeaderStyle HorizontalAlign="Center" ForeColor="#00ffff"></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="TMID" SortExpression="TMID" HeaderText="訓練職類">
                                    <HeaderStyle HorizontalAlign="Center" ForeColor="#00ffff"></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="CJOB_NAME" HeaderText="通俗職類">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="ClassName" SortExpression="ClassName" HeaderText="班別名稱">
                                    <HeaderStyle HorizontalAlign="Center" ForeColor="#00ffff"></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="THours" HeaderText="受訓&lt;BR&gt;時數">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="TRound" SortExpression="TRound" HeaderText="受訓期間">
                                    <HeaderStyle HorizontalAlign="Center" ForeColor="#00ffff"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <%--<asp:BoundColumn DataField="SkillName" HeaderText="技能檢定"><HeaderStyle HorizontalAlign="Center"></HeaderStyle></asp:BoundColumn>--%>
                                <asp:BoundColumn DataField="WEEKS" HeaderText="上課時間">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="TFlag" HeaderText="訓練&lt;BR&gt;狀態">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </asp:BoundColumn>
                                <%--<asp:BoundColumn DataField="RejectDayIn14" HeaderText="遞補期限內&lt;BR&gt;離訓(※註)"><HeaderStyle HorizontalAlign="Center"></HeaderStyle></asp:BoundColumn>--%>
                                <asp:BoundColumn DataField="MEMO1" HeaderText="備註">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                </asp:BoundColumn>
                            </Columns>
                            <PagerStyle Visible="False"></PagerStyle>
                        </asp:DataGrid>
                    </div>
                    <%-- <asp:BoundColumn DataField="JobStatus" HeaderText="訓後&lt;BR&gt;就業&lt;BR&gt;狀況">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
</asp:BoundColumn>
<asp:BoundColumn DataField="JobOrgName" HeaderText="就業單位名稱">
<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
</asp:BoundColumn>--%>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                </td>
            </tr>
            <%--<tr><td class="style1">※註：屬「遞補期限內離訓」(不含遞補期限內退訓)者，不列入不予錄訓規定之「職前訓練參訓紀錄」計算。 </td></tr>--%>
            <tr>
                <td align="center" class="whitecol">
                    <table>
                        <tr id="trRBListExpType" runat="server">
                            <td class="bluecol">匯出檔案格式</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                                    <asp:ListItem Value="ODS">ODS</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol" colspan="2">
                                <input id="Button4" type="button" value="回上一頁" name="Button4" runat="server" class="asp_button_M">
                                <asp:Button ID="but_P" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                                <asp:Button ID="but_f" Style="display: none" runat="server" Text="列印後動作" CssClass="asp_Export_M"></asp:Button>
                                <asp:Button ID="but_S" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                                <input id="Btnclose" type="button" value="關閉" name="Btnclose" runat="server" class="asp_button_M">
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
