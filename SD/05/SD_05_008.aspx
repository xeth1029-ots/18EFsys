<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_05_008.aspx.vb" Inherits="WDAIIP.SD_05_008" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>結訓學員資料卡登錄</title>
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
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        var cst_inline1 = "";
        function choose_class() {
            if (document.getElementById('OCID1').value == '') {
                document.getElementById('Button13').click();
            }
            openClass('../02/SD_02_ch.aspx?RID=' + document.form1.RIDValue.value);
        }

        function SetOneOCID() {
            document.getElementById('Button13').click();
        }

        function CheckPrint() {
            //var flag = false;
            var MyTable = document.getElementById('DataGrid2');
            var HidStudentID = document.getElementById('HidStudentID');
            //HidStudentID.value = '';
            var vStudentID = '';
            for (var i = 1; i < MyTable.rows.length; i++) {
                var MyCheck = MyTable.rows[i].cells[0].children[0];
                if (MyCheck.checked) {
                    if (vStudentID != '') vStudentID += ',';
                    vStudentID += '\'' + MyCheck.value + '\'';
                }
            }
            HidStudentID.value = vStudentID;
            if (vStudentID == '') {
                alert('請勾選要列印的學員!');
                return false;
            }
            return true;
            //e lse {window.open(url + 'DLID=' + document.getElementById('DLID').value + '&OCID=' + document.getElementById('OCID').value + '&StudentID=' + StudentID, 'print', 'toolbar=0,location=0,status=0,menubar=0,resizable=1')}
        }

        function choose(num) {
            //num 1:ClassTable  2:StudentTable
            //查詢出來的DataGrid
            var mytable3 = document.getElementById('ClassTable');
            //根據班級查出的DataGrid
            var mytable4 = document.getElementById('StudentTable');

            var mybut = document.getElementById('Button6');
            var mylable = document.getElementById('LabelMsg1');

            if (mytable3) mytable3.style.display = 'none';
            if (mytable4) mytable4.style.display = 'none';

            mybut.style.display = 'none';
            mylable.innerHTML = '';
            document.getElementById('msg').innerHTML = '';

            var cst_pt1 = 0;
            var cst_pt2 = 1;
            if (document.getElementsByName('RadioButtonList1').length > 2) {
                cst_pt1 = 1; //cst_pt RadioButtonList1
                cst_pt2 = 2;
            }
            if (document.getElementsByName('RadioButtonList1')[cst_pt1].checked) {
                document.getElementById('TR1_1').style.display = cst_inline1; //'inline';
                document.getElementById('TR1_2').style.display = cst_inline1; //'inline';

                document.getElementById('TR1_3').style.display = cst_inline1; //'inline';
                document.getElementById('TR1_4').style.display = cst_inline1; //'inline';
                document.getElementById('TR2_1').style.display = 'none';
                document.getElementById('TR2_2').style.display = 'none';
            }
            else if (document.getElementsByName('RadioButtonList1')[cst_pt2].checked) {
                document.getElementById('TR1_1').style.display = 'none';
                document.getElementById('TR1_2').style.display = 'none';

                document.getElementById('TR1_3').style.display = 'none';
                document.getElementById('TR1_4').style.display = 'none';
                document.getElementById('TR2_1').style.display = cst_inline1; //'inline';
                document.getElementById('TR2_2').style.display = cst_inline1; //'inline';
                mybut.style.display = 'inline';
            }

            if (num == 1) {
                if (mytable3) mytable3.style.display = cst_inline1; //'inline';
            }
            else if (num == 2) {
                if (mytable4) mytable4.style.display = cst_inline1; //'inline';
            }

            <%-- 
            因為這裡會動態變動顯示內容, 造成顯示內容超出 iframe 顯示區域的情況
            顯示內容切換完後, 
            要呼叫主框頁面中的 setMainFrameHeight() 去更新 iframe 顯示區域高度
            --%>
            if (window.top && window.top.setMainFrameHeight != undefined) {
                window.top.setMainFrameHeight();
            }
        }

        function search() {
            var msg = '';
            //if (!isChecked(document.form1.RadioButtonList1)) msg += '請選擇局屬或非局屬\n';
            if (!isChecked(document.form1.RadioButtonList1)) msg += '請選擇署屬或非分署屬\n';
            if (getRadioValue(document.form1.RadioButtonList1) == 1) {
                if (document.getElementById('FTDate1').value != '')
                    if (!checkDate(document.getElementById('FTDate1').value)) msg += '結訓日期起日必須為正確的時間格式\n';
                if (document.getElementById('FTDate2').value != '')
                    if (!checkDate(document.getElementById('FTDate2').value)) msg += '結訓日期迄日必須為正確的時間格式\n';
            }

            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function showFrame() {
            document.getElementById('frTest').style.display = document.getElementById('HistoryList2').style.display;
        }


        function ChangeAll(obj) {
            var objLen = document.form1.length;
            for (var iCount = 0; iCount < objLen; iCount++) {
                if (obj.checked == true) {
                    if (document.form1.elements[iCount].type == "checkbox")
                    { document.form1.elements[iCount].checked = true; }
                }
                else {
                    if (document.form1.elements[iCount].type == "checkbox")
                    { document.form1.elements[iCount].checked = false; }
                }
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <span id="TitleLab1">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;</span>
                                <span id="TitleLab2">結訓學員資料卡登錄</span>
                            </td>
                        </tr>
                    </table>
                    <table id="SearchTable" cellspacing="0" cellpadding="0" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <table class="table_sch" id="Table3" cellpadding="1" cellspacing="1" width="100%">
                                    <tr id="MainTr" runat="server">
                                        <%--<td class="bluecol" style="width: 20%">局/非局屬 </td>--%>
                                        <td class="bluecol" style="width: 20%">署/非署屬 </td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="RadioButtonList1" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                                <%--<asp:ListItem Value="1">局屬</asp:ListItem>--%>
                                                <asp:ListItem Value="1">署屬</asp:ListItem>
                                                <%--<asp:ListItem Value="2">非局屬</asp:ListItem>--%>
                                                <asp:ListItem Value="2">非署屬</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr id="TR1_1" runat="server">
                                        <td class="bluecol">訓練機構 </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Columns="35" Width="60%"></asp:TextBox>
                                            <input id="Button11" type="button" value="..." name="Button11" runat="server">
                                            <input id="RIDValue" type="hidden" runat="server">
                                            <asp:Button ID="Button13" Style="display: none" runat="server"></asp:Button>
                                            <span id="HistoryList2" style="z-index: 2; position: absolute; display: none">
                                                <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                            </span>
                                        </td>
                                    </tr>
                                    <tr id="TR1_2" runat="server">
                                        <td class="bluecol">訓練計畫 </td>
                                        <td class="whitecol">
                                            <asp:DropDownList ID="TPlan" runat="server">
                                            </asp:DropDownList>
                                            <iframe id="frTest" style="z-index: 1; position: absolute; width: 310px; display: none; height: 23px; left: 120px" marginwidth="0" marginheight="0" src="" frameborder="0" scrolling="no"></iframe>
                                        </td>
                                    </tr>
                                    <tr id="TR1_3" runat="server">
                                        <td class="bluecol">職類/班別 </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                            <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                            <input onclick="choose_class()" type="button" value="...">
                                            <input id="TMIDValue1" type="hidden" runat="server">
                                            <input id="OCIDValue1" type="hidden" runat="server">
                                            <span id="HistoryList" style="position: absolute; display: none; left: 270px">
                                                <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                            </span>
                                        </td>
                                    </tr>
                                    <tr id="TR1_4" runat="server">
                                        <td class="bluecol">通俗職類 </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="30%"></asp:TextBox>
                                            <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server">
                                            <input id="cjobValue" type="hidden" name="cjobValue" runat="server">
                                        </td>
                                    </tr>
                                    <tr id="TR1_5" runat="server">
                                        <td class="bluecol">結訓日期 </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="FTDate1" runat="server" Columns="13" Width="15%"></asp:TextBox>
                                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= FTDate1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">～
										    <asp:TextBox ID="FTDate2" runat="server" Columns="13" Width="15%"></asp:TextBox>
                                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= FTDate2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                        </td>
                                    </tr>
                                    <tr id="TR2_1" runat="server">
                                        <td class="bluecol">訓練機構 </td>
                                        <td class="whitecol">
                                            <%--
										<asp:ListItem Value="007">退輔會訓練中心</asp:ListItem>
										<asp:ListItem Value="008">青輔會青年分署</asp:ListItem>
										<asp:ListItem Value="009">農委會漁業署遠洋漁業開發中心</asp:ListItem>
										<asp:ListItem Value="010">台北市分署</asp:ListItem>
										<asp:ListItem Value="011">高雄市訓練就業中心</asp:ListItem>
										<asp:ListItem Value="014">新北市政府職業訓練中心</asp:ListItem>
                                            --%>
                                            <asp:DropDownList ID="UnitCode" runat="server">
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                    <tr id="TR2_2" runat="server">
                                        <td class="bluecol">班別 </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="ClassName" runat="server" Width="60%"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <%--<tr>
                                        <td class="bluecol">匯出檔案格式</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                                                <asp:ListItem Value="ODS">ODS</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>--%>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>&nbsp;
                                <asp:Button ID="But1" runat="server" Text="結訓學員匯出功能" Enabled="False" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label><br>
                                <div align="left">
                                    <asp:Label ID="Label3" runat="server" CssClass="font"></asp:Label>
                                </div>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <table id="ClassTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Width="100%" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                        <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn HeaderText="序號">
                                <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構">
                                <HeaderStyle Width="30%"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="班別(點選可以修改封面)">
                                <HeaderStyle Width="25%"></HeaderStyle>
                                <ItemTemplate>
                                    <asp:Label ID="LabStart" runat="server"></asp:Label>
                                    <%--修改封面--%>
                                    <asp:LinkButton ID="LinkBtn2_edit" runat="server" CommandName="edit"></asp:LinkButton>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn DataField="total" HeaderText="結訓人數">
                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="num" HeaderText="填寫人數">
                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="功能">
                                <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                <ItemTemplate>
                                    <asp:Button ID="Btn1_view_std" runat="server" Text="查詢學員" CommandName="viewstd" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="Btn1_add_std" runat="server" Text="新增學員" CommandName="addstd" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="Btn2_add" runat="server" Text="新增封面" CommandName="add" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="Btn3_print" runat="server" Text="列印封面" CommandName="print" CssClass="asp_button_M"></asp:Button>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn Visible="False" DataField="OCID" HeaderText="OCID"></asp:BoundColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                </td>
            </tr>
            <tr>
                <td align="center" class="whitecol">
                    <asp:Button ID="Button6" runat="server" Text="新增班別封面檔" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
        </table>
        <table id="StudentTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td>
                    <asp:Label ID="LabelMsg1" runat="server" CssClass="font"></asp:Label>
                    <br>
                    <asp:Label ID="LabelMsg2" runat="server" CssClass="font"></asp:Label>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid2" runat="server" CssClass="font" Width="100%" AutoGenerateColumns="False" CellPadding="8">
                        <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:TemplateColumn HeaderText="列印">
                                <HeaderStyle Width="5%" />
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <HeaderTemplate>
                                    選取<input id="chkbox_all" type="checkbox" runat="server">
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <input id="Checkbox2" type="checkbox" runat="server">
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn HeaderText="學號">
                                <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="姓名(離退訓日期)">
                                <HeaderStyle Width="30%" />
                                <ItemTemplate>
                                    <asp:Label ID="LabStart2" runat="server"></asp:Label>
                                    <asp:Label ID="LabName" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn HeaderText="填寫狀態">
                                <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="功能">
                                <HeaderStyle HorizontalAlign="Center" Width="30%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Button ID="Button4" runat="server" Text="修改" CommandName="edit" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="Button9" runat="server" Text="刪除" CommandName="del" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="Button5" runat="server" Text="新增" CommandName="add" CssClass="asp_button_M"></asp:Button>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn Visible="False" DataField="DLID" HeaderText="DLID"></asp:BoundColumn>
                            <asp:BoundColumn Visible="False" DataField="SubNo" HeaderText="SubNo"></asp:BoundColumn>
                            <asp:BoundColumn Visible="False" DataField="OCID" HeaderText="OCID"></asp:BoundColumn>
                            <asp:BoundColumn Visible="False" DataField="StudentID" HeaderText="StudentID"></asp:BoundColumn>
                        </Columns>
                    </asp:DataGrid>
                </td>
            </tr>
            <tr>
                <td align="center" class="whitecol">
                    <asp:Button ID="Button12" runat="server" Text="列印學員空白資料卡" CssClass="asp_button_M"></asp:Button>
                    <asp:Button ID="Button7" runat="server" Text="列印結訓學員資料卡" CssClass="asp_button_M"></asp:Button>
                    <asp:Button ID="Button10" runat="server" Text="回班別列表" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
        </table>
        <input id="DLID" type="hidden" name="DLID" runat="server">
        <input id="OCID" type="hidden" name="OCID" runat="server">
        <input id="HidStudentID" type="hidden" runat="server">
    </form>
</body>
</html>
