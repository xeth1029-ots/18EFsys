<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_03_002.aspx.vb" Inherits="WDAIIP.SD_03_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD html 4.0 Transitional//EN">
<html>
<head>
    <title>學員資料維護</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-confirm.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery.blockUI.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/global.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        //產業人才投資方案與TIMS 計畫用
        function GETvalue() { document.getElementById('Button13').click(); }

        function SetOneOCID() { document.getElementById('Button14').click(); }

        function choose_class(num) {
            var vRID = $('#RIDValue').val();
            //var ImportTable = document.getElementById('ImportTable');
            document.form1.TMID1.value = '';
            document.form1.TMIDValue1.value = '';
            document.form1.OCID1.value = '';
            document.form1.OCIDValue1.value = '';
            document.form1.hidLockTime1.value = '1';
            //document.form1.Button2.disabled=true;
            //匯入學員名冊
            //if (ImportTable) { ImportTable.style.display = 'none'; }
            $('#trImport1').hide();
            //$('#trRBListExpType').hide();
            //$('#trExport1').hide();
            document.getElementById('DataGridTable').style.display = 'none'; //學員資料
            document.getElementById('msg').innerHTML = '';
            if (document.getElementById('OCIDValue1').value == '') { document.getElementById('Button14').click(); }
            openClass('../02/SD_02_ch.aspx?special=12&RWClass=1&RID=' + vRID);
        }

        function CheckPrint() {
            var flag = false;
            var MyTable = document.getElementById('DataGrid1');
            var StudentID = '';
            var OCIDValue1 = document.getElementById('OCIDValue1');
            var PrintValue = document.getElementById('PrintValue');
            for (var i = 1; i < MyTable.rows.length; i++) {
                var MyCheck = null;
                if (MyTable.rows[i].cells[0].children.length > 0) {
                    var ichk = MyTable.rows[i].cells[0].children.length - 1;
                    MyCheck = MyTable.rows[i].cells[0].children[ichk];
                }
                if (MyCheck != null && MyCheck.checked) {
                    flag = true;
                    if (StudentID != '') StudentID += ','
                    StudentID += '\'' + MyCheck.value + '\'';
                }
            }
            PrintValue.value = StudentID;
            //alert(StudentID);
            //alert(document.getElementById('PrintValue').value);
            if (PrintValue.value == '') {
                alert('請勾選要列印的學員!')
                return false;
            }
            else {
                //alert(document.getElementById('PrintValue').value);
                url = '../../SQControl.aspx?'
                url += '&path=TIMS';
                url += '&filename=in_class_stud';
                url += '&StudentID=' + PrintValue.value;
                url += '&OCID=' + OCIDValue1.value;
                window.open(url, 'print', 'toolbar=0,location=0,status=0,menubar=0,resizable=1');
            }
        }

        function InsertValue(Flag, MyValue) {
            //alert(Flag);
            //alert(MyValue);
            var PrintValue = document.getElementById('PrintValue');
            if (Flag) {
                //增加VALUE。
                if (PrintValue.value.indexOf('\'' + MyValue + '\'') == -1) {
                    if (PrintValue.value != '') PrintValue.value += ',';
                    PrintValue.value += '\'' + MyValue + '\'';
                }
            }
            else {
                //VALUE清除。
                if (PrintValue.value.indexOf('\'' + MyValue + '\'') != -1) {
                    //清理異常逗號
                    if (PrintValue.value.indexOf(',,') != -1) {
                        PrintValue.value = PrintValue.value.replace(',,', ',');
                    }
                    //中間
                    if (PrintValue.value.indexOf(',\'' + MyValue + '\',') != -1) {
                        PrintValue.value = PrintValue.value.replace(',\'' + MyValue + '\',', '');
                    }
                    //尾部
                    if (PrintValue.value.indexOf(',\'' + MyValue + '\'') != -1) {
                        PrintValue.value = PrintValue.value.replace(',\'' + MyValue + '\'', '');
                    }
                    //前面
                    if (PrintValue.value.indexOf('\'' + MyValue + '\',') != -1) {
                        PrintValue.value = PrintValue.value.replace('\'' + MyValue + '\',', '');
                    }
                    //任何位置
                    if (PrintValue.value.indexOf('\'' + MyValue + '\'') != -1) {
                        PrintValue.value = PrintValue.value.replace('\'' + MyValue + '\'', '');
                    }
                    //清理異常逗號
                    if (PrintValue.value.indexOf(',,') != -1) {
                        PrintValue.value = PrintValue.value.replace(',,', ',');
                    }
                }
            }
        }

        function ChangeAll(obj) {
            //alert(document.form1.Checkbox3.checked);
            //alert(document.getElementById('Checkbox3').value);
            //var MyTable=document.getElementById('DataGrid1');
            //for(i=1;i<MyTable.rows.length;i++){
            //	MyTable.rows(i).cells(0).checked=tf;
            //}
            var objLen = document.form1.length;
            //alert(objLen);
            for (var iCount = 0; iCount < objLen; iCount++) {
                if (obj.checked == true) {
                    if (document.form1.elements[iCount].type == "checkbox") {
                        document.form1.elements[iCount].checked = true;
                    }
                }
                else {
                    if (document.form1.elements[iCount].type == "checkbox") {
                        document.form1.elements[iCount].checked = false;
                    }
                }
            }
        }

        /*個資法js*/
        function showLoginPwdDiv(num) {
            var rblWorkMode_0 = document.getElementById('rblWorkMode_0');   //模糊顯示
            if (!rblWorkMode_0) { return; }
            var rblWorkMode_1 = document.getElementById('rblWorkMode_1');   //正常顯示 
            if (!rblWorkMode_1) { return; }

            //num: 1:查詢 2:匯出 (記錄目前查詢按鈕)
            var hidSchBtnNum = document.getElementById('hidSchBtnNum'); //記錄目前查詢按鈕
            hidSchBtnNum.value = num; //num: 1:查詢 2:匯出 (記錄目前查詢按鈕)

            var hidLockTime1 = document.getElementById('hidLockTime1');   //啟用鎖定
            var hidLockTime2 = document.getElementById('hidLockTime2');
            var OCIDValue1 = document.getElementById('OCIDValue1'); //班級
            var divPwdFrame = document.getElementById('divPwdFrame');
            var txtdivPxssward = document.getElementById('txtdivPxssward');
            //document.getElementById('divFrame').style.display = 'none';
            if (OCIDValue1.value == '') {
                alert('請選擇班級');
                return false;
            }
            var blnPwdFrame = false; //不顯示密碼輸入
            if (rblWorkMode_1.checked != true) { hidLockTime1.value = '1'; }
            if (rblWorkMode_1.checked == true && hidLockTime1.value == '1' && hidLockTime2.value == '1') {
                blnPwdFrame = true; //顯示密碼輸入
            }
            //alert(hidLockTime1.value);
            if (blnPwdFrame) {
                if (divPwdFrame) { divPwdFrame.style.display = 'inline'; } //顯示
                if (txtdivPxssward != null) txtdivPxssward.focus();
                return false;
            }
            else {
                if (divPwdFrame) { divPwdFrame.style.display = 'none'; } //display = 'none'
                return true;
            }
            unblockUI();
        }

        function chkTxtPxssward() {
            var txtdivPxssward = document.getElementById('txtdivPxssward');
            if (!txtdivPxssward) { return; }
            var msg = '';
            if (txtdivPxssward.value == '') msg = '請輸入您的個資安全密碼!';
            if (msg != '') {
                alert(msg);
                return false;
            }
            unblockUI();
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <div style="position: absolute; top: -333px">
            <input type="text" title="Chaff for Chrome Smart Lock" /><input type="password" title="Chaff for Chrome Smart Lock" />
        </div>
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;學員資料管理&gt;&gt;學員資料維護</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable2" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table id="table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <table class="table_nw" id="searchtable" cellspacing="1" cellpadding="1" width="100%">
                                    <tr>
                                        <td class="bluecol" width="20%">訓練機構 </td>
                                        <td class="whitecol" width="80%">
                                            <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                            <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                            <input id="Button8" type="button" value="..." name="Button8" runat="server" class="asp_button_Mini" />
                                            <asp:Button ID="Button14" Style="display: none" runat="server"></asp:Button>
                                            <asp:Button ID="Button13" Style="display: none" runat="server" Text="Button13"></asp:Button>
                                            <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                                <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                            </span>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need" width="20%">職類/班級 </td>
                                        <td class="whitecol" width="80%">
                                            <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                            <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                            <input id="button5" type="button" value="..." name="button5" runat="server" class="asp_button_Mini" />
                                            <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                            <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                            <span id="HistoryList" style="position: absolute; display: none; left: 28%">
                                                <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                            </span>
                                        </td>
                                    </tr>
                                    <tr id="tr_rblWorkMode" runat="server">
                                        <td class="bluecol" width="20%">資料顯示模式 </td>
                                        <td colspan="3" class="whitecol" width="80%">
                                            <asp:RadioButtonList ID="rblWorkMode" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                                <asp:ListItem Value="1" Selected="True">模糊顯示</asp:ListItem>
                                                <asp:ListItem Value="2">正常顯示</asp:ListItem>
                                            </asp:RadioButtonList>
                                            <%--<asp:HiddenField ID="hidWorkMode" runat="server" />--%>
                                        </td>
                                    </tr>
                                    <tr id="trImport1" runat="server">
                                        <td class="bluecol" width="20%">匯入學員名冊 </td>
                                        <td class="whitecol" width="80%">
                                            <input id="File1" type="file" name="File1" runat="server" size="60" accept=".xls,.ods" />
                                            <asp:Button ID="button7" runat="server" Text="匯入名冊" CssClass="asp_button_M"></asp:Button>
                                            <asp:Label ID="label1" runat="server">(必須為ods或xls格式)</asp:Label>
                                            <asp:HyperLink ID="hyperlink1" runat="server" ForeColor="#8080ff" CssClass="font">下載整批上載格式檔</asp:HyperLink>
                                        </td>
                                    </tr>
                                    <tr id="trRBListExpType" runat="server">
                                        <td class="bluecol" width="20%">匯出檔案格式</td>
                                        <td class="whitecol" width="80%">
                                            <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                                                <asp:ListItem Value="ODS">ODS</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr id="trExport1" runat="server">
                                        <td class="bluecol">匯出學員名冊 </td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="shiftsort" runat="server" CssClass="font" RepeatLayout="flow" RepeatDirection="horizontal">
                                                <asp:ListItem Value="1" Selected="true">以代號匯出</asp:ListItem>
                                                <asp:ListItem Value="2">以名稱匯出</asp:ListItem>
                                            </asp:RadioButtonList>
                                            <asp:Button ID="Button6" runat="server" Text="匯出名冊" CssClass="asp_Export_M"></asp:Button>
                                        </td>
                                    </tr>
                                    <tr id="tr_ddl_INQUIRY_S" runat="server">
                                        <td class="bluecol_need">查詢原因</td>
                                        <td class="whitecol" colspan="3">
                                            <asp:DropDownList ID="ddl_INQUIRY_Sch" runat="server"></asp:DropDownList>
                                        </td>
                                    </tr>
                                </table>
                                <table width="100%">
                                    <tr>
                                        <td align="center" class="whitecol">
                                            <asp:Label ID="labpagesize" runat="server" ForeColor="slateblue">顯示列數</asp:Label>
                                            <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                            <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                            <%--<asp:Button ID="button2" runat="server" Text="新增" Enabled="false" CssClass="asp_button_S"></asp:Button>--%>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <%-- <tr><td><table class="" id="ImportTable" cellspacing="1" width="100%" runat="server"></table></td></tr>--%>
                        <tr>
                            <td>
                                <div style="margin-top: 3px; margin-bottom: 3px" align="center">
                                    <asp:Label ID="msg" runat="server" ForeColor="red" CssClass="font"></asp:Label>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                    <tr>
                                        <td>
                                            <table class="font" id="table4" cellspacing="1" cellpadding="1" border="0" width="100%">
                                                <tr>
                                                    <td width="12%">受訓日期：</td>
                                                    <td width="24%">
                                                        <asp:Label ID="dateround" runat="server" CssClass="font"></asp:Label></td>
                                                    <td width="12%">開班人數：</td>
                                                    <td width="20%">
                                                        <asp:Label ID="tnum" runat="server" CssClass="font"></asp:Label></td>
                                                    <td width="12%">學員人數：</td>
                                                    <td width="20%">
                                                        <asp:Label ID="stdnum" runat="server" CssClass="font"></asp:Label></td>
                                                </tr>
                                                <tr>
                                                    <td width="12%">導師：</td>
                                                    <td colspan="5" width="88%">
                                                        <asp:Label ID="ctname" runat="server" CssClass="font"></asp:Label></td>
                                                </tr>
                                            </table>
                                            <font size="2"><font color="#ff0000" size="2">*表示為該學員有必填資料未填</font><br />
                                                <asp:Label ID="annotate" runat="server" Width="392px" ForeColor="red" Visible="false">#表示為該學員資料尚未確認或為退件修正狀態</asp:Label></font>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="false" AllowPaging="true" AllowCustomPaging="true" CellPadding="8">
                                                <AlternatingItemStyle BackColor="#F5F5F5" />
                                                <HeaderStyle CssClass="head_navy" />
                                                <Columns>
                                                    <asp:TemplateColumn HeaderText="選取">
                                                        <HeaderStyle Width="5%"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center" />
                                                        <HeaderTemplate>
                                                            選取<input id="checkbox3" type="checkbox" runat="server" />
                                                        </HeaderTemplate>
                                                        <ItemTemplate>
                                                            <asp:Label ID="star2" Visible="false" runat="server"><font color="#ff0000">#</font></asp:Label>
                                                            <asp:Label ID="star1" Visible="false" runat="server"><font color="#ff0000">*</font></asp:Label>
                                                            <input id="checkbox2" type="checkbox" runat="server" />
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:BoundColumn DataField="studentid" HeaderText="學號">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="name" HeaderText="姓名">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="idno" HeaderText="身分證號碼">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="sex" HeaderText="性別">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="birthday" HeaderText="出生日期" DataFormatString="{0:d}">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn HeaderText="報名路徑">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn HeaderText="學員狀態">
                                                        <HeaderStyle Width="5%"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="actno" HeaderText="投保單位<br>保險證號">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:TemplateColumn HeaderText="預算別">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                        <ItemStyle CssClass="whitecol" />
                                                        <ItemTemplate>
                                                            <asp:DropDownList ID="budid" runat="server"></asp:DropDownList>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="功能">
                                                        <HeaderStyle HorizontalAlign="Center" Width="10%" />
                                                        <ItemStyle HorizontalAlign="Center" Font-Size="Small" />
                                                        <ItemTemplate>
                                                            <asp:LinkButton ID="Button3" runat="server" Text="修改" CommandName="edit" CssClass="linkbutton"></asp:LinkButton>
                                                            <%--<asp:LinkButton ID="btn10Delete" runat="server" Text="刪除" CommandName="del" CssClass="linkbutton"></asp:LinkButton>--%>
                                                            <asp:HiddenField ID="Hid_idno" runat="server" />
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                </Columns>
                                                <PagerStyle Visible="false"></PagerStyle>
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
                                            <asp:Button ID="button9" runat="server" Text="查詢參訓歷史" CssClass="asp_button_M"></asp:Button>
                                            <asp:Button ID="Button4" runat="server" Text="列印資料卡" CssClass="asp_Export_M"></asp:Button>
                                            <asp:Button ID="Button11" runat="server" Text="學員資料確認" CssClass="asp_button_M"></asp:Button>
                                            <asp:Button ID="edit_but" runat="server" Text="學員資料審核" CommandName="edit" CssClass="asp_button_M"></asp:Button>
                                            <asp:Button ID="Button12" runat="server" Text="回上頁" Visible="false" CssClass="asp_button_M"></asp:Button>
                                            <%--<asp:LinkButton ID="btn10Delete" runat="server" Text="刪除" CommandName="del" CssClass="linkbutton"></asp:LinkButton>--%>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <input id="PrintValue" type="hidden" runat="server" />
        <div id="divPwdFrame" runat="server" style="position: absolute; border-width: 6px; border-style: double; border-color: #4682B4; display: none; width: 350px; height: 300px; left: 195px; top: 200px; background-color: #FFFAF0; padding-left: 30px; padding-top: 30px;">
            <table align="center">
                <tr>
                    <td>請輸入個資安全密碼 </td>
                </tr>
                <tr>
                    <td>
                        <asp:TextBox ID="txtdivPxssward" runat="server" TextMode="Password"></asp:TextBox></td>
                </tr>
                <tr>
                    <td></td>
                </tr>
                <tr>
                    <td align="center">
                        <asp:Button ID="btndivPwdSubmit" runat="server" Text="確定" OnClientClick="return chkTxtPxssward();" CssClass="asp_button_S" CommandName="btndivPwdSubmit" />&nbsp;
                        <input id="btn_close" type="button" value="關閉" onclick="document.getElementById('divPwdFrame').style.display = 'none'; document.getElementById('labChkMsg').text = '';" class="button_b_S" />
                    </td>
                </tr>
                <%--<tr><td></td></tr>--%>
                <tr>
                    <td align="center">
                        <asp:Label ID="labChkMsg" runat="server" CssClass="needFont"></asp:Label></td>
                </tr>
            </table>
        </div>
        <input id="hidLockTime1" type="hidden" name="hidLockTime1" runat="server" value="1" />
        <input id="hidSchBtnNum" type="hidden" name="hidSchBtnNum" runat="server" value="1" />
        <input id="hidLockTime2" type="hidden" name="hidLockTime2" runat="server" value="1" />
        <input id="Hid_show_actno_budid" type="hidden" runat="server" />
        <input id="Hid_nouse_SupplyID" type="hidden" runat="server" />
        <%--<asp:Button ID="save_but" runat="server" Text="儲存" Visible="false" CssClass="asp_button_S"></asp:Button>--%>
        <%--<asp:Panel ID="panelLoginDiv" runat="server" Style="position: absolute; width: 300; height: 300; left: 190px; top: 200px;"><div style="border-width: 6px; border-style: double; border-color: #4682B4; background-color: #FFFAF0; padding-left: 30px; padding-top: 30px;" id="divFrame" runat="server"></div></asp:Panel>--%>
    </form>
</body>
</html>
