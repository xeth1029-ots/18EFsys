<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_02_001.aspx.vb" Inherits="WDAIIP.SD_02_001" %>

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>甄試成績登錄</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
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
        //其主要的功能為當使用者按下Enter時改成Tab
        function EnterToTab() {
            if (event.keyCode == 13) { event.keyCode = 9; }
            //按 左鍵37//按 右鍵39 //想要按上鍵能夠跳上一個TextBox則可以在上述的Function中加下下列判斷
            if (event.keyCode == 38) {
                //按 上鍵
                var inputs = document.getElementsByTagName("input"); //取得所有的input
                for (i = 0; i < inputs.length; i++) //對每一個input
                    if (/^text|checkbox/.test(inputs[i].type)) //如果是 Text Box
                        if (inputs[i] == event.srcElement) //如果是按下按鍵的 Text Box
                        {
                            inputs[i - 1].focus(); //則上一個TextBox取得駐點
                            break;
                        }
            }
            if (event.keyCode == 40) {
                //按 下鍵
                var inputs = document.getElementsByTagName("input"); //取得所有的input
                for (i = 0; i < inputs.length; i++) //對每一個input
                    if (/^text|checkbox/.test(inputs[i].type)) //如果是 Text Box
                        if (inputs[i] == event.srcElement) //如果是按下按鍵的 Text Box
                        {
                            inputs[i + 1].focus(); //則上一個TextBox取得駐點
                            break;
                        }
            }
        }

        //想要每一個 Text Box 的 onkeydown 事件都 Handle 則可以
        window.onload = function init() {
            //視窗載入完成時
            var inputs = document.getElementsByTagName("input"); //取得所有的input
            for (i = 0; i < inputs.length; i++) //對每一個input
                if (/^text|checkbox/.test(inputs[i].type)) {
                    //如果是 Text box
                    if (typeof (inputs[i].attachEvent) != "undefined") {
                        inputs[i].attachEvent("onkeydown", EnterToTab);
                        //加入 onkeydown 事件時做 EnterToTab
                    }
                    if (typeof (inputs[i].addEventListener) != "undefined") {
                        inputs[i].addEventListener("onkeydown", EnterToTab);
                        //加入 onkeydown 事件時做 EnterToTab
                    }
                }
        }

        function GETvalue() {
            document.getElementById('Button5').click();
        }

        function choose_class() {
            var RIDValue = document.getElementById('RIDValue');
            openClass('../02/SD_02_ch.aspx?RID=' + RIDValue.value);
        }

        function search() {
            var OCIDValue1 = document.getElementById('OCIDValue1');
            if (OCIDValue1.value == '') {
                alert('請先選擇班別\n');
                return false;
            }
        }

        var cst_writecol = 4; //筆
        var cst_oralcol = 5; //口
        var cst_Totalcol = 6; //總
        //var cst_EXAMPLUScol = 6; //甄試加分(加權3%) //var cst_IDENTITYID = 7; //身分別 //var cst_Totalcol = 8; //總 //免試狀態處理 //計算分數
        function Grade() {
            var msg = '';
            var Mytable = document.getElementById('DataGrid1');
            var vItemVar1 = document.getElementById('ItemVar1').value;
            var vItemVar2 = document.getElementById('ItemVar2').value;
            var MyGETTRAIN3 = document.getElementById('Hid_GETTRAIN3').value; //甄試方式
            for (i = 1; i < Mytable.rows.length; i++) {
                var MyGrade1 = Mytable.rows[i].cells[cst_writecol].children[0];
                var MyGrade2 = Mytable.rows[i].cells[cst_oralcol].children[0];
                var MyGrade3 = Mytable.rows[i].cells[cst_Totalcol].children[0];
                var vMyGrade1 = MyGrade1.value;
                var vMyGrade2 = MyGrade2.value;
                //var Flag = true;
                /*
                var Flag1 = false; //都要試算，有勾選加分
                if (Mytable.rows[i].cells.length >= (cst_EXAMPLUScol + 1)) {
                    var MyCheck = Mytable.rows[i].cells[cst_EXAMPLUScol].children[0];
                    Flag1 = MyCheck.checked; //有勾選加分
                }
                */
                var vMyGrade3 = 0;
                var blFlag2 = (MyGETTRAIN3.indexOf("2") >= 0); //需筆試
                var blFlag3 = (MyGETTRAIN3.indexOf("3") >= 0); //需口試
                //if (vMyGrade1 == -1 || vMyGrade2 == -1) { vMyGrade3 = -1; }
                //if (vMyGrade1 == -1) { vMyGrade1 = 0; } //如果是缺考以0分計算(筆試)
                //if (vMyGrade2 == -1) { vMyGrade2 = 0; } //如果是缺考以0分計算(口試)
                //需筆試但註記缺考or需口試但註記缺考
                if ((vMyGrade1 == -1 && blFlag2) || (vMyGrade2 == -1 && blFlag3)) { vMyGrade3 = -1; }
                if (vMyGrade1 == -1 || !blFlag2) { vMyGrade1 = 0; } //如果是缺考或免試以0分計算(筆試)
                if (vMyGrade2 == -1 || !blFlag3) { vMyGrade2 = 0; } //如果是缺考或免試以0分計算(口試)
                if (vMyGrade3 == 0) {
                    //若不為負1進入此功能
                    /*
                    if (Flag1) {
                        //有勾選加分1.03
                        //甄試成績算到小數點第1位，小數點第二位四捨五入
                        vMyGrade3 = ((vMyGrade1 * vItemVar1 / 100 + vMyGrade2 * vItemVar2 / 100) * 1.03).toFixed(1);
                    }
                    else {
                        vMyGrade3 = (vMyGrade1 * vItemVar1 / 100 + vMyGrade2 * vItemVar2 / 100).toFixed(1);
                    }
                    */
                    vMyGrade3 = (vMyGrade1 * vItemVar1 / 100 + vMyGrade2 * vItemVar2 / 100).toFixed(1);
                }
                //最高分數為100;
                if (vMyGrade3 > 100) { vMyGrade3 = 100; }
                MyGrade3.value = vMyGrade3;
            }
            if (msg != '') { alert(msg); }
            else {
                //alert('試算成功!');
                blockAlert("試算成功!");
            }
            return false;
        }

        function chkdata() {
            var msg = '';
            var mytext1;
            var mytext2;
            var mytext3;
            var OCIDValue1 = document.getElementById('OCIDValue1');
            var TMIDValue1 = document.getElementById('TMIDValue1');
            if (OCIDValue1.value == '') msg += '請選擇班別!\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
            if (TMIDValue1.value == '') msg += '請選擇班別!\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
            var Mytable = document.getElementById('DataGrid1');
            for (var i = 1; i < Mytable.rows.length; i++) {
                mytext1 = Mytable.rows[i].cells[cst_writecol].children[0];
                mytext2 = Mytable.rows[i].cells[cst_oralcol].children[0];
                mytext3 = Mytable.rows[i].cells[cst_Totalcol].children[0];
                if (mytext1.value != '') {
                    //if(!IsNumeric(mytext1.value))msg+='筆試成績必須為正數!(第'+i+'行)\n';
                    //else
                    if (mytext1.value > 100) { msg += '筆試成績不能超過100!(第' + i + '行)\n'; }
                }
                if (mytext2.value != '') {
                    if (mytext2.value > 100) { msg += '口試成績不能超過100!(第' + i + '行)\n'; }
                }
                if (mytext3.value != '') {
                    if (mytext3.value > 100) { msg += '總成績不能超過100!(第' + i + '行)\n'; }
                }
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        /*
        function SelectAll(Flag) {
            var Mytable = document.getElementById('DataGrid1');
            for (i = 1; i < Mytable.rows.length; i++) {
                if (Mytable.rows(i).cells.length >= (cst_EXAMPLUScol + 1)) {
                    if (Mytable.rows(i).cells(cst_EXAMPLUScol).children(0)) {
                        Mytable.rows(i).cells(cst_EXAMPLUScol).children(0).checked = Flag;
                        //NotExam(Flag, i);
                    }
                }
            }
        }
        */

        var TimerID1;
        var TimerID2;
        function ShowEStudList(obj) {
            //alert(''); //debugger; //if(document.getElementById('DataGrid2').style.display=='inline'){
            if (document.getElementById('EStudList').style.display == 'none') {
                document.getElementById('EStudList').style.display = ''; //inline
                document.getElementById('LinkButton1').innerhtml = '*檢視目前e網報名尚未審核學員名單(關閉)';
                document.getElementById('DataGrid1').style.filter = 'alpha(opacity=100)';
                TimerID1 = setInterval("highlightit(30)", 50)
            }
            else {
                document.getElementById('EStudList').style.display = 'none';
                document.getElementById('LinkButton1').innerhtml = '*檢視目前e網報名尚未審核學員名單(開啟)';
                document.getElementById('DataGrid1').style.filter = 'alpha(opacity=30)';
                TimerID2 = setInterval("highlightit(100)", 50)
            }
            //fix 動態變動顯示內容, 會造成顯示內容超出 iframe 顯示區域被遮掉的情況 
            //if (!_isIE) { parent.setMainFrameHeight(); } //重新調整頁面高度
            if (parent && parent.setMainFrameHeight != undefined) { parent.setMainFrameHeight(); }
        }

        function highlightit(num) {
            var oDataGrid1 = document.getElementById('DataGrid1');
            if (num != 100) { //表示要透明化
                if (oDataGrid1.filters.alpha.opacity > num)
                    oDataGrid1.filters.alpha.opacity -= 15;
                else
                    clearInterval(TimerID1);
            }
            else {
                if (oDataGrid1.filters.alpha.opacity < 100)
                    oDataGrid1.filters.alpha.opacity += 15;
                else
                    clearInterval(TimerID2);
            }
        }

        function CheckTab(e) {
            if (e.keyCode != '9' && e.keyCode != '13') {
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" id="table2" cellspacing="1" cellpadding="1" border="0">
                        <tr>
                            <td class="font">
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;甄試成績登錄</asp:Label>
                                <%--<font color="#000000">(如需匯入甄試成績，請先匯出名冊套用)</font>--%>
                            </td>
                        </tr>
                    </table>
                    <table class="table_nw" id="table3" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol" width="20%">訓練機構 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" Width="70%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="Button8" type="button" value="..." name="Button8" runat="server" class="asp_button_Mini" />
                                <asp:Button ID="Button5" Style="display: none" runat="server" Text="Button5"></asp:Button>
                                <span id="HistoryList2" style="display: none; z-index: 100; position: absolute" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="20%">職類/班級 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="40%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="asp_button_Mini" />
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                <span id="HistoryList" style="display: none; z-index: 102; left: 28%; position: absolute">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">匯入甄試成績 </td>
                            <td class="whitecol">
                                <input id="File1" type="file" size="40" name="File1" runat="server" accept=".csv,.xls,.ods" />
                                <asp:Button ID="Button7" runat="server" Text="匯入名冊" CssClass="asp_Export_M"></asp:Button>
                                <asp:HyperLink ID="HyperLink2" runat="server" NavigateUrl="../../Doc/Result_v21.zip" ForeColor="#8080FF">(匯入檔案必須為csv、ods或xls格式)</asp:HyperLink>
                                <asp:Button ID="btnExportIdentity" runat="server" Text="身分別代碼" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">新式讀卡機-甄試成績匯入 </td>
                            <td class="whitecol">
                                <input id="File2" type="file" size="40" name="File1" runat="server" accept=".xls,.ods" />
                                <asp:Button ID="Button6" runat="server" Text="匯入成績" CssClass="asp_Export_M"></asp:Button>
                                <asp:HyperLink ID="Hyperlink1" runat="server" NavigateUrl="../../Doc/spi1_v21.zip" ForeColor="#8080FF">(匯入檔案必須為ods或xls格式)</asp:HyperLink>
                            </td>
                        </tr>
                        <tr id="Trwork2013a" runat="server">
                            <td class="bluecol">就服單位協助報名 </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="rblEnterPathW" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="A" Selected="True">不區分</asp:ListItem>
                                    <asp:ListItem Value="Y">是</asp:ListItem>
                                    <asp:ListItem Value="N">否</asp:ListItem>
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
                        <tr>
                            <td align="center" class="whitecol" colspan="2">
                                <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="Button4" runat="server" Text="匯出名冊" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <table style="z-index: 3; width: 400px">
                        <tr>
                            <td></td>
                            <td valign="top">
                                <asp:LinkButton ID="LinkButton1" runat="server" Width="300px" ForeColor="Blue" Visible="False">*檢視目前e網報名尚未審核學員名單</asp:LinkButton></td>
                        </tr>
                        <tr>
                            <td></td>
                            <td valign="top">
                                <div id="EStudList" style="display: none; z-index: 3; height: 64px; background-color: white">
                                    <asp:DataGrid ID="DataGrid2" runat="server" Width="256px" Font-Size="X-Small" AutoGenerateColumns="False" BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" BorderWidth="1px" CellPadding="3">
                                        <AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
                                        <HeaderStyle CssClass="head_navy" />
                                        <Columns>
                                            <asp:TemplateColumn HeaderText="序號">
                                                <ItemStyle></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:TextBox ID="stud1" runat="server" Width="20px" onfocus="this.blur()"></asp:TextBox>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:BoundColumn DataField="IDNO_MK" HeaderText="身分證號碼">
                                                <ItemStyle></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="Name" HeaderText="姓名">
                                                <ItemStyle></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="RelEnterDate" HeaderText="報名日" DataFormatString="{0:d}">
                                                <ItemStyle></ItemStyle>
                                            </asp:BoundColumn>
                                        </Columns>
                                        <PagerStyle HorizontalAlign="Left" ForeColor="#000066" BackColor="White" Mode="NumericPages"></PagerStyle>
                                    </asp:DataGrid>
                                </div>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="DataGridTable" style="z-index: 1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>點選標題可以排序 <font color="red">學員缺考者請在缺考欄位輸入-1</font>
                                <asp:Label ID="ArgRole" runat="server"></asp:Label><br />
                                <asp:Label ID="labmsg219" runat="server" ForeColor="Blue"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AutoGenerateColumns="False" AllowSorting="true" CssClass="font" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:BoundColumn DataField="ExamNo" SortExpression="ExamNo" HeaderText="准考證號/報名序號">
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Name" SortExpression="Name" HeaderText="姓名">
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="IDNO_MK" SortExpression="IDNO" HeaderText="身分證號碼">
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="RelEnterDate" SortExpression="RelEnterDate" HeaderText="報名日" DataFormatString="{0:d}">
                                            <HeaderStyle HorizontalAlign="Center" Width="3%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn SortExpression="WriteResult" HeaderText="筆試成績">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" CssClass="whitecol"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:TextBox ID="TextBox1" runat="server" Width="100%" MaxLength="3" Text='<%# DataBinder.Eval(Container.DataItem,"WriteResult")%>'>
                                                </asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn SortExpression="OralResult" HeaderText="口試成績">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" CssClass="whitecol"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:TextBox ID="TextBox2" runat="server" Width="100%" MaxLength="3" Text='<%# DataBinder.Eval(Container.DataItem,"OralResult")%>'>
                                                </asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <%--<asp:TemplateColumn HeaderText="加權3%">
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" CssClass="whitecol"></ItemStyle>
                                            <HeaderTemplate>
                                                加權3%
											<input type="checkbox" onclick="SelectAll(this.checked);" />
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <input id="EXAMPLUS" type="checkbox" runat="server" />
                                                <input id="Hid_SETID" runat="server" type="hidden" />
                                                <input id="Hid_EnterDate" runat="server" type="hidden" />
                                                <input id="Hid_SerNum" runat="server" type="hidden" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>--%>
                                        <%--<asp:TemplateColumn HeaderText="身分別">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" CssClass="whitecol"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:DropDownList ID="TYPE_EIdentityID" runat="server">
                                                </asp:DropDownList>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>--%>
                                        <asp:TemplateColumn SortExpression="TotalResult" HeaderText="總成績">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" CssClass="whitecol"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:TextBox ID="TextBox3" runat="server" Width="100%" onkeydown="return CheckTab(event);" Text='<%# DataBinder.Eval(Container.DataItem,"TotalResult")%>'>
                                                </asp:TextBox>
                                                <input id="Hid_SETID" runat="server" type="hidden" />
                                                <input id="Hid_EnterDate" runat="server" type="hidden" />
                                                <input id="Hid_SerNum" runat="server" type="hidden" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <%--<asp:TemplateColumn HeaderText="券別">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="L_TRNDType" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>--%>
                                        <%--<asp:TemplateColumn HeaderText="甄試加分">
										<HeaderStyle HorizontalAlign="Center" Width="40px"></HeaderStyle>
										<ItemStyle HorizontalAlign="Center"></ItemStyle>
										<HeaderTemplate>
											甄試<br />
											加分<br />
											<input type="checkbox" onclick="SelectAll(this.checked);" />
										</HeaderTemplate>
										<ItemTemplate>
											<input id="EXAMPLUS" type="checkbox" runat="server" />
										</ItemTemplate>
									</asp:TemplateColumn>--%>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <div align="center" class="whitecol">
                                    <asp:Button ID="Button2" runat="server" Text="試算" CssClass="asp_button_M"></asp:Button>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <div align="center" class="whitecol">
                                    <asp:Button ID="Button3" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                                </div>
                            </td>
                        </tr>
                    </table>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
                    </div>
                </td>
            </tr>
        </table>
        <input id="ItemVar1" type="hidden" name="ItemVar1" runat="server" />
        <input id="ItemVar2" type="hidden" name="ItemVar2" runat="server" />
        <asp:HiddenField ID="Hid_OCID1" runat="server" />
        <asp:HiddenField ID="Hid_GETTRAIN3" runat="server" />
    </form>
</body>
</html>
