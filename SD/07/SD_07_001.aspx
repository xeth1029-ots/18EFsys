<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_07_001.aspx.vb" Inherits="WDAIIP.SD_07_001" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>技能檢定作業</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script language="javascript" type="text/javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script language="javascript" type="text/javascript" src="../../js/common.js"></script>
    <script language="javascript" type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script language="javascript" type="text/javascript">
        function GETvalue() {
            var btnGETvalue1 = document.getElementById('btnGETvalue1');
            btnGETvalue1.click();
        }
        function SetOneOCID() {
            var btnSetOneOCID = document.getElementById('btnSetOneOCID');
            btnSetOneOCID.click();
        }
        function choose_class() {
            var RID = document.form1.RIDValue;
            openClass('../02/SD_02_ch.aspx?RID=' + RID.value);
        }

        function search() {
            var msg = '';
            var OCIDValue1 = document.getElementById('OCIDValue1');
            if (OCIDValue1.value == '') {
                msg += '請選擇班級職類\n';
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function addnew1() {
            var msg = '';
            var OCIDValue1 = document.getElementById('OCIDValue1');
            var tExamName = document.getElementById('tExamName');
            if (OCIDValue1.value == '') {
                msg += '請選擇班級職類\n';
            }
            if (tExamName.value == '') {
                msg += '[新增]請輸入選擇檢定職類\n';
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        //num=1表示全選   3表示全選個人 --2表示全選日期
        function select_all(num, nn, index) {
            //nn: this.checked
            var Mytable = document.getElementById("DataGrid1b");
            var Mycheckbox;
            var Mytextbox;
            var Myimg;
            var Mydrop;
            var k, j //k~j

            //如果全選個人，調整k j到該列
            if (num == 3) {
                k = index - 1;
                j = index;
            }
            else {
                k = 1
                j = Mytable.rows.length;
            }

            //跑所有的rows
            for (var i = k; i < j; i++) {
                Mytextbox = Mytable.rows[i].cells[0].children[1];
                var PassValue = Mytextbox.value;
                MyCheck = Mytable.rows[i].cells[0].children[0];

                if ((num == 1 || num == 3) && !MyCheck.disabled) {
                    if (nn) {
                        if (PassValue != 'Y' && PassValue != 'y') {
                            Mycheckbox = Mytable.rows[i].cells[0].children[0];
                            Mycheckbox.checked = nn;
                            Mydrop = Mytable.rows[i].cells[3].children[0];
                            Mydrop.disabled = false;
                            Mytextbox = Mytable.rows[i].cells[4].children[0];
                            Mytextbox.disabled = false;
                            if (Mytextbox.value == '') Mytextbox.value = (new Date()).getFullYear() + '/' + ((new Date()).getMonth() + 1) + '/' + (new Date()).getDate();
                            Myimg = Mytable.rows[i].cells[4].children[1];
                            Myimg.style.display = 'inline';
                        }
                    }
                    else {
                        if (PassValue != 'Y' && PassValue != 'y') {
                            Mycheckbox = Mytable.rows[i].cells[0].children[0];
                            Mycheckbox.checked = nn;
                            Mydrop = Mytable.rows[i].cells[3].children[0];
                            Mydrop.disabled = true;
                            Mytextbox = Mytable.rows[i].cells[4].children[0];
                            Mytextbox.disabled = true;
                            Mytextbox.value = '';
                            Myimg = Mytable.rows[i].cells[4].children[1];
                            Myimg.style.display = 'none';
                        }
                    }
                }
            }
        }

        //檢定級別
        function GetAllLevel(obj) {
            var Mytable = document.getElementById("DataGrid1b");
            var cell_i = 3;
            var j = 0;
            var MyLevel = null;
            var ListValue = '';
            //CheckBoxList: getCheckboxValue//getCheckBoxListValue
            //hasCheckBoxListIndex//chkValue//setRadioValue
            if (obj.checked) {
                for (var i = 1; i < Mytable.rows.length; i++) {
                    var MyLevel = Mytable.rows[i].cells[cell_i].children[0];
                    if (!MyLevel.disabled) {
                        j = i + 1;
                        ListValue = getCheckBoxListValue(MyLevel);
                        break;
                    }
                }

                //if (MyLevel.selectedIndex != 0) {
                if (ListValue != '00000') {
                    if (confirm('您要以第一位未離退訓學員當預設值嗎?')) {
                        for (j; j < Mytable.rows.length; j++) {
                            var MyDrop = Mytable.rows[j].cells[cell_i].children[0];
                            if (!MyDrop.disabled) {
                                setSPANValue(MyDrop, ListValue);
                                //alert(vMsg); break;
                                //MyDrop.selectedIndex = MyLevel.selectedIndex;
                            }
                        }
                    }
                    else {
                        obj.checked = false;
                    }
                }
                else {
                    alert('未設定第一位未離退訓學員的檢定級別');
                    obj.checked = false;
                }
            }
        }

        //申請日 GetAllApplyDate
        function GetAllApplyDate(obj) {
            var Mytable = document.getElementById("DataGrid1b");
            var cell_i = 4;
            var applydate1 = '';
            if (obj.checked) {
                //取得值applydate1
                for (var i = 1; i < Mytable.rows.length; i++) {
                    var applydate2 = Mytable.rows[i].cells[cell_i].children[0];
                    if (!applydate2.disabled && applydate2.value != '') {
                        applydate1 = applydate2.value;
                        var j = i + 1;
                        break;
                    }
                }
                if (applydate1 != '') {    //複製第一個申請日欄位
                    if (confirm('您要以第一位未離退訓學員當預設值嗎?')) {
                        for (j; j < Mytable.rows.length; j++) {
                            var applydate3 = Mytable.rows[j].cells[cell_i].children[0];
                            if (!applydate3.disabled && applydate3.value == '') {
                                applydate3.value = applydate1;
                            }
                        }
                    }
                    else {
                        obj.checked = false;
                    }
                }
                else {    //如果沒有輸入第一個申請日,則代今天日期
                    for (var k = 1; k < Mytable.rows.length; k++) {
                        var applydatee4 = Mytable.rows[k].cells[cell_i].children[0];
                        var d = new Date();
                        if (!applydatee4.disabled && applydatee4.value == '') {
                            applydatee4.value = d.getFullYear() + '/' + (d.getMonth() + 1) + '/' + d.getDate();
                        }
                    }
                }
            }
        }

        //全選檢定結果
        function GetAllExamresult(obj) {
            var Mytable = document.getElementById("DataGrid1c");
            var MyLevel = Mytable.rows[1].cells[6].children[0];
            if (obj.checked) {
                if (confirm('您要以第一位學員當預設值嗎?')) {
                    for (var i = 2; i < Mytable.rows.length; i++) {
                        var MyDrop = Mytable.rows[i].cells[6].children[0];
                        if (!MyDrop.disabled) {
                            MyDrop.selectedIndex = MyLevel.selectedIndex;
                        }
                    }
                }
                else {
                    obj.checked = false;
                }
            }
        }
        //全選檢定日期
        function GetAllExamDate(obj) {
            var Mytable = document.getElementById("DataGrid1c");
            var MyLevel = Mytable.rows[1].cells[7].children[0];
            if (obj.checked) {
                if (MyLevel.value != '') {
                    if (confirm('您要以第一位學員當預設值嗎?')) {
                        for (var i = 2; i < Mytable.rows.length; i++) {
                            var MyText = Mytable.rows[i].cells[7].children[0];
                            if (!MyText.disabled) {
                                MyText.value = MyLevel.value;
                            }
                        }
                    }
                    else {
                        obj.checked = false;
                    }
                }
                else {
                    alert('未設定第一位學員的檢定日');
                    obj.checked = false;
                }
            }
        }


        //全選製證號日期
        function Getlicensedate(obj) {
            var Mytable = document.getElementById("DataGrid1c");   //取得第一個輸入製證日期的欄位
            if (obj.checked) {
                for (var i = 1; i < Mytable.rows.length; i++) {
                    var isnotpass = Mytable.rows[i].cells[6].children[0];
                    var licensedate2 = Mytable.rows[i].cells[8].children[0]
                    if (isnotpass.value == 'Y' && licensedate2.value != '') {
                        var licensedatef = licensedate2.value;
                        var j = i + 1;
                        break;
                    }
                }
                var NGk = true;
                if (licensedate2.value != '') {
                    //判斷學員是合格的並複製第一個製證日期欄位
                    if (confirm('您要以此位合格學員當預設值嗎?')) {
                        //var NGk = true;
                        for (j; j < Mytable.rows.length; j++) {
                            var isnotpass3 = Mytable.rows[j].cells[6].children[0];
                            var licensedate3 = Mytable.rows[j].cells[8].children[0]
                            if (isnotpass3.value == 'Y') {
                                NGk = false;
                                licensedate3.value = licensedatef;
                            }
                        }
                        if (NGk) { alert('查無合格學員!!'); }
                    }
                    else {
                        obj.checked = false;
                    }
                }
                else {    //如果沒有輸入第一個製證日期,則代今天日期
                    //var NGk = true;
                    for (var k = 1; k < Mytable.rows.length; k++) {
                        var isnotpass4 = Mytable.rows[k].cells[6].children[0];
                        var licensedate4 = Mytable.rows[k].cells[8].children[0]
                        var d = new Date();
                        if (isnotpass4.value == 'Y') {
                            NGk = false;
                            licensedate4.value = d.getFullYear() + '/' + (d.getMonth() + 1) + '/' + d.getDate();
                        }
                    }
                    if (NGk) { alert('查無合格學員!!'); }
                }
            }
        }

        function GetlicenseNO(obj) {
            //使證號連續,取得第一個輸入證號
            var Mytable = document.getElementById("DataGrid1c");
            if (obj.checked) {
                for (var i = 1; i < Mytable.rows.length; i++) {
                    var isnotpass = Mytable.rows[i].cells[6].children[0];
                    var licensedate2 = Mytable.rows[i].cells[9].children[0];
                    if (isnotpass.value == 'Y' && licensedate2.value != '') {
                        var licensedatef = licensedate2.value;
                        var j = i + 1;
                        break;
                    }
                }
                if (licensedate2.value != '') {  //產生連續證號
                    if (confirm('您要以此位合格學員當預設值嗎?')) {
                        var NGk = true;
                        for (j; j < Mytable.rows.length; j++) {
                            var isnotpass3 = Mytable.rows[j].cells[6].children[0];
                            var licensedate3 = Mytable.rows[j].cells[9].children[0];
                            if (isnotpass3.value == 'Y') {
                                if (isNaN(parseInt(licensedatef, 10))) {
                                    alert(licensedate2.value + '不是數字,請輸入數字,或改用單一輸入');
                                    return false;
                                }
                                licensedatef = parseInt(licensedatef, 10) + 1;
                                var wlen = (licensedatef + "").length;
                                for (k = 0; k < licensedate2.value.length - wlen; k++) {
                                    licensedatef = "0" + licensedatef;
                                }
                                licensedate3.value = licensedatef;
                                NGk = false;
                            }
                        }
                        if (NGk) { alert('查無合格學員!!'); }
                    }
                    else {
                        obj.checked = false;
                    }
                }
                else {
                    alert('未設定第一位合格學員的證號');
                    obj.checked = false;
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
                    <table id="Table2" class="font" cellspacing="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label><asp:Label ID="TitleLab2" runat="server">
							首頁&gt;&gt;學員動態管理&gt;&gt;技能檢定管理&gt;&gt;技能檢定作業
                                </asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="table_nw" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="bluecol" style="width: 20%">訓練機構 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" Width="55%" onfocus="this.blur()"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                <input id="btnSetLevOrg" type="button" value="..." runat="server" class="asp_button_Mini" />
                                <asp:Button ID="btnSetOneOCID" Style="display: none" runat="server"></asp:Button>
                                <asp:Button ID="btnGETvalue1" Style="display: none" runat="server"></asp:Button>
                                <span onclick="GETvalue()" id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span></td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">職類/班別 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="25%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input id="TMIDValue1" type="hidden" runat="server">
                                <input id="OCIDValue1" type="hidden" runat="server" />
                                <input onclick="choose_class()" type="button" value="..." class="asp_button_Mini" />
                                <span id="HistoryList" style="position: absolute; left: 270px; display: none">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span></td>
                        </tr>
                        <tr>
                            <td class="bluecol">檢定職類/名稱 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="tExamKind" runat="server" Width="25%" MaxLength="5"></asp:TextBox>
                                <asp:TextBox onfocus="this.blur()" ID="tExamName" runat="server" Width="30%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol" colspan="2">
                                <p align="center">
                                    <asp:Button ID="btnSearch" Style="z-index: 0" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button><font face="新細明體">&nbsp;</font>
                                    <asp:Button ID="btnAdd1" Style="z-index: 0" runat="server" Text="申請設定" CssClass="asp_button_M" Width="88px"></asp:Button>
                                </p>
                            </td>
                        </tr>
                    </table>
                    <table id="ShowDataTable" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" Style="z-index: 0" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" DataKeyField="OCID" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號">
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassCName" HeaderText="班級名稱">
                                            <HeaderStyle HorizontalAlign="Center" Width="19%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDate" HeaderText="開訓日期" DataFormatString="{0:d}">
                                            <HeaderStyle HorizontalAlign="Center" Width="19%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ExamKind" HeaderText="檢定職類">
                                            <HeaderStyle HorizontalAlign="Center" Width="19%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ExamName" HeaderText="檢定名稱">
                                            <HeaderStyle HorizontalAlign="Center" Width="19%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <HeaderStyle HorizontalAlign="Center" Width="19%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" Wrap="false"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Button ID="btnEdit2" runat="server" Text="結果輸入" CommandName="btnEdit2" CssClass="asp_button_M"></asp:Button>
                                                <asp:Button ID="btnDel2" runat="server" Text="刪除" CommandName="btnDel2" CssClass="asp_button_M"></asp:Button>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                                <p style="margin-bottom: 3px; margin-top: 3px" align="center">
                                    <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                                </p>
                            </td>
                        </tr>
                    </table>
                    <asp:DataGrid ID="DataGrid1b" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" CellPadding="8">
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <HeaderStyle CssClass="head_navy" HorizontalAlign="Center" />
                        <Columns>
                            <asp:TemplateColumn>
                                <HeaderStyle Width="5%" />
                                <ItemStyle HorizontalAlign="Center" />
                                <HeaderTemplate>
                                    <input onclick="select_all(1, this.checked)" type="checkbox">
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <input id="Checkbox2" type="checkbox" runat="server" name="Checkbox2">
                                    <input id="hidSOCID" type="hidden" value='<%#DataBinder.Eval(Container.DataItem,"SOCID")%>' runat="server" name="hidSOCID">
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn DataField="STUDID2" HeaderText="學號">
                                <HeaderStyle Width="15%" />
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Name" HeaderText="學員">
                                <HeaderStyle Width="15%" />
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="檢定級別">
                                <HeaderStyle Width="50%" />
                                <ItemStyle HorizontalAlign="Center" />
                                <HeaderTemplate>
                                    檢定級別<input id="TrainLevel" onclick="GetAllLevel(this);" type="checkbox">
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <asp:CheckBoxList ID="cbl1TrainLevel" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                        <asp:ListItem Value="1">甲級&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</asp:ListItem>
                                        <asp:ListItem Value="2">乙級&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</asp:ListItem>
                                        <asp:ListItem Value="3">丙級&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</asp:ListItem>
                                        <asp:ListItem Value="4">單一級&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</asp:ListItem>
                                        <asp:ListItem Value="5">不分級</asp:ListItem>
                                    </asp:CheckBoxList>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="申請日">
                                <HeaderStyle Width="15%"></HeaderStyle>
                                <ItemStyle CssClass="whitecol" HorizontalAlign="Center" />
                                <HeaderTemplate>
                                    <input type="checkbox" onclick="GetAllApplyDate(this);">申請日
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <asp:TextBox ID="tApplyDate" runat="server" Width="60%" CssClass="whitecol"></asp:TextBox>
                                    <img id="IMG1" style="cursor: pointer" alt="" src="../../images/show-calendar.gif" align="top" runat="server">
                                    <%--<img id="IMG1" style="cursor: pointer" onclick="javascript:show_calendar('<%= tApplyDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />--%>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                    </asp:DataGrid>
                    <p style="margin-bottom: 3px; margin-top: 3px" align="center">
                        <asp:Label ID="msg2" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                    </p>
                    <asp:DataGrid ID="DataGrid1c" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" CellPadding="8">
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <HeaderStyle CssClass="head_navy" Width="10%" HorizontalAlign="Center" />
                        <Columns>
                            <asp:BoundColumn DataField="STUDID2" HeaderText="學號">
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Name" HeaderText="學員">
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="ExamKind" HeaderText="檢定職類">
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="檢定級別">
                                <HeaderTemplate>
                                    檢定級別
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <asp:Label ID="lExamLevel" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn DataField="ApplyDate" HeaderText="申請日" DataFormatString="{0:d}">
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="ExamName" HeaderText="檢定名稱"></asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="檢定結果">
                                <ItemStyle CssClass="whitecol" HorizontalAlign="Center" />
                                <HeaderTemplate>
                                    檢定結果<input id="Examresult" type="checkbox" onclick="GetAllExamresult(this);">
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <asp:DropDownList ID="ddlExamPass" runat="server">
                                        <asp:ListItem Value="">缺考</asp:ListItem>
                                        <asp:ListItem Value="Y">合格</asp:ListItem>
                                        <asp:ListItem Value="N">不合格</asp:ListItem>
                                        <%--<asp:ListItem>未測試</asp:ListItem>--%>
                                    </asp:DropDownList>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="檢定日">
                                <ItemStyle CssClass="whitecol" HorizontalAlign="Center" Font-Size="Small" />
                                <HeaderTemplate>
                                    檢定日<input id="Examdate" type="checkbox" onclick="GetAllExamDate(this);">
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <asp:TextBox ID="Textbox6" runat="server" Text='<%# DataBinder.Eval(Container.DataItem, "ExamDate") %>' Width="60%">
                                    </asp:TextBox>
                                    <%--<img style="cursor: pointer" onclick="javascript:show_calendar('<%= Textbox6.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />--%>
                                    <img id="Img3" style="cursor: pointer" alt="" src="../../images/show-calendar.gif" align="top" runat="server">
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="製證日">
                                <ItemStyle CssClass="whitecol" HorizontalAlign="Center" Font-Size="Small" />
                                <HeaderTemplate>
                                    製證日<input id="licensedate" onclick="Getlicensedate(this);" type="checkbox">
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <asp:TextBox ID="TextBox4" runat="server" Text='<%# DataBinder.Eval(Container.DataItem,"SendoutCertDate") %>' Width="60%">
                                    </asp:TextBox><img id="IMG2" style="cursor: pointer" alt="" src="../../images/show-calendar.gif" align="top" runat="server">
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="證號">
                                <ItemStyle CssClass="whitecol" HorizontalAlign="Center" />
                                <HeaderTemplate>
                                    證號<input id="licenseNO" onclick="GetlicenseNO(this);" type="checkbox">
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <asp:TextBox ID="TextBox5" runat="server" Text='<%# DataBinder.Eval(Container.DataItem,"ExamNo") %>' Width="60%">
                                    </asp:TextBox>
                                    <input id="hidSOCID" type="hidden" value='<%#DataBinder.Eval(Container.DataItem,"SOCID")%>' runat="server">
                                    <input id="HidSTXID" type="hidden" value='<%#DataBinder.Eval(Container.DataItem,"STXID")%>' runat="server">
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                    </asp:DataGrid>
                    <p style="margin-bottom: 3px; margin-top: 3px" align="center">
                        <asp:Label ID="msg3" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                    </p>
                </td>
            </tr>
            <tr>
                <td>
                    <p align="center" class="whitecol">
                        <asp:Button ID="btnSaveData1" runat="server" Text="存檔" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="btnExport1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                        <asp:Button ID="btnBack1" runat="server" Text="回上頁" CssClass="asp_button_M"></asp:Button>
                    </p>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="hid_CTEID" runat="server" />
        <asp:HiddenField ID="hid_OCID" runat="server" />
        <asp:HiddenField ID="hid_ExamTime" runat="server" />
        <asp:HiddenField ID="hid_EXAMKIND" runat="server" />
    </form>
</body>
</html>
