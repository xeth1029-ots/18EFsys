<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_06_001.aspx.vb" Inherits="WDAIIP.SD_06_001" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>加退保申請</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">

        var cst_sid = 0;
        var cst_sname = 1;
        var cst_sidno = 2;
        var cst_salary = 3;
        var cst_aday1 = 4;
        var cst_aday2 = 5;
        var cst_dday1 = 6;
        var cst_dday2 = 7;
        var cst_reason = 8;

        //薪資全選控制
        function select_all2(obj) {
            var mytable = document.getElementById("DataGrid1");
            var myvalue = mytable.rows[1].cells[cst_salary].children[0];
            if (obj.checked) {
                if (myvalue.value != '') {
                    if (confirm('您要以第一位學員當預設值嗎?') == true) {
                        for (var i = 2; i < mytable.rows.length; i++) {
                            var myvalue2 = mytable.rows[i].cells[cst_salary].children[0];
                            myvalue2.value = myvalue.value;
                        }
                    }
                }
                else {
                    alert("未設定第一位學員的薪資");
                    obj.checked = false;
                }
            }
            else {
                obj.checked = false;
            }
        }

        function GETvalue() {
            document.getElementById('Button3').click();
        }

        function search() {
            var check1 = '';
            if (document.form1.OCIDValue1.value != '') {
                check1 += '1';
                //alert('請選擇班級職類!');
                //return false;
            }
            if (document.form1.FTDate1.value != '' || document.form1.FTDate2.value != '') {
                check1 += '1';
            }
            if (check1 == '') {
                alert('請選擇班級職類或結訓日期區間!');
                return false;
            }
        }

        function choose_class() {
            openClass('../02/SD_02_ch.aspx?RID=' + document.getElementById('RIDValue').value);
        }

        function Apply(num) {
            var mytable = document.getElementById("DataGrid1");
            var chkimg = mytable.rows[num - 1].cells[cst_aday1].children[0];
            var imgbut = mytable.rows[num - 1].cells[cst_aday2].children[1];
            var chkdate = mytable.rows[num - 1].cells[cst_aday2].children[0];

            if (chkimg.checked) {
                chkdate.disabled = false;
                imgbut.style.display = "inline";
                if (chkdate.value == '') chkdate.value = (new Date()).getFullYear() + '/' + ((new Date()).getMonth() + 1) + '/' + (new Date()).getDate();
            }
            else {
                chkdate.disabled = true;
                imgbut.style.display = "none";
                chkdate.value = '';
            }
        }
        function Dropout(num) {
            var mytable = document.getElementById("DataGrid1");
            var chkimg = mytable.rows[num - 1].cells[cst_aday1].children[0];
            var imgbut = mytable.rows[num - 1].cells[cst_aday2].children[1];
            var chkdate = mytable.rows[num - 1].cells[cst_aday2].children[0];

            if (chkimg.checked) {
                chkdate.disabled = false;
                imgbut.style.display = "inline";
                if (chkdate.value == '') chkdate.value = (new Date()).getFullYear() + '/' + ((new Date()).getMonth() + 1) + '/' + (new Date()).getDate();
            }
            else {
                chkdate.disabled = true;
                imgbut.style.display = "none";
                chkdate.value = '';
            }

        }

        //加、退保全選控制
        function select_all(num, nn) {
            var mytable = document.getElementById("DataGrid1");
            var chkdateM; 			//日期TextBox Value(取得有值第1位)
            chkdateM = '';
            for (var i = 1; i < mytable.rows.length; i++) {
                var mycheckbox; 			//每一列之前的checkbox
                var imgbut; 				//日期按鈕									
                var chkdate; 			//日期TextBox
                var gethid; 				//影藏欄位的值，判斷資料庫內是否有資料

                if (num == 1) {
                    mycheckbox = mytable.rows[i].cells[cst_aday1].children[0];
                    imgbut = mytable.rows[i].cells[cst_aday2].children[1];
                    chkdate = mytable.rows[i].cells[cst_aday2].children[0];
                    gethid = mytable.rows[i].cells[cst_aday2].children[2];
                    //document.write(mytable.rows(i).cells(5).children.length)
                }
                else if (num == 2) {
                    mycheckbox = mytable.rows[i].cells[cst_dday1].children[0];
                    imgbut = mytable.rows[i].cells[cst_dday2].children[1];
                    chkdate = mytable.rows[i].cells[cst_dday2].children[0];
                    gethid = mytable.rows[i].cells[cst_dday2].children[2];
                }

                if (chkdateM == '' && chkdate.value != '') {
                    chkdateM = chkdate.value;
                }

                //if (gethid.value=='N'){}
                if (mycheckbox.disabled == false) {
                    if (nn) {
                        imgbut.style.display = "inline";
                        chkdate.disabled = false;
                        if (chkdateM != '') {
                            if (chkdate.value == '') {
                                mycheckbox.checked = nn;
                                chkdate.value = chkdateM;
                            }
                        }
                        //if(chkdate.value=='') chkdate.value=(new Date()).getFullYear()+'/'+((new Date()).getMonth()+1)+'/'+(new Date()).getDate();
                    }
                    else {
                        if (chkdate.value == '') {
                            mycheckbox.checked = nn;
                            imgbut.style.display = "none";
                            chkdate.disabled = true;
                            chkdate.value = '';
                        }
                    }
                }

            }
        }

        //儲存檢查
        function chkdata() {
            var mytable = document.getElementById("DataGrid1");
            var msg = '';
            for (var i = 1; i < mytable.rows.length; i++) {
                var mytextbox1 = mytable.rows[i].cells[cst_salary].children[0];
                if (!isUnsignedInt(mytextbox1.value) && mytextbox1.value != '') msg += '薪資必須為數字(第' + i + '列)\n';

                var mytextbox2 = mytable.rows[i].cells[cst_aday2].children[0];
                var mytextbox3 = mytable.rows[i].cells[cst_dday2].children[0];
                if (mytextbox1.value != '' && mytextbox2.value == '') msg += '有填入薪資必須填入加保日(第' + i + '列)\n';
                if (mytextbox1.value == '' && mytextbox2.value != '') msg += '有填入加保日必須填入薪資(第' + i + '列)\n';
                if (mytextbox2.value == '' && mytextbox3.value != '') msg += '有填入退保日必須填入加保日(第' + i + '列)\n';
                if (!checkDate(mytextbox2.value) && mytextbox2.value != '') msg += '加保日的日期格式不正確(第' + i + '列)\n';
                if (!checkDate(mytextbox3.value) && mytextbox3.value != '') msg += '退保日的日期格式不正確(第' + i + '列)\n';
                //if (mytextbox2.value==mytextbox3.value && mytextbox2.value!='' && mytextbox3.value!='') msg+='加保日與退保日不能為同一天(第'+i+'列)\n';
                if (compareDate(mytextbox2.value, mytextbox3.value) == 1 && mytextbox2.value != '' && mytextbox3.value != '') msg += '退保日不能在加保日之前(第' + i + '列)\n';
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
            else {
                return true;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td align="center">
                    <%--<table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            首頁&gt;&gt;學員動態管理&gt;&gt;加退保管理&gt;&gt;<font color="#990000">加退保申請</font>
                        </td>
                    </tr>
                </table>--%>
                    <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;加退保管理&gt;&gt;加退保申請</asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="table_nw" id="Table3" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol">訓練機構
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="Button8" type="button" value="..." name="Button8" runat="server" class="asp_button_Mini" /><br />
                                <asp:Button ID="Button3" Style="display: none" runat="server" Text="Button3"></asp:Button>
                                <span id="HistoryList2" style="display: none; position: absolute" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td width="100" class="bluecol_need">職類/班別
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="asp_button_Mini" />
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                <br />
                                <span id="HistoryList" style="display: none; left: 270px; position: absolute">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">結訓日期
                            </td>
                            <td class="whitecol">
                                <span id="date" runat="server">
                                    <asp:TextBox ID="FTDate1" runat="server" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= FTDate1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                                    ～<asp:TextBox ID="FTDate2" runat="server" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= FTDate2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                                </span>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td class="whitecol" align="center">
                                <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
                </td>
            </tr>
            <tr>
                <td align="left">
                    <asp:Label Style="z-index: 0" ID="labMsg2" runat="server"></asp:Label>
                </td>
            </tr>
        </table>
        <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <HeaderStyle CssClass="head_navy" />
                        <Columns>
                            <asp:BoundColumn DataField="StudentID" HeaderText="學號">
                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Name" HeaderText="學員">
                                <HeaderStyle Width="10%"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="IDNO" HeaderText="身分證號碼">
                                <HeaderStyle Width="10%"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="薪資">
                                <HeaderStyle Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center" CssClass="whitecol" />
                                <HeaderTemplate>
                                    <input onclick="select_all2(this)" type="checkbox" />薪資
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <asp:TextBox ID="TextBox2" runat="server" Width="60%" Text='<%# DataBinder.Eval(Container.DataItem,"InsureSalary")%>'>
                                    </asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn>
                                <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <HeaderTemplate>
                                    <input type="checkbox" onclick="select_all(1, this.checked)">
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <input id="Checkbox1" type="checkbox" runat="server">
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="加保日">
                                <HeaderStyle Width="15%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center" CssClass="whitecol"></ItemStyle>
                                <ItemTemplate>
                                    <asp:TextBox ID="start_date" runat="server" Width="70%" Text='<%# DataBinder.Eval(Container.DataItem,"ApplyInsurance")%>'>
                                    </asp:TextBox>
                                    <img id="IMG1" style="cursor: pointer" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30">
                                    <input id="Hidden1" type="hidden" runat="server">
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn>
                                <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <HeaderTemplate>
                                    <input type="checkbox" onclick="select_all(2, this.checked)">
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <input id="Checkbox2" type="checkbox" runat="server">
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="退保日">
                                <HeaderStyle Width="15%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center" CssClass="whitecol"></ItemStyle>
                                <ItemTemplate>
                                    <asp:TextBox ID="end_date" runat="server" Width="70%" Text='<%# DataBinder.Eval(Container.DataItem,"DropoutInsurance")%>'>
                                    </asp:TextBox>
                                    <img id="IMG2" style="cursor: pointer" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30">
                                    <input id="Hidden2" type="hidden" runat="server">
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="原因">
                                <HeaderStyle Width="20%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center" CssClass="whitecol"></ItemStyle>
                                <ItemTemplate>
                                    <asp:TextBox ID="TextBox4" runat="server" Text='<%# DataBinder.Eval(Container.DataItem,"AppliedReason")%>' TextMode="MultiLine" Columns="20">
                                    </asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn Visible="False" DataField="StudStatus" HeaderText="StudStatus"></asp:BoundColumn>
                            <asp:BoundColumn Visible="False" DataField="SOCID" HeaderText="SOCID"></asp:BoundColumn>
                        </Columns>
                    </asp:DataGrid>
                </td>
            </tr>
            <tr>
                <td align="center" class="whitecol">
                    <asp:Button ID="Button2" runat="server" Text="存檔" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
