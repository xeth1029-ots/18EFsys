<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_03_001.aspx.vb" Inherits="WDAIIP.SD_03_001" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>學員參訓</title>
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
    </script>
    <script type="text/javascript" language="javascript">
        function GETvalue() { document.getElementById('Button3').click(); }

        function SetOneOCID() { document.getElementById('Button4').click(); }

        function choose_class() {
            var Button4 = document.getElementById('Button4');
            var OCID1 = document.getElementById('OCID1');
            var RIDValue = document.getElementById('RIDValue');
            if (OCID1.value == '') { Button4.click(); }
            openClass('../02/SD_02_ch.aspx?RID=' + RIDValue.value);
        }

        //全選的checkbox
        //要是有更改過畫面上的物件順序則要把elements中的陣列位置重新排列
        //預設從第8個到elements-1個為DataGrid中的checkbox
        function chall(Choose1) {
            //var Choose1 = document.getElementById('Choose1');
            var mytable = document.getElementById('DataGrid1');
            /*debugger;*/
            for (var i = 1; i < mytable.rows.length; i++) {
                var mycheck = mytable.rows[i].cells[0].children[0];
                if (mycheck.disabled == false) { mycheck.checked = Choose1.checked; }
            }
        }

        function search() {
            var OCIDValue1 = document.getElementById('OCIDValue1');
            if (OCIDValue1.value == '') {
                alert('請選擇班別職類');
                return false;
            }
        }

        //必須填寫「放棄報到(原因)」原因
        function chkABANDONReason(textReason, myChk1) {
            var oMyChk1 = document.getElementById(myChk1);
            if (oMyChk1.checked == true) {
                alert('放棄報到 不可勾選已報到!');
                return false;
            }
            var oHid_ABANDONReasonSub = document.getElementById("Hid_ABANDONReasonSub");
            var oTextReason = document.getElementById(textReason);
            oHid_ABANDONReasonSub.value = oTextReason.value;
            if (oTextReason.value == "") {
                alert('必須填寫「放棄報到(原因)」原因!');
                return false;
            }
            return confirm('為維護其他學員參訓權益，已放棄報到者，將無法再報名此班,請確認!');
        }

        function open_History(EncIdnoVal, rqID) {
            //window.open('../01/SD_01_001_old.aspx?IDNO=' + idnoValue, 'history', 'width=700,height=500,scrollbars=1')
            //window.open('../05/SD_05_010.aspx?SD_01_004_Type=Student&IDNO=' + idnoValue, 'history', 'width=900,height=700,scrollbars=1');
            window.open('../05/SD_05_010_pop.aspx?ID=' + rqID + '&SD_01_004_Type=Student&ENCIDNO=' + EncIdnoVal, 'history', 'width=1400,height=820,scrollbars=1');
            return false;
        }

        function Button2_Send() {
            var hTPlanID2854 = document.getElementById('hTPlanID2854');
            var Button2 = document.getElementById('Button2');
            var Button2B = document.getElementById('Button2B');

            var HidLID = document.getElementById('HidLID');
            var txtEnterDate = document.getElementById('txtEnterDate');
            var hSTDate = document.getElementById('hSTDate');
            var hSTDate14 = document.getElementById('hSTDate14');
            var hdatenow = document.getElementById('hdatenow');
            //var txtToday = document.getElementById('txtToday');
            var HidToday = document.getElementById('HidToday');
            var Hidtestflag = document.getElementById('Hidtestflag');
            var Hid_ignoreflag = document.getElementById('Hid_ignoreflag');
            var msg = '';

            if (msg == '' && hSTDate.value == '') { msg += '資料有誤，請重新查詢!!\n'; }
            if (msg == '' && hSTDate14.value == '') { msg += '資料有誤，請重新查詢!!\n'; }
            if (msg != '') {
                alert(msg);
                return false;
            }

            //debugger;
            if (txtEnterDate != null) {
                if (txtEnterDate.value == '') msg += '請輸入報到日期\n';
                if (msg == '' && !checkDate(txtEnterDate.value)) { msg += '報到日期必須是正確的日期格式\n'; }
                if (txtEnterDate.value != '' && msg == '') {
                    //BASIC CHECK;
                    if (msg == '' && getDiffDay(txtEnterDate.value, hdatenow.value) < 0) { msg += '報到日期(只能是系統時間今天或今天以前)請確認\n'; }
                    if (msg == '' && hTPlanID2854.value == "1" && getDiffDay(hSTDate.value, txtEnterDate.value) < 0) { msg += '「報到日期」僅能選擇開訓日當天~開訓日後14日內,(為開訓日(含)之後14天內)請確認\n'; }
                    if (msg == '' && hTPlanID2854.value == "1" && getDiffDay(txtEnterDate.value, hSTDate14.value) < 0) { msg += '「報到日期」僅能選擇開訓日當天~開訓日後14日內,(為開訓日(含)之後14天內)請確認\n'; }
                }

                if (txtEnterDate.value != '' && msg == '' && HidLID.value != '0') {
                    //LOGIC CHECK; //增加可能錯誤的邏輯判斷
                    if (getDiffDay(txtEnterDate.value, hdatenow.value) > 30) {
                        if (Hid_ignoreflag.value == "") { msg += '(非署)報到日期(已超過系統時間30天)請確認\n'; }
                        //msg += txtEnterDate.value + '!\n'; //msg += hdatenow.value + '!\n';
                    }
                }
            }

            //debugger;
            var Cst_DataGrid1_Hstar3 = 1;
            var Cst_DataGrid1_Hstar4 = 2;
            var hidstar3 = document.getElementById('hidstar3');
            var hidstar4 = document.getElementById('hidstar4');
            var mytable = document.getElementById('DataGrid1');
            if (mytable == null) {
                msg += '請重新查詢有效資料。!!!\n';
                alert(msg);
                return false;
            }
            for (var i = 1; i < mytable.rows.length; i++) {
                var Cells0 = mytable.rows[i].cells[0];
                //核取且可使用
                if (Cells0.children[0].checked && !Cells0.children[0].disabled) {
                    if (Cells0.children[Cst_DataGrid1_Hstar3] != null) {
                        if (Cells0.children[Cst_DataGrid1_Hstar3].value != '') {
                            hidstar3.value = Cells0.children[Cst_DataGrid1_Hstar3].value;
                            break;
                        }
                    }
                    if (Cells0.children[Cst_DataGrid1_Hstar4] != null) {
                        if (Cells0.children[Cst_DataGrid1_Hstar4].value != '') {
                            hidstar4.value = Cells0.children[Cst_DataGrid1_Hstar4].value;
                            break;
                        }
                    }
                }
            }
            if (hidstar3.value != '') {
                if (!confirm('本次參訓之學員,仍有學員在訓中,是否儲存,請確認!')) { msg += '尚有學員,仍在訓中\n' };
            }
            if (hidstar4.value != '') {
                //flag_TrainITS
                msg += '因有學員同時參加職前課程，無法完成報到，請重新確認!。!!!\n';
                alert(msg);
                return false;
                //if (!confirm('因有學員同時參加職前課程，無法完成報到，請確認!')) { msg += '尚有學員,參加職前課程\n' };
            }

            //debugger;
            //return false;//測試環境，清空MSG 
            //if (msg != '' && Hidtestflag.value == "Y") { msg = ''; }

            if (msg != '') {
                Button2.style.display = ''; //'inline';
                msg += '!!!\n';
                alert(msg);
                //txtEnterDate.value = HidToday.value;
                //alert('false');
                return false;
            }
            else {
                Button2.style.display = 'none';
                //document.getElementById('Button2').disabled = true;
                Button2B.click();
                return false;
            }
        }

        /*
		function Button5_Send(){
		document.getElementById('Button2').disabled=true;
		document.getElementById('Button5').disabled=true;
		document.getElementById('Button5B').click();
		}
		*/
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;學員資料管理&gt;&gt;學員參訓</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>

                    <table class="table_nw" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td width="20%" class="bluecol">訓練機構 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                <input id="Button8" type="button" value="..." name="Button8" runat="server">
                                <asp:Button ID="Button4" Style="display: none" runat="server"></asp:Button>
                                <asp:Button ID="Button3" Style="display: none" runat="server"></asp:Button>
                                <%-- <asp:Button ID="Button4" runat="server"></asp:Button> <asp:Button ID="Button3" runat="server"></asp:Button>--%>
                                <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td width="20%" class="bluecol_need">職類/班別 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="...">
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                <span id="HistoryList" style="position: absolute; display: none; left: 28%">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td width="20%" class="bluecol">顯示結果 </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="SelResultID" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="99" Selected="True">不區分</asp:ListItem>
                                    <asp:ListItem Value="01">僅顯示正取</asp:ListItem>
                                    <asp:ListItem Value="02">僅顯示備取</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td width="20%" class="bluecol">檢核</td>
                            <td class="whitecol">
                                <asp:CheckBox ID="CheckBoxITS1" runat="server" Text="是否同時參加職前課程 (僅未完成報到的學員) " Checked="True" /></td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td class="whitecol" align="center">
                                <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button></td>
                        </tr>
                    </table>
                    <table class="font" id="DataGridTable1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td class="whitecol">報到日期：<asp:TextBox ID="txtEnterDate" Width="15%" runat="server" MaxLength="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= txtEnterDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top"></span>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:Label ID="Label2" runat="server" ForeColor="Green" CssClass="font">* 表該學員有報名其他班級仍在訓中,請查詢學員參訓歷史</asp:Label><br />
                                <asp:Label ID="LabMsg3" runat="server" ForeColor="Red" CssClass="font">* 表該學員有同時參加職前課程</asp:Label><br />
                                <asp:Label ID="Labmsg6" runat="server" ForeColor="Red" CssClass="font">* 為維護其他學員參訓權益，已放棄報到者，將無法再報名此班!</asp:Label><br />
                                <asp:Label ID="labmsg219" runat="server" ForeColor="Blue" Visible="false"></asp:Label>
                                <asp:Label ID="Label1" runat="server" CssClass="font"></asp:Label><br />
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <HeaderTemplate>
                                                <input onclick="chall(this);" type="checkbox" name="Choose1">
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <input id="Checkbox1" type="checkbox" name="student" runat="server" />
                                                <input id="Hstar3" type="hidden" runat="server" />
                                                <input id="Hstar4" type="hidden" runat="server" />
                                                <asp:Label ID="star3" runat="server" CssClass="font" ForeColor="Green">*</asp:Label>
                                                <asp:Label ID="star4" runat="server" CssClass="font" ForeColor="Red">*</asp:Label>
                                                <input id="SETID" type="hidden" runat="server" />
                                                <input id="EnterDate" type="hidden" runat="server" />
                                                <input id="SerNum" type="hidden" runat="server" />
                                                <input id="HidCFIRE1" type="hidden" runat="server" />
                                                <input id="HidCFIRE1NS" type="hidden" runat="server" />
                                                <input id="HidCMASTER1" type="hidden" runat="server" />
                                                <input id="HidCMASTER1NS" type="hidden" runat="server" />
                                                <input id="HidCMASTER1NT" type="hidden" runat="server" />
                                                <input id="HidIsStdBlack" type="hidden" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="ExamNo" HeaderText="學號">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="姓名">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:Label ID="Label3name" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="WriteResult" HeaderText="筆試成績">
                                            <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="OralResult" HeaderText="口試成績">
                                            <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TotalResult" HeaderText="總成績">
                                            <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="RelEnterDate" HeaderText="報名日期" DataFormatString="{0:d}">
                                            <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="政府已補助經費">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="LabGovCost" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn HeaderText="名次">
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TRNDType" HeaderText="卷別">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="SelResultName" HeaderText="錄訓結果">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="放棄報到(原因)" ItemStyle-HorizontalAlign="Center">
                                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="middle" Width="15%"></HeaderStyle>
                                            <ItemTemplate>
                                                <div>
                                                    <asp:TextBox ID="ABANDONReason" runat="server" MaxLength="150" TextMode="MultiLine" Width="95%" Height="60px"></asp:TextBox><br />
                                                    <asp:Button ID="BtnABA" runat="server" Text="放棄報到" CommandName="ABA" CssClass="asp_button_M" />
                                                    <%--(Restore) Give up the registration--%>
                                                    <asp:Button ID="BtnRESTO" runat="server" Text="(還原)放棄報到" CommandName="RESTO" CssClass="asp_button_M" />
                                                    <input id="HidABANDON" type="hidden" runat="server" />
                                                </div>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol">
                                <div align="center" class="whitecol">
                                    <%--<asp:Button ID="BtnCheckITS1" runat="server" Text="檢核是否同時參加職前課程" CssClass="asp_button_M"></asp:Button>--%>
                                    <asp:Button ID="Button2" runat="server" Text="完成報到" CssClass="asp_button_M"></asp:Button>
                                    <%--<asp:button id="Button5" runat="server" Text="重整學號" Visible="False"></asp:button>--%>
                                </div>
                            </td>
                        </tr>
                    </table>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label><br />
                        <br />
                        <asp:Label ID="Labmsg4" runat="server" ForeColor="Red" CssClass="font">注意：「檢核是否同時參加職前課程」僅針對未完成報到的學員，且因批次勾稽職前系統可能會較為耗時，請耐心等候。</asp:Label><br />
                    </div>
                </td>
            </tr>
            <tr>
                <td class="whitecol" align="center">
                    <asp:Button ID="Button2B" Style="position: absolute; display: none" runat="server" Text="完成報到"></asp:Button>
                </td>
            </tr>
        </table>
        <input id="hTPlanID2854" type="hidden" runat="server" />
        <input id="hdatenow" type="hidden" runat="server" />
        <input id="HidToday" type="hidden" runat="server" />
        <input id="hSTDate" type="hidden" runat="server" />
        <input id="hSTDate14" type="hidden" runat="server" />
        <input id="HidLID" type="hidden" runat="server" />
        <input id="hidstar3" type="hidden" runat="server" />
        <input id="hidstar4" type="hidden" runat="server" />
        <input id="isBlack" type="hidden" runat="server" />
        <input id="Blackorgname" type="hidden" runat="server" />
        <input id="HidOCID1" type="hidden" runat="server" />
        <input id="Hid_ABANDONReasonSub" type="hidden" runat="server" />
        <asp:HiddenField ID="HidTNum" runat="server" />
        <%--Hidtestflag--%>
        <asp:HiddenField ID="Hidtestflag" runat="server" />
        <asp:HiddenField ID="Hid_ignoreflag" runat="server" />
        <asp:HiddenField ID="HidEnterDate" runat="server" />
        <asp:HiddenField ID="Hid_CAN_IGNORE_RULE1_CNT" runat="server" />
    </form>
</body>
</html>
