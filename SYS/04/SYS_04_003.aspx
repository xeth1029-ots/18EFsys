<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_04_003.aspx.vb" Inherits="WDAIIP.SYS_04_003" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>參數設定</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        function checkTextLength(obj, length) { //限定textbox的欄位長度                                            
            if (obj.value.length > length) {
                obj.value = obj.value.substring(0, length);
                alert("限欄位長度不能大於" + length + "個字元(含空白字元)，超出字元將自動截斷");
            }
        }
        function ctrlrblDD4() {
            var rblDD4 = document.getElementById('rblDD4');
            var vRblDD4 = getRadioValue(document.form1.rblDD4);
            //var D1a = document.getElementById('D1a');
            //var D2a = document.getElementById('D2a');
            var spanD1x = document.getElementById('spanD1x');
            var spanD2x = document.getElementById('spanD2x');
            var labD1y = document.getElementById('labD1y');
            var labD2y = document.getElementById('labD2y');
            //var D1b = document.getElementById('D1b');
            //var D2b = document.getElementById('D2b');
            if (rblDD4) {
                labD1y.style.display = 'none';
                labD2y.style.display = 'none';
                spanD1x.style.display = '';
                spanD2x.style.display = '';
                if (vRblDD4 == '2') {
                    labD1y.style.display = '';
                    labD2y.style.display = '';
                    spanD1x.style.display = 'none';
                    spanD2x.style.display = 'none';
                }
            }
        }

        function chgOpen(obj) {
            var B1 = document.getElementById('B1');
            var B2 = document.getElementById('B2');

            if (obj.checked) {
                B1.value = '';
                B2.value = '';

                B1.disabled = true;
                B2.disabled = true;

            } else {
                B1.disabled = false;
                B2.disabled = false;
            }
        }

        function chkdata() {
            var msg = ''
            var Hid_DistID = document.getElementById('Hid_DistID');
            var rblDD4 = document.getElementById('rblDD4');
            var vRblDD4 = getRadioValue(document.form1.rblDD4);
            var D1a = document.getElementById('D1a');
            var D1b = document.getElementById('D1b');
            var D2a = document.getElementById('D2a');
            var D2b = document.getElementById('D2b');

            if (document.form1.TPlan.selectedIndex == 0) {
                msg += '請先選擇訓練計畫\n';
                alert(msg);
                return false;
            }
            if (Hid_DistID.value == '000') {
                return true;
            }

            /*if (document.form1.Checkbox1.checked){
			if (document.form1.A1.value=='') msg+='請輸入行政管理費\n';
			else if(!isUnsignedInt(document.form1.A1.value)) msg+='行政管理費必須為數字\n';
			if (document.form1.A1.value>100) msg+='行政管理費百分比不能超過100\n';
			if (document.form1.A1.value<0) msg+='行政管理費百分比不能小於0\n';
			}*/

            if (document.form1.TNum.value != '' && !isUnsignedInt(document.form1.TNum.value)) msg += '訓練人數必須為數字\n';
            if (document.form1.Thours1.value != '' && !isUnsignedInt(document.form1.Thours1.value)) msg += '非學分班訓練時數上限必須為數字\n';
            if (document.form1.Thours2.value != '' && !isUnsignedInt(document.form1.Thours2.value)) msg += '非學分班訓練時數下限必須為數字\n';
            if (document.form1.B1.value != '' && !isUnsignedInt(document.form1.B1.value)) msg += '筆試成績計算比例必須為數字\n';
            if (document.form1.B1.value > 100) msg += '筆試成績百分比不能超過100\n';
            if (document.form1.B1.value < 0) msg += '筆試成績百分比不能小於0\n';
            if (document.form1.B2.value != '' && !isUnsignedInt(document.form1.B2.value)) msg += '口試成績計算比例必須為數字\n';
            if (document.form1.B2.value > 100) msg += '口試成績百分比不能超過100\n';
            if (document.form1.B2.value < 0) msg += '口試成績百分比不能小於0\n';
            if (document.form1.B1.value != '' || document.form1.B2.value != '') {
                if (parseInt(document.form1.B1.value, 10) + parseInt(document.form1.B2.value, 10) != 100 && parseInt(document.form1.B1.value, 10) + parseInt(document.form1.B2.value, 10) != 0) msg += '筆試成績與口試成績總合必須為100%\n';
            }
            if (document.form1.C1.value == '') msg += '請輸入操行成績底分\n';
            else if (!isUnsignedInt(document.form1.C1.value)) msg += '操行成績底分必須為數字(正整數)\n';
            if (document.form1.C1.value > 100) msg += '操行成績底分不能超過100\n';
            if (document.form1.C1.value < 0) msg += '操行成績底分不能小於0\n';

            if (vRblDD4 == '1') {
                if (D1a.value != '' || D1b.value != '') {
                    if (D1a.value == '') msg += '請輸入出缺勤警示1(分子)\n';
                    else if (!isUnsignedInt(D1a.value)) msg += '出缺勤警示1(分子)必須為數字(正整數)\n';
                    if (D1b.value == '') msg += '請輸入出缺勤警示1(分母)\n';
                    else if (!isUnsignedInt(D1b.value)) msg += '出缺勤警示1(分母)必須為數字(正整數)\n';
                    if (parseInt(D1a.value, 10) > parseInt(D1b.value, 10)) msg += '出缺勤警示1分子不能超過分母\n';
                    if (D1b.value < 0) msg += '出缺勤警示1分母不能為0\n';
                }
                if (msg == '' && (D2a.value != '' || D2b.value != '')) {
                    if (D2a.value == '') msg += '請輸入出缺勤警示2(分子)\n';
                    else if (!isUnsignedInt(D2a.value)) msg += '出缺勤警示2(分子)必須為數字(正整數)\n';
                    if (D2b.value == '') msg += '請輸入出缺勤警示2(分母)\n';
                    else if (!isUnsignedInt(D2b.value)) msg += '出缺勤警示2(分母)必須為數字(正整數)\n';
                    if (parseInt(D2a.value, 10) > parseInt(D2b.value, 10)) msg += '出缺勤警示2分子不能超過分母\n';
                    if (D2b.value < 0) msg += '出缺勤警示2分母不能為0\n';

                    if (msg == '' && (D1a.value != '' && D1b.value != '' && D2a.value != '' && D2b.value != '')) {
                        if ((parseInt(D1a.value, 10) / parseInt(D1b.value, 10)) > (parseInt(D2a.value, 10) / parseInt(D2b.value, 10))) msg += '出缺勤警示1不能大於出缺勤警示2\n';
                    }
                }
            }

            if (vRblDD4 == '2') {
                if (D1a.value != '') {
                    if (D1a.value == '') msg += '請輸入出缺勤警示1(分子)\n';
                    else if (!isUnsignedInt(D1a.value)) msg += '出缺勤警示1(分子)必須為數字(正整數)\n';
                    if (parseInt(D1a.value, 10) > 100) msg += '出缺勤警示1分子不能超過100\n';
                }
                if (msg == '' && D2a.value != '') {
                    if (D2a.value == '') msg += '請輸入出缺勤警示2(分子)\n';
                    else if (!isUnsignedInt(D2a.value)) msg += '出缺勤警示2(分子)必須為數字(正整數)\n';
                    if (parseInt(D2a.value, 10) > 100) msg += '出缺勤警示2分子不能超過100\n';
                    if (msg == '' && (D1a.value != '' && D2a.value != '')) {
                        if ((parseInt(D1a.value, 10) / 100) > (parseInt(D2a.value, 10) / 100)) msg += '出缺勤警示1不能大於出缺勤警示2\n';
                    }
                }
            }

            if (document.form1.E1.value == '') msg += '請輸入在訓證明字號\n';
            if (document.form1.J1.value == '') msg += '請輸入受訓證明字號\n';
            if (document.form1.K1.value == '') msg += '請輸入結訓證書字號\n';
            if (document.form1.L1.value == '') msg += '請輸入獎狀字號\n';
            if (document.form1.F1.value == '') msg += '請輸入甄試通知單內容\n';
            if (document.form1.G1.value == '') msg += '請輸入甄試結果通知單內容\n';
            if (document.form1.H1.value == '') msg += '請輸入代扣所得稅\n';
            //if (document.form1.I1){
            //if (document.form1.I1.value=='') msg+='請輸入核定%數[一般身分別]\n';
            //else if(!isUnsignedInt(document.form1.I1.value)) msg+='核定%數[一般對象]必須為數字\n';
            //if (document.form1.I2.value=='') msg+='請輸入核定%數[特定身分別]\n';
            //else if(!isUnsignedInt(document.form1.I2.value)) msg+='核定%數[特定對象]必須為數字\n';
            //}
            if (!isChecked(document.form1.M1)) msg += '請選擇成績計算方式\n';

            if (document.form1.N1.value != '' && !isUnsignedInt(document.form1.N1.value)) msg += '學科成績計算比例必須為整數數字\n';
            if (document.form1.N1.value > 100) msg += '學科成績百分比不能超過100\n';
            if (document.form1.N1.value < 0) msg += '學科成績百分比不能小於0\n';
            if (document.form1.N2.value != '' && !isUnsignedInt(document.form1.N2.value)) msg += '術科成績計算比例必須為整數數字\n';
            if (document.form1.N2.value > 100) msg += '術科成績百分比不能超過100\n';
            if (document.form1.N2.value < 0) msg += '術科成績百分比不能小於0\n';
            if (document.form1.N1.value != '' || document.form1.N2.value != '') {
                if (parseInt(document.form1.N1.value, 10) + parseInt(document.form1.N2.value, 10) != 100 && parseInt(document.form1.N1.value, 10) + parseInt(document.form1.N2.value, 10) != 0) msg += '學科成績與術科成績總合必須為100%\n';
            }

            if (msg != '') {
                alert(msg);
                return false;
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;系統參數管理&gt;&gt;參數設定</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="Table1" cellspacing="1" width="100%" border="0">
            <tr>
                <td>
                    <%--<table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>首頁&gt;&gt;系統管理&gt;&gt;系統參數管理&gt;&gt;<font color="#990000">參數設定</font> </td>
					</tr>
				</table>--%>
                    <table class="table_sch" id="Table3" cellpadding="1" cellspacing="1">
                        <tr>
                            <td class="bluecol" style="width: 20%">轄區分署 </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="ddlDISTID" runat="server" AutoPostBack="True">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" style="width: 20%">訓練計畫 </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="TPlan" runat="server" AutoPostBack="True">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">開放成績計算比例 </td>
                            <td class="whitecol">
                                <asp:CheckBox ID="chkOpen" Text="開放單位設定" runat="server"></asp:CheckBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">成績計算比例 </td>
                            <td class="whitecol">筆試
							<asp:TextBox ID="B1" runat="server" Width="8%" MaxLength="10"></asp:TextBox>%<br>
                                口試
							<asp:TextBox ID="B2" runat="server" Width="8%" MaxLength="10"></asp:TextBox>% </td>
                        </tr>
                        <tr>
                            <td class="bluecol">操行成績底分 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="C1" runat="server" Width="8%" MaxLength="10"></asp:TextBox>分 </td>
                        </tr>
                        <tr>
                            <td class="bluecol">出缺勤警示 </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="rblDD4" runat="server" CssClass="font12" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="1" Selected="True">依比例</asp:ListItem>
                                    <asp:ListItem Value="2">依百分比</asp:ListItem>
                                </asp:RadioButtonList>
                                <br />
                                1.第一次缺課警告：出缺勤總時數超過訓練時數的
							<asp:TextBox ID="D1a" runat="server" Width="8%" MaxLength="4"></asp:TextBox>
                                <span id="spanD1x" runat="server">/<asp:TextBox ID="D1b" runat="server" Width="8%" MaxLength="4"></asp:TextBox>&nbsp;&nbsp;比例(例:1/15)</span>
                                <asp:Label ID="labD1y" runat="server">%</asp:Label><br>
                                2.第二次缺課警告：出缺勤總時數超過訓練時數的
							<asp:TextBox ID="D2a" runat="server" Width="8%" MaxLength="4"></asp:TextBox>
                                <span id="spanD2x" runat="server">/<asp:TextBox ID="D2b" runat="server" Width="8%" MaxLength="4"></asp:TextBox>&nbsp;&nbsp;比例(例:1/5)</span>
                                <asp:Label ID="labD2y" runat="server">%</asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">在訓證明字號 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="E1" runat="server" Width="50%" MaxLength="250"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">受訓證明字號 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="J1" runat="server" Width="50%" MaxLength="250"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">結訓證書字號 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="K1" runat="server" Width="50%" MaxLength="250"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">獎狀字號 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="L1" runat="server" Width="50%" MaxLength="250"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">操行字號 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="MA1" runat="server" Width="50%" MaxLength="250"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">全勤字號 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="PA1" runat="server" Width="50%" MaxLength="250"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">安全衛生教育字號 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="SA1" runat="server" Width="50%" MaxLength="250"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                <p>甄試通知單</p>
                                <p>內容</p>
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="F1" onblur="checkTextLength(this,250)" onkeyup="checkTextLength(this,250);" runat="server" Width="50%" TextMode="MultiLine" Height="80px" onchange="checkTextLength(this,250)"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">甄試結果<br />
                                通知單內容 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="G1" onblur="checkTextLength(this,250)" onkeyup="checkTextLength(this,250);" runat="server" Width="50%" TextMode="MultiLine" Height="80px" onchange="checkTextLength(this,250)"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">操行獎狀內容(中文) </td>
                            <td class="whitecol">
                                <asp:TextBox ID="MVC" onblur="checkTextLength(this,250)" onkeyup="checkTextLength(this,250);" runat="server" Width="50%" TextMode="MultiLine" Height="80px" onchange="checkTextLength(this,250)"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">操行獎狀內容(英文) </td>
                            <td class="whitecol">
                                <asp:TextBox ID="MVE" onblur="checkTextLength(this,250)" onkeyup="checkTextLength(this,250);" runat="server" Width="50%" TextMode="MultiLine" Height="80px" onchange="checkTextLength(this,250)"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">全勤獎狀內容(中文) </td>
                            <td class="whitecol">
                                <asp:TextBox ID="PVC" onblur="checkTextLength(this,250)" onkeyup="checkTextLength(this,250);" runat="server" Width="50%" TextMode="MultiLine" Height="80px" onchange="checkTextLength(this,250)"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">全勤獎狀內容(英文) </td>
                            <td class="whitecol">
                                <asp:TextBox ID="PVE" onblur="checkTextLength(this,250)" onkeyup="checkTextLength(this,250);" runat="server" Width="50%" TextMode="MultiLine" Height="80px" onchange="checkTextLength(this,250)"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">安全衛生獎狀內容(中文) </td>
                            <td class="whitecol">
                                <asp:TextBox ID="SVC" onblur="checkTextLength(this,250)" onkeyup="checkTextLength(this,250);" runat="server" Width="50%" TextMode="MultiLine" Height="80px" onchange="checkTextLength(this,250)"></asp:TextBox><br>
                                <font color="red">如果名次的話,請在輸入資料時一併輸入</font> </td>
                        </tr>
                        <tr>
                            <td class="bluecol">安全衛生獎狀內容(英文) </td>
                            <td class="whitecol">
                                <asp:TextBox ID="SVE" onblur="checkTextLength(this,250)" onkeyup="checkTextLength(this,250);" runat="server" Width="50%" TextMode="MultiLine" Height="80px" onchange="checkTextLength(this,250)"></asp:TextBox><br>
                                <font color="red">如果名次的話,請在輸入資料時一併輸入</font> </td>
                        </tr>
                        <tr>
                            <td class="bluecol">代扣所得稅 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="H1" runat="server" Width="8%" MaxLength="10"></asp:TextBox>%</td>
                        </tr>
                        <tr id="TPlan23" runat="server">
                            <td class="bluecol">核銷%數 </td>
                            <td class="whitecol">&nbsp;機構別&nbsp;&nbsp;&nbsp;
							<asp:DropDownList ID="OrgType" runat="server">
                            </asp:DropDownList>
                                <br>
                                &nbsp;一般身分<asp:TextBox ID="I1" runat="server" Width="8%" MaxLength="10"></asp:TextBox>%
							<br>
                                &nbsp;特定對象<asp:TextBox ID="I2" runat="server" Width="8%" MaxLength="10"></asp:TextBox>%
							<asp:Button ID="AddBtn" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="BtnSave" runat="server" Text="存檔" CssClass="asp_button_M"></asp:Button>
                                <br>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="WhiteSmoke" />
                                    <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn DataField="Name" HeaderText="機構別">
                                            <HeaderStyle Width="70%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ItemVar1B" HeaderText="一般身分">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ItemVar2B" HeaderText="特定對象">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:Button ID="BtnEdit" runat="server" Text="修改" CommandName="Edit"></asp:Button>
                                                <asp:Button ID="BtnDel" runat="server" Text="刪除" CommandName="Del"></asp:Button>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle BackColor="#FFFFFF"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">計算成績方式 </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="M1" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="1">各科平均法</asp:ListItem>
                                    <asp:ListItem Value="2">訓練時數權重法</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">學、術科百分比 </td>
                            <td class="whitecol">學科
							<asp:TextBox ID="N1" runat="server" Width="8%" MaxLength="10"></asp:TextBox>%<br>
                                術科
							<asp:TextBox ID="N2" runat="server" Width="8%" MaxLength="10"></asp:TextBox>% </td>
                        </tr>
                        <tr id="Tplan28_1" runat="server">
                            <td class="bluecol">訓練人數設定 </td>
                            <td class="whitecol">訓練人數<asp:TextBox ID="TNum" runat="server" Width="8%" MaxLength="10"></asp:TextBox>
                            </td>
                        </tr>
                        <tr id="Tplan28_2" runat="server">
                            <td class="bluecol">時數設定 </td>
                            <td class="whitecol">非學分班訓練時數上限<asp:TextBox ID="Thours1" runat="server" Width="8%" MaxLength="10"></asp:TextBox><br>
                                非學分班訓練時數下限<asp:TextBox ID="Thours2" runat="server" Width="8%" MaxLength="10"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">准考證號碼 </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="rdolist21" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="Y">列印</asp:ListItem>
                                    <asp:ListItem Value="N">不列印</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                <p>
                                    在職進修學員資料維護&nbsp;&nbsp;<br>
                                    取消必填
                                </p>
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="checkboxList22a" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatLayout="Flow" RepeatColumns="4">
                                </asp:CheckBoxList>
                                <asp:CheckBoxList ID="checkboxList22b" runat="server" RepeatDirection="Horizontal" CssClass="font" Enabled="False" Visible="False" RepeatLayout="Flow" RepeatColumns="4">
                                </asp:CheckBoxList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">是否可改備取名次設定 </td>
                            <td class="whitecol">
                                <asp:CheckBox ID="GVID24" runat="server"></asp:CheckBox>錄取作業是否可改備取名次&nbsp; </td>
                        </tr>
                    </table>
                    <p align="center" class="whitecol">
                        <asp:Button ID="Button1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                    </p>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="Hid_DistID" runat="server" />
    </form>
</body>
</html>
