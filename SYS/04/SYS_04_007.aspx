<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_04_007.aspx.vb" Inherits="WDAIIP.SYS_04_007" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>年度訪視率設定</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript">

        function set_lab(rows) {
            //debugger;
            var i = rows;
            var MyTable = document.getElementById('DataGrid1');
            var radio1 = MyTable.rows(i).cells(1).children(0).childNodes(0);
            var radio3 = MyTable.rows(i).cells(1).children(2).childNodes(0);
            var radio2 = MyTable.rows(i).cells(1).children(4).childNodes(0);
            var radio4 = MyTable.rows(i).cells(1).children(6).childNodes(0);
            var lab = MyTable.rows(i).cells(2).children(1);
            //var radiolength=MyTable.rows(i).cells(2).children.length;
            if (radio1.checked) {
                //alert('lab:'+i+':'+radiolength);
                lab.style.display = 'none';
            }
            else {
                //alert('lab2:'+i+':'+radiolength);
                lab.style.display = 'inline';
            }
        }

        function set_lab2(rows) {
            //debugger;
            var i = rows;
            var MyTable = document.getElementById('DataGrid1');
            var radio1 = MyTable.rows(i).cells(1).children(0).childNodes(0);
            var radio3 = MyTable.rows(i).cells(1).children(2).childNodes(0);
            var radio2 = MyTable.rows(i).cells(1).children(4).childNodes(0);
            var radio4 = MyTable.rows(i).cells(1).children(6).childNodes(0);
            var Lradio1 = MyTable.rows(i).cells(1).children(0);
            var Lradio3 = MyTable.rows(i).cells(1).children(2);
            var Lradio2 = MyTable.rows(i).cells(1).children(4);
            var Lradio4 = MyTable.rows(i).cells(1).children(6);
            //var lab = MyTable.rows(i).cells(2).children(1);
            var value1 = MyTable.rows(i).cells(1).children(8);

            if (value1.value == 1) {
                Lradio1.style.display = 'none';
                Lradio2.style.display = 'none';
                Lradio3.style.display = 'none';
                Lradio4.style.display = 'none';
                if (radio1.checked) {
                    Lradio1.style.display = 'inline';
                }
                if (radio3.checked) {
                    Lradio3.style.display = 'inline';
                }
                if (radio2.checked) {
                    Lradio2.style.display = 'inline';
                }
                if (radio4.checked) {
                    Lradio4.style.display = 'inline';
                }
                value1.value = '';
            }
            else {
                if (value1.value == '' && ((Lradio1.style.display == 'none') || (Lradio2.style.display == 'none') || (Lradio3.style.display == 'none') || (Lradio4.style.display == 'none'))) {
                    Lradio1.style.display = 'inline';
                    Lradio3.style.display = 'inline';
                    Lradio2.style.display = 'inline';
                    Lradio4.style.display = 'inline';
                    value1.value = 1;
                }
            }
        }

        function Check_Data() {
            //debugger; 
            var msg = '';
            var MyTable = document.getElementById('DataGrid1');
            for (var i = 1; i < MyTable.rows.length; i++) {
                //var radiolength=MyTable.rows(i).cells(1).children.length;
                var radio1 = MyTable.rows(i).cells(1).children(0).childNodes(0);
                var radio3 = MyTable.rows(i).cells(1).children(2).childNodes(0);
                var radio2 = MyTable.rows(i).cells(1).children(4).childNodes(0);
                var radio4 = MyTable.rows(i).cells(1).children(6).childNodes(0);
                var text1 = MyTable.rows(i).cells(2).children(0);
                var text2 = MyTable.rows(i).cells(3).children(0);
                //alert('radiolength:'+radiolength+'radio1:'+radio1.checked+',radio2:'+radio2.checked+',radio3:'+radio3.checked);

                if (radio1.checked || radio2.checked || radio3.checked) {
                    //if (radio1.checked && text1.value == '') msg += '請填寫[' + MyTable.rows(i).cells(0).innerHTML + ']中心訪視比率值\n'
                    if (radio1.checked && text1.value == '') msg += '請填寫[' + MyTable.rows(i).cells(0).innerHTML + ']分署訪視比率值\n'
                    //if (radio3.checked && text1.value == '') msg += '請填寫[' + MyTable.rows(i).cells(0).innerHTML + ']中心訪視比率值\n'
                    if (radio3.checked && text1.value == '') msg += '請填寫[' + MyTable.rows(i).cells(0).innerHTML + ']分署訪視比率值\n'
                    //if (radio2.checked && text1.value == '') msg += '請填寫[' + MyTable.rows(i).cells(0).innerHTML + ']中心訪視比率值\n'
                    if (radio2.checked && text1.value == '') msg += '請填寫[' + MyTable.rows(i).cells(0).innerHTML + ']分署訪視比率值\n'
                }
                if (radio4.checked) {
                    //if (text1.value != '') msg += '[' + MyTable.rows(i).cells(0).innerHTML + '] 選項為 (尚未選擇) 請清空中心訪視比率值\n'
                    if (text1.value != '') msg += '[' + MyTable.rows(i).cells(0).innerHTML + '] 選項為 (尚未選擇) 請清空分署訪視比率值\n'
                    if (text2.value != '') msg += '[' + MyTable.rows(i).cells(0).innerHTML + '] 選項為 (尚未選擇) 請清空備註\n'
                }
            }

            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        //2009年訪視計畫用，儲存規則改變
        function set_lab2b(rows) {
            //debugger;
            var rowsi = rows;
            var Cst_cells1 = 1;  //受訪者 功能位置 
            var Cst_cells = 2;   //物件顯示或隱藏之存取 功能位置 value2b
            var Cst_cells3 = 3;  //分署(中心)訪視 功能位置 
            var Cst_rb14 = 30;   //尚未選擇 RadioButton 在HTML中的位置
            var MyTable = document.getElementById('DataGrid2');

            var value2 = MyTable.rows(rowsi).cells(Cst_cells).children(MyTable.rows(rowsi).cells(Cst_cells).children.length - 1);
            var value2ok = '';

            if (value2.value == 1) {
                var _Checkbox1 = MyTable.rows(rowsi).cells(Cst_cells1).children(0);
                var _Checkbox2 = MyTable.rows(rowsi).cells(Cst_cells1).children(2);
                var Lradio1a = MyTable.rows(rowsi).cells(Cst_cells).children(0);
                var Lradio1b = MyTable.rows(rowsi).cells(Cst_cells).children(10);
                var Lradio1ab = MyTable.rows(rowsi).cells(Cst_cells).children(1);
                var Lradio1bb = MyTable.rows(rowsi).cells(Cst_cells).children(11);

                for (i = 0; i < MyTable.rows(rowsi).cells(Cst_cells).children.length - 2; i += 2) {
                    var Lradio1 = MyTable.rows(rowsi).cells(Cst_cells).children(i);
                    var Lradio2 = MyTable.rows(rowsi).cells(Cst_cells).children(i + 1);
                    var radio1 = MyTable.rows(rowsi).cells(Cst_cells).children(i).childNodes(0);
                    Lradio1.style.display = 'none';
                    Lradio2.style.display = 'none';
                    if (radio1.checked) {
                        Lradio1.style.display = 'inline';
                        Lradio2.style.display = 'inline';
                        switch (i) {
                            case 2:
                            case 4:
                            case 6:
                            case 8:
                                //依訓練時數(補助、委外)
                                Lradio1a.style.display = 'inline';
                                Lradio1ab.style.display = 'inline';
                                break;
                            case 12:
                            case 14:
                            case 16:
                            case 18:
                                //依訓練時數(自辦、職前)
                                Lradio1b.style.display = 'inline';
                                Lradio1bb.style.display = 'inline';
                                break;
                            case Cst_rb14: //尚未選擇 清除其他參數
                                _Checkbox1.checked = false;
                                _Checkbox2.checked = false;
                                //debugger;
                                for (i2 = 0; i2 < MyTable.rows(rowsi).cells(Cst_cells3).children.length; i2 += 3) {
                                    var radio22 = MyTable.rows(rowsi).cells(Cst_cells3).children(i2 + 1);
                                    if (radio22.checked == true) {
                                        radio22.checked = false;
                                    }
                                }
                                break;
                        }

                    }
                }
                value2.value = '';
            }
            else {
                value2ok = '';
                for (i = 0; i < MyTable.rows(rowsi).cells(Cst_cells).children.length - 2; i += 2) {
                    var Lradio1 = MyTable.rows(rowsi).cells(Cst_cells).children(i);
                    //var radio1=MyTable.rows(rowsi).cells(Cst_cells).children(i).childNodes(0);
                    if (Lradio1.style.display == 'none') {
                        value2ok = '1';
                        break;
                    }
                }
                if (value2ok == '1') {
                    value2.value = value2ok;
                    for (i = 0; i < MyTable.rows(rowsi).cells(Cst_cells).children.length - 2; i += 2) {
                        var Lradio1 = MyTable.rows(rowsi).cells(Cst_cells).children(i);
                        var Lradio2 = MyTable.rows(rowsi).cells(Cst_cells).children(i + 1);
                        Lradio1.style.display = 'inline';
                        Lradio2.style.display = 'inline';
                    }
                }
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;系統參數管理&gt;&gt;年度訪視率設定</asp:Label>
                </td>
            </tr>
        </table>
    <table class="font" id="FrameTable1" cellspacing="1" cellpadding="1" width="100%" border="0">
        <tr>
            <td>
                <%--<table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            <font face="新細明體">首頁&gt;&gt;系統管理&gt;&gt;系統參數管理&gt;&gt;年度訪視率設定</font>
                        </td>
                    </tr>
                </table>--%>
                <table class="table_sch" id="Table2" cellpadding="1" cellspacing="1" width="100%">
                    <tr>
                        <td class="bluecol" style="width:20%">
                            年度
                        </td>
                        <td class="whitecol">
                            <asp:DropDownList ID="Syear" runat="server" AutoPostBack="True">
                            </asp:DropDownList>
                        </td>
                    </tr>
                </table>
                <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                    <tr class="font">
                        <td colspan="2">
                            <%--<font color="red">*2007年度前中心訪視率依次數,2008年度中心訪視率改依百分比<br>
                                2009年度後中心訪視率依次數選項、依百分比選項(若訪視方式選尚未選擇，此筆訓練計畫將不儲存)
                                <br>若須修改，請點選訪視方式中的選項，即會出現其他選項 </font>--%>
                            <font color="red">*2007年度前分署訪視率依次數,2008年度分署訪視率改依百分比<br>
                                2009年度後分署訪視率依次數選項、依百分比選項(若訪視方式選尚未選擇，此筆訓練計畫將不儲存)
                                <br>若須修改，請點選訪視方式中的選項，即會出現其他選項 </font>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:DataGrid ID="DataGrid1" runat="server" AutoGenerateColumns="False" Width="100%" CssClass="font" CellPadding="8">
                                <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                <Columns>
                                    <asp:BoundColumn DataField="PlanName" HeaderText="訓練計畫">
                                        <HeaderStyle Width="55%"></HeaderStyle>
                                    </asp:BoundColumn>
                                    <asp:TemplateColumn HeaderText="訪視方式">
                                        <HeaderStyle Width="15%"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="left"/>
                                        <ItemTemplate>
                                            <asp:RadioButton ID="RadioButton1" runat="server" Text="依開班數" GroupName="Mode2">
                                            </asp:RadioButton><br>
                                            <asp:RadioButton ID="RadioButton3" runat="server" Text="依機構數" GroupName="Mode2">
                                            </asp:RadioButton><br>
                                            <asp:RadioButton ID="RadioButton2" runat="server" Text="依機構開班數" GroupName="Mode2">
                                            </asp:RadioButton><br>
                                            <asp:RadioButton ID="Radiobutton4" runat="server" Text="　(尚未選擇)" GroupName="Mode2">
                                            </asp:RadioButton><br>
                                            <input id="value1" runat="server" style="width: 46px; height: 22px" type="hidden" size="2">
                                        </ItemTemplate>
                                    </asp:TemplateColumn>
                                    <%--<asp:TemplateColumn HeaderText="中心訪視比率">--%>
                                    <asp:TemplateColumn HeaderText="分署訪視比率">
                                        <HeaderStyle Width="15%"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Center"/>
                                        <ItemTemplate>
                                            <asp:TextBox ID="TextBox1" runat="server" Columns="5" Width="80%"></asp:TextBox>
                                            <asp:Label ID="Lab1" runat="server">%</asp:Label>
                                        </ItemTemplate>
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="備註">
                                        <HeaderStyle Width="15%"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Center"/>
                                        <ItemTemplate>
                                            <asp:TextBox ID="Note" runat="server" Width="100%" MaxLength="200" TextMode="MultiLine"
                                                Rows="3" Columns="15"></asp:TextBox>
                                        </ItemTemplate>
                                    </asp:TemplateColumn>
                                </Columns>
                            </asp:DataGrid>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:DataGrid ID="DataGrid2" runat="server" AutoGenerateColumns="False" Width="100%" CssClass="font" CellPadding="8">
                                <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                                <Columns>
                                    <asp:BoundColumn DataField="PlanName" HeaderText="訓練計畫">
                                        <ItemStyle Width="40%"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:TemplateColumn HeaderText="受訪者">
                                        <HeaderStyle Width="10%"></HeaderStyle>
                                        <ItemTemplate>
                                            <input id="Checkbox1" runat="server" type="checkbox" name="Checkbox1">培訓單位<br />
                                            <input id="Checkbox2" runat="server" type="checkbox" name="Checkbox2">參訓學員
                                        </ItemTemplate>
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="訪視方式">
                                        <HeaderStyle Width="15%"></HeaderStyle>
                                        <ItemTemplate>
                                            <div id="d2lab1" runat="server">
                                                依訓練時數(補助、委外)：</div>
                                            <br>
                                            <asp:RadioButton ID="rb1" runat="server" Text="　180小時以下<br>" GroupName="Mode2B"></asp:RadioButton>
                                            <asp:RadioButton ID="rb2" runat="server" Text="　180至360小時<br>" GroupName="Mode2B"></asp:RadioButton>
                                            <asp:RadioButton ID="rb3" runat="server" Text="　361至540小時<br>" GroupName="Mode2B"></asp:RadioButton>
                                            <asp:RadioButton ID="rb4" runat="server" Text="　541小時以上<br>" GroupName="Mode2B"></asp:RadioButton>
                                            <div id="d2lab2" runat="server">
                                                依訓練時數(自辦、職前)：</div>
                                            <br>
                                            <asp:RadioButton ID="rb5" runat="server" Text="　450小時以下<br>" GroupName="Mode2B"></asp:RadioButton>
                                            <asp:RadioButton ID="rb6" runat="server" Text="　451至900小時<br>" GroupName="Mode2B"></asp:RadioButton>
                                            <asp:RadioButton ID="rb7" runat="server" Text="　901至1200小時<br>" GroupName="Mode2B"></asp:RadioButton>
                                            <asp:RadioButton ID="rb8" runat="server" Text="　1201小時以上<br>" GroupName="Mode2B"></asp:RadioButton>
                                            <asp:RadioButton ID="rb9" runat="server" Text="依機構開班數<br>" GroupName="Mode2B"></asp:RadioButton>
                                            <asp:RadioButton ID="rb10" runat="server" Text="依機構單位數<br>" GroupName="Mode2B"></asp:RadioButton>
                                            <asp:RadioButton ID="rb11" runat="server" Text="依受補助計畫總數<br>" GroupName="Mode2B"></asp:RadioButton>
                                            <asp:RadioButton ID="rb12" runat="server" Text="依訓練申請單位數<br>" GroupName="Mode2B"></asp:RadioButton>
                                            <asp:RadioButton ID="rb13" runat="server" Text="依訓練學員總人數<br>" GroupName="Mode2B"></asp:RadioButton>
                                            <asp:RadioButton ID="rb14" runat="server" Text="(尚未選擇)<br>" GroupName="Mode2B"></asp:RadioButton>
                                            <input id="value2b" runat="server" style="width: 46px; height: 22px" type="hidden"
                                                size="2" name="value2b" value="value2bvalue">
                                        </ItemTemplate>
                                    </asp:TemplateColumn>
                                    <%--<asp:TemplateColumn HeaderText="中心訪視">--%>
                                    <asp:TemplateColumn HeaderText="分署訪視">
                                        <HeaderStyle Width="15%"></HeaderStyle>
                                        <ItemTemplate>
                                            <%--中心訪視次數：<br />--%>
                                            分署訪視次數：<br />
                                            <asp:RadioButton ID="rbc1" runat="server" Text="　1次" GroupName="Mode2C"></asp:RadioButton><br>
                                            <asp:RadioButton ID="rbc2" runat="server" Text="　2次" GroupName="Mode2C"></asp:RadioButton><br>
                                            <asp:RadioButton ID="rbc3" runat="server" Text="　3次" GroupName="Mode2C"></asp:RadioButton><br>
                                            <asp:RadioButton ID="rbc4" runat="server" Text="　4次" GroupName="Mode2C"></asp:RadioButton><br>
                                            <%--中心訪視比率：<br />--%>
                                            分署訪視比率：<br />
                                            <asp:RadioButton ID="rbc5" runat="server" Text="　5%以上" GroupName="Mode2C"></asp:RadioButton><br>
                                            <asp:RadioButton ID="rbc6" runat="server" Text="　15%以上" GroupName="Mode2C"></asp:RadioButton><br>
                                            <asp:RadioButton ID="rbc7" runat="server" Text="　25%以上" GroupName="Mode2C"></asp:RadioButton><br>
                                            <asp:RadioButton ID="rbc8" runat="server" Text="　50%以上" GroupName="Mode2C"></asp:RadioButton><br>
                                        </ItemTemplate>
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="備註">
                                        <HeaderStyle Width="20%"></HeaderStyle>
                                        <ItemTemplate>
                                            <asp:TextBox ID="txtNote" runat="server" Width="100%" MaxLength="200" TextMode="MultiLine"
                                                Rows="3" Columns="15"></asp:TextBox>
                                        </ItemTemplate>
                                    </asp:TemplateColumn>
                                </Columns>
                            </asp:DataGrid>
                        </td>
                    </tr>
                    <tr>
                        <td align="center" class="whitecol">
                            <asp:Button ID="Button1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>&nbsp;
                            <asp:Button ID="Button2" runat="server" Text="套用上年度資料" CssClass="asp_button_M"></asp:Button>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
