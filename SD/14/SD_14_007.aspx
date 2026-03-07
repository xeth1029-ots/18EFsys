<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_14_007.aspx.vb" Inherits="WDAIIP.SD_14_007" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>師資/助教名冊</title>
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
        function CheckSearch() {
            var vSTDate1 = document.getElementById('STDate1').value;
            var vSTDate2 = document.getElementById('STDate2').value;
            var vFTDate1 = document.getElementById('FTDate1').value;
            var vFTDate2 = document.getElementById('FTDate2').value;
            var msg = '';
            if (!checkDate(vSTDate1) && vSTDate1 != '') msg += '開訓起始日期必須為正確日期格式\n';
            if (!checkDate(vSTDate2) && vSTDate2 != '') msg += '開訓結束日期必須為正確日期格式\n';
            if (!checkDate(vFTDate1) && vFTDate1 != '') msg += '結訓起始日期必須為正確日期格式\n';
            if (!checkDate(vFTDate2) && vFTDate2 != '') msg += '結訓結束日期必須為正確日期格式\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function SelectAll(Flag) {
            var MyTable = document.getElementById('DataGrid1');
            for (i = 1; i < MyTable.rows.length; i++) {
                //OCID
                MyTable.rows[i].cells[0].children[0].checked = Flag;
                SelectItem(Flag, MyTable.rows[i].cells[0].children[0].value);
            }
        }

        //OCID
        function SelectItem(Flag, MyValue) {
            var SelectValue = document.getElementById('SelectValue');
            if (Flag) {
                if (SelectValue.value != '') SelectValue.value += ',';
                SelectValue.value += MyValue;
            }
            else {
                if (SelectValue.value.indexOf(',' + MyValue) != -1)
                    SelectValue.value = SelectValue.value.replace(',' + MyValue, '')
                else if (SelectValue.value.indexOf(MyValue + ',') != -1)
                    SelectValue.value = SelectValue.value.replace(MyValue + ',', '')
                else if (SelectValue.value.indexOf(MyValue) != -1)
                    SelectValue.value = SelectValue.value.replace(MyValue, '')
            }
        }

        function SelectAll2(Flag) {
            var MyTable = document.getElementById('DataGrid2');
            for (i = 1; i < MyTable.rows.length; i++) {
                MyTable.rows[i].cells[0].children[0].checked = Flag;
            }
        }

        //組合PPIPK
        function getselstr(vRadio1) {
            //vRadio1 0:未轉班(PPIPK) x1:已轉班 2:變更待審(PPIPK)
            var MyTable = document.getElementById('DataGrid2');
            var selsqlstr = document.getElementById('selsqlstr'); //.value
            var vselsqlstr = "";
            for (var i = 1; i < MyTable.rows.length; i++) {
                var TSeqNo = MyTable.rows[i].cells[0].children[0].value; //0
                var TPlanID = MyTable.rows[i].cells[0].children[1].value; //5
                var TComIDNO = MyTable.rows[i].cells[0].children[2].value; //6
                var SubSeqNo = MyTable.rows[i].cells[0].children[3].value; //8
                var TCdate = MyTable.rows[i].cells[0].children[4].value; //9
                //var TechIDstr = MyTable.rows[i].cells[0].children[5].value; //10
                //var PPIPK = MyTable.rows[i].cells[0].children[6].value; //10
                if (MyTable.rows[i].cells[0].children[0].checked) {
                    if (vRadio1 == '0') {
                        if (vselsqlstr != '') vselsqlstr += ',';
                        vselsqlstr += '\'' + TPlanID + '-' + TComIDNO + '-' + TSeqNo + '\'';
                    }
                    if (vRadio1 == '2') {
                        if (vselsqlstr != '') vselsqlstr += ',';
                        vselsqlstr += '\'' + TPlanID + '-' + TComIDNO + '-' + TSeqNo + '-' + SubSeqNo + '-' + TCdate + '\'';
                    }
                }
                //document.getElementById('selsqlstr').value = selsqlstr;
            }
            selsqlstr.value = vselsqlstr;
        }

        function ChkTechID() {
            var MyTable = document.getElementById('DataGrid2');
            var TechIDvalue = document.getElementById('TechIDvalue');
            var vTechIDvalue = '';
            for (var i = 1; i < MyTable.rows.length; i++) {
                if (MyTable.rows[i].cells[0].children[0].checked) {
                    if (vTechIDvalue != '') { vTechIDvalue += ','; }
                    vTechIDvalue += MyTable.rows[i].cells[0].children[5].value;
                }
            }
            TechIDvalue.value = vTechIDvalue;
        }

        function CheckPrint() {
            //var PLANIDValue = document.getElementById('PLANIDValue');
            //vRadio1 0:未轉班 1:已轉班 2:變更待審
            var vRadio1 = '0'; //0:未轉班
            if (document.form1.Radio1_1.checked) vRadio1 = '1'; //1:已轉班
            if (document.form1.Radio1_2.checked) vRadio1 = '2'; //2:變更審核
            //var ComIDNOValue = document.getElementById('ComIDNOValue');
            //var SubSeqNoValue = document.getElementById('SubSeqNoValue');
            //var CDateItem = document.getElementById('CDateItem');
            //function CheckPrint(SMpath) {
            //var SMpath=document.getElementById('SMpath').value;
            //0:未轉班
            if (vRadio1 == '0') {
                var selsqlstr = document.getElementById('selsqlstr'); //.value
                getselstr(vRadio1);
                if (selsqlstr.value == '') {
                    alert('請選擇班級！');
                    return false;
                }
            }
            //1:已轉班
            if (vRadio1 == '1') {
                if (document.getElementById('SelectValue').value == '') {
                    alert('請選擇班級');
                    return false;
                }
            }
            //2:變更審核
            if (vRadio1 == '2') {
                var selsqlstr = document.getElementById('selsqlstr'); //.value
                ChkTechID();
                getselstr(vRadio1);
                if (selsqlstr.value == '') {
                    alert('請選擇班級！');
                    return false;
                }
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tbody>
                <tr>
                    <td>
                        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <tr>
                                <td>
                                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;表單列印&gt;&gt;訓練計畫師資/助教名冊</asp:Label>
                                </td>
                            </tr>
                        </table>
                        <table class="table_sch">
                            <tr>
                                <td class="bluecol" width="20%">訓練機構</td>
                                <td class="whitecol" colspan="3">
                                    <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                    <input id="RIDValue" type="hidden" runat="server">
                                    <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                                    <span id="HistoryList2" style="position: absolute; display: none">
                                        <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">班級狀態 </td>
                                <td class="whitecol" colspan="3">
                                    <asp:RadioButtonList ID="Radio1" runat="server" Width="300px" AutoPostBack="True" RepeatDirection="Horizontal" RepeatLayout="Flow" CssClass="font">
                                        <asp:ListItem Value="0">未轉班</asp:ListItem>
                                        <asp:ListItem Value="1">已轉班</asp:ListItem>
                                        <asp:ListItem Value="2">變更待審</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">開訓期間 </td>
                                <td class="whitecol" colspan="3">
                                    <asp:TextBox ID="STDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                    <span runat="server">
                                        <img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span> ~
                                    <asp:TextBox ID="STDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                    <span runat="server">
                                        <img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">結訓期間 </td>
                                <td class="whitecol" colspan="3">
                                    <asp:TextBox ID="FTDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                    <span runat="server">
                                        <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span> ~
                                    <asp:TextBox ID="FTDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                    <span runat="server">
                                        <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                </td>
                            </tr>
                            <tr id="TRPlanPoint28" runat="server">
                                <td class="bluecol">計畫 </td>
                                <td class="whitecol" colspan="3">
                                    <asp:RadioButtonList ID="PlanPoint" runat="server" AutoPostBack="True" RepeatDirection="Horizontal" RepeatLayout="Flow" CssClass="font">
                                        <asp:ListItem Value="1" Selected="True">產業人才投資計畫</asp:ListItem>
                                        <asp:ListItem Value="2">提升勞工自主學習計畫</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                        </table>
                        <div align="center" class="whitecol">
                            <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                            <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                            <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                        </div>
                        <div align="center">
                            <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></div>
                        <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td>
                                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AllowPaging="True" AutoGenerateColumns="False" AllowCustomPaging="True" CellPadding="8">
                                        <AlternatingItemStyle BackColor="#f5f5f5" />
                                        <HeaderStyle CssClass="head_navy" />
                                        <Columns>
                                            <asp:TemplateColumn>
                                                <HeaderStyle Width="5%"></HeaderStyle>
                                                <HeaderTemplate>
                                                    <input onclick="SelectAll(this.checked);" type="checkbox">
                                                </HeaderTemplate>
                                                <ItemTemplate>
                                                    <input id="OCID" type="checkbox" name="OCID" runat="server">
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構">
                                                <HeaderStyle Width="40%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="ClassCName" HeaderText="班級名稱">
                                                <HeaderStyle Width="35%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="STDate" HeaderText="開訓日期" DataFormatString="{0:d}">
                                                <HeaderStyle Width="10%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="FTDate" HeaderText="結訓日期" DataFormatString="{0:d}">
                                                <HeaderStyle Width="10%"></HeaderStyle>
                                            </asp:BoundColumn>
                                        </Columns>
                                        <PagerStyle Visible="False"></PagerStyle>
                                    </asp:DataGrid>
                                    <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AllowPaging="True" AutoGenerateColumns="False" AllowCustomPaging="True" PagerStyle-Visible="false" CellPadding="8">
                                        <AlternatingItemStyle BackColor="#f5f5f5" />
                                        <HeaderStyle CssClass="head_navy" />
                                        <Columns>
                                            <asp:TemplateColumn>
                                                <HeaderStyle Width="5%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" />
                                                <HeaderTemplate>
                                                    <input onclick="SelectAll2(this.checked);" type="checkbox">
                                                </HeaderTemplate>
                                                <ItemTemplate>
                                                    <input id="SeqNo" type="checkbox" runat="server">
                                                    <input id="PlanID" type="hidden" runat="server">
                                                    <input id="ComIDNO" type="hidden" runat="server">
                                                    <input id="SubSeqNO" type="hidden" runat="server">
                                                    <input id="CDateValue" type="hidden" runat="server">
                                                    <input id="TechID" type="hidden" runat="server">
                                                    <input id="HidPPIPK" type="hidden" runat="server">
                                                    <input id="hPTDRID" type="hidden" runat="server">
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構">
                                                <HeaderStyle Width="30%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="ClassCName" HeaderText="班級名稱">
                                                <HeaderStyle Width="30%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="STDate" HeaderText="開訓日期" DataFormatString="{0:d}">
                                                <HeaderStyle Width="10%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="FTDate" HeaderText="結訓日期" DataFormatString="{0:d}">
                                                <HeaderStyle Width="10%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn HeaderText="申請變更時間">
                                                <HeaderStyle Width="15%"></HeaderStyle>
                                            </asp:BoundColumn>
                                        </Columns>
                                        <PagerStyle Visible="False"></PagerStyle>
                                    </asp:DataGrid>
                                </td>
                            </tr>
                            <tr>
                                <td align="center">
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                    <uc1:PageControler ID="PageControler2" runat="server"></uc1:PageControler>
                                </td>
                            </tr>
                            <tr>
                                <td align="center" class="whitecol">
                                    <asp:Button ID="btnPrint1" runat="server" Text="列印" CssClass="asp_Export_M" />
                                    <%--<input id="print" runat="server" type="button" value="列印" class="asp_button_S" />--%>
                                </td>
                            </tr>
                        </table>
                        <%--<input id="SMpath" type="hidden" name="SMpath" runat="server">--%>
                        <input id="ROC_Years" type="hidden" name="Years" runat="server" />
                        <input id="SelectValue" type="hidden" name="SelectValue" runat="server" />
                        <input id="SeqNoValue" type="hidden" name="SeqNoValue" runat="server" />
                        <input id="ComIDNOValue" type="hidden" name="ComIDNOValue" runat="server" />
                        <input id="PLANIDValue" type="hidden" name="PLANIDValue" runat="server" />
                        <input id="KindValue" type="hidden" name="KindValue" runat="server" />
                        <input id="CDateItem" type="hidden" name="CDateItem" runat="server" />
                        <input id="SubSeqNoValue" type="hidden" name="SubSeqNoValue" runat="server" />
                        <input id="selsqlstr" type="hidden" name="selsqlstr" runat="server" />
                        <input id="TechIDvalue" type="hidden" name="TechIDvalue" runat="server" />
                        <%--<input id="Years2" type="hidden" name="Years2" runat="server"/>--%>
                    </td>
                </tr>
            </tbody>
        </table>
    </form>
</body>
</html>
