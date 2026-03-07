<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_14_015.aspx.vb" Inherits="WDAIIP.SD_14_015" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訓練計畫總表</title>
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
    </script>
    <script type="text/javascript">
        function SelectOCID(Flag, MyValue) {
            if (Flag) {
                if (document.getElementById('OCIDValue').value == '') {
                    document.getElementById('OCIDValue').value = MyValue;
                }
                else {
                    document.getElementById('OCIDValue').value += ',' + MyValue;
                }
            }
            else {
                if (document.getElementById('OCIDValue').value.indexOf(',' + MyValue) != -1)
                    document.getElementById('OCIDValue').value = document.getElementById('OCIDValue').value.replace(',' + MyValue, '')
                else if (document.getElementById('OCIDValue').value.indexOf(MyValue + ',') != -1)
                    document.getElementById('OCIDValue').value = document.getElementById('OCIDValue').value.replace(MyValue + ',', '')
                else if (document.getElementById('OCIDValue').value.indexOf(MyValue) != -1)
                    document.getElementById('OCIDValue').value = document.getElementById('OCIDValue').value.replace(MyValue, '')
            }
        }

        function SelectSeqNo(Flag, MyValue, MyValue1, MyValue2) {
            if (Flag) {
                if (document.getElementById('SeqNoValue').value == '') {
                    document.getElementById('SeqNoValue').value = MyValue;
                    document.getElementById('PlanIDValue').value = MyValue1;
                    document.getElementById('ComIDNOValue').value = MyValue2;
                }
                else {
                    document.getElementById('SeqNoValue').value += ',' + MyValue;
                    document.getElementById('PlanIDValue').value += ',' + MyValue1;
                    document.getElementById('ComIDNOValue').value += ',' + MyValue2;
                }
            }
            else {
                if (document.getElementById('SeqNoValue').value.indexOf(',' + MyValue) != -1)
                    document.getElementById('SeqNoValue').value = document.getElementById('SeqNoValue').value.replace(',' + MyValue, '')
                else if (document.getElementById('SeqNoValue').value.indexOf(MyValue + ',') != -1)
                    document.getElementById('SeqNoValue').value = document.getElementById('SeqNoValue').value.replace(MyValue + ',', '')
                else if (document.getElementById('SeqNoValue').value.indexOf(MyValue) != -1)
                    document.getElementById('SeqNoValue').value = document.getElementById('SeqNoValue').value.replace(MyValue, '')

                if (document.getElementById('PlanIDValue').value.indexOf(',' + MyValue1) != -1)
                    document.getElementById('PlanIDValue').value = document.getElementById('PlanIDValue').value.replace(',' + MyValue1, '')
                else if (document.getElementById('PlanIDValue').value.indexOf(MyValue1 + ',') != -1)
                    document.getElementById('PlanIDValue').value = document.getElementById('PlanIDValue').value.replace(MyValue1 + ',', '')
                else if (document.getElementById('PlanIDValue').value.indexOf(MyValue1) != -1)
                    document.getElementById('PlanIDValue').value = document.getElementById('PlanIDValue').value.replace(MyValue1, '')

                if (document.getElementById('ComIDNOValue').value.indexOf(',' + MyValue2) != -1)
                    document.getElementById('ComIDNOValue').value = document.getElementById('ComIDNOValue').value.replace(',' + MyValue2, '')
                else if (document.getElementById('ComIDNOValue').value.indexOf(MyValue2 + ',') != -1)
                    document.getElementById('ComIDNOValue').value = document.getElementById('ComIDNOValue').value.replace(MyValue2 + ',', '')
                else if (document.getElementById('ComIDNOValue').value.indexOf(MyValue2) != -1)
                    document.getElementById('ComIDNOValue').value = document.getElementById('ComIDNOValue').value.replace(MyValue2, '')
            }
        }

        function CheckPrint() {
            if (document.form1.Radio1_1.checked) {
                if (document.getElementById('OCIDValue').value == '') {
                    alert('請選擇班級');
                    return false;
                }
                else {
                    if (document.getElementById('Years').value <= 97) {
                        openPrint('../../SQControl.aspx??SQ_AutoLogout=true&sys=BussinessTrain&filename=SD_14_015_1&OCID=' + document.getElementById('OCIDValue').value + '&Years=' + document.getElementById('Years').value, '', '');
                    }
                    else if (document.getElementById('Years').value == 98) {
                        openPrint('../../SQControl.aspx??SQ_AutoLogout=true&sys=BussinessTrain&filename=SD_14_015_1_2009&OCID=' + document.getElementById('OCIDValue').value + '&Years=' + document.getElementById('Years').value, '', '');
                    }
                    else if (document.getElementById('Years').value >= 99) {
                        openPrint('../../SQControl.aspx??SQ_AutoLogout=true&sys=BussinessTrain&filename=SD_14_015_1_2010&OCID=' + document.getElementById('OCIDValue').value + '&Years=' + document.getElementById('Years').value, '', '');
                    }
                }
            }
            else {
                if (document.getElementById('PlanIDValue').value == '') {
                    alert('請選擇班級');
                    return false;
                }
                else {
                    if (document.getElementById('Years').value <= 97) {
                        openPrint('../../SQControl.aspx??SQ_AutoLogout=true&sys=BussinessTrain&filename=SD_14_015&PLANID=' + document.getElementById('PlanIDValue').value + '&ComIDNO=' + document.getElementById('ComIDNOValue').value + '&SEQNO=' + document.getElementById('SeqNoValue').value + '&Years=' + document.getElementById('Years').value, '', '');
                    }
                        //openPrint('../../SQControl.aspx??SQ_AutoLogout=true&sys=BussinessTrain&filename=SD_14_015_5&path=TIMS&PLANID='+document.getElementById('PlanIDValue').value+'&ComIDNO='+document.getElementById('ComIDNOValue').value+'&SEQNO='+document.getElementById('SeqNoValue').value+'&Years='+document.getElementById('Years').value,'','');
                    else if (document.getElementById('Years').value == 98) {
                        openPrint('../../SQControl.aspx??SQ_AutoLogout=true&sys=BussinessTrain&filename=SD_14_015_2009&PLANID=' + document.getElementById('PlanIDValue').value + '&ComIDNO=' + document.getElementById('ComIDNOValue').value + '&SEQNO=' + document.getElementById('SeqNoValue').value + '&Years=' + document.getElementById('Years').value, '', '');
                    }
                    else if (document.getElementById('Years').value >= 99) {
                        openPrint('../../SQControl.aspx??SQ_AutoLogout=true&sys=BussinessTrain&filename=SD_14_015_2010&PLANID=' + document.getElementById('PlanIDValue').value + '&ComIDNO=' + document.getElementById('ComIDNOValue').value + '&SEQNO=' + document.getElementById('SeqNoValue').value + '&Years=' + document.getElementById('Years').value, '', '');
                    }
                }
            }
        }

        function SelectAll(Flag) {
            var MyTable = document.getElementById('DataGrid1');
            for (i = 1; i < MyTable.rows.length; i++) {
                MyTable.rows[i].cells[0].children[0].checked = Flag;
                SelectOCID(Flag, MyTable.rows[i].cells[0].children[0].value);
            }
        }

        function SelectAll2(Flag) {
            var MyTable = document.getElementById('DataGrid2');
            for (i = 1; i < MyTable.rows.length; i++) {
                var TSeqNo = MyTable.rows[i].cells[0].children[0].value;
                var TPlanID = MyTable.rows[i].cells[0].children[1].value;
                var TComIDNO = MyTable.rows[i].cells[0].children[2].value;
                MyTable.rows[i].cells[0].children[0].checked = Flag;
                SelectSeqNo(Flag, TSeqNo, TPlanID, TComIDNO);
            }
        }

        function SelectAll3(Flag) {
            var MyTable = document.getElementById('DataGrid2');
            document.getElementById('SeqNoValue').value += '';
            document.getElementById('PlanIDValue').value += '';
            document.getElementById('ComIDNOValue').value += '';
            for (i = 1; i < MyTable.rows.length; i++) {
                var TSeqNo = MyTable.rows[i].cells[0].children[0].value;
                var TPlanID = MyTable.rows[i].cells[0].children[1].value;
                var TComIDNO = MyTable.rows[i].cells[0].children[2].value;
                MyTable.rows[i].cells[0].children[0].checked = Flag;
                SelectSeqNo(Flag, TSeqNo, TPlanID, TComIDNO);
            }
        }

        function ClearData() {
            document.getElementById('TMID1').value = '';
            document.getElementById('OCID1').value = '';
            document.getElementById('TMIDValue1').value = '';
            document.getElementById('OCIDValue1').value = '';
        }

        function choose_class() {
            document.getElementById('OCID1').value = '';
            document.getElementById('TMID1').value = '';
            document.getElementById('OCIDValue1').value = '';
            document.getElementById('TMIDValue1').value = '';
            openClass('../02/SD_02_ch.aspx?&RID=' + document.getElementById('RIDValue').value);
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;表單列印&gt;&gt;訓練計畫總表</asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch">
                        <tr>
                            <td class="bluecol" width="20%">訓練機構 </td>
                            <td class="whitecol" colspan="3" width="80%">
                                <asp:TextBox ID="center" runat="server" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="Hidden2" runat="server">
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="310px"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">班級狀態 </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="Radio1" runat="server"  RepeatLayout="Flow" RepeatDirection="Horizontal" AutoPostBack="True">
                                    <asp:ListItem Value="0">未轉班</asp:ListItem>
                                    <asp:ListItem Value="1">已轉班</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="ClassTR" runat="server">
                            <td class="bluecol">職類/班別 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox><asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="button_b_Mini">
                                <input id="Button4" type="button" value="清除" name="Button4" runat="server" class="asp_button_S">
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                <span id="HistoryList" style="position: absolute; display: none; left: 270px">
                                    <asp:Table ID="HistoryTable" runat="server" Width="310"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">申請階段 </td>
                            <td class="whitecol" style="height: 17px">
                                <asp:DropDownList ID="AppStage" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr id="TRPlanPoint28" runat="server">
                            <td class="bluecol">計畫 </td>
                            <td class="whitecol" style="height: 38px" colspan="3">
                                <asp:RadioButtonList ID="PlanPoint" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                                    <asp:ListItem Value="1" Selected="True">產業人才投資計畫</asp:ListItem>
                                    <asp:ListItem Value="2">提升勞工自主學習計畫</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                    <div align="center">
                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" Width="27px" MaxLength="2">10</asp:TextBox>
                        <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                    </div>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></div>
                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True">
                                    <AlternatingItemStyle BackColor="#f5f5f5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="25px"></HeaderStyle>
                                            <HeaderTemplate>
                                                <input type="checkbox" onclick="SelectAll(this.checked);">
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <input id="OCID" type="checkbox" runat="server" name="OCID">
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassCName" HeaderText="班級名稱"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDate" HeaderText="開訓日期" DataFormatString="{0:d}">
                                            <HeaderStyle Width="100px"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="FTDate" HeaderText="結訓日期" DataFormatString="{0:d}">
                                            <HeaderStyle Width="100px"></HeaderStyle>
                                        </asp:BoundColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                                <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True">
                                    <AlternatingItemStyle BackColor="#f5f5f5" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="25px"></HeaderStyle>
                                            <HeaderTemplate>
                                                <input type="checkbox" onclick="SelectAll2(this.checked);">
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <input id="SeqNo" type="checkbox" runat="server">
                                                <input id="PlanID" type="hidden" runat="server">
                                                <input id="ComIDNO" type="hidden" runat="server">
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassCName" HeaderText="班級名稱"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDate" HeaderText="開訓日期" DataFormatString="{0:d}">
                                            <HeaderStyle Width="100px"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="FTDate" HeaderText="結訓日期" DataFormatString="{0:d}">
                                            <HeaderStyle Width="100px"></HeaderStyle>
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
                            <td align="center">
                                <asp:CheckBox ID="AllPrint" runat="server" Font-Size="X-Small" Text="全部列印"></asp:CheckBox><input id="Button3" type="button" value="列印" name="Button3" runat="server" class="asp_Export_M"></td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <input id="Years" type="hidden" name="Years" runat="server">
        <input id="OCIDValue" type="hidden" name="OCIDValue" runat="server">
        <input id="SeqNoValue" type="hidden" name="SeqNoValue" runat="server">
        <input id="ComIDNOValue" type="hidden" name="ComIDNOValue" runat="server">
        <input id="PlanIDValue" type="hidden" name="PlanIDValue" runat="server">
        <input id="orgid" type="hidden" name="orgid" runat="server">
        <input id="isPointYN" type="hidden" name="isPointYN" runat="server">
    </form>
</body>
</html>
