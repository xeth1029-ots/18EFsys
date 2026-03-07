<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_14_014.aspx.vb" Inherits="WDAIIP.SD_14_014" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>教學環境資料表(SD_14_014)</title>
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

        function SelectAll2(v_Flag) {
            var MyTable = document.getElementById('DataGrid1');
            var SelectValue1 = document.getElementById('SelectValue1'); //0:未轉班 1:已轉班 2:變更待審
            for (i = 1; i < MyTable.rows.length; i++) {
                var v_chkSeqNo = MyTable.rows[i].cells[0].children[0].value; //0:chkSeqNo
                var v_hid_PCS = MyTable.rows[i].cells[0].children[1].value; //1:hid_PCS
                var v_hid_PCS2 = MyTable.rows[i].cells[0].children[2].value; //2:hid_PCS2
                var v_hid_Radio1 = MyTable.rows[i].cells[0].children[3].value; //3:hid_Radio1
                MyTable.rows[i].cells[0].children[0].checked = v_Flag;
                if (v_hid_Radio1 == '2') {
                    //0:未轉班 1:已轉班 2:變更待審
                    SelectSeqNo(v_Flag, v_hid_PCS2, v_hid_Radio1);
                }
                else {
                    //0:未轉班 1:已轉班 2:變更待審
                    SelectSeqNo(v_Flag, v_hid_PCS, v_hid_Radio1);
                }
            }
        }

        function SelectSeqNo(v_Flag, v_hid_PCS, v_hid_Radio1) {
            var SelectValue1 = document.getElementById('SelectValue1'); //0:未轉班 1:已轉班 2:變更待審
            var Hid_PCSALL = document.getElementById('Hid_PCSALL');
            var Hid_PCS2ALL = document.getElementById('Hid_PCS2ALL');
            if (v_hid_Radio1 == '2') {
                //0:未轉班 1:已轉班 2:變更待審
                //SelectSeqNo(v_Flag, v_hid_PCS2, v_hid_Radio1);
                if (v_Flag) {
                    if (Hid_PCS2ALL.value.indexOf(v_hid_PCS) == -1) {
                        if (Hid_PCS2ALL.value != '') { Hid_PCS2ALL.value += ","; }
                        Hid_PCS2ALL.value += v_hid_PCS;
                    }
                }
                else {
                    if (Hid_PCS2ALL.value.indexOf(v_hid_PCS) != -1) {
                        Hid_PCS2ALL.value = Hid_PCS2ALL.value.replace(',' + v_hid_PCS + ',', ',')
                        Hid_PCS2ALL.value = Hid_PCS2ALL.value.replace(',' + v_hid_PCS, '')
                        Hid_PCS2ALL.value = Hid_PCS2ALL.value.replace(v_hid_PCS + ',', '')
                        Hid_PCS2ALL.value = Hid_PCS2ALL.value.replace(v_hid_PCS, '')
                    }
                }
            }
            else {
                //0:未轉班 1:已轉班 2:變更待審
                //SelectSeqNo(v_Flag, v_hid_PCS, v_hid_Radio1);
                if (v_Flag) {
                    if (Hid_PCSALL.value.indexOf(v_hid_PCS) == -1) {
                        if (Hid_PCSALL.value != '') { Hid_PCSALL.value += ","; }
                        Hid_PCSALL.value += v_hid_PCS;
                    }
                } else {
                    if (Hid_PCSALL.value.indexOf(v_hid_PCS) != -1) {
                        Hid_PCSALL.value = Hid_PCSALL.value.replace(',' + v_hid_PCS + ',', ',')
                        Hid_PCSALL.value = Hid_PCSALL.value.replace(',' + v_hid_PCS, '')
                        Hid_PCSALL.value = Hid_PCSALL.value.replace(v_hid_PCS + ',', '')
                        Hid_PCSALL.value = Hid_PCSALL.value.replace(v_hid_PCS, '')
                    }
                }
            }
        }

        function CheckPrint() {
            //0:未轉班 1:已轉班 2:變更待審
            var SelectValue1 = document.getElementById('SelectValue1');
            var Hid_PCSALL = document.getElementById('Hid_PCSALL');
            var Hid_PCS2ALL = document.getElementById('Hid_PCS2ALL');
            if (SelectValue1.value == '') {
                alert('請選擇班級');
                return false;
            }
            if (SelectValue1.value == '2') {
                if (Hid_PCS2ALL.value == '') {
                    alert('請選擇班級');
                    return false;
                }
            }
            else {
                if (Hid_PCSALL.value == '') {
                    alert('請選擇班級');
                    return false;
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;表單列印&gt;&gt;教學環境資料表</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch">
                        <tr>
                            <td class="bluecol" width="20%">訓練機構 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="Hidden2" runat="server">
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">班級狀態 </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="Radio1" runat="server" AutoPostBack="True" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
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
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span> ～
                                <asp:TextBox ID="FTDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                    </table>
                    <div align="center" class="whitecol">
                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                        <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                    </div>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                    </div>
                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <AlternatingItemStyle BackColor="#f5f5f5" />
                                    <Columns>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構" HeaderStyle-Width="30%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassCName" HeaderText="班級名稱" HeaderStyle-Width="30%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDate" HeaderText="開訓日期" DataFormatString="{0:d}">
                                            <HeaderStyle Width="10%" HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="FTDate" HeaderText="結訓日期" DataFormatString="{0:d}">
                                            <HeaderStyle Width="10%" HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="10%" HorizontalAlign="Center"></HeaderStyle>
                                            <HeaderTemplate>申請變更時間</HeaderTemplate>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:Label ID="LabModifyDate" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="10%" HorizontalAlign="Center" />
                                            <HeaderTemplate>功能</HeaderTemplate>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <%--<input id="chkSeqNo" type="checkbox" runat="server" />--%>
                                                <input id="hid_PCS" runat="server" type="hidden" />
                                                <input id="hid_PCS2" runat="server" type="hidden" />
                                                <input id="hid_Radio1" runat="server" type="hidden" />
                                                <asp:Button ID="BtnPrint2" runat="server" Text="列印" CssClass="asp_Export_M" DisabledCssClass="asp_button_S_disabled" CommandName="PRINT2" />
                                                <asp:Button ID="BtnPrint2R" runat="server" Text="列印(遠距)" CssClass="asp_Export_M" DisabledCssClass="asp_button_S_disabled" CommandName="PRINT2R" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
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
                            <td align="center">
                                <%--
							    <asp:Button ID="BtnPrint1" runat="server" Text="列印" CssClass="asp_button_S" DisabledCssClass="asp_button_S_disabled" />
							    <input id="print" type="button" value="列印" runat="server" class="asp_button_S" onclick="return print_onclick()">
                                --%>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <%----%>
        <input id="ROC_Years" type="hidden" runat="server" />
        <input id="SelectValue1" type="hidden" runat="server">
        <input id="Hid_PCSALL" type="hidden" runat="server">
        <input id="Hid_PCS2ALL" type="hidden" runat="server">
    </form>
</body>
</html>
