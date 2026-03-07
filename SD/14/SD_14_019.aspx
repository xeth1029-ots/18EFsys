<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_14_019.aspx.vb" Inherits="WDAIIP.SD_14_019" %>
 

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>銀行匯款資料匯出</title>
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
    <script type="text/javascript" >
        /*function printkind(){if (getValue("print_type")=='2') {document.getElementById('print_orderyby').value ='c.StudentID'}else {document.getElementById('print_orderyby').value = 'd.IDNO'}}function SelectAll(Flag){var MyTable=document.getElementById('DataGrid1');for(i=1;i<MyTable.rows.length;i++){MyTable.rows[i].cells[0].children[0].checked=Flag;SelectItem(Flag ,MyTable.rows[i].cells[0].children[0].value);}}function SelectItem(Flag,MyValue){if(Flag){if(document.getElementById('SelectValue').value==''){document.getElementById('SelectValue').value=MyValue;}else{document.getElementById('SelectValue').value+=','+MyValue;}}else{if(document.getElementById('SelectValue').value.indexOf(','+MyValue)!=-1)document.getElementById('SelectValue').value=document.getElementById('SelectValue').value.replace(','+MyValue,'')else if(document.getElementById('SelectValue').value.indexOf(MyValue+',')!=-1)document.getElementById('SelectValue').value=document.getElementById('SelectValue').value.replace(MyValue+',','')else if(document.getElementById('SelectValue').value.indexOf(MyValue)!=-1)document.getElementById('SelectValue').value=document.getElementById('SelectValue').value.replace(MyValue,'')}}*/
        function choose_class() {
            var RID = document.form1.RIDValue.value;
            document.form1.TMID1.value = '';
            document.form1.TMIDValue1.value = '';
            document.form1.OCID1.value = '';
            document.form1.OCIDValue1.value = '';
            openClass('../02/SD_02_ch.aspx?RID=' + RID);
        }
        function CheckSearch() {
            var msg = '';
            var STDate1 = document.getElementById('STDate1').value;
            var STDate2 = document.getElementById('STDate2').value;
            var FTDate1 = document.getElementById('FTDate1').value;
            var FTDate2 = document.getElementById('FTDate2').value;
            if (!checkDate(STDate1) && STDate1 != '') msg += '開訓起始日期必須為正確日期格式\n';
            if (!checkDate(STDate2) && STDate2 != '') msg += '開訓結束日期必須為正確日期格式\n';
            if (!checkDate(FTDate1) && FTDate1 != '') msg += '結訓起始日期必須為正確日期格式\n';
            if (!checkDate(FTDate2) && FTDate2 != '') msg += '結訓結束日期必須為正確日期格式\n';
            if (msg != '') {
                alert(msg);
                return false;
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
                                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;表單列印&gt;&gt;銀行匯款資料匯出</asp:Label>
                                </td>
                            </tr>
                        </table>
                        <table class="table_sch">
                            <tbody>
                                <tr>
                                    <td class="bluecol" style="width: 20%">訓練機構 </td>
                                    <td class="whitecol" colspan="3">
                                        <asp:TextBox ID="center" runat="server" Width="70%"></asp:TextBox>
                                        <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                        <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                                        <span id="HistoryList2" style="position: absolute; display: none"><asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table></span>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol">職類/班別 </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                        <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="40%"></asp:TextBox>
                                        <input onclick="choose_class()" type="button" value="..." class="button_b_Mini">
                                        <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                        <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                        <span id="HistoryList" style="position: absolute; display: none; left: 28%"><asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table></span>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol">開訓期間 </td>
                                    <td class="whitecol" colspan="3">
                                        <asp:TextBox ID="STDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                        <span runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span> ～
                                        <asp:TextBox ID="STDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                        <span runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol">結訓期間 </td>
                                    <td class="whitecol" colspan="3">
                                        <asp:TextBox ID="FTDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                        <span runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span> ～
                                        <asp:TextBox ID="FTDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                        <span runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                    </td>
                                </tr>
                                <tr id="TRPlanPoint28" runat="server">
                                    <td class="bluecol">計畫 </td>
                                    <td class="whitecol" colspan="3">
                                        <asp:RadioButtonList ID="PlanPoint" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="0" Selected="True">不區分</asp:ListItem>
                                            <asp:ListItem Value="1">產業人才投資計畫</asp:ListItem>
                                            <asp:ListItem Value="2">提升勞工自主學習計畫</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </td>
                                </tr>
                            </tbody>
                        </table>
                        <div align="center" class="whitecol">
                            <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                            <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                            <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                        </div>
                        <div align="center"><asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></div>
                        <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server" class="font">
                            <tr>
                                <td class="whitecol" align="center">
                                    <asp:RadioButtonList ID="ExportType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1" Selected="True">依【中國信託文字檔】格式&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</asp:ListItem>
                                        <asp:ListItem Value="2">依【臺灣銀行XML檔】格式&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</asp:ListItem>
                                        <asp:ListItem Value="3">依【ach】格式</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol">付款日期：
                                    <asp:TextBox ID="txtNowDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                    <span runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('txtNowDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol"><asp:Label ID="lab_news1" runat="server" ForeColor="Red">※尚未完成補助審核之班級，無法匯出銀行匯款資料</asp:Label></td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                                        <AlternatingItemStyle BackColor="#f5f5f5" />
                                        <HeaderStyle CssClass="head_navy" />
                                        <Columns>
                                            <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構" HeaderStyle-Width="35%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="ClassCName" HeaderText="班級名稱" HeaderStyle-Width="35%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="STDate" HeaderText="開訓日期" DataFormatString="{0:d}" HeaderStyle-Width="10%">
                                                <HeaderStyle Width="80px"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="FTDate" HeaderText="結訓日期" DataFormatString="{0:d}" HeaderStyle-Width="10%"></asp:BoundColumn>
                                            <asp:TemplateColumn HeaderStyle-Width="10%">
                                                <HeaderTemplate>匯出</HeaderTemplate>
                                                <ItemStyle HorizontalAlign="Center" />
                                                <ItemTemplate>
                                                    <input id="OCID" type="hidden" runat="server" name="OCID">
                                                    <asp:Button ID="btnExport" runat="server" class="asp_Export_M" Text="匯出" CommandName="Export"></asp:Button>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                        <PagerStyle Visible="False"></PagerStyle>
                                    </asp:DataGrid>
                                </td>
                            </tr>
                            <tr>
                                <td align="center"><uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler></td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </tbody>
        </table>
        <%--<asp:Button ID="Button6" Style="display: none" runat="server" Text="Button6"></asp:Button>--%>
    </form>
</body>
</html>