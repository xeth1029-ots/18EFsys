<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_07_002.aspx.vb" Inherits="WDAIIP.CP_07_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>受訓學員座談紀錄表</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script language="javascript" src="../../js/common.js"></script>
    <script language="javascript" type="text/javascript">
        function chkOrg() {
            if (document.getElementById("center").value == '') {
                alert('請選擇機構!');
                return false;
            }
            if (document.getElementById("OCID1").value == '') {
                alert('請選擇班別!');
                return false;
            }
        }
        function SetOneOCID() {
            document.getElementById('Button7').click();
        }

        function choose_class() {
            var RID = document.form1.RIDValue.value;
            if (document.getElementById('OCID1').value == '') {
                document.getElementById('Button7').click();
            }
            openClass('../../SD/02/SD_02_ch.aspx?RID=' + RID);
        }

        function ChkSOCID() {
            var MyTable = document.getElementById('DG_ClassInfo');
            var SOCIDvalue = '';
            for (var i = 1; i < MyTable.rows.length; i++) {
                if (MyTable.rows(i).cells(0).children(0).checked) {
                    //debugger;
                    //CB_SOCID
                    var socid1 = MyTable.rows(i).cells(0).children(0).value;

                    if (SOCIDvalue != '') { SOCIDvalue += ','; }
                    SOCIDvalue += socid1;
                }
            }
            document.getElementById('SOCIDvalue').value = SOCIDvalue;
        }

        function SelectAll(Flag) {
            var MyTable = document.getElementById('DG_ClassInfo');
            for (var i = 1; i < MyTable.rows.length; i++) {
                MyTable.rows(i).cells(0).children(0).checked = Flag;
            }
        }

        function chkprint(type) {
            if (document.getElementById('DG_ClassInfo') == null) {
                alert('目前無資料可供列印！');
                return false;
            }

            switch (type) {
                case 1:
                    // return true; //CP_07_002_blank
                    openPrint('../../SQControl.aspx?filename=CP_07_002_blank&Years=' + document.getElementById('years').value + '&PlanID=' + document.getElementById('PlanID').value);
                    break;
                default:
                    ChkSOCID();
                    //alert(document.getElementById('OCIDValue1').value);
                    //type : 2
                    var msg = '';
                    if (document.getElementById("center").value == '') { msg += '請選擇機構!\n'; }
                    if (document.getElementById("OCID1").value == '') { msg += '請選擇班別!\n'; }
                    if (document.getElementById('SOCIDvalue').value == '') { msg += '請選擇 『學員』！\n'; }
                    if (msg != '') {
                        alert(msg);
                        return false;
                    }
                    //CP_07_002
                    openPrint('../../SQControl.aspx?filename=CP_07_002&SOCID=' + document.getElementById('SOCIDvalue').value + '&Years=' + document.getElementById('years').value + '&OCID=' + document.getElementById('OCIDValue1').value + '&RID=' + document.getElementById('RIDValue').value + '&PlanID=' + document.getElementById('PlanID').value);
                    break;
            }


            if (type == 1) {
            }
            else {
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <table class="font" id="Table1" width="100%">
                        <tr>
                            <td class="font">
                                <p>
                                    首頁&gt;&gt;查核/績效管理&gt;&gt;<font color="#990000">受訓學員座談紀錄表</font>&nbsp;
                                </p>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch" id="Table2" cellpadding="1" cellspacing="1">
                        <tr>
                            <td width="100" class="bluecol">訓練機構
                            </td>
                            <td bgcolor="#ecf7ff" class="whitecol">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="410px"></asp:TextBox>
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                                <input id="RIDValue" type="hidden" name="Hidden2" runat="server">
                                <asp:Button ID="Button7" Style="display: none" runat="server" Text="Button7"></asp:Button>
                                <input id="years" type="hidden" name="years" runat="server">
                                <input id="PlanID" type="hidden" name="PlanID" runat="server">
                                <input id="distid" type="hidden" name="distid" runat="server">
                                <input id="SOCIDvalue" type="hidden" name="SOCIDvalue" runat="server">
                                <span id="HistoryList2" style="display: none; left: 117px; width: 152px; position: absolute; top: 72px; height: 38px">
                                    <asp:Table ID="HistoryRID" runat="server" Width="152px">
                                    </asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr class="SD_title">
                            <td width="100" class="bluecol">職類/班別
                            </td>
                            <td bgcolor="#ecf7ff" class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="200px"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="210px"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="button_b_Mini">
                                <input id="TMIDValue1" style="width: 40px; height: 22px" type="hidden" name="Hidden1" runat="server">
                                <input id="OCIDValue1" style="width: 32px; height: 22px" type="hidden" name="Hidden3" runat="server">
                                <input id="hidSearchTag" style="width: 23px; height: 22px" type="hidden" name="hidSearchTag" runat="server">
                                <span id="HistoryList" style="display: none; left: 270px; width: 208px; position: absolute; top: 104px; height: 38px">
                                    <asp:Table ID="HistoryTable" runat="server" Width="310">
                                    </asp:Table>
                                </span>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch" id="Table3" cellpadding="1" cellspacing="1">
                        <tr>
                            <td style="width: 101px" width="101" class="bluecol">開訓日期
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="STDate1" runat="server" onfocus="this.blur()" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">~
                            <asp:TextBox ID="STDate2" runat="server" onfocus="this.blur()" Columns="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            </td>
                        </tr>
                        <tr style="display: none">
                            <td style="width: 101px" width="101" class="bluecol">學號
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="StudentID" runat="server" Columns="10"></asp:TextBox>
                            </td>
                        </tr>
                    </table>
                    <p align="center">
                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label><asp:TextBox ID="TxtPageSize" Width="28px" runat="server">10</asp:TextBox>
                        <asp:Button ID="bt_search" Text="查詢" runat="server" CssClass="asp_button_S"></asp:Button>
                    </p>
                    <table class="font" id="Table4" width="100%" runat="server">
                        <tr>
                            <td align="center">
                                <asp:DataGrid ID="DG_ClassInfo" runat="server" Width="100%" CssClass="font" Visible="False" AllowSorting="True" AllowPaging="True" AutoGenerateColumns="False">
                                    <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderTemplate>
                                                <input onclick="SelectAll(this.checked);" type="checkbox">
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <input id="CB_SOCID" type="checkbox" name="CB_SOCID" runat="server">
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="StudentID" HeaderText="學號">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Name" HeaderText="姓名">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn>
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <HeaderTemplate>
                                                管控<br />
                                                單位
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:Label ID="lOrgName2" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="OrgName" SortExpression="OrgName" HeaderText="訓練機構">
                                            <HeaderStyle HorizontalAlign="Center" ForeColor="Black"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassCName" HeaderText="班別名稱">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn>
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <HeaderTemplate>
                                                問卷期間起日
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:Label ID="lQaySDate" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn>
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <HeaderTemplate>
                                                問卷期間迄日
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:Label ID="lQayFDate" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="CreateDate" HeaderText="填寫日期">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Button ID="btnedit" CommandName="edit" runat="server" Text="修改" ToolTip="在問卷調查起迄期間內且起迄日期均有值才可新增或修改"></asp:Button>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                <%--       <asp:BoundColumn Visible="False" DataField="OCID" HeaderText="OCID"></asp:BoundColumn>
                                    <asp:BoundColumn Visible="False" DataField="StudentID" HeaderText="StudentID"></asp:BoundColumn>
                                    <asp:BoundColumn Visible="true" DataField="socid"></asp:BoundColumn>--%>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <asp:Button ID="bt_PrintRpt" runat="server" Text="列印報表" CssClass="asp_Export_M"></asp:Button><asp:Button ID="bt_blankRpt" runat="server" Text="列印空白報表" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
