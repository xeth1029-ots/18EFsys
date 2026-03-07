<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_014_add.aspx.vb" Inherits="WDAIIP.SD_05_014_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>學員出缺勤作業(產投)</title>
    <meta content="False" name="vs_snapToGrid">
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <style type="text/css">
        .FixedTitleRow { z-index: 10; position: relative; background-color: #0eabd6; color: white; top: expression(this.offsetParent.scrollTop); }
        .FixedTitleColumn { position: relative; left: expression(this.parentElement.offsetParent.scrollLeft); }
        .FixedDataColumn { position: relative; left: expression(this.parentElement.offsetParent.parentElement.scrollLeft); }
        .DivWidth { position: static; width: 600px; display: inline; height: 350px; overflow: auto; cursor: default; }
        .DivHeight { position: static; overflow-y: scroll; display: inline; height: 350px; cursor: default; }
    </style>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        function choose_class() {
            openClass('../02/SD_02_ch.aspx?RID=' + document.getElementById('RIDValue').value + '&BtnName=Button1');
        }

        function CheckData() {
            var Cst_Name = 1;
            var Cst_LeaveHours = 3; //hidLeaveDate3 /LeaveDate3
            var LeaveDateS1 = '';
            var NowDate = document.getElementById('LeaveDate').value;
            var msg = '';

            var MyTable = document.getElementById('DataGrid1');
            if (MyTable) {
                LeaveDateS1 = ''
                for (i = 1; i < MyTable.rows.length; i++) {
                    var NameValue = MyTable.rows[i].cells[Cst_Name].innerHTML;
                    var cellsLeaveHours = MyTable.rows[i].cells[Cst_LeaveHours];
                    //cellsLeaveHours.children[0].value:Hours
                    //cellsLeaveHours.children[1].value:hidLeaveDate3
                    //cellsLeaveHours.children[2].value:HidSOCID
                    if (cellsLeaveHours.children[0].value != '') {
                        if (cellsLeaveHours.children[1].value != '') {
                            LeaveDateS1 = cellsLeaveHours.children[1].value;
                            if (LeaveDateS1.indexOf(",") > -1) {
                                var LeaveDateA1 = LeaveDateS1.split(",");
                                for (i2 = 0; i2 < LeaveDateA1.length; i2++) {
                                    var TempDate = LeaveDateA1[i2];
                                    if (compareDate(TempDate, NowDate) == 0) {
                                        msg += NameValue + ' 該學員請假當日已有紀錄,請確認是否要儲存此紀錄?\n';
                                        i2 = LeaveDateA1.length;
                                    }
                                }
                            }
                            else {
                                var TempDate = LeaveDateS1;
                                if (compareDate(TempDate, NowDate) == 0) {
                                    msg += NameValue + ' 該學員請假當日已有紀錄,請確認是否要儲存此紀錄?\n';
                                }
                            }
                        }
                    }
                }
            }

            if (msg != '') {
                return confirm(msg);
            }
            else {
                return true;
            }
        }

        function IsDate(MyDate) {
            if (MyDate != '' && !checkDate(MyDate)) { return false; }
            return true;
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <%--<table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0"> <tr><td>首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">學員出缺勤作業</font> </td></tr> </table>--%>
                    <table class="table_sch" id="AddTable" runat="server" width="100%" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" width="20%">訓練機構 </td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="55%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="Hidden2" runat="server">
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="asp_button_Mini">
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span></td>
                        </tr>
                        <tr>
                            <td class="bluecol">職類/班別 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="25%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class();" type="button" value="..." class="asp_button_Mini">
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                <span id="HistoryList" style="position: absolute; display: none; left: 270px">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                                <asp:Button ID="Button1" runat="server" Text="查詢學員(隱藏)" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">未出席日期 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="LeaveDate" runat="server" Columns="10" MaxLength="20" Width="15%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('LeaveDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="2">
                                <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                    <tr>
                                        <td>點選姓名可以查詢個人所有請假歷程 </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <div id="scrollDiv" runat="server" class="whitecol">
                                                <asp:DataGrid ID="DataGrid1" runat="server" AutoGenerateColumns="False" CssClass="font" Width="100%" CellPadding="8">
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <%--<AlternatingItemStyle BackColor="#F5F5F5" />--%>
                                                    <Columns>
                                                        <asp:BoundColumn DataField="StudentID" HeaderText="學號" HeaderStyle-Width="18%"></asp:BoundColumn>
                                                        <asp:BoundColumn DataField="Name" HeaderText="姓名" HeaderStyle-Width="10%"></asp:BoundColumn>
                                                        <asp:BoundColumn DataField="CountHours" HeaderText="目前累積缺席時數" HeaderStyle-Width="18%"></asp:BoundColumn>
                                                        <asp:TemplateColumn HeaderText="缺席時數" HeaderStyle-Width="18%" ItemStyle-HorizontalAlign="Center">
                                                            <ItemStyle CssClass="whitecol" Wrap="false" />
                                                            <ItemTemplate>
                                                                <asp:TextBox ID="Hours" runat="server" Columns="10" MaxLength="3"></asp:TextBox>
                                                                <input id="hidLeaveDate3" type="hidden" runat="server">
                                                                <asp:HiddenField ID="HidSOCID" runat="server" />
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="不列入缺席時數" HeaderStyle-Width="18%" ItemStyle-HorizontalAlign="Center">
                                                            <ItemStyle CssClass="whitecol" Wrap="false" />
                                                            <ItemTemplate>
                                                                <asp:TextBox ID="NIHOURS" runat="server" Columns="10" MaxLength="3"></asp:TextBox>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="不列入缺席原因" HeaderStyle-Width="18%" ItemStyle-HorizontalAlign="Center">
                                                            <ItemStyle CssClass="whitecol" Wrap="false" />
                                                            <ItemTemplate>
                                                                <%--<asp:CheckBox ID="chk_LEAVEID05" runat="server" Text="喪假" />--%>
                                                                <asp:TextBox ID="NIREASONS" runat="server" Columns="12" MaxLength="10"></asp:TextBox>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                    </Columns>
                                                </asp:DataGrid>
                                            </div>
                                        </td>
                                    </tr>
                                </table>
                                <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="2" class="whitecol">
                                <asp:Button ID="Button3" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                                &nbsp;<asp:Button ID="Button4" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="EditTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td class="bluecol" width="20%">班級名稱 </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="ClassCName2" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">學號 </td>
                            <td class="whitecol">
                                <asp:Label ID="StudentID2" runat="server"></asp:Label>
                            </td>
                            <td class="bluecol">姓名 </td>
                            <td class="whitecol">
                                <asp:Label ID="Name2" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練時數 </td>
                            <td class="whitecol">
                                <asp:Label ID="THours" runat="server"></asp:Label>
                            </td>
                            <td class="bluecol">目前累積缺席時數 </td>
                            <td class="whitecol">
                                <asp:Label ID="CountHours" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4">個人缺席歷程(依照日期排序)-紫色的表示您所選擇的修改資料 </td>
                        </tr>
                        <tr>
                            <td colspan="4" class="whitecol">
                                <asp:DataGrid ID="DataGrid2" runat="server" AutoGenerateColumns="False" CssClass="font" Width="100%" CellPadding="8">
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <Columns>
                                        <asp:TemplateColumn HeaderText="未出席日期" HeaderStyle-Width="20%" ItemStyle-HorizontalAlign="Center">
                                            <ItemTemplate>
                                                <asp:TextBox ID="LeaveDate2" runat="server" Columns="10" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                                <img id="IMG1" style="cursor: pointer" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30">
                                                <input id="Change" type="hidden" value="0" runat="server">
                                                <asp:HiddenField ID="HidSTOID" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="時數" HeaderStyle-Width="20%" ItemStyle-HorizontalAlign="Center">
                                            <ItemTemplate>
                                                <asp:TextBox ID="Hours2" runat="server" Columns="5" Width="60%"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="不列入缺席原因" HeaderStyle-Width="60%" ItemStyle-HorizontalAlign="Center">
                                            <ItemTemplate>
                                                <%--<asp:CheckBox ID="chk_LEAVEID05" runat="server" Text="喪假" />--%>
                                                <asp:TextBox ID="NIREASONS" runat="server" Columns="12" MaxLength="10" Width="80%"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="4" class="whitecol">
                                <asp:Button ID="Button5" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                                &nbsp;<asp:Button ID="Button6" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <input id="STOIDvalue" type="hidden" name="STOIDvalue" runat="server" />
        <input id="SOCIDvalue" type="hidden" name="SOCIDvalue" runat="server" />
        <%--<asp:HiddenField ID="Hid_SUBSIDYCOST_FLG" runat="server" />--%>
    </form>
</body>
</html>
