<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_04_001.aspx.vb" Inherits="WDAIIP.TC_04_001" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>班級審核作業</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        //unction fnOpen(myvalue, planid,ComIDNO){,if (myvalue=='Y') {,win=window.open("TC_04_Trans.aspx?PlanID="+planid+"&ComIDNO="+ComIDNO,"","height=250,width=450,mentbar=yes,scrollbars=yes,resizable=yes");,win.moveTo(50,60);,},},

        function SavaData() {
            var MyTable = document.getElementById('DataGrid2');
            var msg = '';
            for (i = 1; i < MyTable.rows.length; i++) {
                var Result = MyTable.rows[i].cells[7].children[0];
                if (Result.selectedIndex != 0) {
                    msg += MyTable.rows[i].cells[5].innerHTML + '\n';
                }
            }
            if (msg == '') {
                alert('請選擇要取消審核的班級')
                return false;
            }
            else {
                return confirm('您確定要取消審核以下班級?\n\n' + msg);
            }
        }

        function ChangeAll(j) {
            var MyTable = document.getElementById('dgPlan');
            for (i = 1; i < MyTable.rows.length; i++) {
                if (MyTable.rows[i].cells[8] != null) {
                    MyTable.rows[i].cells[8].children[0].selectedIndex = j;
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;班級審核</asp:Label>
                    <%--<font color="#990000">-新增</font> (<font color="#ff0000">*</font>為必填欄位)--%>
                </td>
            </tr>
        </table>
        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch" id="Table3" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol" style="width: 20%">訓練機構</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="Org" type="button" value="..." name="Org" runat="server">
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">
                                <asp:Label ID="LabTMID" runat="server">訓練職類</asp:Label></td>
                            <td class="whitecol" width="30%">
                                <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Columns="30" Width="60%"></asp:TextBox>
                                <input id="btu_sel" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." name="btu_sel" runat="server">
                                <input id="TPlanid" style="width: 27px; height: 22px" type="hidden" name="TPlanid" runat="server">
                                <input id="trainValue" style="width: 43px; height: 22px" type="hidden" name="trainValue" runat="server">
                                <input id="jobValue" style="width: 43px; height: 22px" type="hidden" name="jobValue" runat="server">
                            </td>
                            <td width="20%" class="bluecol">
                                <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label></td>
                            <td class="whitecol" width="30%">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="60%"></asp:TextBox>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server">
                                <input id="cjobValue" type="hidden" name="cjobValue" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">班別名稱</td>
                            <td class="whitecol">
                                <asp:TextBox ID="ClassName" runat="server" Columns="30" Width="50%"></asp:TextBox></td>
                            <td class="bluecol">期別</td>
                            <td class="whitecol">
                                <asp:TextBox ID="CyclType" runat="server" MaxLength="2" Columns="5" Width="40%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">申請日期</td>
                            <td class="whitecol" colspan="3">
                                <span runat="server">
                                    <asp:TextBox ID="UNIT_SDATE" runat="server" Width="15%" MaxLength="10" ToolTip="日期格式:99/01/31"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= UNIT_SDATE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                    ~
                                    <asp:TextBox ID="UNIT_EDATE" runat="server" Width="15%" MaxLength="10" ToolTip="日期格式:99/01/31"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= UNIT_EDATE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">審核類型</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="PlanMode" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="1" Selected="True">審核中</asp:ListItem>
                                    <asp:ListItem Value="2">已通過</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td class="whitecol">
                                <div align="center">
                                    <asp:Button ID="btnQuery" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                </div>
                            </td>
                        </tr>
                    </table>
                    <table id="DataGridTable1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:Label ID="Label1" runat="server" Visible="False" CssClass="font">可點選班別名稱,查看計畫內容</asp:Label></td>
                        </tr>
                        <tr>
                            <td class="whitecol">
                                <asp:DataGrid ID="dgPlan" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <Columns>
                                        <asp:TemplateColumn HeaderText="編號">
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="labSEQNO" runat="server" Text=""></asp:Label>
                                                <asp:HiddenField ID="Hid_PLANID" runat="server" />
                                                <asp:HiddenField ID="Hid_COMIDNO" runat="server" />
                                                <asp:HiddenField ID="Hid_SEQNO" runat="server" />
                                                <asp:HiddenField ID="Hid_RIDV" runat="server" />
                                                <input id="KeyValue" type="hidden" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <%--<asp:BoundColumn Visible="False" DataField="PlanID" HeaderText="PlanID"></asp:BoundColumn>
                                        <asp:BoundColumn Visible="False" DataField="SeqNo" HeaderText="SeqNO"></asp:BoundColumn>--%>
                                        <asp:BoundColumn DataField="PlanYear" HeaderText="計畫年度">
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="班別名稱">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="LinkButton1" runat="server" ForeColor="Black"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="ComIDNO" HeaderText="統一編號">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="訓練單位">
                                            <HeaderStyle Width="15%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Address" HeaderText="公司地址">
                                            <HeaderStyle Width="15%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ContactName" HeaderText="聯絡人">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Phone" HeaderText="電話">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="審核">
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <HeaderTemplate>
                                                審核:
                                            <asp:DropDownList ID="SelectAll" runat="server">
                                                <asp:ListItem Value="==請選擇==">==請選擇==</asp:ListItem>
                                                <asp:ListItem Value="Y">通過</asp:ListItem>
                                                <asp:ListItem Value="M">請修正資料</asp:ListItem>
                                                <asp:ListItem Value="N">不通過</asp:ListItem>
                                            </asp:DropDownList>
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:DropDownList ID="AppliedResult" runat="server">
                                                    <asp:ListItem>==請選擇==</asp:ListItem>
                                                    <asp:ListItem Value="Y">通過</asp:ListItem>
                                                    <asp:ListItem Value="M">請修正資料</asp:ListItem>
                                                    <asp:ListItem Value="N">不通過</asp:ListItem>
                                                </asp:DropDownList>
                                                <asp:DropDownList ID="AppliedResult1" runat="server">
                                                    <asp:ListItem>==請選擇==</asp:ListItem>
                                                    <asp:ListItem Value="Y">通過</asp:ListItem>
                                                    <asp:ListItem Value="O">審核後修正</asp:ListItem>
                                                </asp:DropDownList>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="原因">
                                            <HeaderStyle Width="15%"></HeaderStyle>
                                            <ItemTemplate>
                                                <textarea id="Reason" name="VerReason" rows="5" runat="server" cols="20"></textarea>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="bntAdd" runat="server" Text="儲存" Enabled="False" CssClass="asp_button_M"></asp:Button></td>
                        </tr>
                    </table>
                    <table class="font" id="DataGridTable2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號" HeaderStyle-Width="5%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="AppliedDate" HeaderText="申請日期" DataFormatString="{0:d}" HeaderStyle-Width="15%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDate" HeaderText="訓練日期" DataFormatString="{0:d}" HeaderStyle-Width="15%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="FDDate" HeaderText="訓練迄日" DataFormatString="{0:d}" HeaderStyle-Width="15%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構" HeaderStyle-Width="15%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassName" HeaderText="班別名稱" HeaderStyle-Width="15%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="TransFlag" HeaderText="轉班" HeaderStyle-Width="15%"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="取消審核" HeaderStyle-Width="5%">
                                            <ItemTemplate>
                                                <asp:DropDownList ID="AppliedResult2" runat="server">
                                                    <asp:ListItem>==請選擇==</asp:ListItem>
                                                    <asp:ListItem Value="取消審核">取消審核</asp:ListItem>
                                                </asp:DropDownList>
                                                <input id="KeyValue" type="hidden" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <uc1:PageControler ID="PageControler2" runat="server"></uc1:PageControler>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="Button1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button></td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <input id="isBlack" type="hidden" name="isBlack" runat="server" />
        <input id="orgname" type="hidden" name="orgname" runat="server" />
        <asp:HiddenField ID="Hid_UserComIDNO" runat="server" />
        <asp:HiddenField ID="Hid_RID1" runat="server" />
    </form>
</body>
</html>
