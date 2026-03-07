<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_04_002.aspx.vb" Inherits="WDAIIP.TC_04_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>班級審核作業</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        $(document).ready(function () {
            //console.log("ready!");
            doChkDistType1();
        });

        function doChkDistType1() {
            var DT_val = $("input:radio[name='DistType']:checked").val();
            if (DT_val == undefined) { $("#trDist").hide(); $("#trOrg").show(); return; }
            //console.log("DistType value: " + DT_val);
            $("#trDist").hide();
            $("#trOrg").hide();
            if (DT_val == "1") { $("#trOrg").show(); }
            if (DT_val == "0") { $("#trDist").show(); }
        }

        //[全選／全不選]
        function SelectAll_J() {
            //debugger;
            var sub_SelectAll1 = $("#dgPlan").find("select[Name$=SelectAll1]");
            var v_SelectAll1 = "";
            if (sub_SelectAll1) { v_SelectAll1 = sub_SelectAll1.val(); }
            if (v_SelectAll1 == "") { return; }
            $('#dgPlan tr').each(function () {
                //debugger;
                var sub_AppliedResult1 = $(this).find("select[Name$=AppliedResult1]");
                if (sub_AppliedResult1) {
                    if (sub_AppliedResult1.prop("disabled") == false) { sub_AppliedResult1.val(v_SelectAll1); }
                }
            });
        }

        //選擇全部
        function SelectAll(obj, hidobj) {
            //debugger;
            var num = getCheckBoxListValue(obj).length;
            var myallcheck = document.getElementById(obj + '_' + 0);

            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
        }

        function SetDistType(obj1, obj2, obj3, obj3b) {
            //debugger;
            //DistType.Attributes("onclick") = "SetDistType('DistType','DistrictList','center','Org');"
            var Hid_radio1 = document.getElementById("Hid_radio1");
            var radio1 = document.getElementById(obj1).firstChild.checked;
            Hid_radio1.value = "N";
            if (radio1 == true) { Hid_radio1.value = "Y"; }
            document.getElementById(obj2).style.display = "none";//"inline";
            document.getElementById(obj3).style.display = "none";
            document.getElementById(obj3b).style.display = "none";
            if (radio1 == true) {
                document.getElementById(obj2).style.display = "";
            }
            else {
                //document.getElementById(obj2).style.display = "none";
                document.getElementById(obj3).style.display = "";//"inline";
                document.getElementById(obj3b).style.display = "";//"inline";
            }
        }


        //unction fnOpen(myvalue, planid,ComIDNO){
        //	if (myvalue=='Y') {
        //		win=window.open("TC_04_Trans.aspx?PlanID="+planid+"&ComIDNO="+ComIDNO,"","height=250,width=450,mentbar=yes,scrollbars=yes,resizable=yes");
        //		win.moveTo(50,60);
        //	}
        //}

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
                alert('請選擇要取消審核的班級');
                return false;
            }
            else {
                return confirm('您確定要取消審核以下班級?\n\n' + msg);
            }
        }

        /*
		function ChangeAll(j){
		var MyTable=document.getElementById('dgPlan');
		for(i=1;i<MyTable.rows.length;i++){
		MyTable.rows(i).cells(9).children(0).selectedIndex=j;
		}
		}
		*/

        function ChangeAll1(j) {
            //var cells_num = (11 - 1); //有一個欄位隱藏了。
            var cells_num = 11;
            var MyTable = document.getElementById('dgPlan');
            //alert('選擇'+j);
            for (i = 1; i < MyTable.rows.length; i++) {
                //alert(MyTable.rows(i).cells(9).children(0).disabled);
                if (MyTable.rows[i].cells[cells_num]) {
                    var cells = MyTable.rows[i].cells[cells_num];
                    if (!cells.children[0].disabled) {
                        cells.children[0].selectedIndex = j;
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;班級審核</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch" id="Table3">
                        <tr id="trDistType" runat="server">
                            <td class="bluecol" width="16%">搜尋型態</td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="DistType" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font" AutoPostBack="True">
                                    <asp:ListItem Value="0" Selected="True">依轄區</asp:ListItem>
                                    <asp:ListItem Value="1">依訓練機構</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="trDist" runat="server">
                            <td class="bluecol">轄區
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:CheckBoxList ID="DistrictList" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                                </asp:CheckBoxList>
                                <input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server" />
                            </td>
                        </tr>
                        <tr id="trOrg" runat="server">
                            <td class="bluecol">訓練機構
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="center" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox>
                                <input id="Org" type="button" value="..." name="Org" runat="server" />
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" /><br>
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td width="16%" class="bluecol">
                                <asp:Label ID="LabTMID" runat="server">訓練職類</asp:Label>
                            </td>
                            <td width="34%" class="whitecol">
                                <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Columns="30" Width="66%"></asp:TextBox>
                                <input id="btu_sel" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." name="btu_sel" runat="server">
                                <input id="TPlanid" type="hidden" name="TPlanid" runat="server">
                                <input id="trainValue" type="hidden" name="trainValue" runat="server">
                                <input id="jobValue" type="hidden" name="jobValue" runat="server">
                                <%--<asp:Button ID="Button2" runat="server" Visible="False" Text="95" CssClass="asp_button_M"></asp:Button>--%>
                            </td>
                            <td width="16%" class="bluecol">
                                <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label>
                            </td>
                            <td width="34%" class="whitecol">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="66%"></asp:TextBox>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server">
                                <input id="cjobValue" type="hidden" name="cjobValue" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">&nbsp;班別名稱
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="ClassName" runat="server" Columns="30" Width="66%"></asp:TextBox>
                            </td>
                            <td class="bluecol">期別
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="CyclType" runat="server" Columns="5" MaxLength="2" Width="40%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">申請日期</td>
                            <td class="whitecol">
                                <span runat="server">
                                    <asp:TextBox ID="UNIT_SDATE" runat="server" Width="35%" MaxLength="10" ToolTip="日期格式:yyyy/MM/dd"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= UNIT_SDATE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">~
        							<asp:TextBox ID="UNIT_EDATE" runat="server" Width="35%" MaxLength="10" ToolTip="日期格式:yyyy/MM/dd"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= UNIT_EDATE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                </span>
                            </td>
                            <td class="bluecol">開訓日期</td>
                            <td class="whitecol">
                                <span runat="server">
                                    <asp:TextBox ID="start_date" Width="35%" runat="server" MaxLength="10" ToolTip="日期格式:yyyy/MM/dd"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">～
		        					<asp:TextBox ID="end_date" Width="35%" runat="server" MaxLength="10" ToolTip="日期格式:yyyy/MM/dd"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">計畫範圍</td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="OrgKind2" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="A" Selected="True">不區分</asp:ListItem>
                                    <asp:ListItem Value="G">產業人才投資計畫</asp:ListItem>
                                    <asp:ListItem Value="W">提升勞工自主學習計畫</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <%--<tr id="tr_AppStage_TP28" runat="server"><td class="bluecol">申請階段</td><td class="whitecol" colspan="3"><asp:DropDownList ID="AppStage" runat="server"></asp:DropDownList></td></tr>--%>
                        <tr id="tr_AppStage_TP28" runat="server">
                            <td class="bluecol">申請階段</td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="AppStage2" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font"></asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">審核類型</td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="PlanMode" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" AutoPostBack="True">
                                    <asp:ListItem Value="S" Selected="True">審核中</asp:ListItem>
                                    <asp:ListItem Value="Y">已通過</asp:ListItem>
                                    <asp:ListItem Value="R">退件修正(含不通過的)</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                         <tr>
                            <td class="bluecol">&nbsp;檢送資料</td>
                            <td colspan="3" class="whitecol"><%--檢送資料-未檢送--%>
                                <asp:CheckBox ID="CB_DataNotSent_SCH" runat="server" Text="未檢送資料" ToolTip="(未勾選)排除未檢送資料" />
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch" id="TRA" runat="server" width="100%">
                        <tr>
                            <td class="bluecol" width="16%">已通過審核功能</td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="AdvanceMode" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="S" Selected="True">審核狀態</asp:ListItem>
                                    <asp:ListItem Value="C">取消審核</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                    <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td colspan="4" class="whitecol">
                                <div align="center">
                                    <asp:Button ID="btnQuery" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                </div>
                            </td>
                        </tr>
                    </table>
                    <table id="DataGridTable1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:Label ID="Label1" runat="server" CssClass="font" Visible="False">可點選班別名稱,查看計畫內容</asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">

                                <asp:DataGrid ID="dgPlan" runat="server" CssClass="font" Width="100%" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <Columns>
                                        <asp:TemplateColumn HeaderText="編號">
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="labNumber1" runat="server" Text="labNumber1"></asp:Label>
                                                <%--<asp:HiddenField ID="Hid_PCS" runat="server" />--%>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn Visible="False" DataField="PlanID" HeaderText="PlanID"></asp:BoundColumn>
                                        <asp:BoundColumn Visible="False" DataField="SeqNo" HeaderText="SeqNO"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="PlanYear" HeaderText="計畫年度">
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="AppliedDate" HeaderText="申請日期" DataFormatString="{0:d}">
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDate" HeaderText="訓練起日" DataFormatString="{0:d}">
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="FDDate" HeaderText="訓練迄日" DataFormatString="{0:d}">
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="班別名稱">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_RESULTBUTTON" runat="server"></asp:Label>
                                                <asp:LinkButton ID="LinkButton1" runat="server" ForeColor="Black"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="Point" HeaderText="課程種類">
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn Visible="False" DataField="ComIDNO" HeaderText="統一編號">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="訓練單位">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn Visible="False" DataField="Address" HeaderText="公司地址">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ContactName" HeaderText="聯絡人">
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Phone" HeaderText="電話">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="資格初審">
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <HeaderTemplate>
                                                <asp:Label ID="Label12" runat="server">資格複審</asp:Label>
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:Button ID="BtnEdit" runat="server" Text="編輯" CausesValidation="false" CommandName="Edit" CssClass="asp_button_M"></asp:Button>
                                                <asp:Button ID="BtnAdd" runat="server" Text="新增" CausesValidation="false" CommandName="Add" CssClass="asp_button_M"></asp:Button>
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                &nbsp;
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="審核">
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <HeaderTemplate>
                                                <asp:Label ID="Label15" runat="server">複審狀況</asp:Label><br>
                                                <asp:DropDownList ID="SelectAll1" runat="server">
                                                    <asp:ListItem>==請選擇==</asp:ListItem>
                                                    <asp:ListItem Value="Y">通過</asp:ListItem>
                                                    <asp:ListItem Value="N">不通過</asp:ListItem>
                                                    <asp:ListItem Value="R">退件修正</asp:ListItem>
                                                </asp:DropDownList>
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:DropDownList ID="AppliedResult1" runat="server">
                                                    <asp:ListItem>==請選擇==</asp:ListItem>
                                                    <asp:ListItem Value="Y">通過</asp:ListItem>
                                                    <asp:ListItem Value="N">不通過</asp:ListItem>
                                                    <asp:ListItem Value="R">退件修正</asp:ListItem>
                                                </asp:DropDownList>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="原因">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                            <ItemTemplate>
                                                &nbsp;
												<textarea id="Reason" style="width: 100%; height: 66px" name="VerReason" rows="3" cols="17" runat="server"></textarea>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>

                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <uc1:PageControler ID="Pagecontroler1" runat="server"></uc1:PageControler>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="bntAdd" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <table id="DataGridTable2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td class="whitecol">
                                <asp:DataGrid ID="DataGrid2" runat="server" CssClass="font" Width="100%" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號" HeaderStyle-Width="4%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="AppliedDate" HeaderText="申請日期" DataFormatString="{0:d}" HeaderStyle-Width="12%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDate" HeaderText="訓練起日" DataFormatString="{0:d}" HeaderStyle-Width="12%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="FDDate" HeaderText="訓練迄日" DataFormatString="{0:d}" HeaderStyle-Width="12%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構" HeaderStyle-Width="12%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="CLASSNAME2" HeaderText="班別名稱" HeaderStyle-Width="12%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="TransFlag" HeaderText="轉班" HeaderStyle-Width="12%"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="取消審核" HeaderStyle-Width="12%">
                                            <ItemTemplate>
                                                <asp:DropDownList ID="AppliedResult2" runat="server" Width="100%">
                                                    <asp:ListItem>==請選擇==</asp:ListItem>
                                                    <asp:ListItem Value="O">取消審核</asp:ListItem>
                                                </asp:DropDownList>
                                                <input id="KeyValue" type="hidden" runat="server">
                                                <input id="KPlanID" type="hidden" runat="server">
                                                <input id="KComIDNO" type="hidden" runat="server">
                                                <input id="KSeqNo" type="hidden" runat="server">
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
                                <asp:Button ID="Button1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="Hid_radio1" runat="server" />
    </form>
</body>
</html>
