<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_010_R.aspx.vb" Inherits="WDAIIP.SD_15_010" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>檢閱審核結果列表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        //選擇全部
        function SelectAll(obj, hidobj) {
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

        //unction fnOpen(myvalue, planid,ComIDNO){
        //	if (myvalue=='Y') {
        //		win=window.open("TC_04_Trans.aspx?PlanID="+planid+"&ComIDNO="+ComIDNO,"","height=250,width=450,mentbar=yes,scrollbars=yes,resizable=yes");
        //		win.moveTo(50,60);
        //	}
        //}

        /*	function SavaData(){
		var MyTable=document.getElementById('DataGrid2');
		var msg='';
		for(i=1;i<MyTable.rows.length;i++){
		var Result=MyTable.rows(i).cells(7).children(0);
		if (Result.selectedIndex!=0){
		msg+=MyTable.rows(i).cells(5).innerHTML+'\n';
		}
		}
				
		if(msg==''){
		alert('請選擇要取消審核的班級')
		return false;
		}
		else{
		return confirm('您確定要取消審核以下班級?\n\n'+msg);
		}
		}
		*/
        /*
		function ChangeAll(j){
		var MyTable=document.getElementById('dgPlan');
		for(i=1;i<MyTable.rows.length;i++){
		MyTable.rows(i).cells(9).children(0).selectedIndex=j;
		}
		}
         function ChangeAll1(j) {
            var MyTable = document.getElementById('dgPlan');
            for (i = 1; i < MyTable.rows.length; i++) {
                //alert(MyTable.rows(i).cells(9).children(0).disabled);
                if (!MyTable.rows(i).cells(10).children(0).disabled) {
                    MyTable.rows(i).cells(10).children(0).selectedIndex = j;
                }
            }
        }
		*/
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">
					首頁&gt;&gt;學員動態管理&gt;&gt;產學訓統計表&gt;&gt;<FONT color="#800000">檢閱審核結果列表</FONT>
                    </asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>

                    <table class="table_sch" id="Table3">
                        <tr id="Tr1" runat="server">
                            <td class="bluecol" width="20%">搜尋型態 </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="DistType" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                                    <asp:ListItem Value="0" Selected="True">依轄區</asp:ListItem>
                                    <asp:ListItem Value="1">依訓練機構</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="Dist" runat="server">
                            <td class="bluecol">轄區 </td>
                            <td class="whitecol" colspan="3">
                                <asp:CheckBoxList ID="DistrictList" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font"></asp:CheckBoxList>
                                <input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練機構 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="410px"></asp:TextBox>
                                <input id="Org" type="button" value="..." name="Org" runat="server" class="button_b_Mini">
                                <input id="RIDValue" style="width: 32px; height: 22px" type="hidden" name="RIDValue" runat="server">
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="310px">
                                    </asp:Table>
                                </span></td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                <asp:Label ID="LabTMID" runat="server">訓練業別</asp:Label>
                            </td>
                            <td class="whitecol" colspan="3"><font face="新細明體">
                                <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Width="210px"></asp:TextBox><input id="btu_sel" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." name="btu_sel" runat="server" class="button_b_Mini">
                                <input id="TPlanid" style="width: 27px; height: 22px" type="hidden" name="TPlanid" runat="server"><input id="trainValue" style="width: 43px; height: 22px" type="hidden" name="trainValue" runat="server"><input id="jobValue" style="width: 43px; height: 22px" type="hidden" name="jobValue" runat="server">
                                <asp:Button ID="Button2" runat="server" Text="95" CssClass="asp_button_Mini"></asp:Button></font> </td>
                        </tr>
                        <tr>
                            <td class="bluecol">班別名稱 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="ClassName" runat="server" Width="200px"></asp:TextBox>
                            </td>
                            <td class="bluecol">期別 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="CyclType" runat="server" MaxLength="2" Columns="5"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">申請日期 </td>
                            <td class="whitecol" runat="server">
                                <asp:TextBox ID="UNIT_SDATE" runat="server" Width="70px" MaxLength="8" ToolTip="日期格式:99/01/31"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= UNIT_SDATE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">~
							<asp:TextBox ID="UNIT_EDATE" runat="server" Width="70px" MaxLength="8" ToolTip="日期格式:99/01/31"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= UNIT_EDATE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            </td>
                            <td class="bluecol" id="td5" runat="server">開訓日期 </td>
                            <td class="whitecol" runat="server">
                                <asp:TextBox ID="start_date" onfocus="this.blur()" Width="80" runat="server"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">～
							<asp:TextBox ID="end_date" onfocus="this.blur()" Width="80" runat="server"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            </td>
                        </tr>
                        <tr id="tr_AppStage_TP28" runat="server">
                            <td class="bluecol">申請階段 </td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="AppStage" runat="server" Width="100px">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr id="trPlanKind" runat="server">
                            <td class="bluecol">計畫範圍 </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="SearchPlan" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="A">不區分</asp:ListItem>
                                    <asp:ListItem Value="G" Selected="True">產業人才投資計畫</asp:ListItem>
                                    <asp:ListItem Value="W">提升勞工自主學習計畫</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="trPackageType" runat="server">
                            <td class="bluecol">包班種類 </td>
                            <td class="whitecol" colspan="3">
                                <%--<asp:ListItem Value="1">非包班</asp:ListItem>--%>
                                <asp:RadioButtonList ID="PackageType" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="A">全部</asp:ListItem>
                                    <asp:ListItem Value="2">企業包班</asp:ListItem>
                                    <asp:ListItem Value="3">聯合企業包班</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4" align="center">
                                <asp:Button ID="Print" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>

                    </table>

                </td>
            </tr>
        </table>
    </form>
</body>
</html>
