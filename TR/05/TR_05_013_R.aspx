<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_05_013_R.aspx.vb" Inherits="WDAIIP.TR_05_013_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>ECFA執行數據統計表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <%-- <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>--%>
    <script type="text/javascript" language="javascript" src="../../js/date-picker2.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        //決定date-picker元件使用的是西元年or民國年，by:20181022
        <%--
        var bl_rocYear = "<%=ConfigurationManager.AppSettings("REPLACE2ROC_YEARS") %>";
        var scriptTag = document.createElement('script');
        var jsPath = (bl_rocYear == "Y" ? "../../js/date-picker2.js" : "../../js/date-picker.js");
        scriptTag.src = jsPath;
        document.head.appendChild(scriptTag);
        --%>

        function chkSearch() {
            var msg = '';

            /**
			var obj='DistID';
			var num=getCheckBoxListValue(obj).length
			var j=0;
			document.form1.hidDistID.value="";
			debugger;
			for(var i=1;i<num;i++){
			var mycheck=document.getElementById(obj+'_'+i);
			if (mycheck.checked) {
			if(document.form1.hidDistID.value!="") document.form1.hidDistID.value +=","
			document.form1.hidDistID.value +="'" +mycheck.value+"'"
			}
			//if (mycheck.checked) { j+=1; }
			}
			//var DistID=getRadioValue(document.getElementsByName('DistID'));
			//if(document.form1.DistID.selectedIndex==0) msg+='請選擇轄區中心\n';
			//if(DistID=='') msg+='請選擇轄區中心\n';
			//if(j==0) msg+='請選擇轄區中心\n';
			if(document.form1.hidDistID.value=="") msg+='請選擇轄區中心\n';
			**/

            //if (document.form1.Syear.selectedIndex == 0) msg += '請選擇年度\n';

            //checkRocDate
            function checkDate(dateValue) { return checkRocDate(dateValue); }

            if (document.form1.STDate1.value != '') {
                if (!checkDate(document.form1.STDate1.value)) msg += '開訓期間 的起始日不是正確的日期格式\n';
            }
            if (document.form1.STDate2.value != '') {
                if (!checkDate(document.form1.STDate2.value)) msg += '開訓期間 的迄止日不是正確的日期格式\n';
            }
            if (document.form1.FTDate1.value != '') {
                if (!checkDate(document.form1.FTDate1.value)) msg += '結訓期間 的起始日不是正確的日期格式\n';
            }
            if (document.form1.FTDate2.value != '') {
                if (!checkDate(document.form1.FTDate2.value)) msg += '結訓期間 的迄止日不是正確的日期格式\n';
            }

            var vSTDate1 = document.getElementById("STDate1").value;
            var vSTDate2 = document.getElementById("STDate2").value;
            var vFTDate1 = document.getElementById("FTDate1").value;
            var vFTDate2 = document.getElementById("FTDate2").value;
            if (document.form1.Syear.selectedIndex == 0 && vSTDate1 == "" && vSTDate2 == "" && vFTDate1 == "" && vFTDate2 == "") msg += '請選擇年度或開結訓日期\n';

            obj = 'TPlanID';
            num = getCheckBoxListValue(obj).length
            j = 0;
            for (var i = 1; i < num; i++) {
                var mycheck = document.getElementById(obj + '_' + i);
                if (mycheck.checked) { j += 1; }
            }
            if (j == 0) msg += '請選擇訓練計畫\n';
            obj = 'BudgetList';
            num = getCheckBoxListValue(obj).length
            j = 0;
            for (var i = 0; i < num; i++) {
                var mycheck = document.getElementById(obj + '_' + i);
                if (mycheck.checked) { j += 1; }
            }
            if (j == 0) msg += '請選擇預算來源\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        //選擇全部
        function SelectAll(obj, hidobj) {
            var num = getCheckBoxListValue(obj).length; //長度
            var myallcheck = document.getElementById(obj + '_' + 0); //第1個
            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
            else {
                for (var i = 1; i < num; i++) {
                    if ('0' == getCheckBoxListValue(obj).charAt(i)) {
                        document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(i);
                        var mycheck = document.getElementById(obj + '_' + i);
                        myallcheck.checked = mycheck.checked;
                        break;
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;統計分析&gt;&gt;ECFA執行數據統計表</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tbody>
                <tr>
                    <td>
                        <table class="table_sch" id="Table2" cellspacing="1" cellpadding="1">
                            <tbody>
                                <tr>
                                    <td class="bluecol" style="width: 20%">年度</td>
                                    <td class="whitecol">
                                        <asp:DropDownList ID="Syear" runat="server"></asp:DropDownList></td>
                                </tr>
                                <%--<tr>
                                    <td class="bluecol_need">轄區分署</td>
                                    <td class="whitecol"><asp:DropDownList ID="DistID" runat="server"></asp:DropDownList></td>
                                </tr>--%>

                                <tr>
                                    <td class="bluecol_need">轄區分署</td>
                                    <td class="whitecol">
                                        <asp:CheckBoxList ID="CBLDISTID" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="4">
                                        </asp:CheckBoxList>
                                        <input id="HidCBLDISTID" type="hidden" value="0" name="HidCBLDISTID" runat="server">
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol">開訓期間</td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="STDate1" runat="server" Columns="10" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar2('STDate1','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                        <font color="#000000">～</font>
                                        <asp:TextBox ID="STDate2" runat="server" Columns="10" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar2('STDate2','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol">結訓期間</td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="FTDate1" runat="server" Columns="10" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar2('FTDate1','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                        <font color="#000000">～</font>
                                        <asp:TextBox ID="FTDate2" runat="server" Columns="10" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar2('FTDate2','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">訓練計畫</td>
                                    <td class="whitecol">
                                        <asp:CheckBoxList ID="TPlanID" runat="server" CellPadding="0" CellSpacing="0" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3"></asp:CheckBoxList>
                                        <input id="TPlanHidden" type="hidden" value="0" name="TPlanHidden" runat="server">
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">預算來源</td>
                                    <td class="whitecol">
                                        <asp:CheckBoxList ID="BudgetList" runat="server" RepeatLayout="Flow" CssClass="font" RepeatDirection="Horizontal" Enabled="False"></asp:CheckBoxList></td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">查詢方式</td>
                                    <td class="whitecol">
                                        <asp:RadioButtonList ID="rblSearchStyle" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                            <asp:ListItem Value="1" Selected="True">統計資料</asp:ListItem>
                                            <asp:ListItem Value="2">明細資料</asp:ListItem>
                                            <asp:ListItem Value="3">班級別</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol">匯出檔案格式</td>
                                    <td class="whitecol">
                                        <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                                            <asp:ListItem Value="ODS">ODS</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </td>
                                </tr>
                            </tbody>
                        </table>
                        <div align="center" class="whitecol">
                            <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                            <asp:TextBox ID="TxtPageSize" runat="server" MaxLength="2" Width="6%">10</asp:TextBox>
                            <asp:Button ID="btnSearch" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                            <asp:Button ID="btnExport" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                        </div>
                        <div align="center">
                            <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                        </div>
                    </td>
                </tr>
            </tbody>
        </table>
        <table id="ResultTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <div id="Div1" runat="server">
                        <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Width="100%" AllowPaging="True" CellPadding="8">
                            <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                            <PagerStyle Visible="False"></PagerStyle>
                        </asp:DataGrid>
                    </div>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
