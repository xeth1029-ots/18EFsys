<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_05_014_R.aspx.vb" Inherits="WDAIIP.TR_05_014_R" %>

 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>TR_05_013_R</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        function chkSearch() {
            var msg = '';
            if (document.form1.Syear.selectedIndex == 0) msg += '請選擇結訓年度\n';
            if (document.form1.ddlMonths.selectedIndex == 0) msg += '請選擇結訓月分\n';

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
            if (document.form1.STDate1.value !='') {
            if(!checkDate(document.form1.STDate1.value)) msg+='開訓期間 的起始日不是正確的日期格式\n';
            }
            if (document.form1.STDate2.value !='') {
            if(!checkDate(document.form1.STDate2.value)) msg+='開訓期間 的迄止日不是正確的日期格式\n';
            }
				
            if (document.form1.FTDate1.value !='') {
            if(!checkDate(document.form1.FTDate1.value)) msg+='結訓期間 的起始日不是正確的日期格式\n';
            }
            if (document.form1.FTDate2.value !='') {
            if(!checkDate(document.form1.FTDate2.value)) msg+='結訓期間 的迄止日不是正確的日期格式\n';
            }
				
            obj='TPlanID';
            num=getCheckBoxListValue(obj).length
            j=0;
            for(var i=1;i<num;i++){
            var mycheck=document.getElementById(obj+'_'+i);
            if (mycheck.checked) { j+=1; }
            }
            if(j==0) msg+='請選擇訓練計畫\n';
				
            obj='BudgetList';
            num=getCheckBoxListValue(obj).length
            j=0;
            for(var i=0;i<num;i++){
            var mycheck=document.getElementById(obj+'_'+i);
            if (mycheck.checked) { j+=1; }
            }
            if(j==0) msg+='請選擇預算來源\n';
            **/

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
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
            <tbody>
                <tr>
                    <td>
                        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <tr>
                                <td>
                                    <asp:Label ID="TitleLab1" runat="server"></asp:Label><asp:Label ID="TitleLab2" runat="server">
										首頁&gt;&gt;訓練與就業需求管理&gt;&gt;統計分析&gt;&gt;<FONT color="#990000">勞保勾稽查詢</FONT>
                                    </asp:Label>
                                </td>
                            </tr>
                        </table>
                        <table class="table_sch" id="Table2" cellspacing="1" cellpadding="1">
                            <tbody>
                                <tr>
                                    <td class="bluecol_need" width="100">轄區
                                    </td>
                                    <td class="whitecol">
                                        <asp:CheckBoxList ID="DistID" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="4">
                                        </asp:CheckBoxList>
                                        <input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server">
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">縣市
                                    </td>
                                    <td class="whitecol">
                                        <asp:CheckBoxList ID="CTID" runat="server" RepeatDirection="Horizontal" RepeatColumns="6" CssClass="font">
                                        </asp:CheckBoxList>
                                        <input id="CTIDHidden" type="hidden" value="0" name="CTIDHidden" runat="server">
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">訓練計畫
                                    </td>
                                    <td class="whitecol">
                                        <asp:CheckBoxList ID="TPlanID" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3">
                                        </asp:CheckBoxList>
                                        <input id="TPlanHidden" type="hidden" value="0" name="TPlanHidden" runat="server">
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">結訓年月
                                    </td>
                                    <td class="whitecol">
                                        <asp:DropDownList Style="z-index: 0" ID="Syear" runat="server">
                                        </asp:DropDownList>
                                        <asp:DropDownList Style="z-index: 0" ID="ddlMonths" runat="server">
                                        </asp:DropDownList>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">就業區間
                                    </td>
                                    <td class="whitecol">
                                        <asp:CheckBoxList ID="cbl_JOBMDATE_MM" runat="server" RepeatLayout="Flow" CssClass="font" RepeatDirection="Horizontal">
                                        </asp:CheckBoxList>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">資料來源
                                    </td>
                                    <td class="whitecol">
                                        <asp:RadioButtonList Style="z-index: 0" ID="rblJobMode" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                            <asp:ListItem Value="A">不區分</asp:ListItem>
                                            <asp:ListItem Value="1" Selected="True">系統</asp:ListItem>
                                            <asp:ListItem Value="2">人工</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </td>
                                </tr>
                            </tbody>
                        </table>
                        <p align="center">
                            <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label><asp:TextBox ID="TxtPageSize" runat="server" MaxLength="2" Width="23px">10</asp:TextBox>
                            <asp:Button ID="btnSearch" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>&nbsp;
                        <asp:Button ID="btnExport" runat="server" Text="匯出Excel" CssClass="asp_Export_M"></asp:Button>
                        </p>
                        <p align="center">
                            <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                        </p>
                        </p>
                    </td>
                </tr>
            </tbody>
        </table>
        <table id="ResultTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <div id="Div1" runat="server">
                        <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Width="100%" AllowPaging="True">
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
