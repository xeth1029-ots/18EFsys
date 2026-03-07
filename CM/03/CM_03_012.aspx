<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CM_03_012.aspx.vb" Inherits="WDAIIP.CM_03_012" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>屆退官兵人數統計表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        //function ClearData(){
        //document.getElementById('PlanID').value='';
        //document.getElementById('center').value='';
        //document.getElementById('RIDValue').value='';
        //for(var i=document.form1.OCID.options.length-1;i>=0;i--){
        //	document.form1.OCID.options[i]=null;
        //}
        //document.getElementById('OCID').style.display='none';
        //document.getElementById('msg').innerHTML='請先選擇機構';
        //}

        //檢查列印條件為
        function CheckPrint() {

            var msg = '';

            if (document.form1.STDate1.value != '') {
                if (!checkDate(document.form1.STDate1.value)) msg += '[開訓區間的起日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }
            if (document.form1.FTDate1.value != '') {
                if (!checkDate(document.form1.FTDate1.value)) msg += '[結訓區間的起日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }


            if (document.form1.STDate2.value != '') {
                if (!checkDate(document.form1.STDate2.value)) msg += '[開訓區間的迄日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }

            if (document.form1.FTDate2.value != '') {
                if (!checkDate(document.form1.FTDate2.value)) msg += '[結訓區間的迄日]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';

            }

            if (document.form1.STDate2.value != '' && document.form1.STDate1.value != '' && document.form1.STDate2.value < document.form1.STDate1.value)
            { msg += '[開訓區間的迄日]必需大於[開訓區間的起日]\n'; }
            if (document.form1.FTDate2.value != '' && document.form1.FTDate1.value != '' && document.form1.FTDate2.value < document.form1.FTDate1.value)
            { msg += '[結訓區間的迄日]必需大於[結訓區間的起日]\n'; }


            var Identity1 = getCheckBoxListValue('Identity');
            var DistID1 = getCheckBoxListValue('DistID');
            var TPlanID1 = getCheckBoxListValue('TPlanID');

            if (parseInt(DistID1) == 0)
            { msg += '請選擇轄區\n'; }
            if (parseInt(TPlanID1) == 0)
            { msg += '請選擇計畫\n'; }

            if (msg != '') {
                alert(msg);
                return false;
            }
        }

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
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">
                    首頁&gt;&gt;訓練需求管理&gt;&gt;統計分析&gt;&gt;屆退官兵人數統計表 
                    </asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch" id="Table2" runat="server" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" width="20%">年度
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="Syear" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">開訓區間
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="STDate1" runat="server" Columns="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                ~<asp:TextBox ID="STDate2" runat="server" Columns="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">結訓區間
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="FTDate1" runat="server" Columns="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                ~<asp:TextBox ID="FTDate2" runat="server" Columns="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">轄區
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="DistID" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                </asp:CheckBoxList>
                                <input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">訓練計畫
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="TPlanID" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3" CellPadding="0" CellSpacing="0">
                                </asp:CheckBoxList>
                                <input id="TPlanHidden" type="hidden" value="0" name="TPlanHidden" runat="server">
                            </td>
                        </tr>
                        <%--<TR>
									<TD class="CM_TD1">&nbsp;&nbsp;&nbsp; 訓練機構</TD>
									<TD class="CM_TD2"><asp:textbox id="center" runat="server" Width="310px"></asp:textbox><INPUT id="Button2" type="button" value="..." name="Button2" runat="server"><INPUT id="RIDValue" type="hidden" name="RIDValue" runat="server"><INPUT id="PlanID" type="hidden" name="PlanID" runat="server">
										<asp:button id="Button3" runat="server" Text="查詢班級"></asp:button>(勾選班級後會省略[年度]、[開訓區間]、[結訓區間]的條件)</TD>
								</TR>
								<TR>
									<TD class="CM_TD1">&nbsp;&nbsp;&nbsp; 班別</TD>
									<TD class="CM_TD2"><asp:listbox id="OCID" runat="server" Width="300px" SelectionMode="Multiple" Rows="6"></asp:listbox><asp:label id="msg" runat="server" ForeColor="Red"></asp:label>(按Ctrl可以複選班級)</TD>
								</TR>
								<TR id="IdentityTR" runat="server">
									<TD class="CM_TD1"><FONT face="新細明體">&nbsp;&nbsp;&nbsp; 身分別&nbsp;</FONT></TD>
									<TD class="CM_TD2"><asp:checkboxlist id="Identity" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow"
											RepeatColumns="3"></asp:checkboxlist><INPUT id="Identity_List" type="hidden" value="0" name="Identity_List" runat="server">
									</TD>
								</TR>--%>
                    </table>
                    <p align="center">
                        <asp:Button ID="Print" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                    </p>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
