<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_009.aspx.vb" Inherits="WDAIIP.SD_15_009" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_15_009</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function OpenOrg(vTPlanID) {
            if (document.getElementById('DistID').selectedIndex == 0) {
                alert('請先選擇轄區');
                return false;
            }
            else {
                wopen('../../common/MainOrg.aspx?DistID=' + document.getElementById('DistID').value + '&TPlanID=' + vTPlanID, '', 400, 400, 'yes');
            }
        }

        function choose_class() {
            document.getElementById('OCID1').value = '';
            document.getElementById('TMID1').value = '';
            document.getElementById('OCIDValue1').value = '';
            document.getElementById('TMIDValue1').value = '';

            openClass('../02/SD_02_ch.aspx?&RID=' + document.getElementById('RIDValue').value);
        }

        /*			
        function CheckPrint(){
        //var STDate1=document.getElementById('STDate1').value;
        //var STDate2=document.getElementById('STDate2').value;
        //var DistID=document.getElementById('DistID').value;
        //var PlanID=document.getElementById('PlanID').value;
        var RID=document.getElementById('RIDValue').value;
				
        var msg='';
        //if(!checkDate(STDate1) && STDate1!='') msg+='開訓起始日期必須為正確日期格式\n';
        //if(!checkDate(STDate2) && STDate2!='') msg+='開訓結束日期必須為正確日期格式\n';
				
        if(msg!=''){
        alert(msg);
        return false;
        }
				
        //openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=BussinessTrain&filename=SD_15_009&path=TIMS&STDate1='+STDate1+'&STDate2='+STDate2+'&DistID='+DistID+'&PlanID='+PlanID+'&RID='+RID+'&OCID='+document.getElementById('OCIDValue1').value,'','');
        }
        */

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

        //openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=BussinessTrain&filename=SD_15_009&path=TIMS&Years='+Years+'&DistID='+DistID+'&CTID'+CTID+'&JobID'+JobID+'&CCID'+CCID+'&Orgname='Orgname);

    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">

        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">
						首頁&gt;&gt;學員動態管理&gt;&gt;產學訓統計表&gt;&gt;<font color="#800000">各類課程明細表</font>
                    </asp:Label>
                </td>
            </tr>
        </table>

        <table class="table_nw" id="FrameTable3" width="100%">
            <tr>
                <td class="bluecol" width="20%">年度
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="yearlist" runat="server">
                    </asp:DropDownList>
                    <asp:RequiredFieldValidator ID="MustYear" runat="server" CssClass="font" ControlToValidate="yearlist" Display="Dynamic" ErrorMessage="請選擇年度"></asp:RequiredFieldValidator></FONT></FONT>
                </td>
            </tr>
            <tr>
                <td class="bluecol">轄區
                </td>
                <td class="whitecol">
                    <font face="新細明體">
                        <asp:DropDownList ID="DistID" runat="server">
                        </asp:DropDownList>
                    </font>
                </td>
            </tr>
            <tr>
                <td class="bluecol">訓練機構
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="410px"></asp:TextBox>
                    <input id="Button1" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                    <input id="PlanID" type="hidden" name="PlanID" runat="server">
                    <input id="RIDValue" type="hidden" name="Hidden2" runat="server">
                    <span id="HistoryList2" style="position: absolute; display: none">
                        <asp:Table ID="HistoryRID" runat="server" Width="310px">
                        </asp:Table>
                    </span>
                </td>
            </tr>
            <!--
				<TR align="center">
					<TD class="SD_TD1"><FONT class="font" face="新細明體" color="#ffffff" size="2">轄區</FONT></TD>
					<TD class="SD_TD2">
						<TABLE id="Table2" style="WIDTH: 100%; HEIGHT: 52px" cellSpacing="1" cellPadding="1" width="536"
							border="0">
							<TR>
							</TR>
						</TABLE>
						<asp:checkboxlist id="DistrictList" runat="server" CssClass="font" Height="11px" Width="512px" RepeatDirection="Horizontal"></asp:checkboxlist><INPUT id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server"></TD>
				</TR>
				-->
            <tr>
                <td class="bluecol">縣市
                </td>
                <td class="whitecol">
                    <asp:CheckBoxList ID="CityList" runat="server" CssClass="font" RepeatColumns="8" RepeatDirection="Horizontal">
                    </asp:CheckBoxList>
                    <input id="CityHidden" type="hidden" value="0" name="CityHidden" runat="server">
                </td>
            </tr>
            <!--
				<TR>
					<TD class="SD_TD1"><FONT color="#ffffff">機構名稱</FONT></TD>
					<TD class="SD_TD2"><asp:textbox id="OrgName" runat="server"></asp:textbox></TD>
				</TR>
				-->
            <tr>
                <td class="bluecol">課程類別
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="ClassCate" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">訓練行業別</td>
                <td class="whitecol">
                    <asp:DropDownList ID="JobID" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr id="tr_AppStage_TP28" runat="server">
                <td class="bluecol">申請階段</td>
                <td class="whitecol">
                    <asp:DropDownList ID="AppStage" runat="server" Width="100px">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr id="trPlanKind" runat="server">
                <td class="bluecol">計畫範圍
                </td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="SearchPlan" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                        <asp:ListItem Value="G" Selected="True">產業人才投資計畫</asp:ListItem>
                        <asp:ListItem Value="W">提升勞工自主學習計畫</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr id="trPackageType" runat="server">
                <td class="bluecol">包班種類
                </td>
                <td class="whitecol" colspan="4">
                    <asp:RadioButtonList ID="PackageType" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                        <asp:ListItem Value="A" Selected="True">全部</asp:ListItem>
                        <%--<asp:ListItem Value="1">非包班</asp:ListItem>--%>
                        <asp:ListItem Value="2">企業包班</asp:ListItem>
                        <asp:ListItem Value="3">聯合企業包班</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">課程審核狀況
                </td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="rdlResult" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                        <asp:ListItem Value="A" Selected="True">不拘</asp:ListItem>
                        <asp:ListItem Value="Y">通過</asp:ListItem>
                        <asp:ListItem Value="N">不通過</asp:ListItem>
                        <asp:ListItem Value="T">其他(排除通過及不通過的其他狀態)</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td align="center" colspan="2">
                    <input id="Button3" type="button" value="列印" runat="server" class="asp_Export_M">
                </td>
            </tr>
        </table>

        <input id="Years" type="hidden" name="Years" runat="server">
    </form>
</body>
</html>
