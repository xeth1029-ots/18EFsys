<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_14_013.aspx.vb" Inherits="WDAIIP.SD_14_013" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>補助經費申請書</title>
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
    <script type="text/javascript">
        function GETvalue() {
            document.getElementById('Button5').click();
        }

        function ClearData() {
            document.getElementById('TMID1').value = '';
            document.getElementById('OCID1').value = '';
            document.getElementById('TMIDValue1').value = '';
            document.getElementById('OCIDValue1').value = '';
        }

        function CheckSearch() {
            var SOCIDValue = document.getElementById('SOCIDValue');
            var OCIDValue1 = document.getElementById('OCIDValue1');
            var IDNO = document.getElementById('IDNO');
            var Name = document.getElementById('Name');
            SOCIDValue.value = '';
            if (OCIDValue1.value == '' && IDNO.value == '' && Name.value == '') {
                //alert('至少要輸入一項條件');
                //return false;
            }
        }

        function choose_class() {
            var RIDValue = document.getElementById('RIDValue');
            openClass('../02/SD_02_ch.aspx?RID=' + RIDValue.value);
        }

        //function SelectItem(Flag, MyValue) {
        //    var SOCIDValue = document.getElementById('SOCIDValue');
        //    if (Flag) {
        //        if (SOCIDValue.value != '') { SOCIDValue.value += ','; }
        //        SOCIDValue.value += MyValue;
        //    }
        //    else {
        //        if (SOCIDValue.value.indexOf(',' + MyValue + ',') != -1)
        //            SOCIDValue.value = SOCIDValue.value.replace(',' + MyValue, '')
        //        else if (SOCIDValue.value.indexOf(',' + MyValue) != -1)
        //            SOCIDValue.value = SOCIDValue.value.replace(',' + MyValue, '')
        //        else if (SOCIDValue.value.indexOf(MyValue + ',') != -1)
        //            SOCIDValue.value = SOCIDValue.value.replace(MyValue + ',', '')
        //        else if (SOCIDValue.value.indexOf(MyValue) != -1)
        //            SOCIDValue.value = SOCIDValue.value.replace(MyValue, '')
        //    }
        //}

        //function SelectAll(Flag, idx) {
        //    var MyTable1 = document.getElementById('DataGrid1');
        //    var MyTable = MyTable1.rows[idx].cells[0].children[0].rows[1].cells[1].children[0];
        //    for (i = 1; i < MyTable.rows.length; i++) {
        //        if (MyTable.rows[i].cells[0].children[0].checked != Flag) {
        //            MyTable.rows[i].cells[0].children[0].checked = Flag;
        //            SelectItem(MyTable.rows[i].cells[0].children[0].checked, MyTable.rows[i].cells[0].children[0].value);
        //        }
        //    }
        //}

        //function SelectAll2(Flag) {
        //    var MyTable = document.getElementById('DataGrid1');
        //    for (i = 1; i < MyTable.rows.length; i++) {
        //        if (MyTable.rows[i].cells[0].children[0].checked != Flag) {
        //            MyTable.rows[i].cells[0].children[0].checked = Flag;
        //            SelectItem(MyTable.rows[i].cells[0].children[0].checked, MyTable.rows[i].cells[0].children[0].value);
        //        }
        //    }
        //}

        //function ShowDetail(obj, obj2) {
        //    if (document.getElementById(obj)) {
        //        if (document.getElementById(obj).style.display == 'none') {
        //            document.getElementById(obj).style.display = '';
        //            document.getElementById(obj2).src = '../../images/n02.gif';
        //        }
        //        else {
        //            document.getElementById(obj).style.display = 'none';
        //            document.getElementById(obj2).src = '../../images/n01.gif';
        //        }
        //    }
        //}

        function ChangeAll(obj) {
            var objLen = document.form1.length;
            for (var iCount = 0; iCount < objLen; iCount++) {
                if (document.form1.elements[iCount].type == "checkbox") {
                    document.form1.elements[iCount].checked = obj.checked;//true/false;
                }
            }
        }

    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;產學訓表單列印&gt;&gt;補助經費申請書</asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch" id="Table2">
                        <tr id="OrgTR" runat="server">
                            <td class="bluecol" style="width: 20%">訓練機構</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                                <asp:Button ID="Button5" Style="display: none" runat="server" Text="Button5"></asp:Button>
                                <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">職類/班別</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="button_b_Mini">
                                <input id="Button4" type="button" value="清除" name="Button4" runat="server" class="asp_button_M">
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                <br>
                                <%--<asp:CheckBox ID="Detail" runat="server" Text="查詢此班級學員的詳細歷程"></asp:CheckBox>--%>
                                <span id="HistoryList" style="position: absolute; display: none; left: 270px">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">開訓期間</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="STDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">～
                                <asp:TextBox ID="STDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            </td>
                        </tr>
                        <%-- A29ACD67218B121F132FF90ED282A0ABF36022D28B9BCE33CE0F3D703B7E1AD168D56A7294F05591E9D8960760ECF1196E2C5718FDDB8ED15EA9B28A
B114F84C2228A1FFFE62BAFD886239683492ECFD47282FFAAAC16AD35FCE650B9B100114EA894ADDC80C62A86FA41755AD4E2C2F021E4940D807F9CF
ACC0F69D35F543250553BE8AFCEEB1C9BA97518CE15C85A7F712FD843EFFB18050A0389C6571704E371C38A52626173DA9FB6D95ADE1756617A410C2
23EB7E64D51B27455C54418D2476E5C49D6A32267E8F2F1D589ADA073F5A91345B9B92B1BFDC7C5495D19C1EE74007CF19F6866AFFF97D8DFD43FB65
18093A4AB726932BB6C556420C800B67F7ECB6EA7077C6E64A9382CDC8FD9092A76329094F0E98F7A4FAC1CC06EF22AE60C48CF647BCA4BF8A4B2E11
7B5E57CCFE28910B3613E0103024313EC4987742839782180651EE6A50973938956AE1C6FB601CF3376078DB994526756339F15AC6D48D18041DEC51
41F34EA1CD5AE13AD21B81FED8AE63531BF93A55864B2F64--%>
                        <tr>
                            <td class="bluecol" style="width: 20%">身分證號碼</td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="IDNO" runat="server" Width="40%"></asp:TextBox></td>
                            <td class="bluecol" style="width: 20%">學員姓名</td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="Name" runat="server" Width="40%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">學員狀態</td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="Radio1" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="0">全部狀態</asp:ListItem>
                                    <asp:ListItem Value="1">已結訓且已申請經費</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="TRPlanPoint28" runat="server">
                            <td class="bluecol">計畫</td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="PlanPoint" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" CssClass="font">
                                    <asp:ListItem Value="1" Selected="True">產業人才投資計畫</asp:ListItem>
                                    <asp:ListItem Value="2">提升勞工自主學習計畫</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="tr_ddl_INQUIRY_S" runat="server">
                            <td class="bluecol_need">查詢原因</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="ddl_INQUIRY_Sch" runat="server"></asp:DropDownList>
                            </td>
                        </tr>
                    </table>
                    <div align="center" class="whitecol">
                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">40</asp:TextBox>
                        <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                    </div>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                    </div>
                    <%--<tr><td>點選箭頭可以展開該學員的參訓班級資料</td></tr>--%>
                    <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr id="trlimitedmax999" runat="server">
                            <td>列印資料筆數限定最多為999筆</td>
                        </tr>
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#f5f5f5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn HeaderText="選取">
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <HeaderTemplate>
                                                選取<input id="checkbox3" type="checkbox" runat="server" />
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <input id="checkbox2" type="checkbox" runat="server" />
                                                <asp:HiddenField ID="hf_SOCID" runat="server" />
                                                <asp:HiddenField ID="hf_SID" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="STUDID2" HeaderText="學號">
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="NAME" HeaderText="姓名">
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="IDNO" HeaderText="身分證號碼">
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="CLASSCNAME2" HeaderText="班名">
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDATE" HeaderText="訓練起日">
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="FTDATE" HeaderText="訓練迄日">
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundColumn>
                                        <%--72DDE4D70FE2883492160D833D382EDC2A5198B70717503A5FA406E78D60B2F81E7286AD407F04FFE71B351ADEA0ED37768232AF80E71DFA14F1632B
                                            DB3308E551DA5E21FCF6BD88A4B892FEE96C8C0C99DD568A621F74BEB4715742577FBE8D2B58A648E0F73207915B6E4066B0823C1140715F85163CAE
                                            89629DA52224938FA814A7DDD138217D3226A159106FD7AF3831512CE19485C90B8401B624E1449DA5B7C855B56171C908B3A59F3990D53FE0D23593
                                            DC10641E2D871FC7CFFF7677DBC47DA8F7134701012323F7408908A42757A13244D7E99BFA29FE2671CBDE4AA226AB038FD3AA1853206A6EDFF83514
                                            A1C25D2D6ABD3C52B6BE6EABE04AB65DB48C4AAF77A57CF62C8529C9C1906CB2C3B603C75938F1940E0CFBA48152C8A4060616B7CAB77850056E88AF
                                            E8C3C6F468C780DC7170C2059DAB6D200A20B4498F1395A2FC19E3A76BA938F653A239E5AE8DC8BF748CA931D492B3058EC78B9185C1B17504544A33
                                            AF27D8D3CD05D0DAE7388783CF8C0492DF917F01AD144DF75D45EEDE90073B2C5D9B869D7F98C8401DEE95D9B27BC2940C043BB1BA3AA70EC79856EB
                                            8CA79DEA7F38D0AD528E1826A03D171B095BBD02D0E3E8E9F5C640927917557401AACFBFBFFBDF77FA5BFE57D6A06F4A3296539AA1AC7702A283F8E1
                                            6E87084221A517C91B2989CC16F16CBF12DB43D9376C0DDDF91A7AC843A13FAB2D000DD497462DD09EFE6962E117433D1178549B6E3790EF7D026FC3
                                            D0396D196BC58D0D3A24BCF89CAF38251E4317AC7C8C171CF8AB05EE3C936B04F574D42757BCAEB5289958C2920DD1B856154D591656F426F8FD9AF7
                                            A2D8D902C8BC9C7D7CDB9DBC8A603AECCA26BEEAA065A7FA1F822E1E92A2F29B3B187E24F761EA20C8C7BC3C729A2B491A9270B5396D8D739E0D9845
                                            8B8CB9FD368E0EAC0338CCADD0FC7186A07F08EFB688EEA3DE63CCBA85450646F4B8B3919ACFF4EA5D5403D6349E089B08FA0AFE800C56AE5F054D7F
                                            4AE3503F7B22EE5C890DA5F29D3549C5464B47F5F21CCA3B01B6DE9C910947FBBAC082CB79DEA4459890034B1EFF752FF15B7EAE89782A13FE15FC38
                                            DF7A2B98D1F0ED42F2EC5B18BCD391B0F17AACA52A922A046EC18674C7C7CC5CF944FE6790C137EF1DE84AE1E1E2AEB5DF0FD035B9451716C3BD5E8A
                                            E2A8E935282140102E16B4FEB993827817CEFB47F1B019B0BB62DB86AE0D1EBAF073020D208EF72908FAC077CF65BB55D64C505110A0251151718E5B
                                            7F5EA435F18468792C17919742F3886E33D0DB5518CFE1B538CC920101511AA38997FDCC22142E332E60AB9BFE95A2E7C395A695DD4BDC6A4885A0D5
                                            87CA9720562D08CF465D88A5D1B85ED7C6F10EE0E95B43F498BD295AEE5071D32EBF1680CD9C4F3B3B6A779BE0E8BA9CE6255A88B6EECB635903C6A8
                                            8B047D26F65E56AA45496E23A8CE9527FA5062F0
                                        --%>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <%--<input id="Button3" type="button" value="列印" name="Button3" runat="server" class="asp_button_S" onclick="return Button3_onclick()">--%>
                                <asp:Button ID="BtnPrint3" runat="server" Text="列印" CssClass="asp_Export_M" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <input id="SOCIDValue" type="hidden" runat="server" />
        <input id="Years" type="hidden" name="Years" runat="server" />
        <input id="HidYears" type="hidden" name="HidYears" runat="server" />
        <input id="KindValue" type="hidden" name="KindValue" runat="server" />
        <%--<input id="hidgTestflag" type="hidden" name="hidgTestflag" runat="server">--%>
    </form>
</body>
</html>
