<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_016_add.aspx.vb" Inherits="WDAIIP.TC_01_016_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訓練機構計畫調動</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function SetLastYearExeRate() {
            //debugger;
            //obj1=document.getElementById("txtLastYearExeRate");
            obj2 = document.form1.txtLastYearExeRate;
            obj3 = document.form1.ExeRate;
            if (obj3) {
                if (document.getElementById("LastYearExeRate").children[0].checked) {
                    obj2.disabled = true;
                    obj3.disabled = true;
                    //obj1.style.display='none';
                }
                else {
                    if (obj2.value == '') { obj2.value = 0; }
                    obj2.disabled = false;
                    obj3.disabled = false;
                    //obj1.style.display='inline';
                }
            }
            else {
                if (document.getElementById("LastYearExeRate").children[0].checked) {
                    obj2.disabled = true;
                    //obj1.style.display='none';
                }
                else {
                    if (obj2.value == '') { obj2.value = 0; }
                    obj2.disabled = false;
                }
            }
        }

        function ChangeMode(num) {
            var div1 = $("#div1");
            var div2 = $("#div2");
            div1.removeClass();
            div2.removeClass();
            if (document.getElementById('Table1') && document.getElementById('HistoryTable')) {
                if (num == 1) {
                    //document.getElementById('Table1').style.display = 'inline';
                    div1.addClass("active");
                    document.getElementById('Table1').style.display = '';
                    document.getElementById('HistoryTable').style.display = 'none';
                }
                if (num == 2) {
                    div2.addClass("active");
                    document.getElementById('Table1').style.display = 'none';
                    //document.getElementById('HistoryTable').style.display = 'inline';
                    document.getElementById('HistoryTable').style.display = '';
                }
                if (document.body) { window.scroll(0, document.body.scrollHeight); }
                if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
            }
        }

        function CheckAll() {
            var Cst_cellsNum = 1;
            var mytable = document.getElementById('DataGrid2');
            for (var i = 1; i < mytable.rows.length; i++) {
                //設定 DataGrid2中的 checkbox 與 Choose1的 checkbox 相同
                var mycheck = mytable.rows[i].cells[Cst_cellsNum].children[0];
                if (mycheck) {
                    if (mycheck.disabled == false)
                        mycheck.checked = document.form1.Choose1.checked;
                }
            }
        }
    </script>
    <%-- <style type="text/css">
        .auto-style1 { color: Black; text-align: right; padding: 4px 6px; background-color: #f1f9fc; border-right: 3px solid #49cbef; height: 45px; }
        .auto-style2 { color: #333333; padding: 4px; height: 45px; }
    </style>--%>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <%--
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;基本資料設定&gt;&gt;訓練機構設定-</asp:Label>
                    <font color="#990000"><asp:Label ID="lblProecessType" runat="server"></asp:Label></font>
                </td>
            </tr>
        </table>
        --%>
        <font color="#990000">
            <asp:Label ID="lblProecessType" runat="server" Visible="false"></asp:Label></font>
        <input id="PlanIDValue" type="hidden" name="PlanIDValue" runat="server">
        <input id="Re_ID" type="hidden" name="Re_ID" runat="server">
        <input id="OrgIDValue" type="hidden" name="OrgIDValue" runat="server">
        <table class="font" id="Tableform" cellspacing="0" cellpadding="0" border="0" runat="server" width="100%">
            <tr>
                <td align="center" colspan="4">
                    <table class="table_nw" width="100%">
                        <tr>
                            <td class="bluecol" width="20%">訓練機構調動至</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="BtnOrg" type="button" value="..." name="BtnOrg" runat="server" class="button_b_Mini">
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                <input id="btnPlanSearch" style="display: none" type="button" value="搜尋" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">計畫調動至</td>
                            <td colspan="3" class="whitecol" width="80%">
                                <asp:DropDownList ID="OrgPlanNameList" runat="server"></asp:DropDownList></td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td align="center" colspan="4" width="100%" class="whitecol">
                    <asp:Button ID="bt_save" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                    <asp:Button ID="Button1" runat="server" Text="回上一頁" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
            <tr>
                <td>
                    <br />
                    <table class="font" id="MenuTable" style="cursor: pointer" cellspacing="0" cellpadding="0" border="0" runat="server" width="50%">
                        <tr class="newlink newlink-blue">
                            <%--<td onclick="ChangeMode(1);" width="1%" background="../../images/BookMark_01.gif"><font size="2"></font></td>--%>
                            <td onclick="ChangeMode(1);" align="center" id="div1">基本資料</td>
                            <%--<td onclick="ChangeMode(1);" width="1%" background="../../images/BookMark_03.gif"><font size="2"></font></td>--%>
                            <%--<td onclick="ChangeMode(2);" width="1%" background="../../images/BookMark_01.gif"> <font size="2"></font></td>--%>
                            <td onclick="ChangeMode(2);" align="center" id="div2">辦訓記錄</td>
                            <%--<td onclick="ChangeMode(2);" width="1%" background="../../images/BookMark_03.gif"><font size="2"></font></td>--%>
                            <%--<td width="64%"></td>--%>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table class="table_nw" id="Table1" width="100%" runat="server">
                        <tr>
                            <td class="table_title" colspan="4">
                                <center>訓練機構共同資料<asp:Label ID="GWOrgKind" runat="server"></asp:Label></center>
                            </td>
                        </tr>
                        <tr>
                            <%--<td class="bluecol" width="20%">轄區中心</td>--%>
                            <td class="bluecol" width="20%">轄區分署</td>
                            <td colspan="3" class="whitecol" width="80%">
                                <asp:DropDownList ID="DistrictList" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td id="td2" runat="server" class="bluecol" width="20%">機構名稱全銜</td>
                            <td colspan="3" class="whitecol" width="70%">
                                <asp:TextBox ID="TBtitle" runat="server" Width="60%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">機構別</td>
                            <td colspan="3" class="whitecol" width="80%">
                                <asp:DropDownList ID="OrgKindList" runat="server" CssClass="font"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">統一編號</td>
                            <td class="whitecol" width="30%">
                                <asp:TextBox ID="TBID" runat="server" MaxLength="10" Width="80%"></asp:TextBox></td>
                            <td id="Td3" runat="server" class="bluecol" width="20%">立案證號</td>
                            <td class="whitecol" width="30%">
                                <asp:TextBox ID="TBseqno" runat="server" Width="80%"></asp:TextBox></td>
                        </tr>
                        <tr id="TPlanID28A" runat="server">
                            <td class="bluecol" width="20%">
                                <asp:Label ID="LabLastYear" runat="server"></asp:Label></td>
                            <td class="whitecol" width="30%">
                                <asp:RadioButtonList ID="LastYearExeRate" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="-1" Selected="True">否</asp:ListItem>
                                    <asp:ListItem Value="1">是</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                            <td class="bluecol" width="20%">
                                <asp:Label ID="LabLastYear2" runat="server"></asp:Label></td>
                            <td class="whitecol" width="30%">
                                <asp:TextBox ID="txtLastYearExeRate" runat="server" Width="28%"></asp:TextBox>%
                                <input id="LastYear" style="width: 10%" type="hidden" name="LastYear" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4" class="table_title" width="100%">
                                <center>訓練機構承辦人資料</center>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">階層</td>
                            <td colspan="3" class="whitecol" width="80%">
                                <asp:DropDownList ID="level_list" runat="server" AutoPostBack="True">
                                    <asp:ListItem Value="2">委訓</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">隸屬機構</td>
                            <td colspan="3" class="whitecol" width="80%">
                                <asp:TextBox ID="TBplan" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">是否為管控單位</td>
                            <td colspan="3" class="whitecol" width="80%">
                                <asp:RadioButtonList ID="IsConUnit" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="1">是</asp:ListItem>
                                    <asp:ListItem Value="0" Selected="True">否</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">分支單位名稱</td>
                            <td colspan="3" class="whitecol" width="80%">
                                <asp:TextBox ID="TB_OrgPName" runat="server" Width="60%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">地址</td>
                            <td colspan="3" class="whitecol" width="80%">
                                <input id="city_code" onfocus="this.blur()" name="city_code" runat="server">－
                                <input id="ZipCODEB3" maxlength="3" runat="server">
                                <input id="hidZipCODE6W" type="hidden" runat="server" />
                                <asp:Literal ID="Litcity_code" runat="server"></asp:Literal><%--郵遞區號查詢--%>
                                <br>
                                <asp:TextBox ID="TBCity" runat="server" Width="34%"></asp:TextBox>
                                <asp:TextBox ID="TBaddress" runat="server" Width="56%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr id="TPlanID28" runat="server">
                            <td class="bluecol" width="20%">計畫主持人</td>
                            <td class="whitecol" width="30%">
                                <asp:TextBox ID="PlanMaster" runat="server" Width="45%"></asp:TextBox></td>
                            <td class="bluecol" width="20%">主持人電話</td>
                            <td class="whitecol" width="30%">
                                <asp:TextBox ID="PlanMasterPhone" runat="server" Width="56%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">負責人姓名</td>
                            <td class="whitecol" width="30%">
                                <asp:TextBox ID="TBm_name" runat="server" Width="45%"></asp:TextBox></td>
                            <td id="Td4" runat="server" class="bluecol" width="20%">聯絡人姓名</td>
                            <td class="whitecol" width="30%">
                                <asp:TextBox ID="TBContactName" runat="server" Width="45%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">聯絡人電話</td>
                            <td class="whitecol" width="30%">
                                <asp:TextBox ID="TBtel" runat="server" Width="56%"></asp:TextBox></td>
                            <td id="Td6" runat="server" class="bluecol" width="20%">聯絡人行動電話</td>
                            <td class="whitecol" width="30%">
                                <asp:TextBox ID="TBcontact_cellphone" runat="server" Width="56%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">聯絡人E-MAIL</td>
                            <td class="whitecol" width="30%">
                                <asp:TextBox ID="TBmail" runat="server" Width="90%"></asp:TextBox></td>
                            <td class="bluecol" width="20%">聯絡人傳真</td>
                            <td class="whitecol" width="30%">
                                <asp:TextBox ID="ContactFax" runat="server" Width="56%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">保險證號</td>
                            <td colspan="3" class="whitecol" width="80%">
                                <asp:TextBox ID="TB_ActNo" runat="server" Width="30%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">訓練容量</td>
                            <td class="whitecol" width="30%">
                                <asp:TextBox ID="TB_TrainCap" runat="server" Width="80%"></asp:TextBox></td>
                            <td id="Td7" runat="server" class="bluecol" width="20%">消防安檢狀況</td>
                            <td class="whitecol" width="30%">
                                <asp:TextBox ID="TB_FireControlState" runat="server" Width="60%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">專長訓練職類</td>
                            <td colspan="3" class="whitecol" width="80%">
                                <asp:TextBox ID="TB_ProTrainKind" runat="server" Width="70%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">機構簡介</td>
                            <td colspan="3" class="whitecol" width="80%">
                                <asp:TextBox ID="ComSumm" runat="server" Width="70%" Columns="20" Rows="8" TextMode="MultiLine"></asp:TextBox></td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td colspan="4" width="100%">
                    <table class="font" id="HistoryTable" width="100%" align="left" border="0" runat="server">
                        <tr>
                            <td align="center" colspan="4">
                                <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號">
                                            <HeaderStyle HorizontalAlign="Center" Width="6%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn ItemStyle-Width="4%" ItemStyle-HorizontalAlign="Center">
                                            <HeaderTemplate>
                                                <input onclick="CheckAll();" type="checkbox" checked name="Choose1">
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <input id="Checkbox1" type="checkbox" checked name="Checkbox1" runat="server">
                                                <input id="SeqNO" type="hidden" runat="server" name="SeqNO">
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <%--<asp:BoundColumn DataField="DistName" HeaderText="轄區&lt;BR&gt;中心">--%>
                                        <%--<asp:BoundColumn DataField="DistName" HeaderText="轄區中心">--%>
                                        <asp:BoundColumn DataField="DistName" HeaderText="轄區分署">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="PlanYear" HeaderText="年度">
                                            <HeaderStyle Width="8%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="PlanName" HeaderText="訓練計畫">
                                            <HeaderStyle Width="11%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構">
                                            <HeaderStyle Width="11%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="COMIDNO" HeaderText="訓練機構統編">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TrainName" HeaderText="訓練職類">
                                            <HeaderStyle Width="11%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassName" HeaderText="班別名稱">
                                            <HeaderStyle Width="13%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TRound" HeaderText="受訓期間">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="4">
                                <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <br>
        <asp:HiddenField ID="Hid_Errmsg" runat="server" />
        <%--ViewState("Errmsg")--%>
    </form>
</body>
</html>
