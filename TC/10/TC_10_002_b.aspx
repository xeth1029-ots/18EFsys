<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_10_002_b.aspx.vb" Inherits="WDAIIP.TC_10_002_b" %>

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>會議與出席場次管理</title>
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
        if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        //開啟視窗-查詢審查委員
        function openEXAMINER(nText, nValue) {
            wopen('TC_10_MBR.aspx?TextField=' + nText + '&ValueField=' + nValue, 'sendEXAMINER', 700, 800, 0);
        }

        //CheckboxAll
        function ChangeAll(obj) {
            var objLen = document.form1.length;
            for (var iCount = 0; iCount < objLen; iCount++) {
                if (document.form1.elements[iCount].type == "checkbox") {
                    var mycheck = document.form1.elements[iCount];
                    if (!mycheck.disabled && mycheck.name.endsWith('cbATTEND')) {
                        mycheck.checked = (obj.checked == true ? true : false);
                    }
                    //if (mycheck.checked) { debugger;; }
                }
            }
        }

        function Click_cbATTEND(obj1, obj2) {
            var mycheck = document.getElementById(obj1);
            var mycheck2 = document.getElementById(obj2);
            if (!mycheck.disabled) { mycheck2.checked = (mycheck.checked == true ? false : true); }
            mycheck2.disabled = (mycheck.checked ? true : false);
        }

        function Click_cbNOTINABS(obj1, obj2) {
            var mycheck = document.getElementById(obj1);
            var mycheck2 = document.getElementById(obj2);
            if (!mycheck2.disabled) { mycheck.checked = (mycheck2.checked == true ? false : true); }
            mycheck.disabled = (mycheck2.checked ? true : false);
        }

        function chkSaveData1() {
            var msg = '';
            var ddlDISTID = document.getElementById('ddlDISTID');
            var ddlMYEARS = document.getElementById('ddlMYEARS');
            //var rblCATEGORY = document.getElementById('rblCATEGORY'); 審查會議類別
            //var v_rblCATEGORY = getRBLValue("rblCATEGORY"); //取得 RadioButtonList 值 審查會議類別
            //var v2_rblCATEGORY = getRadioValue(document.form1.rblCATEGORY); //取得 RadioButtonList 值 審查會議類別
            //debugger;;

            var cblORGPLANKIND = document.getElementById('cblORGPLANKIND');
            var ddlACCEPTSTAGE = document.getElementById('ddlACCEPTSTAGE');

            var SMEETDATE = document.getElementById('SMEETDATE');
            var FMEETDATE = document.getElementById('FMEETDATE');
            var MEETPLACE = document.getElementById('MEETPLACE');
            var MODERATOR = document.getElementById('MODERATOR');

            if (ddlDISTID.value == '') { msg += '請選擇 轄區分署\n'; }
            if (ddlMYEARS.value == '') { msg += '請選擇 年度\n'; }
            //if (rblCATEGORY.value == '') { msg += '請選擇 審查會議類別\n'; } 
            //if (v_rblCATEGORY == '') { msg += '請選擇 審查會議類別\n'; }

            var v_cblORGPLANKIND = getCheckBoxListValue('cblORGPLANKIND');
            if (parseInt(v_cblORGPLANKIND, 10) == 0) { msg += '請選擇 計畫別(至少一筆)\n'; }

            if (ddlACCEPTSTAGE.value == '') { msg += '請選擇 受理階段\n'; }

            if (SMEETDATE.value == '') { msg += '請選擇輸入 會議日期/時間-起始日\n'; }
            if (FMEETDATE.value == '') { msg += '請選擇輸入 會議日期/時間-迄止日\n'; }
            if (MEETPLACE.value == '') { msg += '請輸入 會議地點\n'; }
            if (MODERATOR.value == '') { msg += '請輸入 主席\n'; }

            if (msg != '') {
                msg += '!!!\n';
                alert(msg);
                return false;
            }

            //if (ddlACCEPTSTAGE.value == 'B1') { msg += '受理階段不可選政策性 \n'; }
            //if (ddlACCEPTSTAGE.value == 'B2') { msg += '受理階段不可選政策性申復 \n'; }
            //if (ddlACCEPTSTAGE.value == 'D1') { msg += '受理階段不可選進階政策性 \n'; }
            //if (ddlACCEPTSTAGE.value == 'D2') { msg += '受理階段不可選進階政策性申復 \n'; }
            //if (v_rblCATEGORY == '2' && ddlACCEPTSTAGE.value == 'B1') { msg += '審查會議類別-跨區 受理階段不可選政策性 \n'; }
            //if (v_rblCATEGORY == '2' && ddlACCEPTSTAGE.value == 'B2') { msg += '審查會議類別-跨區 受理階段不可選政策性申復 \n'; }
            //if (v_rblCATEGORY == '2' && ddlACCEPTSTAGE.value == 'D1') { msg += '審查會議類別-跨區 受理階段不可選進階政策性 \n'; }
            //if (v_rblCATEGORY == '2' && ddlACCEPTSTAGE.value == 'D2') { msg += '審查會議類別-跨區 受理階段不可選進階政策性申復 \n'; }

            if (msg != '') {
                msg += '!!!\n';
                alert(msg);
                return false;
            }
            return true;
        }

    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;審查委員管理&gt;&gt;會議與出席場次管理</asp:Label>
                </td>
            </tr>
        </table>

        <asp:Panel ID="panelView4" runat="server">
            <table class="table_nw" cellspacing="1" cellpadding="1" width="100%">
                <tr>
                    <td class="head_navy" colspan="2">審查會議</td>
                </tr>
                <tr>
                    <td class="bluecol_need" width="20%">轄區分署</td>
                    <td class="whitecol">
                        <asp:Label ID="labDISTID4" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol_need">年度</td>
                    <td class="whitecol">
                        <asp:Label ID="labMYEARS4" runat="server"></asp:Label></td>
                </tr>
                <%--<tr>
                    <td class="bluecol_need">審查會議類別</td>
                    <td class="whitecol">
                        <asp:Label ID="labCATEGORY4" runat="server"></asp:Label></td>
                </tr>--%>
                <tr>
                    <td class="bluecol_need">計畫別</td>
                    <td class="whitecol">
                        <asp:Label ID="labORGPLANKIND4" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol_need">審查課程職類</td>
                    <td class="whitecol">
                        <asp:Label ID="labGOVCODE3_4" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol_need">受理階段</td>
                    <td class="whitecol">
                        <asp:Label ID="labACCEPTSTAGE4" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol_need">日期/時間</td>
                    <td class="whitecol">
                        <asp:Label ID="labSFMEETDATE4" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol_need">會議地點</td>
                    <td class="whitecol">
                        <asp:Label ID="labMEETPLACE4" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol">主席</td>
                    <td class="whitecol">
                        <asp:Label ID="labMODERATOR4" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td align="center" colspan="2" class="head_navy">審查委員名單</td>
                </tr>

                <tr>
                    <td colspan="2" width="100%">
                        <table id="DataGrid4Table" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                            <tr>
                                <td>
                                    <asp:DataGrid ID="DataGrid4" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False">
                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <Columns>
                                            <asp:BoundColumn HeaderText="序號" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="5%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="RECRUIT_N" HeaderText="遴聘類別"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="MBRNAME" HeaderText="審查委員姓名"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="UNITNAME" HeaderText="現職服務機構"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="JOBTITLE" HeaderText="職稱"></asp:BoundColumn>
                                            <%--<asp:BoundColumn DataField="PUSHDISTID" HeaderText="推薦分署"></asp:BoundColumn>--%>
                                            <asp:TemplateColumn HeaderText="推薦分署">
                                                <ItemTemplate>
                                                    <asp:HiddenField ID="Hid_EMSEQ" runat="server" />
                                                    <asp:Label ID="labPUSHDISTID_N" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                        <PagerStyle Visible="False"></PagerStyle>
                                    </asp:DataGrid>
                                </td>
                            </tr>
                            <tr>
                                <td></td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td colspan="2" class="whitecol" align="center" width="100%">
                        <div align="center">
                            <asp:Label ID="labmsg_DG4" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
                        </div>
                        <%--Button1_Click Button1--%>
                        <asp:Button ID="btnBack4" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                    </td>
                </tr>
            </table>
        </asp:Panel>

        <asp:Panel ID="panelEdit3" runat="server">
            <table class="table_nw" cellspacing="1" cellpadding="1" width="100%">
                <tr>
                    <td class="head_navy" colspan="2">審查會議</td>
                </tr>
                <tr>
                    <td class="bluecol_need" width="20%">轄區分署</td>
                    <td class="whitecol">
                        <asp:Label ID="labDISTID" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol_need">年度</td>
                    <td class="whitecol">
                        <asp:Label ID="labMYEARS" runat="server"></asp:Label></td>
                </tr>
                <%--<tr>
                    <td class="bluecol_need">審查會議類別</td>
                    <td class="whitecol">
                        <asp:Label ID="labCATEGORY" runat="server"></asp:Label></td>
                </tr>--%>
                <tr>
                    <td class="bluecol_need">計畫別</td>
                    <td class="whitecol">
                        <asp:Label ID="labORGPLANKIND" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol_need">審查課程職類</td>
                    <td class="whitecol">
                        <asp:Label ID="labGOVCODE3" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol_need">受理階段</td>
                    <td class="whitecol">
                        <asp:Label ID="labACCEPTSTAGE" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol_need">日期/時間</td>
                    <td class="whitecol">
                        <asp:Label ID="labSFMEETDATE" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol_need">會議地點</td>
                    <td class="whitecol">
                        <asp:Label ID="labMEETPLACE" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol">主席</td>
                    <td class="whitecol">
                        <asp:Label ID="labMODERATOR" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td align="center" colspan="2" class="head_navy">審查委員出席狀況</td>
                </tr>
                <tr>
                    <td colspan="2" width="100%">
                        <table id="DataGrid3Table" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                            <tr>
                                <td>
                                    <asp:DataGrid ID="DataGrid3" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False">
                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <Columns>
                                            <asp:BoundColumn HeaderText="序號" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="5%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="RECRUIT_N" HeaderText="遴聘類別"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="MBRNAME" HeaderText="審查委員姓名"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="UNITNAME" HeaderText="現職服務機構"></asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="出席">
                                                <HeaderTemplate>
                                                    出席<br />
                                                    <input id="CheckboxAll" type="checkbox" runat="server">
                                                </HeaderTemplate>
                                                <ItemStyle Width="10%" HorizontalAlign="Center" />
                                                <ItemTemplate>
                                                    <input id="cbATTEND" type="checkbox" runat="server" />
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="不列入缺席">
                                                <ItemStyle Width="10%" HorizontalAlign="Center" />
                                                <ItemTemplate>
                                                    <input id="cbNOTINABS" type="checkbox" runat="server" />
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="不列入缺席原因">
                                                <HeaderStyle HorizontalAlign="Center" VerticalAlign="middle"></HeaderStyle>
                                                <ItemStyle Width="20%" HorizontalAlign="Center" />
                                                <ItemTemplate>
                                                    <asp:HiddenField ID="Hid_EMSEQ" runat="server" />
                                                    <asp:TextBox ID="REASON" runat="server" MaxLength="300" TextMode="MultiLine" Width="95%" Height="60px"></asp:TextBox>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                        <PagerStyle Visible="False"></PagerStyle>
                                    </asp:DataGrid>
                                </td>
                            </tr>
                            <tr>
                                <td></td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td colspan="2" class="whitecol" align="center" width="100%">
                        <div align="center">
                            <asp:Label ID="labmsg_DG3" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
                        </div>
                        <%--Button1_Click Button1--%>
                        <asp:Button ID="btnSave3" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="btnBack3" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                        <br />
                        (若有勾選，離開時要按儲存)</td>
                </tr>
            </table>
        </asp:Panel>

        <asp:Panel ID="panelEdit4" runat="server">
            <table class="table_nw" cellspacing="1" cellpadding="1" width="100%">
                <tr>
                    <td class="head_navy" colspan="2">審查會議</td>
                </tr>
                <tr>
                    <td class="bluecol_need" width="20%">轄區分署</td>
                    <td class="whitecol">
                        <asp:Label ID="labDISTIDp4" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol_need">年度</td>
                    <td class="whitecol">
                        <asp:Label ID="labMYEARSp4" runat="server"></asp:Label></td>
                </tr>
                <%--<tr><td class="bluecol_need">審查會議類別</td><td class="whitecol"><asp:Label ID="labCATEGORY4" runat="server"></asp:Label></td></tr>--%>
                <tr>
                    <td class="bluecol_need">計畫別</td>
                    <td class="whitecol">
                        <asp:Label ID="labORGPLANKINDp4" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol_need">審查課程職類</td>
                    <td class="whitecol">
                        <asp:Label ID="labGOVCODE3_p4" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol_need">受理階段</td>
                    <td class="whitecol">
                        <asp:Label ID="labACCEPTSTAGEp4" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol_need">日期/時間</td>
                    <td class="whitecol">
                        <asp:Label ID="labSFMEETDATEp4" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol_need">會議地點</td>
                    <td class="whitecol">
                        <asp:Label ID="labMEETPLACEp4" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol">主席</td>
                    <td class="whitecol">
                        <asp:Label ID="labMODERATORp4" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td align="center" colspan="2" class="head_navy">審查委員計畫職類</td>
                </tr>
                <tr>
                    <td colspan="2" width="100%">
                        <table id="tbDataGrid5" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                            <tr>
                                <td>
                                    <asp:DataGrid ID="DataGrid5" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False">
                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <Columns>
                                            <asp:BoundColumn HeaderText="序號" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="5%">
                                                <HeaderStyle Width="5%" />
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="RECRUIT_N" HeaderText="遴聘類別" HeaderStyle-Width="10%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="MBRNAME" HeaderText="審查委員姓名" HeaderStyle-Width="10%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="UNITNAME" HeaderText="現職服務機構" HeaderStyle-Width="15%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="JOBTITLE" HeaderText="職稱" HeaderStyle-Width="10%"></asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="審查計畫">
                                                <HeaderStyle Width="15%" />
                                                <ItemTemplate>
                                                    <input id="cbORGKIND_G" type="checkbox" runat="server" />產業人才投資計畫<br />
                                                    <input id="cbORGKIND_W" type="checkbox" runat="server" />提升勞工自主學習計畫
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="審查職類">
                                                <HeaderStyle Width="15%" />
                                                <ItemTemplate>
                                                    <asp:HiddenField ID="Hid_EMSEQ" runat="server" />
                                                    <asp:CheckBoxList ID="cblGOVCODE3_dg5" runat="server" Width="100%"></asp:CheckBoxList>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                        <PagerStyle Visible="False"></PagerStyle>
                                    </asp:DataGrid>
                                </td>
                            </tr>
                            <tr>
                                <td></td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td colspan="2" class="whitecol" align="center" width="100%">
                        <div align="center">
                            <asp:Label ID="labmsg_DG5" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
                        </div>
                        <%--Button1_Click Button1--%>
                        <asp:Button ID="BtnSAVEDATAp4" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="BtnBACKp4" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                        <br />
                        (若有勾選，離開時要按儲存)</td>
                </tr>
            </table>
        </asp:Panel>


        <asp:Panel ID="panelEdit" runat="server">
            <table class="table_nw" cellspacing="1" cellpadding="1" width="100%">
                <tr>
                    <td class="head_navy" colspan="4">審查會議</td>
                </tr>
                <tr>
                    <td class="bluecol_need" width="20%">轄區分署</td>
                    <td colspan="3" class="whitecol">
                        <asp:DropDownList ID="ddlDISTID" runat="server"></asp:DropDownList></td>
                </tr>
                <tr>
                    <td class="bluecol_need">年度</td>
                    <td colspan="3" class="whitecol">
                        <asp:DropDownList ID="ddlMYEARS" runat="server"></asp:DropDownList></td>
                </tr>
                <%-- <tr>
                    <td class="bluecol_need">審查會議類別</td>
                    <td colspan="3" class="whitecol">
                        <asp:RadioButtonList ID="rblCATEGORY" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                            <asp:ListItem Value="1">轄區</asp:ListItem>
                            <asp:ListItem Value="2">跨區</asp:ListItem>
                        </asp:RadioButtonList></td>
                </tr>--%>
                <tr>
                    <td class="bluecol_need">計畫別</td>
                    <td colspan="3" class="whitecol">
                        <asp:CheckBoxList ID="cblORGPLANKIND" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                            <asp:ListItem Value="G">產業人才投資計畫</asp:ListItem>
                            <asp:ListItem Value="W">提升勞工自主學習計畫</asp:ListItem>
                        </asp:CheckBoxList></td>
                </tr>
                <tr>
                    <td class="bluecol_need">審查課程職類</td>
                    <td colspan="3" class="whitecol">
                        <input id="cblGOVCODE3Hidden" type="hidden" value="0" name="cblGOVCODE3Hidden" runat="server" />
                        <asp:CheckBoxList ID="cblGOVCODE3" runat="server" RepeatDirection="Horizontal" RepeatColumns="5" Width="100%"></asp:CheckBoxList></td>
                </tr>
                <tr>
                    <td class="bluecol_need">受理階段</td>
                    <td colspan="3" class="whitecol">
                        <%--<asp:ListItem Value="">==請選擇==</asp:ListItem><asp:ListItem Value="A1">上半年</asp:ListItem><asp:ListItem Value="A2">上半年申復</asp:ListItem>
                          <asp:ListItem Value="B1">政策性</asp:ListItem><asp:ListItem Value="B2">政策性申復</asp:ListItem>
                          <asp:ListItem Value="C1">下半年</asp:ListItem><asp:ListItem Value="C2">下半年申復</asp:ListItem>--%>
                        <asp:DropDownList ID="ddlACCEPTSTAGE" runat="server"></asp:DropDownList></td>
                </tr>
                <tr>
                    <td class="bluecol_need">日期/時間</td>
                    <td colspan="3" class="whitecol">
                        <asp:TextBox ID="SMEETDATE" runat="server" Columns="16" MaxLength="11"></asp:TextBox>
                        <span runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= SMEETDATE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= SMEETDATE.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" />
                        </span>&nbsp;
                        <asp:DropDownList ID="HR1" runat="server"></asp:DropDownList>時：<asp:DropDownList ID="MM1" runat="server"></asp:DropDownList>分～
                        <asp:TextBox ID="FMEETDATE" runat="server" Columns="16" MaxLength="11"></asp:TextBox>
                        <span runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= FMEETDATE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= FMEETDATE.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" />
                        </span>
                        &nbsp;
                        <asp:DropDownList ID="HR2" runat="server">
                        </asp:DropDownList>
                        時：<asp:DropDownList ID="MM2" runat="server">
                        </asp:DropDownList>
                        分
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need">會議地點</td>
                    <td colspan="3" class="whitecol">
                        <asp:TextBox ID="MEETPLACE" runat="server" MaxLength="200" Columns="70"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol">主席</td>
                    <td colspan="3" class="whitecol">
                        <asp:TextBox ID="MODERATOR" runat="server" MaxLength="30" Columns="44"></asp:TextBox></td>
                </tr>
                <tr>
                    <td align="center" colspan="4" class="head_navy">預計參加審查委員名單</td>
                </tr>
                <tr>
                    <td colspan="4" width="100%">
                        <table id="Table3" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td align="center" class="bluecol" width="70%">查詢審查委員 </td>
                                <td align="center" class="bluecol" width="20%">功能</td>
                            </tr>
                            <tr>
                                <td class="whitecol" align="right">
                                    <asp:TextBox ID="txtEXAMINER" runat="server" onfocus="this.blur()" Columns="90"></asp:TextBox>
                                    <input id="btn_sel3" onclick="openEXAMINER('txtEXAMINER', 'hid_EMSEQVAL');" value="..." type="button" runat="server" class="button_b_Mini" />
                                    <input id="hid_EMSEQVAL" type="hidden" runat="server" /></td>
                                <td class="whitecol" align="center">
                                    <asp:Button ID="Button29" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_M"></asp:Button></td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td colspan="4" width="100%">
                        <table id="DataGrid2Table" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                            <tr>
                                <td>
                                    <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False">
                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <Columns>
                                            <asp:BoundColumn HeaderText="序號" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="5%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="RECRUIT_N" HeaderText="遴聘類別"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="MBRNAME" HeaderText="審查委員姓名"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="UNITNAME" HeaderText="現職服務機構"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="JOBTITLE" HeaderText="職稱"></asp:BoundColumn>
                                            <%--<asp:BoundColumn DataField="PUSHDISTID" HeaderText="推薦分署"></asp:BoundColumn>--%>
                                            <asp:TemplateColumn HeaderText="推薦分署">
                                                <ItemTemplate>
                                                    <asp:Label ID="labPUSHDISTID_N" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="功能">
                                                <ItemStyle Width="20%" HorizontalAlign="Center" />
                                                <ItemTemplate>
                                                    <asp:HiddenField ID="Hid_EMSEQ" runat="server" />
                                                    <asp:Button ID="BTNDEL1" runat="server" CausesValidation="False" Text="刪除" CommandName="DEL1" CssClass="asp_button_M"></asp:Button>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                        <PagerStyle Visible="False"></PagerStyle>
                                    </asp:DataGrid>
                                </td>
                            </tr>
                            <tr>
                                <td></td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td colspan="4" class="whitecol" align="center" width="100%">
                        <%--Button1_Click Button1--%>
                        <asp:Button ID="btnSave1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="btnBack1" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                        <br />
                        (若有更動，離開時要按儲存)</td>
                </tr>
            </table>
        </asp:Panel>

        <asp:Panel ID="panelSch" runat="server">
            <table width="100%" cellpadding="1" cellspacing="1" class="table_sch">
                <tr>
                    <td class="bluecol" width="20%">轄區分署</td>
                    <td colspan="3" class="whitecol">
                        <asp:DropDownList ID="ddlDISTID_SCH" runat="server"></asp:DropDownList></td>
                </tr>
                <tr>
                    <td class="bluecol_need">年度</td>
                    <td colspan="3" class="whitecol">
                        <asp:DropDownList ID="ddlMYEARS_SCH" runat="server"></asp:DropDownList></td>
                </tr>
                <%--<tr><td class="bluecol">審查會議類別</td>
                    <td colspan="3" class="whitecol">
                        <asp:RadioButtonList ID="rblCATEGORY_SCH" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                            <asp:ListItem Value="" Selected="True">不區分</asp:ListItem>
                            <asp:ListItem Value="1">轄區</asp:ListItem>
                            <asp:ListItem Value="2">跨區</asp:ListItem>
                        </asp:RadioButtonList></td></tr>--%>
                <tr>
                    <td class="bluecol">計畫別</td>
                    <td colspan="3" class="whitecol">
                        <asp:CheckBoxList ID="cblORGPLANKIND_sch" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                            <asp:ListItem Value="G">產業人才投資計畫</asp:ListItem>
                            <asp:ListItem Value="W">提升勞工自主學習計畫</asp:ListItem>
                        </asp:CheckBoxList></td>
                </tr>
                <tr>
                    <td class="bluecol">審查課程職類</td>
                    <td colspan="3" class="whitecol">
                        <input id="cblGOVCODE3_schHidden" type="hidden" value="0" name="cblGOVCODE3_schHidden" runat="server" />
                        <asp:CheckBoxList ID="cblGOVCODE3_sch" runat="server" RepeatDirection="Horizontal" RepeatColumns="5" Width="100%"></asp:CheckBoxList></td>
                </tr>
                <tr>
                    <td class="bluecol">受理階段</td>
                    <td colspan="3" class="whitecol">
                        <%-- <asp:ListItem Value="">==請選擇==</asp:ListItem><asp:ListItem Value="A1">上半年</asp:ListItem><asp:ListItem Value="A2">上半年申復</asp:ListItem>
                            <asp:ListItem Value="B1">政策性</asp:ListItem><asp:ListItem Value="B2">政策性申復</asp:ListItem>
                            <asp:ListItem Value="C1">下半年</asp:ListItem><asp:ListItem Value="C2">下半年申復</asp:ListItem>--%>
                        <asp:DropDownList ID="ddlACCEPTSTAGE_sch" runat="server"></asp:DropDownList></td>
                </tr>
                <tr>
                    <td class="bluecol">會議日期</td>
                    <td colspan="3" class="whitecol">
                        <asp:TextBox ID="SMEETDATE_sch1" runat="server" Columns="20" MaxLength="11"></asp:TextBox>
                        <span runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= SMEETDATE_sch1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= SMEETDATE_sch1.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" />
                        </span>～<asp:TextBox ID="SMEETDATE_sch2" runat="server" Columns="20"></asp:TextBox>
                        <span runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= SMEETDATE_sch2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= SMEETDATE_sch2.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" />
                        </span></td>
                </tr>
                <tr>
                    <td class="bluecol">會議地點</td>
                    <td colspan="3" class="whitecol">
                        <asp:TextBox ID="MEETPLACE_sch" runat="server" MaxLength="200" Columns="66"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol">主席</td>
                    <td colspan="3" class="whitecol">
                        <asp:TextBox ID="MODERATOR_sch" runat="server" MaxLength="30" Columns="44"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol">匯出檔案格式</td>
                    <td colspan="3" class="whitecol">
                        <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                            <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                            <asp:ListItem Value="ODS">ODS</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td class="whitecol" align="center" colspan="4">
                        <asp:Button ID="BtnSearch" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="BtnAddnew" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>
                    </td>
                </tr>
            </table>
            <div align="center">
                <asp:Label ID="msg1" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
            </div>
            <table id="tbDataGrid1" runat="server" cellspacing="1" cellpadding="1" width="100%" border="0">
                <tr>
                    <td align="center">
                        <%--序號、遴聘類別、姓名、現職服務機構、職稱、推薦分署--%>
                        <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" AllowPaging="True" Width="100%" AutoGenerateColumns="False" CellPadding="8">
                            <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                            <Columns>
                                <asp:BoundColumn HeaderText="序號" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="5%"></asp:BoundColumn>
                                <%--<asp:BoundColumn DataField="RECRUIT" HeaderText="遴聘類別"></asp:BoundColumn>--%>
                                <%--<asp:TemplateColumn HeaderText="遴聘類別">
                                    <ItemTemplate>
                                        <asp:Label ID="labRECRUIT_N" runat="server"></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateColumn>--%>
                                <asp:BoundColumn DataField="DISTNAME" HeaderText="轄區分署"></asp:BoundColumn>
                                <%--<asp:BoundColumn DataField="CATEGORY_N" HeaderText="會議類別"></asp:BoundColumn>--%>
                                <asp:BoundColumn DataField="ORGPLANKIND_N" HeaderText="計畫別"></asp:BoundColumn>
                                <asp:BoundColumn DataField="GOVCODE3_N" HeaderText="課程職類"></asp:BoundColumn>
                                <asp:BoundColumn DataField="ACCEPTSTAGE_N" HeaderText="受理階段"></asp:BoundColumn>
                                <asp:BoundColumn DataField="SFMEETDATE_N" HeaderText="會議時間"></asp:BoundColumn>
                                <asp:BoundColumn DataField="MEETPLACE" HeaderText="地點"></asp:BoundColumn>
                                <%--<asp:BoundColumn DataField="PUSHDISTID" HeaderText="推薦分署"></asp:BoundColumn>--%>
                                <%--<asp:TemplateColumn HeaderText="轄區分署">
                                    <ItemTemplate>
                                        <asp:Label ID="labPUSHDISTID_N" runat="server"></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateColumn>--%>
                                <asp:TemplateColumn HeaderText="功能" ItemStyle-HorizontalAlign="Center">
                                    <HeaderStyle HorizontalAlign="center" Width="22%"></HeaderStyle>
                                    <ItemTemplate>
                                        <asp:Button ID="BTNUPD1" runat="server" Text="修改" CommandName="UPD1" CssClass="asp_button_M" />
                                        <asp:Button ID="BTNDEL1" runat="server" Text="刪除" CommandName="DEL1" CssClass="asp_button_M" />
                                        <asp:Button ID="BTNEDIT4" runat="server" Text="委員計畫職類" CommandName="EDIT4" CssClass="asp_button_M" />
                                        <asp:Button ID="BTNEDIT3" runat="server" Text="管理出席狀況" CommandName="EDIT3" CssClass="asp_button_M" />
                                        <asp:Button ID="BTNVIEW4" runat="server" Text="查看出席狀況" CommandName="VIEW4" CssClass="asp_button_M" />
                                        <asp:Button ID="BTNEXP1" runat="server" Text="匯出名單" CommandName="EXP1" CssClass="asp_Export_M" />
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                            </Columns>
                            <PagerStyle Visible="False"></PagerStyle>
                        </asp:DataGrid>
                        <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:HiddenField ID="Hid_MTSEQ" runat="server" />
        <asp:HiddenField ID="hid_GOVCODE3" runat="server" />
        <asp:HiddenField ID="hid_EXAMINER_TABLE_GUID1" runat="server" />
    </form>
    <%--序號、遴聘類別、姓名、現職服務機構、職稱、推薦分署--%>
</body>
</html>
