<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_07_001.aspx.vb" Inherits="WDAIIP.CP_07_001" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>受訓期間學員滿意度</title>
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script language="javascript" type="text/javascript">
        //function search1() {
        //    document.form1.hidSearchTag.value = 'search';
        //    if (document.form1.OCIDValue1.value == '') {
        //        alert('請選擇職類班別!');
        //        return false;
        //    }
        //}
        function chkOrg() {
            if (document.getElementById("RIDValue").value == '') {
                alert('請選擇機構!');
                return false;
            }
            //if (document.getElementById("OCIDValue1").value == '') {
            //    alert('請選擇班別!');
            //    return false;
            //}
        }

        function SetOneOCID() {
            document.getElementById('Button7').click();
        }

        function choose_class() {
            var RID = document.form1.RIDValue.value;
            if (document.getElementById('OCID1').value == '')
            { document.getElementById('Button7').click(); }
            openClass('../../SD/02/SD_02_ch.aspx?RID=' + RID);
        }

        function ChkSOCID() {
            var MyTable = document.getElementById('DG_ClassInfo');
            var hid_SOCIDvalue = '';
            for (var i = 1; i < MyTable.rows.length; i++) {
                var Mycells0 = MyTable.rows[i].cells[0];
                if (!Mycells0) { return; }
                if (Mycells0.children[0] && Mycells0.children[0].checked && Mycells0.children[0].value != "") {
                    //debugger; //CB_SOCID
                    //var socid1 = MyTable.rows(i).cells(0).children(0).value;
                    if (hid_SOCIDvalue != '') { hid_SOCIDvalue += ','; }
                    hid_SOCIDvalue += Mycells0.children[0].value;
                }
            }
            document.getElementById('hid_SOCIDvalue').value = hid_SOCIDvalue;
        }

        function SelectAll(Flag) {
            var MyTable = document.getElementById('DG_ClassInfo');
            for (var i = 1; i < MyTable.rows.length; i++) {
                var Mycells0 = MyTable.rows[i].cells[0];
                if (!Mycells0) { return; }
                Mycells0.children[0].checked = Flag;
            }
        }

        //function chkprint(SMpath, rptType) {
        //    var InquireType = document.getElementsByName("InquireType");
        //    var selectType;
        //    var years = '';
        //    years = document.getElementById('years').value;

        //    if (InquireType[1].checked)
        //    { selectType = 3; }
        //    else if (InquireType[2].checked)
        //    { selectType = 2; }
        //    else if (InquireType[3].checked)
        //    { selectType = 1; }

        //    if (document.getElementById('DG_ClassInfo') == null) {
        //        alert('目前無資料可供列印！');
        //        return false;
        //    }

        //    ChkSOCID();

        //    var SOCIDvalue = '';
        //    SOCIDvalue = document.getElementById('SOCIDvalue').value;

        //    if (selectType == null) {
        //        alert('請選擇調查方式！');
        //        return false;
        //    }
        //    else {
        //        if (InquireType[1].checked) {
        //            alert('調查方式 請選擇 『電話訪查』或『系統登打』！');
        //            return false;
        //        }
        //    }
        //    if (SOCIDvalue == '') {
        //        alert('請選擇 『學員』！');
        //        return false;
        //    }
        //    if (rptType != 1) {
        //        if (document.getElementById('IsDataload').recordNumber == 'n') {
        //            alert('目前無資料可供列印！');
        //            return false;
        //        }
        //        if (document.getElementById("center").value == '') {
        //            alert('請選擇機構!');
        //            return false;
        //        }
        //        if (document.getElementById("OCID1").value == '') {
        //            alert('請選擇班別!');
        //            return false;
        //        }
        //        if (InquireType[2].checked) {   //電話
        //            openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=member&filename=CP_07_003&path=' + SMpath + '&SOCID=' + SOCIDvalue + '&Years=' + document.getElementById('years').value + '&OCID=' + document.getElementById('OCIDValue1').value + '&RID=' + document.getElementById('RIDValue').value + '&PlanID=' + document.getElementById('PlanID').value);
        //        }
        //        else if (InquireType[3].checked)  //系統登打
        //        {
        //            if (years == '2010') {
        //                openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=member&filename=CP_07_001_2010&path=' + SMpath + '&SOCID=' + SOCIDvalue + '&Years=' + document.getElementById('years').value + '&OCID=' + document.getElementById('OCIDValue1').value + '&RID=' + document.getElementById('RIDValue').value + '&PlanID=' + document.getElementById('PlanID').value);
        //            } else {
        //                openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=member&filename=CP_07_001&path=' + SMpath + '&SOCID=' + SOCIDvalue + '&Years=' + document.getElementById('years').value + '&OCID=' + document.getElementById('OCIDValue1').value + '&RID=' + document.getElementById('RIDValue').value + '&PlanID=' + document.getElementById('PlanID').value);
        //            }
        //        }
        //    }
        //    else {
        //        if (InquireType[2].checked) //電話
        //        { openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=member&filename=CP_07_003_blank&path=' + SMpath + '&Years=' + document.getElementById('years').value + '&distid=' + document.getElementById('distid').value); }
        //        else if (InquireType[3].checked)  //系統登打
        //        {
        //            if (years == '2010') {
        //                openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=member&filename=CP_07_001_blank_2010&path=' + SMpath + '&Years=' + document.getElementById('years').value + '&distid=' + document.getElementById('distid').value);
        //            }
        //            else {
        //                openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=member&filename=CP_07_001_blank&path=' + SMpath + '&Years=' + document.getElementById('years').value + '&distid=' + document.getElementById('distid').value);
        //            }
        //        }
        //    }
        //}

        //function printBlank(rptType) {
        //    var years = '';
        //    years = document.getElementById('years').value;
        //    //電話
        //    if (rptType == 2) { openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=member&filename=CP_07_003_blank&path=' + SMpath + '&Years=' + document.getElementById('years').value + '&distid=' + document.getElementById('distid').value); }
        //    //系統登打
        //    else if (rptType == 1) {
        //        if (years == '2010') {
        //            openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=member&filename=CP_07_001_blank_2010&path=' + SMpath + '&Years=' + document.getElementById('years').value + '&distid=' + document.getElementById('distid').value);
        //        }
        //        else {
        //            openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=member&filename=CP_07_001_blank&path=' + SMpath + '&Years=' + document.getElementById('years').value + '&distid=' + document.getElementById('distid').value);
        //        }
        //    }
        //}

        //function PrintRpt(rptType, SOCIDvalue, OCID) {
        //    var years = '';
        //    years = document.getElementById('years').value;
        //    //系統登打
        //    if (rptType == 1) {
        //        if (years == '2010') {
        //            openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=member&filename=CP_07_001_2010&path=' + SMpath + '&SOCID=' + SOCIDvalue + '&Years=' + document.getElementById('years').value + '&OCID=' + OCID + '&PlanID=' + document.getElementById('PlanID').value);
        //        }
        //        else {
        //            openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=member&filename=CP_07_001&path=' + SMpath + '&SOCID=' + SOCIDvalue + '&Years=' + document.getElementById('years').value + '&OCID=' + OCID + '&PlanID=' + document.getElementById('PlanID').value);
        //        }
        //    }
        //    //電話
        //    else {
        //        openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=member&filename=CP_07_003&path=' + SMpath + '&SOCID=' + SOCIDvalue + '&Years=' + document.getElementById('years').value + '&OCID=' + OCID + '&PlanID=' + document.getElementById('PlanID').value);
        //    }
        //}		

    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;訓練成效與滿意度&gt;&gt;受訓期間學員滿意度</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Frametable3" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td width="16%" class="bluecol">訓練機構 </td>
                <td width="84%" class="whitecol" colspan="3">
                    <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                    <input id="Button2" value="..." type="button" name="Button2" runat="server" class="button_b_Mini">
                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                    <asp:Button Style="display: none" ID="Button7" runat="server" Text="Button7"></asp:Button>
                    <%--<input id="years" type="hidden" name="years" runat="server" />
                    <input id="PlanID" type="hidden" name="PlanID" runat="server" />
                    <input id="distid" type="hidden" name="distid" runat="server" />
                    <input id="IsDataload" type="hidden" name="IsDataload" runat="server">--%>
                    <span id="HistoryList2" style="position: absolute; display: none">
                        <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                    </span>

                </td>
            </tr>
            <tr>
                <td class="bluecol">職類/班別</td>
                <td class="whitecol" colspan="3">
                    <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                    <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                    <input onclick="choose_class()" value="..." type="button" class="button_b_Mini">
                    <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                    <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                    <%-- <input id="hidSearchTag" type="hidden" name="hidSearchTag" runat="server" />
                    <input id="SOCIDvalue" type="hidden" name="SOCIDvalue" runat="server" />--%>
                    <span id="HistoryList" style="position: absolute; display: none;">
                        <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                    </span>
                    <%--<span style="position: absolute; width: 208px; display: none; height: 38px; top: 104px; left: 270px" id="HistoryList">
                        <asp:Table ID="HistoryTable" runat="server" Width="310">
                        </asp:Table>
                    </span>--%>
                </td>
            </tr>
            <tr>
                <td class="bluecol">開訓日期 </td>
                <td class="whitecol">
                    <span id="span01" runat="server">
                        <asp:TextBox ID="STDate1" runat="server" onfocus="this.blur()" Columns="10" Width="15%"></asp:TextBox>
                        <img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">~
				        <asp:TextBox ID="STDate2" runat="server" onfocus="this.blur()" Columns="10" Width="15%"></asp:TextBox>
                        <img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">
                    </span>
                </td>
            </tr>
             <tr id="trbtnImport1" runat="server">
                <td class="bluecol">匯入問卷資料</td>
                <td class="whitecol">
                    <input id="File1" type="file" name="File1" runat="server"  accept=".ods,.xls" size="77" />
                    <asp:Button ID="Btn_XlsImport" runat="server" Text="匯入問卷" CssClass="asp_button_M"></asp:Button>(必須為ods或xls格式)<br />
                    <asp:HyperLink ID="HyperLink1" runat="server" CssClass="font" ForeColor="#8080FF">下載整批上載格式檔</asp:HyperLink>
                     <%--NavigateUrl="../../Doc/StudQues_v21.zip"--%>
                </td>
            </tr>

            <%--style="display: none"--%>
            <%--<tr id="tr_StudentID" runat="server">
                <td width="16%" class="bluecol">學號 </td>
                <td width="84%" class="whitecol" colspan="3">
                    <asp:TextBox ID="StudentID" runat="server" Columns="20"></asp:TextBox><font color="#000033"> &nbsp; </font>

                </td>
            </tr>
            <tr>
                <td class="bluecol">調查方式 </td>
                <td class="whitecol" colspan="3">
                    <asp:RadioButtonList ID="InquireType" runat="server" CssClass="font" RepeatDirection="Horizontal">
                        <asp:ListItem Value="0" Selected="True">全部</asp:ListItem>
                        <asp:ListItem Value="1">系統登打</asp:ListItem>
                        <asp:ListItem Value="2">電話訪查</asp:ListItem>
                    </asp:RadioButtonList>
                    <asp:Button ID="bt_blankRpt" runat="server" Text="列印空白報表(系統登打)" CssClass="asp_Export_M"></asp:Button>
                    <asp:Button ID="bt_blankRpt2" runat="server" Text="列印空白報表(電話訪查)" CssClass="asp_Export_M"></asp:Button>
                </td>
            </tr>--%>
            <tr id="tr_bt_search" runat="server">
                <td align="center" class="whitecol" colspan="4">
                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                    <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                    <asp:Button ID="bt_search" runat="server" Text="查詢" CssClass="asp_Export_M"></asp:Button>
                    <asp:Button ID="btnPrintB2" runat="server" Text="列印空白調查表" CssClass="asp_Export_M"></asp:Button>
                </td>
            </tr>
        </table>
        <div align="center">
            <asp:Label ID="msg1" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
        </div>
        <table id="tb_DG_ClassInfo" runat="server" width="100%">
            <tr>
                <td><%--Visible="False"--%>
                    <asp:DataGrid ID="DG_ClassInfo" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False"
                        AllowPaging="True" AllowSorting="True">
                        <AlternatingItemStyle BackColor="WhiteSmoke"></AlternatingItemStyle>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn HeaderText="序號">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <%--<HeaderStyle HorizontalAlign="Center" Width="40px"></HeaderStyle>--%>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="ClassCName" HeaderText="班別名稱">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="CNT5" HeaderText="結訓人數">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="CNT6" HeaderText="填寫人數">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="功能">
                                <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Button ID="btnQUIRE" runat="server" Text="查詢" CommandName="QUIRE" CssClass="asp_button_M"></asp:Button>
                                    <%--<asp:Button ID="btnPrint" Text="列印" runat="server" CommandName="print" CssClass="asp_Export_M"></asp:Button>--%>
                                    <asp:Button ID="btnPrintB1" Text="列印空白調查表" runat="server" CommandName="PrintB1" CssClass="asp_Export_M"></asp:Button>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                </td>
            </tr>
            <tr>
                <td>

                    <div align="center">
                        <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                    </div>
                </td>
            </tr>
        </table>

        <div align="center">
            <asp:Label ID="msg2" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
        </div>

        <table id="tb_StudentTable" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
            <tr>
                <td>
                    <asp:Label ID="ClassLabel1" runat="server" CssClass="font"></asp:Label>
                    <br />
                    <asp:Label ID="ClassLabel2" runat="server" CssClass="font"></asp:Label>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:DataGrid ID="DG_stud" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                        <AlternatingItemStyle BackColor="WhiteSmoke"></AlternatingItemStyle>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn DataField="STUDID2" HeaderText="學號">
                                <HeaderStyle HorizontalAlign="Center" Width="25%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="STDNAME" HeaderText="姓名(離退訓日期)">
                                <HeaderStyle HorizontalAlign="Center" Width="25%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="FILLSTATUS_N" HeaderText="填寫狀態">
                                <HeaderStyle HorizontalAlign="Center" Width="25%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="功能">
                                <HeaderStyle HorizontalAlign="Center" Width="25%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Button ID="btnInsert" runat="server" Text="新增" CommandName="Insert" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="btnEdit" runat="server" Text="修改" CommandName="Edit" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="btnCheck" runat="server" Text="查詢" CommandName="Check" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="btnPrint" runat="server" Text="列印" CommandName="Print" ToolTip="填寫狀態為「是」，才可列印" CssClass="asp_Export_M"></asp:Button>
                                    <asp:Button ID="btnClear" runat="server" Text="清除重填" CommandName="Clear" CssClass="asp_button_M"></asp:Button>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <%--<asp:BoundColumn Visible="False" DataField="OCID" HeaderText="OCID"></asp:BoundColumn>
                            <asp:BoundColumn Visible="False" DataField="StudentID" HeaderText="StudentID"></asp:BoundColumn>
                            <asp:BoundColumn Visible="False" DataField="SOCID" HeaderText="SOCID"></asp:BoundColumn>--%>
                        </Columns>
                    </asp:DataGrid>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <asp:Button ID="BtnBack1" runat="server" Text="回上一頁" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="Hid_show_table" runat="server" />
        <input id="Hid_OCID" type="hidden" runat="server" />
        <input id="hid_SOCIDvalue" type="hidden" runat="server" />
    </form>
</body>
</html>
