<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CO_01_001.aspx.vb" Inherits="WDAIIP.CO_01_001" %>

<%--<!DOCTYPE html>--%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>函送資料登錄作業</title>
    <%--<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />--%>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-confirm.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery.blockUI.js"></script>
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="../../js/TIMS.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        //function choose_class() {
        //    var jRIDValue = $("#RIDValue");
        //    wopen('../../SD/02/SD_02_ch.aspx?RID=' + jRIDValue.val(), 'Class', 540, 520, 1);
        //}
        function changSTATUS1() {
            $('[name="STATUS1"]').removeAttr('checked');
            $('[name="OVERWEEK1"]').removeAttr('checked');
            return false;
        }

        function changSTATUS2() {
            var Hid_stdate = $("#Hid_stdate");
            var Hid_stdate14 = $("#Hid_stdate14");
            var lab_msg_2 = $("#lab_msg_2");
            lab_msg_2.html("未檢核"); //lab_msg_2.val("未檢核");
            if (Hid_stdate14.val() != '' && $("#SENDDATE2").val() != '') {
                var stDate = new Date(Hid_stdate.val());
                var beginDate = new Date(Hid_stdate14.val());
                var endDate = new Date($("#SENDDATE2").val());
                $('[name="STATUS2"]').removeAttr('checked');
                $('[name="OVERWEEK2"]').removeAttr('checked');
                if (beginDate >= endDate) {
                    lab_msg_2.html("未逾期"); //lab_msg_2.val("未逾期");
                    $("input[name=STATUS2][value=1]").prop('checked', true);
                }
                else {
                    lab_msg_2.html("已逾期"); //lab_msg_2.val("已逾期");
                    $("input[name=STATUS2][value=2]").prop('checked', true);
                }
            }
            //alert(lab_msg_2.val());
            return false;
        }

        function changSTATUS3() {
            var Hid_ftdate = $("#Hid_ftdate");
            var Hid_ftdate21 = $("#Hid_ftdate21");
            var lab_msg_3 = $("#lab_msg_3");
            lab_msg_3.html("未檢核"); //lab_msg_2.val("未檢核");
            if (Hid_ftdate21.val() != '' && $("#SENDDATE3").val() != '') {
                var ftDate = new Date(Hid_ftdate.val());
                var beginDate = new Date(Hid_ftdate21.val());
                var endDate = new Date($("#SENDDATE3").val());
                $('[name="STATUS3"]').removeAttr('checked');
                $('[name="OVERWEEK3"]').removeAttr('checked');
                if (beginDate >= endDate) {
                    lab_msg_3.html("未逾期"); //lab_msg_2.val("未逾期");
                    $("input[name=STATUS3][value=1]").prop('checked', true);
                }
                else {
                    lab_msg_3.html("已逾期"); //lab_msg_2.val("已逾期");
                    $("input[name=STATUS3][value=2]").prop('checked', true);
                }
            }
            //alert(lab_msg_3.val());
            return false;
        }

        function changSTATUS4() {
            $('#ChkboxSave_1').prop('checked', true);
            $('[name="SENDDATE1"]').val("");
            $('[name="STATUS1"]').removeAttr('checked');
            //$('[name="ISPASS1"]').removeAttr('checked');
            $('[name="OVERWEEK1"]').removeAttr('checked');
            return false;
        }

        $(function () {
            //$("#SENDDATE3").change(function () {
            //    //alert('c');
            //    debugger;
            //    $('[name="STATUS3"]').removeAttr('checked');
            //    $("input[name=STATUS3][value=2]").prop('checked', true);
            //});
            //$("#SENDDATE3").on("change", function () {
            //    //alert('c');
            //    debugger;
            //    $('[name="STATUS3"]').removeAttr('checked');
            //    $("input[name=STATUS3][value=2]").prop('checked', true);
            //});
            $("#btnClsOption1").on("click", function () { return changSTATUS4(); });
            $("#SENDDATE1").on("click", function () { return changSTATUS1(); });
            //$("#SENDDATE2").on("change", function () { changSTATUS2(); });
            //$("#SENDDATE2").on("input", function () { changSTATUS2(); });
            $("#SENDDATE2").on("click", function () { return changSTATUS2(); });
            $("#btnChkSend2").on("click", function () { return changSTATUS2(); });

            //$("#SENDDATE3").on("change", function () { changSTATUS3(); });
            //$("#SENDDATE3").on("input", function () { changSTATUS3(); });
            $("#SENDDATE3").on("click", function () { return changSTATUS3(); });
            $("#btnChkSend3").on("click", function () { return changSTATUS3(); });
            //debugger;

        });
        //$(document).ready(function () {
        //});
    </script>
    <style type="text/css">
        .red-style1 { color: #FF0000; padding: 4px; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;審查計分表&gt;&gt;函送資料登錄作業</asp:Label>
                </td>
            </tr>
        </table>
        <%--style="display: none"--%>
        <div id="divSch1" runat="server">
            <table class="table_sch" cellpadding="1" cellspacing="1" width="100%" border="0">
                <tr>
                    <td class="bluecol" style="width: 20%">訓練機構</td>
                    <td colspan="3" class="whitecol">
                        <asp:TextBox ID="center" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox>
                        <input id="Org" type="button" value="..." name="Org" runat="server">
                        <%--<input onclick="choose_class()" type="button" value="..." class="button_b_Mini">--%>
                        <input id="RIDValue" style="width: 32px; height: 22px" type="hidden" name="RIDValue" runat="server">
                        <input id="Orgidvalue" style="width: 32px; height: 22px" type="hidden" name="Orgidvalue" runat="server">
                        <span id="HistoryList2" style="position: absolute; display: none">
                            <asp:Table ID="HistoryRID" runat="server" Width="100%">
                            </asp:Table>
                        </span>
                    </td>
                </tr>
                <%--<tr>
                    <td class="bluecol" style="width: 20%">
                        <asp:Label ID="LabTMID" runat="server">訓練職類</asp:Label>
                    </td>
                    <td colspan="3" class="whitecol">
                        <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Columns="30" Width="40%"></asp:TextBox>
                        <input id="btu_sel" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." name="btu_sel" runat="server">&nbsp;
							<input id="trainValue" type="hidden" name="trainValue" runat="server">
                        <input id="jobValue" type="hidden" name="jobValue" runat="server">
                    </td>
                </tr>
                <tr>
                    <td class="bluecol" style="width: 20%">
                        <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label>
                    </td>
                    <td colspan="3" class="whitecol">
                        <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="40%"></asp:TextBox>
                        <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server">
                        <input id="cjobValue" type="hidden" name="cjobValue" runat="server">
                    </td>
                </tr>--%>
                <tr>
                    <td class="bluecol" style="width: 20%">班級名稱
                    </td>
                    <td class="whitecol" style="width: 30%">
                        <asp:TextBox ID="ClassName" runat="server" Columns="50" MaxLength="50"></asp:TextBox>
                    </td>
                    <td class="bluecol" style="width: 20%">期別
                    </td>
                    <td class="whitecol" style="width: 30%">
                        <asp:TextBox ID="CyclType" runat="server" Columns="5" MaxLength="3" Width="40%"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">開訓日期                    </td>
                    <td colspan="3" class="whitecol">
                        <asp:TextBox ID="STDATE1" Width="18%" onfocus="this.blur()" runat="server" MaxLength="10"></asp:TextBox>
                        <span runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= STDATE1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= STDATE1.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" />
                        </span>
                        ～
					        <asp:TextBox ID="STDATE2" Width="18%" onfocus="this.blur()" runat="server" MaxLength="10"></asp:TextBox>
                        <span runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= STDATE2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= STDATE2.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" />
                        </span>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need">申請階段</td>
                    <td class="whitecol" colspan="3">
                        <asp:DropDownList ID="sch_ddlAPPSTAGE" runat="server"></asp:DropDownList></td>
                </tr>
            </table>
            <table width="100%">
                <tr>
                    <td class="whitecol" align="center">
                        <div align="center">
                            <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                            <asp:TextBox ID="TxtPageSize" runat="server" MaxLength="2" Width="55px">10</asp:TextBox>
                            <asp:Button ID="btnQuery" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                        </div>
                    </td>
                </tr>
            </table>
            <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                <tr>
                    <td align="center">
                        <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td align="center">
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <%--<HeaderStyle ForeColor="#00ffff"></HeaderStyle>--%>
                                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AllowSorting="True" PagerStyle-HorizontalAlign="Left"
                                        PagerStyle-Mode="NumericPages" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                                        <Columns>
                                            <asp:BoundColumn HeaderText="序號" HeaderStyle-Width="5%">
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="APPSTAGE_N" HeaderText="申請階段"  HeaderStyle-Width="10%">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="STDate" HeaderText="訓練起日" HeaderStyle-Width="10%">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="FTDate" HeaderText="訓練迄日" HeaderStyle-Width="10%">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="OrgName" HeaderText="機構名稱" HeaderStyle-Width="22%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="CLASSCNAME2" HeaderText="班名" HeaderStyle-Width="25%"></asp:BoundColumn>
                                            <%-- <asp:BoundColumn DataField="CJOBNAME" HeaderText="通俗職類" HeaderStyle-Width="10%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="JOBNAME" HeaderText="訓練業別" HeaderStyle-Width="10%"></asp:BoundColumn>--%>
                                            <asp:TemplateColumn HeaderText="功能" HeaderStyle-Width="8%">
                                                <ItemStyle HorizontalAlign="Center" Font-Size="Small" />
                                                <ItemTemplate>
                                                    <asp:LinkButton ID="lbtEdit" runat="server" Text="編輯" CommandName="btnEdit" CssClass="linkbutton"></asp:LinkButton>
                                                    <%--<asp:LinkButton ID="lbtDel1" runat="server" Text="刪除" CommandName="btnDel" CssClass="linkbutton"></asp:LinkButton>--%>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>

                                        </Columns>
                                        <PagerStyle Visible="False" HorizontalAlign="Left" ForeColor="Blue" Position="Top" Mode="NumericPages"></PagerStyle>
                                    </asp:DataGrid>
                                </td>
                            </tr>

                        </table>
                    </td>
                </tr>
                <tr>
                    <td align="center">
                        <asp:Label ID="msg1" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                    </td>
                </tr>
            </table>
        </div>

        <div id="divEdt1" runat="server">
            <table class="table_nw" cellpadding="1" cellspacing="1" width="100%" border="0">
                <tr>
                    <td class="bluecol" width="18%">訓練機構
                    </td>
                    <td class="whitecol" width="32%">
                        <asp:Label ID="LabOrgName" runat="server"></asp:Label>
                    </td>
                    <td class="bluecol" width="18%">班別代碼
                    </td>
                    <td class="whitecol" width="32%">
                        <asp:Label ID="LabClassID" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">班級名稱
                    </td>
                    <td class="whitecol">
                        <asp:Label ID="LabCLASSCNAME" runat="server"></asp:Label>
                    </td>
                    <td class="bluecol">期別
                    </td>
                    <td class="whitecol">
                        <asp:Label ID="LabCyclType" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">通俗職類
                    </td>
                    <td class="whitecol">
                        <asp:Label ID="LabCJOBNAME" runat="server"></asp:Label>
                    </td>
                    <td class="bluecol">訓練業別
                    </td>
                    <td class="whitecol">
                        <asp:Label ID="LabJOBNAME" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">開結訓日期區間
                    </td>
                    <td class="whitecol" colspan="3">
                        <asp:Label ID="LabSFTDATE" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <%--<asp:LinkButton ID="lbtDel1" runat="server" Text="刪除" CommandName="btnDel" CssClass="linkbutton"></asp:LinkButton>--%>
                    <td colspan="4" class="table_title">函送資料時效及正確性</td>
                </tr>
                <%--<asp:LinkButton ID="lbtDel1" runat="server" Text="刪除" CommandName="btnDel" CssClass="linkbutton"></asp:LinkButton>--%>
                <tr>
                    <td colspan="4" class="class_title2">自製招訓資料</td>
                </tr>
                <tr>
                    <td class="bluecol_need">勾選(儲存)</td>
                    <td class="whitecol">
                        <asp:CheckBox ID="ChkboxSave_1" runat="server" /></td>
                    <td class="whitecol">
                        <asp:Button ID="btnClsOption1" runat="server" Text="清除自製招訓選項" /></td>
                    <td class="whitecol"><span class="red-style1">(點選清除後，請重新勾選儲存)</span></td>
                </tr>
                <tr>
                    <td class="bluecol">自製招訓資料函送日期
                    </td>
                    <td class="whitecol">
                        <span runat="server">
                            <asp:TextBox ID="SENDDATE1" Width="111px" onfocus="this.blur()" runat="server" MaxLength="10"></asp:TextBox>
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= SENDDATE1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                        </span>
                        <label id="labmsg_1"></label>
                    </td>
                    <td class="bluecol">自製招訓資料函送狀態(2-1-1)
                    </td>
                    <td class="whitecol">
                        <asp:RadioButtonList ID="STATUS1" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                            <asp:ListItem Value="1">依規定辦理</asp:ListItem>
                            <asp:ListItem Value="2">逾期(扣分)</asp:ListItem>
                            <asp:ListItem Value="3">逾期(不扣分)</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">資料內容符合規定(2-1-2)
                    </td>
                    <td class="whitecol">
                        <%--<asp:LinkButton ID="lbtDel1" runat="server" Text="刪除" CommandName="btnDel" CssClass="linkbutton"></asp:LinkButton>--%>
                        <asp:DropDownList ID="ddlISPASS1" runat="server"></asp:DropDownList>
                    </td>
                    <td class="bluecol">逾期週數(2-1-1)</td>
                    <td class="whitecol">
                        <asp:RadioButtonList ID="OVERWEEK1" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                            <asp:ListItem Value="1">逾期1週</asp:ListItem>
                            <asp:ListItem Value="2">逾期1週至2週以內</asp:ListItem>
                            <asp:ListItem Value="3">逾期2週</asp:ListItem>
                            <asp:ListItem Value="4">停辦</asp:ListItem>
                            <asp:ListItem Value="9">無逾期</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>


                <%--<asp:LinkButton ID="lbtDel1" runat="server" Text="刪除" CommandName="btnDel" CssClass="linkbutton"></asp:LinkButton>--%>
                <tr>
                    <td colspan="4" class="class_title2">開訓資料</td>
                </tr>
                <tr>
                    <td class="bluecol_need">勾選(儲存)</td>
                    <td colspan="3">
                        <asp:CheckBox ID="ChkboxSave_2" runat="server" />
                        <asp:Label ID="lab_msg_2" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">開訓資料函送日期
                    </td>
                    <td class="whitecol">
                        <span runat="server">
                            <asp:TextBox ID="SENDDATE2" Width="111px" onfocus="this.blur()" runat="server" MaxLength="10"></asp:TextBox>
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= SENDDATE2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                        </span>
                        <asp:Button ID="btnChkSend2" runat="server" Text="檢核" />
                        <%--<asp:LinkButton ID="lbtDel1" runat="server" Text="刪除" CommandName="btnDel" CssClass="linkbutton"></asp:LinkButton>--%>
                    </td>
                    <td class="bluecol">開訓資料函送狀態(2-1-1)
                    </td>
                    <td class="whitecol">
                        <asp:RadioButtonList ID="STATUS2" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                            <asp:ListItem Value="1">依規定辦理</asp:ListItem>
                            <asp:ListItem Value="2">逾期(扣分)</asp:ListItem>
                            <asp:ListItem Value="3">逾期(不扣分)</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">資料內容符合規定(2-1-2)
                    </td>
                    <td class="whitecol">
                        <%--<asp:LinkButton ID="lbtDel1" runat="server" Text="刪除" CommandName="btnDel" CssClass="linkbutton"></asp:LinkButton>--%>
                        <asp:DropDownList ID="ddlISPASS2" runat="server"></asp:DropDownList>
                    </td>
                    <td class="bluecol">逾期週數(2-1-1)</td>
                    <td class="whitecol">
                        <asp:RadioButtonList ID="OVERWEEK2" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                            <asp:ListItem Value="1">逾期1週</asp:ListItem>
                            <asp:ListItem Value="2">逾期1週至2週以內</asp:ListItem>
                            <asp:ListItem Value="3">逾期2週</asp:ListItem>
                            <asp:ListItem Value="4">停辦</asp:ListItem>
                            <asp:ListItem Value="9">無逾期</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>


                <%--table_title--%>
                <tr>
                    <td colspan="4" class="class_title2">結訓資料</td>
                </tr>
                <tr>
                    <td class="bluecol_need">勾選(儲存)</td>
                    <td colspan="3">
                        <asp:CheckBox ID="ChkboxSave_3" runat="server" />
                        <asp:Label ID="lab_msg_3" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">結訓資料函送日期
                    </td>
                    <td class="whitecol">
                        <span runat="server">
                            <asp:TextBox ID="SENDDATE3" Width="111px" onfocus="this.blur()" runat="server" MaxLength="10"></asp:TextBox>
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= SENDDATE3.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                        </span>
                        <asp:Button ID="btnChkSend3" runat="server" Text="檢核" />
                        <%--自製招訓資料--%>
                    </td>
                    <td class="bluecol">結訓資料函送狀態(2-1-1)
                    </td>
                    <td class="whitecol">
                        <asp:RadioButtonList ID="STATUS3" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                            <asp:ListItem Value="1">依規定辦理</asp:ListItem>
                            <asp:ListItem Value="2">逾期(扣分)</asp:ListItem>
                            <asp:ListItem Value="3">逾期(不扣分)</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">資料內容符合規定(2-1-2)
                    </td>
                    <td class="whitecol">
                        <%--<asp:RadioButtonList ID="ISPASS1" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                            <asp:ListItem Value="Y">符合規定</asp:ListItem><asp:ListItem Value="N">不符合規定</asp:ListItem></asp:RadioButtonList>--%>
                        <asp:DropDownList ID="ddlISPASS3" runat="server"></asp:DropDownList>
                    </td>
                    <td class="bluecol">逾期週數(2-1-1)</td>
                    <td class="whitecol">
                        <asp:RadioButtonList ID="OVERWEEK3" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                            <asp:ListItem Value="1">逾期1週</asp:ListItem>
                            <asp:ListItem Value="2">逾期1週至2週以內</asp:ListItem>
                            <asp:ListItem Value="3">逾期2週</asp:ListItem>
                            <asp:ListItem Value="4">停辦</asp:ListItem>
                            <asp:ListItem Value="9">無逾期</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td class="whitecol" colspan="4">
                        <div align="center">
                            <%--開訓資料--%>
                            <asp:Button ID="BtnBack1" runat="server" Text="回上頁" CssClass="asp_button_M"></asp:Button>
                            <asp:Button ID="BtnSaveData1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td class="whitecol" colspan="4">
                        <div align="center">
                            <span style="color: red;">*勾選(儲存)該區塊資料才會儲存</span>
                        </div>
                    </td>
                </tr>
            </table>
        </div>

        <asp:HiddenField ID="Hid_CSCID" runat="server" />
        <asp:HiddenField ID="Hid_OCID" runat="server" />
        <asp:HiddenField ID="Hid_stdate" runat="server" />
        <asp:HiddenField ID="Hid_stdate14" runat="server" />
        <asp:HiddenField ID="Hid_ftdate" runat="server" />
        <asp:HiddenField ID="Hid_ftdate21" runat="server" />

        <%--<label id="labmsg_2"></label>--%><%--<asp:RadioButtonList ID="ISPASS2" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                        <asp:ListItem Value="Y">符合規定</asp:ListItem><asp:ListItem Value="N">不符合規定</asp:ListItem></asp:RadioButtonList>--%>
    </form>
</body>
</html>
