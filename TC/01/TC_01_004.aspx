<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="TC_01_004.aspx.vb" Inherits="WDAIIP.TC_01_004" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>開班資料查詢</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-confirm.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery.blockUI.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        //function but_edit(ocid, id) {
        //    location.href = 'TC_01_004_add.aspx?ocid=' + ocid + '&ProcessType=Update&ID=' + id;
        //}

        //function but_del(ocid, PlanID, ComIDNO, SeqNO, Years, is_parent, id) {
        //    if (is_parent) {
        //        alert("此班級檔尚有與班級學員檔或排課檔或已有報名資料參照,不可刪除!!");
        //        return;
        //    }
        //    if (window.confirm("此動作會刪除班別資料，是否確定刪除?")) {
        //        location.href = 'TC_01_004_del.aspx?ocid=' + ocid + '&PlanID=' + PlanID + '&ComIDNO=' + ComIDNO + '&SeqNO=' + SeqNO + '&Years=' + Years + '&ID=' + id;
        //    }
        //}

        //20181030 (依照承辦人的增修需求,增加"主要職類"清除功能)
        function clearCareer() {
            document.getElementById('TB_career_id').value = '';
            document.getElementById('trainValue').value = '';
            document.getElementById('jobValue').value = '';
        }

        //20181030 (依照承辦人的增修需求,增加"通俗職類"清除功能)
        function clearCjob() {
            document.getElementById('txtCJOB_NAME').value = '';
            document.getElementById('cjobValue').value = '';
        }

        //20181030 (依照承辦人的增修需求,增加"開訓日期"清除功能)
        function clearTrainDate() {
            document.getElementById('start_date').value = '';
            document.getElementById('end_date').value = '';
        }

        //[全選／全不選]
        function doSelectAll(obj) {
            $("input[type=checkbox][data-role='chkItem']").prop("checked", obj.checked);
        }

        //批次設定1
        function doBatchSet1() {
            var blFlag2 = true;
            if ($("#OnShellDate").val() == "") { blFlag2 = false; }
            if (!blFlag2) {
                blockAlert("請先設定上架日期");
                return false;
            }
            var blFlag = false;
            $("input[type=checkbox][data-role='chkItem']").each(function () {
                if ($(this).prop("checked")) {
                    blFlag = true;
                    var tr = $(this).closest("tr");
                    $(tr).find("input[Name$=OnShellDate_i]").val($("#OnShellDate").val());
                    $(tr).find("select[Name$=OnShellDate_HR_i]").val($("#OnShellDate_HR").val());
                    $(tr).find("select[Name$=OnShellDate_MI_i]").val($("#OnShellDate_MI").val());
                }
            });
            if (blFlag) { return true; }
            blockAlert("尚未勾選班級");
            return false;
        }

        //批次設定2
        function doBatchSet2() {
            var blFlag2 = true;
            if ($("#SEnterDate").val() == "") { blFlag2 = false; }
            if (!blFlag2) {
                blockAlert("請先設定報名開始日期");
                return false;
            }
            var blFlag = false;
            $("input[type=checkbox][data-role='chkItem']").each(function () {
                if ($(this).prop("checked")) {
                    blFlag = true;
                    var tr = $(this).closest("tr");
                    $(tr).find("input[Name$=SEnterDate_i]").val($("#SEnterDate").val());
                    $(tr).find("select[Name$=SEnterDate_HR_i]").val($("#SEnterDate_HR").val());
                    $(tr).find("select[Name$=SEnterDate_MI_i]").val($("#SEnterDate_MI").val());
                }
            });
            if (blFlag) { return true; }
            blockAlert("尚未勾選班級");
            return false;
        }


        //批次設定3
        function doBatchSet3() {
            var blFlag2 = true;
            if ($("#FEnterDate").val() == "") { blFlag2 = false; }
            if (!blFlag2) {
                blockAlert("請先設定報名結束日期");
                return false;
            }
            var blFlag = false;
            $("input[type=checkbox][data-role='chkItem']").each(function () {
                if ($(this).prop("checked")) {
                    blFlag = true;
                    var tr = $(this).closest("tr");
                    $(tr).find("input[Name$=FEnterDate_i]").val($("#FEnterDate").val());
                    $(tr).find("select[Name$=FEnterDate_HR_i]").val($("#FEnterDate_HR").val());
                    $(tr).find("select[Name$=FEnterDate_MI_i]").val($("#FEnterDate_MI").val());
                }
            });
            if (blFlag) { return true; }
            blockAlert("尚未勾選班級");
            return false;
        }

    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <%--
        <input id="check_del" type="hidden" size="4" name="check_del" runat="server">
	    <input id="check_mod" type="hidden" size="4" name="check_mod" runat="server">
	    <input id="check_add" type="hidden" size="4" name="check_add" runat="server">
        --%>
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;開班資料設定&gt;&gt;開班資料查詢</asp:Label>
                </td>
            </tr>
        </table>
        <asp:Label ID="Label1" runat="server" Width="100%" CssClass="font">計畫轉入班級後，必須再查詢修改班級才可在學員動態查到班級資料</asp:Label><br>
        <asp:Label ID="Label2" runat="server" Width="100%" CssClass="font">計畫轉入班級後，必須再使用<FONT color="#990000">帳號班級賦予</FONT>才可在開班資料查詢查到班級資料</asp:Label><br>
        <table class="table_nw" width="100%" cellpadding="1" cellspacing="1" border="0">
            <tr>
                <td id="td6" width="14%" runat="server" class="bluecol">訓練機構</td>
                <td colspan="3" width="86%" class="whitecol">
                    <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                    <input id="Org" type="button" value="..." name="Org" runat="server" class="button_b_Mini">
                    <span id="HistoryList2" style="position: absolute; display: none">
                        <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                    </span>
                </td>
            </tr>
            <tr>
                <td id="td7" runat="server" class="bluecol">
                    <asp:Label ID="LabTMID" runat="server">訓練職類</asp:Label></td>
                <td colspan="3" class="whitecol">
                    <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Columns="30" Width="40%"></asp:TextBox>
                    <input id="career" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." name="career" runat="server" class="button_b_Mini">
                    <input id="RIDValue" style="width: 32px; height: 22px" type="hidden" name="RIDValue" runat="server">
                    <input id="trainValue" style="width: 33px; height: 22px" type="hidden" name="trainValue" runat="server">
                    <input id="jobValue" type="hidden" name="jobValue" runat="server">
                    <img style="cursor: pointer" onclick="javascript:clearCareer();" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" />
                </td>
            </tr>
            <tr>
                <td class="bluecol">
                    <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label></td>
                <td colspan="3" class="whitecol">
                    <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="35%"></asp:TextBox>
                    <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="button_b_Mini">
                    <input id="cjobValue" type="hidden" name="cjobValue" runat="server">
                    <img style="cursor: pointer" onclick="javascript:clearCjob();" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" />
                </td>
            </tr>
            <%--
                <tr id="trTrainX" runat="server">
                <td class="bluecol" width="14%">
                    <asp:Label ID="LabTPropertyID" runat="server">訓練性質</asp:Label></td>
                <td class="whitecol" width="36%">
                    --<asp:ListItem Value="0">職前</asp:ListItem><asp:ListItem Value="1">在職</asp:ListItem>--
                    --<asp:ListItem Value="2">接受委託</asp:ListItem>--
                    <asp:RadioButtonList ID="RB_TPropertyID" runat="server" Width="100%" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                        <asp:ListItem Value="1" Selected="True">在職</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>--%>
            <tr>
                <td class="bluecol" width="14%">班別代碼</td>
                <td class="whitecol" width="36%">
                    <asp:TextBox ID="ClassID" runat="server" Columns="20" MaxLength="20" Width="50%"></asp:TextBox></td>
                <td class="bluecol" width="14%">
                    <asp:Label ID="LabTPeriod" runat="server">訓練時段</asp:Label></td>
                <td class="whitecol" width="36%">
                    <asp:DropDownList ID="AppStage" runat="server"></asp:DropDownList>
                    <asp:DropDownList ID="TPeriod_List" runat="server"></asp:DropDownList></td>
            </tr>
            <tr>
                <td class="bluecol" width="14%">班級名稱</td>
                <td class="whitecol" width="36%">
                    <asp:TextBox ID="TB_ClassName" runat="server" Columns="30" MaxLength="30" Width="60%"></asp:TextBox></td>
                <td class="bluecol" width="14%">期別</td>
                <td class="whitecol" width="36%">
                    <asp:TextBox ID="TB_cycltype" runat="server" Columns="5" MaxLength="3" Width="50%"></asp:TextBox></td>
            </tr>
            <tr>
                <td id="td5" runat="server" class="bluecol" width="14%">開訓日期</td>
                <td class="whitecol" width="36%">
                    <asp:TextBox ID="start_date" Width="35%" onfocus="this.blur()" runat="server"></asp:TextBox>
                    <span id="span1" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>～
                    <asp:TextBox ID="end_date" Width="35%" onfocus="this.blur()" runat="server"></asp:TextBox>
                    <span id="span2" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    <img style="cursor: pointer" onclick="javascript:clearTrainDate();" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" />
                </td>
                <td class="bluecol" width="14%">開訓狀態</td>
                <td class="whitecol" width="36%">
                    <asp:RadioButtonList ID="ClassState" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                        <asp:ListItem Value="不區分" Selected="True">不區分</asp:ListItem>
                        <asp:ListItem Value="已開訓">已開訓</asp:ListItem>
                        <asp:ListItem Value="未開訓">未開訓</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">開班狀態</td>
                <td colspan="3" class="whitecol">
                    <asp:RadioButtonList ID="NotOpen" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                        <asp:ListItem Value="N" Selected="True">開班</asp:ListItem>
                        <asp:ListItem Value="Y">不開班</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
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
                <td colspan="4" class="whitecol">
                    <div align="center">
                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                        <asp:Button ID="bt_search" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="bt_EXPORT" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                    </div>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                    </div>
                </td>
            </tr>
            <%-- 2018 add 批次設定上架日 --%>
            <tr id="tr_SetOnShellDate1" runat="server">
                <td class="bluecol">上架日期</td>
                <td colspan="3" class="whitecol">
                    <asp:TextBox ID="OnShellDate" Columns="20" onfocus="this.blur()" runat="server"></asp:TextBox>
                    <span id="sp_imgOnShellDate" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= OnShellDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                    </span>
                    <asp:DropDownList ID="OnShellDate_HR" runat="server" AppendDataBoundItems="True"></asp:DropDownList>時：
                        <asp:DropDownList ID="OnShellDate_MI" runat="server" AppendDataBoundItems="True"></asp:DropDownList>分
                        <input type="button" class="asp_button_M" onclick="doBatchSet1()" value="批次設定" />
                </td>
            </tr>
            <tr id="tr_setSEnterDate" runat="server">
                <td class="bluecol">報名開始日期</td>
                <td colspan="3" class="whitecol">
                    <asp:TextBox ID="SEnterDate" Columns="20" onfocus="this.blur()" runat="server"></asp:TextBox>
                    <span id="sp_imgSEnterDate" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= SEnterDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                    </span>
                    <asp:DropDownList ID="SEnterDate_HR" runat="server" AppendDataBoundItems="True"></asp:DropDownList>時：
                        <asp:DropDownList ID="SEnterDate_MI" runat="server" AppendDataBoundItems="True"></asp:DropDownList>分
                        <input type="button" class="asp_button_M" onclick="doBatchSet2()" value="批次設定" />
                </td>
            </tr>
            <tr id="tr_setFEnterDate" runat="server">
                <td class="bluecol">報名結束日期</td>
                <td colspan="3" class="whitecol">
                    <asp:TextBox ID="FEnterDate" Columns="20" onfocus="this.blur()" runat="server"></asp:TextBox>
                    <span id="sp_imgFEnterDate" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= FEnterDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                    </span>
                    <asp:DropDownList ID="FEnterDate_HR" runat="server" AppendDataBoundItems="True"></asp:DropDownList>時：
                        <asp:DropDownList ID="FEnterDate_MI" runat="server" AppendDataBoundItems="True"></asp:DropDownList>分
                        <input type="button" class="asp_button_M" onclick="doBatchSet3()" value="批次設定" />
                </td>
            </tr>

        </table>

        <%--<table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0"></table>--%>
        <asp:Panel ID="Panel_ClassInfo" runat="server" Width="100%">
            <asp:DataGrid ID="DG_ClassInfo" runat="server" CssClass="font" Width="100%" AllowSorting="True" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                <Columns>
                    <asp:TemplateColumn>
                        <HeaderStyle Width="5%" />
                        <ItemStyle HorizontalAlign="Center" />
                        <HeaderTemplate>
                            <label>選取</label><br />
                            <input type='checkbox' onclick='doSelectAll(this)'>
                        </HeaderTemplate>
                        <ItemTemplate>
                            <input type="checkbox" id="chkItem" data-role="chkItem" runat="server" />
                            <asp:HiddenField ID="hidOCID" runat="server" />
                            <asp:HiddenField ID="hidSTDATE" runat="server" />
                            <asp:HiddenField ID="hidFTDATE" runat="server" />
                        </ItemTemplate>
                    </asp:TemplateColumn>
                    <asp:BoundColumn HeaderText="序號">
                        <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <%--<asp:BoundColumn HeaderText="管控&lt;br&gt;單位"><HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle></asp:BoundColumn>--%>
                    <asp:BoundColumn DataField="OrgName" SortExpression="OrgName" HeaderText="訓練機構">
                        <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn HeaderText="班別代碼">
                        <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                    </asp:BoundColumn>
                    <%--<asp:BoundColumn HeaderText="班數"> <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle> <ItemStyle HorizontalAlign="Center"></ItemStyle> </asp:BoundColumn>--%>
                    <asp:BoundColumn HeaderText="開結訓日">
                        <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ClassCName" HeaderText="班別名稱">
                        <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="TrainName" HeaderText="訓練職類"><HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle></asp:BoundColumn>
                    <asp:BoundColumn DataField="CJOB_NAME" HeaderText="通俗職類"><HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle></asp:BoundColumn>
                    <%--<asp:BoundColumn Visible="False" DataField="TPropertyID" HeaderText="訓練性質"><HeaderStyle HorizontalAlign="Center"></HeaderStyle></asp:BoundColumn>
                       <asp:BoundColumn Visible="False" DataField="HourRanName" HeaderText="訓練時段"><HeaderStyle HorizontalAlign="Center"></HeaderStyle></asp:BoundColumn>--%>
                    <asp:TemplateColumn HeaderText="上架日期">
                        <HeaderStyle Width="15%" />
                        <ItemStyle CssClass="whitecol" />
                        <ItemTemplate>
                            <asp:HiddenField ID="hidOnShellDate" runat="server" />
                            <asp:HiddenField ID="hidOnShellDate_HR" runat="server" />
                            <asp:HiddenField ID="hidOnShellDate_MI" runat="server" />
                            <asp:TextBox ID="OnShellDate_i" Columns="15" onfocus="this.blur()" runat="server"></asp:TextBox>
                            <img id="imgOnShellDate_i" runat="server" style="cursor: pointer" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"><br />
                            <asp:DropDownList ID="OnShellDate_HR_i" runat="server" AppendDataBoundItems="True"></asp:DropDownList>時：
                            <asp:DropDownList ID="OnShellDate_MI_i" runat="server" AppendDataBoundItems="True"></asp:DropDownList>分
                        </ItemTemplate>
                    </asp:TemplateColumn>

                    <asp:TemplateColumn HeaderText="報名開始日期">
                        <HeaderStyle Width="15%" />
                        <ItemStyle CssClass="whitecol" />
                        <ItemTemplate>
                            <asp:HiddenField ID="hidSEnterDate" runat="server" />
                            <asp:HiddenField ID="hidSEnterDate_HR" runat="server" />
                            <asp:HiddenField ID="hidSEnterDate_MI" runat="server" />
                            <asp:TextBox ID="SEnterDate_i" Columns="15" onfocus="this.blur()" runat="server"></asp:TextBox>
                            <img id="imgSEnterDate_i" runat="server" style="cursor: pointer" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"><br />
                            <asp:DropDownList ID="SEnterDate_HR_i" runat="server" AppendDataBoundItems="True"></asp:DropDownList>時：
                            <asp:DropDownList ID="SEnterDate_MI_i" runat="server" AppendDataBoundItems="True"></asp:DropDownList>分
                        </ItemTemplate>
                    </asp:TemplateColumn>

                    <asp:TemplateColumn HeaderText="報名結束日期">
                        <HeaderStyle Width="15%" />
                        <ItemStyle CssClass="whitecol" />
                        <ItemTemplate>
                            <asp:HiddenField ID="hidFEnterDate" runat="server" />
                            <asp:HiddenField ID="hidFEnterDate_HR" runat="server" />
                            <asp:HiddenField ID="hidFEnterDate_MI" runat="server" />
                            <asp:TextBox ID="FEnterDate_i" Columns="15" onfocus="this.blur()" runat="server"></asp:TextBox>
                            <img id="imgFEnterDate_i" runat="server" style="cursor: pointer" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"><br />
                            <asp:DropDownList ID="FEnterDate_HR_i" runat="server" AppendDataBoundItems="True"></asp:DropDownList>時：
                            <asp:DropDownList ID="FEnterDate_MI_i" runat="server" AppendDataBoundItems="True"></asp:DropDownList>分
                        </ItemTemplate>
                    </asp:TemplateColumn>

                    <asp:TemplateColumn HeaderText="功能">
                        <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        <ItemTemplate>
                            <asp:LinkButton ID="lbtEdit" runat="server" Text="修改" CommandName="edit" CssClass="linkbutton"></asp:LinkButton>
                            <asp:LinkButton ID="lbtDel" runat="server" Text="刪除" CommandName="del" CssClass="linkbutton"></asp:LinkButton>
                            <asp:LinkButton ID="lbtExport" runat="server" Text="匯出" CommandName="add" CssClass="asp_Export_M"></asp:LinkButton>
                        </ItemTemplate>
                    </asp:TemplateColumn>
                </Columns>
                <PagerStyle Visible="False"></PagerStyle>
            </asp:DataGrid>
            <table id="search_tbl" class="font" border="0" cellspacing="0" cellpadding="0" width="100%" runat="server">
                <tr>
                    <td>
                        <div align="center">
                            <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                        </div>
                    </td>
                </tr>
            </table>
            <table width="100%" id="tbSave" runat="server">
                <tr>
                    <td class="whitecol">
                        <div align="center">
                            <asp:Button ID="btnSave" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                        </div>
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <input id="isBlack" type="hidden" name="isBlack" runat="server" />
        <input id="orgname" type="hidden" name="orgname" runat="server" />
    </form>
</body>
</html>
