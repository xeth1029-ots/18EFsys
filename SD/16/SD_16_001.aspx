<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_16_001.aspx.vb" Inherits="WDAIIP.SD_16_001" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>學員處分功能</title>
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
    <script type="text/javascript" language="javascript">
        function GETvalue() { document.getElementById('BtnGETvalue2').click(); }
        //document.getElementById('DataGridTable').style.display='none';
        //if (document.getElementById('OCID1').value=='')
        //{  document.getElementById('Button7').click();}
        function choose_class() { openClass('../02/SD_02_ch.aspx?RID=' + document.getElementById('RIDValue').value); }

        function checkSave() {
            //debugger;
            var msg = '';
            var vHid_LID = document.getElementById('Hid_LID').value;
            if (document.getElementById('RIDValue').value == "") { msg += '請選擇【訓練機構】\n'; }
            if (vHid_LID != '0' && document.getElementById('OCIDValue1').value == "") { msg += '請選擇【職類/班別】\n'; }
            if (document.getElementById('txt_idno').value == "") { msg += '請填寫【身分證號碼】\n'; }
            if (document.getElementById('txt_No').value == "") { msg += '請選擇【處分文號】\n'; }
            if (isEmpty(document.form1.ddlSBTERMS)) { msg += '請選擇處分緣由\n'; }

            if (document.getElementById('txt_SBSdate').value == "") { msg += '請選擇【處分日期】\n'; }
            if (document.getElementById('ddl_SBYears').selectedIndex == 0) { msg += '請選擇【處分年限】\n'; }
            if (document.getElementById('txt_SBComment').value == "") { msg += '請填寫【處分事由】\n'; }
            else {
                if (checkMaxLen(document.getElementById('txt_SBComment').value, 300 * 2)) { msg += '【處分事由】長度不可超過300字元\n'; }
            }
            if (msg != '') {
                window.alert(msg);
                return false;
            } else {
                msg = '';
                msg += '\n請確認資料是否無誤,儲存後資料將不可修改\n\n';
                msg += '如確認資料無誤後,請按下確定,謝謝!!\n';
                return confirm(msg);
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" id="Table1" width="740">
                        <tr>
                            <td class="font">
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;學員處分功能</asp:Label>
                                <asp:Label ID="lbl_title" runat="server">學員處分功能</asp:Label>
                            </td>
                        </tr>
                    </table>
                    <asp:Panel ID="Panel1" runat="server" Visible="True">
                        <table id="Table3" class="table_sch" cellspacing="1" cellpadding="1">
                            <tr>
                                <td class="bluecol" style="width: 20%">計畫別 </td>
                                <td class="whitecol" colspan="3">
                                    <asp:DropDownList ID="ddlTPlanIDSch" runat="server">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" style="width: 20%">原處分分署 </td>
                                <td class="whitecol" style="width: 30%">
                                    <asp:DropDownList ID="DistID" runat="server">
                                    </asp:DropDownList>
                                </td>
                                <td class="bluecol" style="width: 20%">處分年度 </td>
                                <td class="whitecol" style="width: 30%">
                                    <asp:DropDownList ID="Years" runat="server">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">身分證號碼 </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="IDNO" runat="server" Width="40%" MaxLength="10"></asp:TextBox>
                                </td>
                                <td class="bluecol">學員姓名 </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="Name" runat="server" Width="40%" MaxLength="50"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">作業顯示模式</td>
                                <td class="whitecol" colspan="3">
                                    <asp:RadioButtonList Style="z-index: 0" ID="rblWorkMode" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1" Selected="True">模糊顯示</asp:ListItem>
                                        <asp:ListItem Value="2">正常顯示</asp:ListItem>
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
                              <tr id="tr_ddl_INQUIRY_S" runat="server">
                            <td class="bluecol_need">查詢原因</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="ddl_INQUIRY_Sch" runat="server"></asp:DropDownList>
                            </td>
                        </tr>
                        </table>
                        <table width="100%" class="font">
                            <tr>
                                <td align="center" class="whitecol">
                                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                    <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                    <asp:Button ID="btn_Sch" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                    &nbsp;<asp:Button ID="btn_Add" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>
                                    <%--<input id="Sch_Mark" value="0" type="hidden" name="DistValue" runat="server">--%>
								&nbsp;<asp:Button ID="btnExport1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                                </td>
                            </tr>
                        </table>
                        <div style="width: 100%; text-align: center">
                            <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
                        </div>
                        <table id="tb_Sch" border="0" cellspacing="0" cellpadding="0" width="100%" runat="server" class="font">
                            <tr>
                                <td align="center">
                                    <div id="Div1" runat="server">
                                        <asp:DataGrid ID="dg_Sch" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                                            <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                            <Columns>
                                                <asp:BoundColumn HeaderText="序號">
                                                    <HeaderStyle HorizontalAlign="Center" Width="3%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="PlanName" HeaderText="計畫別">
                                                    <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="DistName" HeaderText="原處分分署">
                                                    <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="name" HeaderText="學員姓名">
                                                    <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:TemplateColumn HeaderText="身分證號碼">
                                                    <HeaderStyle HorizontalAlign="Center" Width="11%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:Label ID="labIDNO" runat="server"></asp:Label>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:TemplateColumn HeaderText="處分緣由">
                                                    <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:Label ID="labSBTERMS" runat="server"></asp:Label>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:BoundColumn DataField="SBSdate" HeaderText="處分起日">
                                                    <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="SBYears" HeaderText="年限">
                                                    <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="SBComment" HeaderText="事由">
                                                    <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:TemplateColumn HeaderText="功能">
                                                    <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:LinkButton ID="lbtView" runat="server" CommandName="view" CssClass="linkbutton">檢視</asp:LinkButton>
                                                        <asp:LinkButton ID="lbtEdit" runat="server" CommandName="edit" CssClass="linkbutton">修改</asp:LinkButton>
                                                        <asp:LinkButton ID="lbtDel" runat="server" CommandName="del" CssClass="asp_Export_M">刪除</asp:LinkButton>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                            </Columns>
                                            <PagerStyle Visible="False"></PagerStyle>
                                        </asp:DataGrid>
                                    </div>
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                    <%--<asp:BoundColumn Visible="False" DataField="SBSN" HeaderText="流水號"></asp:BoundColumn>--%>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                    <asp:Panel ID="Panel2" runat="server" Visible="False">
                        <table class="table_sch" cellpadding="1" cellspacing="1">
                            <tr>
                                <td class="bluecol" style="width: 20%">計畫別 </td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="ddlTPlanID" runat="server" AutoPostBack="True">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">原處分分署 </td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="ddl_DistID" runat="server"></asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">訓練機構 </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="center" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox>
                                    <input id="RIDValue" type="hidden" runat="server" />
                                    <input id="Button6" value="..." type="button" name="Button6" runat="server" />
                                    <asp:Button Style="display: none" ID="BtnGETvalue2" runat="server"></asp:Button>
                                    <span style="position: absolute; display: none" id="HistoryList2" onclick="GETvalue()">
                                        <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">類別/班別 </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                    <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                    <input onclick="choose_class()" value="..." type="button" />
                                    <input id="TMIDValue1" type="hidden" runat="server" />
                                    <input id="OCIDValue1" type="hidden" runat="server" />
                                    <span style="z-index: 1; position: absolute; display: none; left: 270px" id="HistoryList">
                                        <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                    </span>
                                    <iframe style="position: absolute; display: none; left: 270px" id="FrameObj" height="0" frameborder="0" width="100%"></iframe>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">身分證號碼 </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_idno" runat="server" Width="20%" MaxLength="10"></asp:TextBox>
                                    <asp:Label ID="Label1" runat="server" Text="  (或居留證號碼)"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">處分文號 </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_No" runat="server" Width="300px" MaxLength="30"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">處分緣由 </td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="ddlSBTERMS" runat="server">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">處分日期 </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_SBSdate" runat="server" Width="15%" onfocus="this.blur()" MaxLength="10"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= txt_SBSdate.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">處分年限 </td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="ddl_SBYears" runat="server">
                                        <asp:ListItem Value="-1" Selected="True">請選擇</asp:ListItem>
                                        <asp:ListItem Value="0">0</asp:ListItem>
                                        <asp:ListItem Value="1">1</asp:ListItem>
                                        <asp:ListItem Value="2">2</asp:ListItem>
                                        <asp:ListItem Value="3">3</asp:ListItem>
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">處分事由 </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_SBComment" runat="server" Width="100%" MaxLength="150" Rows="5" TextMode="MultiLine"></asp:TextBox>
                                </td>
                            </tr>
                        </table>
                        <table width="100%">
                            <tr>
                                <td align="center" class="whitecol">
                                    <asp:Button ID="btn_Save" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                                    &nbsp;<asp:Button ID="btn_lev" runat="server" Text="離開" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                    <asp:Panel ID="Panel3" runat="server" Visible="False">
                        <table class="table_sch" cellpadding="1" cellspacing="1">
                            <tr>
                                <td class="bluecol" style="width: 20%">計畫別 </td>
                                <td class="whitecol">
                                    <asp:Label ID="lbl_PlanName" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">原處分分署 </td>
                                <td class="whitecol">
                                    <asp:Label ID="lbl_DistID" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">訓練機構 </td>
                                <td class="whitecol">
                                    <asp:Label ID="lbl_RID" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">職類/班級 </td>
                                <td class="whitecol">
                                    <asp:Label ID="lbl_ClassName" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">身分證號碼 </td>
                                <td class="whitecol">
                                    <asp:Label ID="lbl_idno" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">處分文號 </td>
                                <td class="whitecol">
                                    <asp:Label ID="lbl_No" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">處分緣由 </td>
                                <td class="whitecol">
                                    <asp:Label ID="lbl_SBTERMS" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">處分日期 </td>
                                <td class="whitecol">
                                    <asp:Label ID="lbl_year" runat="server"></asp:Label>
                                    年
								<asp:Label ID="lbl_month" runat="server"></asp:Label>
                                    月
								<asp:Label ID="lbl_day" runat="server"></asp:Label>
                                    日 </td>
                            </tr>
                            <tr>
                                <td class="bluecol">處分年限 </td>
                                <td class="whitecol">
                                    <asp:Label ID="lbl_SBYears" runat="server"></asp:Label>
                                    年 </td>
                            </tr>
                            <tr>
                                <td class="bluecol">處分事由 </td>
                                <td class="whitecol">
                                    <asp:Label ID="lbl_SBCommect" runat="server"></asp:Label>
                                </td>
                            </tr>
                        </table>
                        <table width="100%">
                            <tr>
                                <td align="center" class="whitecol">
                                    <asp:Button ID="btn_lev2" runat="server" Text="離開" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="HidSBSN" runat="server" />
        <asp:HiddenField ID="HidvsType" runat="server" />
        <asp:HiddenField ID="Hid_LID" runat="server" />
    </form>
</body>
</html>
