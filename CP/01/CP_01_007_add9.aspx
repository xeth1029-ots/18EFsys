<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_01_007_add9.aspx.vb" Inherits="WDAIIP.CP_01_007_add9" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>不預告(電話)抽訪學員紀錄表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-confirm.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery.blockUI.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/global.js"></script>
    <%--<script language="javascript" src="../../js/date-picker.js"></script>--%>
    <script type="text/javascript" language="javascript" src="../../js/date-picker2.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        <%--
        //決定date-picker元件使用的是西元年or民國年，by:20181018
        var bl_rocYear = "<%=ConfigurationManager.AppSettings("REPLACE2ROC_YEARS") %>";
        var scriptTag = document.createElement('script');
        var jsPath = (bl_rocYear == "Y" ? "../../js/date-picker2.js" : "../../js/date-picker.js");
        scriptTag.src = jsPath;
        document.head.appendChild(scriptTag);--%>

        function chkdata() {
            var msg = '';
            if (document.form1.RIDValue.value == '') msg += '請選擇機構\n';
            if (document.form1.OCIDValue1.value == '') msg += '請選擇職類\n';
            if (document.form1.ApplyDate.value == '') msg += '請輸入訪查日期\n';
            if (bl_rocYear == "Y") {
                if (document.form1.ApplyDate.value != '' && !checkRocDate(document.form1.ApplyDate.value)) msg += '訪查日期的時間格式不正確\n';
            }
            else {
                if (document.form1.ApplyDate.value != '' && !checkDate(document.form1.ApplyDate.value)) msg += '訪查日期的時間格式不正確\n';
            }
            if (document.form1.Item10_1.checked && document.form1.Item10_Note.value == '') {
                msg += '請輸入附加說明!!\n';
            }
            if (!isChecked(document.form1.Item10)) msg += '請選擇結論1的選項\n';
            //if(document.form1.CurseName.value=='') msg+='請輸入培訓單位人員姓名?\n';
            if (document.form1.VisitorName.value == '') msg += '請輸入訪視人員姓名?\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function clear_rdoAnswer(j_cbl1, j_rdo1) {
            if ($(j_cbl1).length == 0 || $(j_rdo1).length == 0) return false;
            var isChecked = $(j_cbl1).is(':checked');
            if (!isChecked) return false;
            $(j_rdo1 + " input[type='radio']").prop('checked', false);
            //var selectedValue = $('#DataGrid1_ctl10_rdoAnswer1 input[type="radio"]:checked').val();
            //var isChecked = $('#DataGrid1_ctl10_cb1_show').is(':checked');
            //if($("#DataGrid1_ctl10_rdoAnswer1").length==0) return false;
            //$("#DataGrid1_ctl10_rdoAnswer1 input[type='radio']").prop('checked', false);
        }
        function clear_cb1_show(j_cbl1, j_rdo1) {
            if ($(j_cbl1).length == 0 || $(j_rdo1).length == 0) return false;
            var selectedValue = $(j_rdo1 + ' input[type="radio"]:checked').val();
            if (selectedValue == '') { return false; }
            $(j_cbl1).prop("checked", false);
            //$("#DataGrid1_ctl10_cb1_show").prop("checked", false);
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;不預告(電話)抽訪學員紀錄表</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="bluecol" width="20%">機構</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="Button2" type="button" value="..." name="Button2" runat="server">
                                <input id="RIDValue" type="hidden" name="Hidden1" runat="server">
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="20%">職類/班別</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input id="Button3" onclick="javascript: window.open('../CP_01_ch.aspx?RID=' + document.form1.RIDValue.value, '', 'width=860,height=760,location=0,status=0,menubar=0,scrollbars=1,resizable=0');" type="button" value="..." name="Button3" runat="server"><input id="TMIDValue1" type="hidden" name="Hidden1" runat="server">
                                <input id="OCIDValue1" type="hidden" name="Hidden2" runat="server">
                                <span id="HistoryList" style="position: absolute; display: none; left: 30%">
                                    <asp:Table ID="HistoryTable" runat="server"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="20%">抽訪日期</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="ApplyDate" runat="server" Width="14%" MaxLength="11"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('ApplyDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="20%">電話抽訪原因</td>
                            <td class="whitecol" width="80%">
                                <asp:RadioButtonList ID="rbl_TELVISITREASON" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="1" Selected="True">辦訓異常或陳情案件</asp:ListItem>
                                    <asp:ListItem Value="2">抽訪時未到</asp:ListItem>
                                    <asp:ListItem Value="3">例行訪視</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">實地訪視日期</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="REALVISITDATE" runat="server" Width="14%" MaxLength="11"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('REALVISITDATE','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <%--<tr>
                            <td class="bluecol" width="20%">實地／遠距</td>
                            <td class="whitecol" width="80%">&nbsp;(<asp:CheckBox ID="cb1_NotYetENTITY" runat="server" />尚未上實體課程) &nbsp;
                                 &nbsp;(<asp:CheckBox ID="cb2_NotYetDISTANCE" runat="server" />尚未上遠距課程) &nbsp;
                            </td>
                        </tr>--%>
                        <tr>
                            <td width="100%" colspan="2">
                                <table class="font" id="DataGridTable1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                    <tr>
                                        <td>
                                            <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False">
                                                <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                <Columns>
                                                    <asp:TemplateColumn HeaderText="項次">
                                                        <HeaderStyle HorizontalAlign="Center" Width="6%"></HeaderStyle>
                                                        <ItemStyle Wrap="False" HorizontalAlign="Center" VerticalAlign="Middle"></ItemStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="labShowItem" runat="server" Text=""></asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="內容">
                                                        <HeaderStyle HorizontalAlign="Center" Width="24%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="labquestion" runat="server" Text=""></asp:Label>
                                                            <asp:CheckBox ID="cb1_show" runat="server" Visible="false" Width="100%" />
                                                            <asp:HiddenField ID="hid_ckcolumn" runat="server" Value="" />
                                                            <asp:HiddenField ID="hid_dataitem" runat="server" Value="" />
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="訪問一編號：">
                                                        <ItemStyle Wrap="False" Width="12%" CssClass="whitecol"></ItemStyle>
                                                        <ItemTemplate>
                                                            <asp:RadioButtonList ID="rdoAnswer1" runat="server" CssClass="font"></asp:RadioButtonList>
                                                            <asp:TextBox ID="txtAnswer1" runat="server" Width="100%"></asp:TextBox>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="訪問二編號：">
                                                        <ItemStyle Wrap="False" Width="12%" CssClass="whitecol"></ItemStyle>
                                                        <ItemTemplate>
                                                            <asp:RadioButtonList ID="rdoAnswer2" runat="server" CssClass="font"></asp:RadioButtonList>
                                                            <asp:TextBox ID="txtAnswer2" runat="server" Width="100%"></asp:TextBox>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="訪問三編號：">
                                                        <ItemStyle Wrap="False" Width="12%" CssClass="whitecol"></ItemStyle>
                                                        <ItemTemplate>
                                                            <asp:RadioButtonList ID="rdoAnswer3" runat="server" CssClass="font"></asp:RadioButtonList>
                                                            <asp:TextBox ID="txtAnswer3" runat="server" Width="100%"></asp:TextBox>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="訪問四編號：">
                                                        <ItemStyle Wrap="False" Width="12%" CssClass="whitecol"></ItemStyle>
                                                        <ItemTemplate>
                                                            <asp:RadioButtonList ID="rdoAnswer4" runat="server" CssClass="font"></asp:RadioButtonList>
                                                            <asp:TextBox ID="txtAnswer4" runat="server" Width="100%"></asp:TextBox>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="訪問五編號：">
                                                        <ItemStyle Wrap="False" Width="12%" CssClass="whitecol"></ItemStyle>
                                                        <ItemTemplate>
                                                            <asp:RadioButtonList ID="rdoAnswer5" runat="server" CssClass="font"></asp:RadioButtonList>
                                                            <asp:TextBox ID="txtAnswer5" runat="server" Width="100%"></asp:TextBox>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="備註／說明事項">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                        <ItemTemplate>
                                                            <asp:TextBox ID="txtNote" runat="server" Width="100%" CssClass="font" TextMode="MultiLine"></asp:TextBox>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                </Columns>
                                                <PagerStyle Visible="False"></PagerStyle>
                                            </asp:DataGrid>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="center">
                                            <uc1:PageControler ID="PageControler1" runat="server" Visible="False"></uc1:PageControler>
                                            <br>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">結論</td>
                            <td class="whitecol" width="80%">
                                <asp:RadioButtonList ID="Item10" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="1">正常</asp:ListItem>
                                    <asp:ListItem Value="2">不正常，須加以查核</asp:ListItem>
                                </asp:RadioButtonList>
                                &nbsp;<asp:CheckBox ID="Item10_1" runat="server" Text="其他附加說明"></asp:CheckBox>
                                <br />
                                <asp:TextBox ID="Item10_Note" runat="server" Width="50%" TextMode="MultiLine" CssClass="whitecol"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">其他處理方式 </td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="Item10_Other" runat="server" Width="50%" TextMode="MultiLine" CssClass="whitecol"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <table class="font" id="table7" cellspacing="1" width="100%" border="0">
                                    <tr>
                                        <td class="bluecol" width="20%">抽訪人員單位 </td>
                                        <td class="whitecol" width="80%">
                                            <asp:TextBox ID="VisitorOrgNAME" runat="server" onfocus="this.blur()" MaxLength="50" Width="40%"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">抽訪人員姓名 </td>
                                        <td class="whitecol" width="80%">
                                            <asp:TextBox ID="VisitorName" runat="server" MaxLength="50" Width="20%"></asp:TextBox></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                    <%--<p align="center" class="whitecol"></p>--%>
                </td>
            </tr>
            <tr>
                <td align="center" class="whitecol">
                    <asp:Button ID="Button1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>&nbsp;
                        <input id="Button4" type="button" value="回查詢頁面" name="Button4" runat="server" class="button_b_M" />
                    <input id="Button5" type="button" value="回查詢頁面" name="Button5" runat="server" class="button_b_M" />
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
