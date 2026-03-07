<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_01_006_add9t.aspx.vb" Inherits="WDAIIP.CP_01_006_add9t" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>不預告實地訪查紀錄表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <%--<script language="javascript" src="../../js/date-picker.js"></script>--%>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        //決定date-picker元件使用的是西元年or民國年，by:20181018
        var bl_rocYear = "<%=ConfigurationManager.AppSettings("REPLACE2ROC_YEARS") %>";
        var scriptTag = document.createElement('script');
        var jsPath = (bl_rocYear == "Y" ? "../../js/date-picker2.js" : "../../js/date-picker.js");
        scriptTag.src = jsPath;
        document.head.appendChild(scriptTag);

        function save_chkdata() {
            var msg = '';
            //if (document.form1.RIDValue.value == '') msg += '請選擇機構\n';
            //if (document.form1.OCIDValue1.value == '') msg += '請選擇職類\n';
            //if (document.form1.ApplyDate.value == '') msg += '請輸入訪查日期\n';
            //if (bl_rocYear == "Y") {
            //    if (document.form1.ApplyDate.value != '' && !checkRocDate(document.form1.ApplyDate.value)) msg += '訪查日期的時間格式不正確\n';
            //}
            //else {
            //    if (document.form1.ApplyDate.value != '' && !checkDate(document.form1.ApplyDate.value)) msg += '訪查日期的時間格式不正確\n';
            //}
            //if (document.form1.Item10_1.checked && document.form1.Item10_Note.value == '') {
            //    msg += '請輸入附加說明!!\n';
            //}
            //if (!isChecked(document.form1.Item10)) msg += '請選擇結論1的選項\n';
            var $VisitorName = $("#VisitorName");
            if ($VisitorName.val() == '') msg += '請輸入訪視人員姓名?\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }
        function CHANGE_CB1(C1, B1, B2, B3, B4, B5) {
            //alert(C1); debugger;
            if ($("#" + C1).length == 0) { return; }
            var isChecked = $("#" + C1).prop("checked");
            if (isChecked) {
                //var BV1 = $("input[type='radio'][name='" + B1 + "']:checked").val();
                //var BV2 = $("input[type='radio'][name='" + B2 + "']:checked").val();
                //var BV3 = $("input[type='radio'][name='" + B3 + "']:checked").val();
                //var BV4 = $("input[type='radio'][name='" + B4 + "']:checked").val();
                //var BV5 = $("input[type='radio'][name='" + B5 + "']:checked").val();
                //var s_VV1 = BV1 + "," + BV2 + "," + BV3 + "," + BV4 + "," + BV5;
                //$("#" + HN1).val(s_VV1);
                $("input[type='radio'][name='" + B1 + "']").prop("checked", false);
                $("input[type='radio'][name='" + B2 + "']").prop("checked", false);
                $("input[type='radio'][name='" + B3 + "']").prop("checked", false);
                $("input[type='radio'][name='" + B4 + "']").prop("checked", false);
                $("input[type='radio'][name='" + B5 + "']").prop("checked", false);
            }
            //else {
            //    $("input[type='radio'][name='" + B1 + "'][value='yourValue']").prop("checked", true);
            //    $("input[type='radio'][name='" + B2 + "'][value='yourValue']").prop("checked", true);
            //    $("input[type='radio'][name='" + B3 + "'][value='yourValue']").prop("checked", true);
            //    $("input[type='radio'][name='" + B4 + "'][value='yourValue']").prop("checked", true);
            //    $("input[type='radio'][name='" + B5 + "'][value='yourValue']").prop("checked", true);
            //}
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;不預告實地抽訪紀錄表</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="bluecol" width="20%">機構</td>
                            <td class="whitecol" width="80%" colspan="3">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="70%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">職類/班別</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="35%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="35%"></asp:TextBox>
                                <input id="TMIDValue1" type="hidden" runat="server" />
                                &nbsp;<input id="OCIDValue1" type="hidden" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練期間</td>
                            <td class="whitecol">
                                <asp:Label ID="labSFDATE_TW" runat="server"></asp:Label>
                            </td>
                            <td class="bluecol_need">訪查方式</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="rblVISITWAY" runat="server" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="1" Selected="True">實地訪查</asp:ListItem>
                                    <asp:ListItem Value="2">視訊訪查</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="20%">抽訪日期</td>
                            <td class="whitecol" width="30%">
                                <asp:TextBox ID="ApplyDate" runat="server" Width="64%" MaxLength="11"></asp:TextBox>
                            </td>
                            <td class="bluecol_need" width="20%">抽訪原因</td>
                            <td class="whitecol">
                                <asp:RadioButton ID="RBL_VISITREASON" runat="server" Checked="True" Text="實地訪視" />
                            </td>
                        </tr>
                        <tr>
                            <td width="100%" colspan="4">
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
                                                            <asp:HiddenField ID="Hid_rdoVal" runat="server" />
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
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">結論 </td>
                            <td class="whitecol" width="80%" colspan="3">
                                <asp:TextBox ID="Item10_Note" runat="server" Width="70%" TextMode="MultiLine" CssClass="whitecol" Rows="6"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td colspan="4">
                                <table class="font" id="table7" cellspacing="1" width="100%" border="0">
                                    <tr>
                                        <td class="bluecol" width="20%">抽訪人員單位 </td>
                                        <td class="whitecol" width="80%">
                                            <asp:TextBox ID="VisitorOrgNAME" runat="server" onfocus="this.blur()" MaxLength="50" Width="50%"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">抽訪人員姓名 </td>
                                        <td class="whitecol" width="80%">
                                            <asp:TextBox ID="VisitorName" runat="server" MaxLength="50" Width="40%"></asp:TextBox></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                    <p align="center">
                        <asp:Button ID="BtnSave1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>&nbsp;
                        <asp:Button ID="BtnBack1" runat="server" Text="回查詢頁面" CssClass="asp_button_M"></asp:Button>&nbsp;
                        <asp:Button ID="BtnPrint1" runat="server" Text="列印" CssClass="asp_button_M"></asp:Button>
                    </p>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
