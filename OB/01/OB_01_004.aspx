<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="OB_01_004.aspx.vb" Inherits="WDAIIP.OB_01_004" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>投標廠商資料查詢</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/common.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script language="JavaScript">
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label><asp:Label ID="TitleLab2" runat="server">
							<FONT face="新細明體">首頁&gt;&gt;委外訓練管理&gt;&gt;<font color="#990000">投標廠商資料查詢</font></FONT>
                    </asp:Label><font color="#000000">(<font face="新細明體"><font color="#ff0000">*</font>為必填欄位</font>)
                    <asp:Label ID="LabActionType" runat="server" ForeColor="Maroon"></asp:Label></font>
                </td>
            </tr>
        </table>
        <asp:Panel ID="panelSch" runat="server">
            <table class="font" border="0" cellspacing="1" cellpadding="1" width="740">
                <tr>
                    <td>
                        <table class="table_sch" cellspacing="1" cellpadding="1">
                            <tr>
                                <%--<td width="100" class="bluecol">轄區中心</td>--%>
                                <td width="100" class="bluecol">轄區分署</td>
                                <td colspan="3" class="whitecol">
                                    <asp:DropDownList ID="ddlDistID" runat="server">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">機構名稱全銜
                                </td>
                                <td colspan="3" class="whitecol">
                                    <asp:TextBox ID="txtOrgName" runat="server" MaxLength="50"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">統一編號
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txtComIDNO" runat="server" MaxLength="15"></asp:TextBox>
                                </td>
                                <td width="100" class="bluecol">立案證號
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txtComCIDNO" runat="server" MaxLength="50"></asp:TextBox>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td>
                        <p align="center">
                            <asp:Button ID="btnQuery" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>&nbsp;
                        <asp:Button ID="btnAdd" runat="server" Text="新增" CssClass="asp_button_S"></asp:Button>&nbsp;
                        <asp:Button ID="btnBack" runat="server" Text="關閉" CssClass="asp_button_S"></asp:Button>
                        </p>
                    </td>
                </tr>
                <tr>
                    <td>
                        <table border="0" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td align="center">
                                    <p>
                                        <table id="DataGridTable" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                            <tr>
                                                <td>
                                                    <asp:DataGrid ID="DataGrid1" runat="server" AllowSorting="True" PagerStyle-HorizontalAlign="Left" PagerStyle-Mode="NumericPages" AllowPaging="True" AutoGenerateColumns="False" CssClass="font" Width="100%">
                                                        <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                        <Columns>
                                                            <asp:BoundColumn HeaderText="序號">
                                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            </asp:BoundColumn>
                                                            <asp:BoundColumn DataField="OrgName" HeaderText="機構名稱全銜">
                                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            </asp:BoundColumn>
                                                            <asp:BoundColumn DataField="ComIDNO" HeaderText="統一編號">
                                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            </asp:BoundColumn>
                                                            <asp:BoundColumn DataField="ComCIDNO" HeaderText="立案證號">
                                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            </asp:BoundColumn>
                                                            <asp:TemplateColumn HeaderText="功能">
                                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                                <ItemStyle HorizontalAlign="Center" Width="150"></ItemStyle>
                                                                <ItemTemplate>
                                                                    <asp:Button ID="btnView" runat="server" Text="檢視" CommandName="view"></asp:Button>
                                                                    <asp:Button ID="btnEdit" runat="server" Text="修改" CommandName="edit"></asp:Button>
                                                                    <asp:Button ID="btnDel" runat="server" Text="刪除" CommandName="del"></asp:Button>
                                                                </ItemTemplate>
                                                            </asp:TemplateColumn>
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
                                        </table>
                                    </p>
                                    <p>
                                        <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
                                    </p>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="panelEdit" runat="server">
            <table class="table_nw" width="740" cellspacing="1" cellpadding="1">
                <tr id="trPlanName" runat="server">
                    <td colspan="4" class="td_light">
                        <asp:Label ID="Label1" runat="server" ForeColor="Maroon">訓練計畫名稱：</asp:Label>
                        <asp:Label ID="LabPlanName" runat="server"></asp:Label>
                        <asp:Label ID="Label2" runat="server" ForeColor="Maroon">標案名稱：</asp:Label>
                        <asp:Label ID="LabTenderCName" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <%--<td width="100" class="bluecol_need">轄區中心</td>--%>
                    <td width="100" class="bluecol_need">轄區分署</td>
                    <td width="85%" colspan="3" class="whitecol">
                        <asp:DropDownList ID="ddlEDistID" runat="server">
                        </asp:DropDownList>
                        <%--<asp:RequiredFieldValidator ID="rfvDistID" runat="server" ControlToValidate="ddlEDistID" Display="None" ErrorMessage="請選擇轄區中心"></asp:RequiredFieldValidator>--%>
                        <asp:RequiredFieldValidator ID="rfvDistID" runat="server" ControlToValidate="ddlEDistID" Display="None" ErrorMessage="請選擇轄區分署"></asp:RequiredFieldValidator>
                        <asp:Label ID="lblDist" runat="server" CssClass="font"></asp:Label><input id="hidTCsn" type="hidden" name="hidTCsn" runat="server">
                    </td>
                </tr>
                <tr>
                    <td width="100" class="bluecol_need">機構名稱全銜
                    </td>
                    <td width="85%" colspan="3" class="whitecol">
                        <asp:TextBox ID="txtTitle" runat="server" MaxLength="50" Width="392px"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvTitle" runat="server" ControlToValidate="txtTitle" Display="None" ErrorMessage="請輸入機構名稱全銜"></asp:RequiredFieldValidator>
                    </td>
                </tr>
                <tr>
                    <td width="100" class="bluecol_need">機構別
                    </td>
                    <td width="85%" colspan="3" class="whitecol">
                        <asp:DropDownList ID="ddlOrg" runat="server" CssClass="font" Width="336px">
                        </asp:DropDownList>
                        <asp:RequiredFieldValidator ID="rfvOrg" runat="server" ControlToValidate="ddlOrg" Display="None" ErrorMessage="請選擇機構別"></asp:RequiredFieldValidator>
                        <asp:Label ID="lblOrg" runat="server" CssClass="font"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td width="100" class="bluecol_need">統一編號
                    </td>
                    <td width="35%" class="whitecol">
                        <asp:TextBox ID="txtEComIDNO" runat="server" Width="80px"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvEComIDNO" runat="server" ControlToValidate="txtEComIDNO" Display="None" ErrorMessage="請輸入統一編號"></asp:RequiredFieldValidator>
                        <asp:RegularExpressionValidator ID="revEComIDNO" runat="server" ControlToValidate="txtEComIDNO" Display="None" ErrorMessage="統一編號請填寫八位數字" ValidationExpression="[0-9]{8}"></asp:RegularExpressionValidator>
                    </td>
                    <td width="100" class="bluecol_need">立案證號
                    </td>
                    <td width="35%" class="whitecol">
                        <asp:TextBox ID="txtEComCIDNO" runat="server" MaxLength="50" Width="80px"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvEComCIDNO" runat="server" ControlToValidate="txtEComCIDNO" Display="None" ErrorMessage="請輸入立案證號"></asp:RequiredFieldValidator>
                    </td>
                </tr>
                <tr>
                    <td width="100" class="bluecol_need">機構電話
                    </td>
                    <td width="35%" class="whitecol">
                        <asp:TextBox ID="txtTel" runat="server" MaxLength="25"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvTel" runat="server" ControlToValidate="txtTel" Display="None" ErrorMessage="請輸入機構電話"></asp:RequiredFieldValidator>
                    </td>
                    <td width="100" class="bluecol_need">機構傳真
                    </td>
                    <td width="35%" class="whitecol">
                        <asp:TextBox ID="txtFax" runat="server" MaxLength="25"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvFax" runat="server" ControlToValidate="txtFax" Display="None" ErrorMessage="請輸入機構傳真"></asp:RequiredFieldValidator>
                    </td>
                </tr>
                <tr>
                    <td width="100" class="bluecol_need">地址
                    </td>
                    <td width="25%" colspan="3" class="whitecol">
                        <input id="txtZip" maxlength="3" runat="server" />－
                        <input id="txtZIPB3" maxlength="3" runat="server" />
                        <%--<asp:TextBox ID="txtZip" runat="server" onfocus="this.blur()"></asp:TextBox>－
                        <asp:TextBox ID="txtZIPB3" runat="server" MaxLength="3"></asp:TextBox>--%>
                        <input id="hidZIP6W" type="hidden" runat="server" />
                        <asp:Literal ID="LitZip" runat="server"></asp:Literal><%--郵遞區號查詢--%>
                        <br />
                        <asp:TextBox ID="txtCity" runat="server" Width="220px"></asp:TextBox>
                        <input id="btnCityZip" value="..." type="button" name="btnCityZip" runat="server" class="button_b_Mini" />
                        <asp:TextBox ID="txtAddr" runat="server" MaxLength="200" Width="350px"></asp:TextBox>

                        <asp:RequiredFieldValidator ID="rfvZip" runat="server" ControlToValidate="txtZip" Display="None" ErrorMessage="請選擇郵遞區號前3碼"></asp:RequiredFieldValidator>
                        <asp:RequiredFieldValidator ID="rfvZIPB3" runat="server" ControlToValidate="txtZIPB3" Display="None" ErrorMessage="請輸入郵遞區號後2碼"></asp:RequiredFieldValidator>
                        <asp:RegularExpressionValidator ID="revZIPB3" runat="server" ControlToValidate="txtZIPB3" Display="None" ErrorMessage="郵遞區號後2碼格式不正確" ValidationExpression="[0-9]{2,3}"></asp:RegularExpressionValidator>
                        <asp:RequiredFieldValidator ID="rfvCity" runat="server" ControlToValidate="txtCity" Display="None" ErrorMessage="請選擇縣市"></asp:RequiredFieldValidator>
                        <asp:RequiredFieldValidator ID="rfvAddr" runat="server" ControlToValidate="txtAddr" Display="None" ErrorMessage="請輸入地址"></asp:RequiredFieldValidator>
                    </td>
                </tr>
                <tr>
                    <td width="100" class="bluecol_need">負責人姓名
                    </td>
                    <td width="35%" class="whitecol">
                        <asp:TextBox ID="txtMName" runat="server" MaxLength="30"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvMName" runat="server" ControlToValidate="txtMName" Display="None" ErrorMessage="請輸入負責人姓名"></asp:RequiredFieldValidator>
                    </td>
                    <td width="100" class="bluecol_need">負責人身分證
                    </td>
                    <td width="35%" class="whitecol">
                        <asp:TextBox ID="txtMIDNO" runat="server"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvMIDNO" runat="server" ControlToValidate="txtMIDNO" Display="None" ErrorMessage="請輸入負責人身分證"></asp:RequiredFieldValidator>
                        <asp:RegularExpressionValidator ID="revMIDNO" runat="server" ControlToValidate="txtMIDNO" Display="None" ErrorMessage="負責人身分證格式不正確" ValidationExpression="^[A-Z][1-2]{1}\d{8}$"></asp:RegularExpressionValidator>
                    </td>
                </tr>
                <tr>
                    <td width="100" class="bluecol">計畫主持人
                    </td>
                    <td width="35%" class="whitecol">
                        <asp:TextBox ID="txtPlanMaster" runat="server" MaxLength="30"></asp:TextBox>
                    </td>
                    <td class="bluecol">主持人電話
                    </td>
                    <td width="35%" class="whitecol">
                        <asp:TextBox ID="txtPMPhone" runat="server" MaxLength="25"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td width="100" class="bluecol">主持人傳真
                    </td>
                    <td width="35%" colspan="3" class="whitecol">
                        <asp:TextBox ID="txtPMFax" runat="server" MaxLength="25"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td width="100" class="bluecol">聯絡人姓名
                    </td>
                    <td width="35%" class="whitecol">
                        <asp:TextBox ID="txtCName" runat="server" MaxLength="30"></asp:TextBox>
                    </td>
                    <td class="bluecol">聯絡人性別
                    </td>
                    <td width="35%" class="whitecol">
                        <asp:RadioButtonList ID="rblCSex" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                            <asp:ListItem Value="m" Selected="True">男</asp:ListItem>
                            <asp:ListItem Value="f">女</asp:ListItem>
                        </asp:RadioButtonList>
                        <asp:Label ID="lblCSex" runat="server" CssClass="font"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td width="100" class="bluecol">聯絡人電話
                    </td>
                    <td width="35%" class="whitecol">
                        <asp:TextBox ID="txtCPhone" runat="server" MaxLength="25"></asp:TextBox>
                    </td>
                    <td class="bluecol">聯絡人手機
                    </td>
                    <td width="35%" class="whitecol">
                        <asp:TextBox ID="txtCCell" runat="server" MaxLength="25"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td width="100" class="bluecol">聯絡人e-mail
                    </td>
                    <td width="35%" class="whitecol">
                        <asp:TextBox ID="txtCEMail" runat="server" MaxLength="64"></asp:TextBox>
                        <asp:RegularExpressionValidator ID="revMail" runat="server" ControlToValidate="txtCEMail" Display="None" ErrorMessage="請重新輸入e-mail" ValidationExpression="\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*"></asp:RegularExpressionValidator>
                    </td>
                    <td class="bluecol">聯絡人傳真
                    </td>
                    <td width="35%" class="whitecol">
                        <asp:TextBox ID="txtCFax" runat="server" MaxLength="25"></asp:TextBox>
                    </td>
                </tr>
            </table>
            <table width="740">
                <tr>
                    <td class="whitecol">
                        <div align="center">
                            <asp:Button ID="btnSave" runat="server" Text="儲存" CssClass="asp_button_S"></asp:Button>&nbsp;
                        <asp:Button ID="btnExit" runat="server" Text="離開" CausesValidation="False" CssClass="asp_button_S"></asp:Button>
                        </div>
                    </td>
                </tr>
            </table>
            <asp:ValidationSummary ID="Summary" runat="server" Width="104px" Height="28px" ShowMessageBox="True" ShowSummary="False" DisplayMode="List"></asp:ValidationSummary>
        </asp:Panel>
    </form>
</body>
</html>
