<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_02_011.aspx.vb" Inherits="WDAIIP.SD_02_011" %>

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd HTML 4.0 Transitional//EN">
<html>
<head>
    <title>甄試比例設定</title>
    <meta content="microsoft visual studio .net 7.1" name="generator" />
    <meta content="visual basic .net 7.1" name="code_language" />
    <meta content="javascript" name="vs_defaultclientscript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetschema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        //判斷儲存
        function chkSave() {
            var msg = '';
            var txtWrite = document.getElementById('txtWrite');
            var txtOral = document.getElementById('txtOral');
            if (!isBlank(txtWrite) && !isBlank(txtOral)) {
                if (isBlank(txtWrite)) msg += '請輸入筆試比例!\n';
                else if (!isPositiveFloat(txtWrite.value) && !isPositiveInt(txtWrite.value)) {
                    msg += '筆試比例格式錯誤!\n';
                }
                if (isBlank(txtOral)) msg += '請輸入口試比例!\n';
                else if (!isPositiveFloat(txtOral.value) && !isPositiveInt(txtOral.value)) {
                    msg += '口試比例格式錯誤!\n';
                }
                if (msg == '') {
                    if (parseFloat(txtWrite.value) + parseFloat(txtOral.value) != 100) {
                        msg += '筆試加口式比例不等於100!\n';
                    }
                }
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
    <style type="text/css">
        .style1 { height: 22px; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="titlelab1" runat="server"></asp:Label>
                    <asp:Label ID="titlelab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;甄試比例設定</asp:Label>
                </td>
            </tr>
        </table>
        <table cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Panel ID="tbsch" runat="server">
                        <table class="table_nw" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol" style="width: 20%">年度</td>
                                <td class="whitecol" colspan="3"><asp:DropDownList ID="ddlSchYear" runat="server" Enabled="false"></asp:DropDownList></td>
                            </tr>
                            <tr>
                                <td class="bluecol">訓練職類</td>
                                <td class="whitecol" colspan="3">
                                    <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Columns="30" Width="40%"></asp:TextBox>
                                    <input id="btu_sel" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." name="btu_sel" runat="server" class="asp_button_Mini" />&nbsp;
                                    <input id="trainValue" type="hidden" name="trainvalue" runat="server" />
                                    <input id="jobvalue" type="hidden" name="jobvalue" runat="server" />
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">通俗職</td>
                                <td class="whitecol" colspan="3">
                                    <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="40%"></asp:TextBox>
                                    <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="asp_button_Mini" />
                                    <input id="cjobValue" type="hidden" name="cjobValue" runat="server" />
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">訓練機構</td>
                                <td class="whitecol" colspan="3">
                                    <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                    <input id="org" type="button" value="..." name="org" runat="server" class="asp_button_Mini" />
                                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server" style="width: 32px; height: 22px" />
                                    <asp:Button ID="btnOrgSet" runat="server" Text="單位比例設定" CssClass="asp_button_M"></asp:Button>
                                    <span id="HistoryList2" style="display: none; position: absolute"><asp:Table ID="historyrid" runat="server" Width="100%"></asp:Table></span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" style="width: 20%">班級名稱</td>
                                <td class="whitecol" style="width: 30%"><asp:TextBox ID="txtschclass" runat="server" Columns="40"></asp:TextBox></td>
                                <td class="bluecol" style="width: 20%">期別</td>
                                <td class="whitecol" style="width: 30%"><asp:TextBox ID="txtcycltype" runat="server" Columns="5" Width="40%"></asp:TextBox></td>
                            </tr>
                        </table>
                        <table width="100%">
                            <tr>
                                <td align="center" class="whitecol"><asp:Button ID="btnsch" Text="查詢" runat="server" CssClass="asp_button_M"></asp:Button></td>
                            </tr>
                        </table>
                        <table width="100%">
                            <tr>
                                <td align="center" class="whitecol">
                                    <br style="line-height: 5px" />
                                    <asp:Label ID="labMsg" Style="color: red" runat="server" Visible="false">查無資料!!</asp:Label>
                                    <asp:DataGrid ID="DataGrid1" runat="server" PagerStyle-Visible="false" AutoGenerateColumns="false" AllowPaging="true" AllowSorting="true" CssClass="font" Width="100%">
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <HeaderStyle CssClass="head_navy" />
                                        <Columns>
                                            <asp:TemplateColumn HeaderText="序號" ItemStyle-HorizontalAlign="Center">
                                                <HeaderStyle Width="5%" />
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="年度" ItemStyle-HorizontalAlign="Center">
                                                <HeaderStyle Width="10%" />
                                                <ItemTemplate>
                                                    <asp:Label ID="labDYear" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="訓練機構">
                                                <HeaderStyle Width="27%" />
                                                <ItemTemplate>
                                                    <asp:Label ID="labDOrgName" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="班級名稱">
                                                <HeaderStyle Width="27%" />
                                                <ItemTemplate>
                                                    <asp:Label ID="labDClassName" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="筆試比例%" ItemStyle-HorizontalAlign="Center">
                                                <HeaderStyle Width="11%" />
                                                <ItemTemplate>
                                                    <asp:Label ID="labDWrite" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="口試比例%" ItemStyle-HorizontalAlign="Center">
                                                <HeaderStyle Width="11%" />
                                                <ItemTemplate>
                                                    <asp:Label ID="labDOral" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="功能">
                                                <HeaderStyle Width="7%" />
                                                <ItemStyle HorizontalAlign="Center" Font-Size="Small" />
                                                <ItemTemplate>
                                                    <asp:LinkButton ID="btnEdit" Text="設定" CommandName="edt" runat="server" CssClass="linkbutton"></asp:LinkButton>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                    </asp:DataGrid><uc1:PageControler ID="pagecontroler1" runat="server"></uc1:PageControler>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                    <asp:Panel ID="tbEdit" runat="server">
                        <table class="table_sch" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol" style="width: 20%">年度</td>
                                <td class="whitecol"><asp:Label ID="labYear" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">轄區</td>
                                <td class="whitecol"><asp:Label ID="labDist" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">訓練機構</td>
                                <td class="whitecol"><asp:Label ID="labOrg" runat="server"></asp:Label></td>
                            </tr>
                            <tr id="trclass" runat="server">
                                <td class="bluecol">班級名稱</td>
                                <td class="whitecol"><asp:Label ID="labClass" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">成績計算比例</td>
                                <td class="whitecol">
                                    筆試<asp:TextBox ID="txtWrite" MaxLength="5" runat="server" Width="10%"></asp:TextBox>%<br />
                                    口試<asp:TextBox ID="txtOral" MaxLength="5" runat="server" Width="10%"></asp:TextBox>%
                                </td>
                            </tr>
                        </table>
                        <table width="100%">
                            <tr>
                                <td align="center" class="whitecol">
                                    <asp:Button ID="btnSave" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>&nbsp;
                                    <asp:Button ID="btnBack" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                </td>
            </tr>
        </table>
        <input id="hidWOID" type="hidden" runat="server" name="hidWOID">
        <input id="hidYears" type="hidden" runat="server" name="hidYears">
        <input id="hidDistID" type="hidden" runat="server" name="hidDistID">
        <input id="hidPlanID" type="hidden" runat="server" name="hidPlanID">
        <input id="hidOrgID" type="hidden" runat="server" name="hidOrgID">
        <input id="hidOCID" type="hidden" runat="server" name="hidOCID">
    </form>
</body>
</html>