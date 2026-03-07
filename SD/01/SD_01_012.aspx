<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_01_012.aspx.vb" Inherits="WDAIIP.SD_01_012" %>

<!DOCTYPE html PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head id="Head1" runat="server">
    <title>民眾自行取消報名</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function GETvalue() { document.getElementById('Button6').click(); }

        function SetOneOCID() { document.getElementById('Button7').click(); }

        function ClearData() {
            document.getElementById('TMID1').value = '';
            document.getElementById('OCID1').value = '';
            document.getElementById('TMIDValue1').value = '';
            document.getElementById('OCIDValue1').value = '';
        }
        function CheckSearch() {
            if (document.getElementById('OCIDValue1').value == ''
			&& document.getElementById('txtQIDNO').value == ''
			&& document.getElementById('txtQName').value == '') {
                alert('至少要輸入一項條件');
                return false;
            }
        }
        function choose_class() {
            if (document.getElementById('OCID1').value == '') {
                document.getElementById('Button7').click();
            }
            openClass('../02/SD_02_ch.aspx?RID=' + document.getElementById('RIDValue').value);
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;民眾自行取消報名</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Frametable1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table id="tbSch" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <table class="table_nw" id="table2" cellspacing="1" cellpadding="1" width="100%">
                                    <tr id="Orgtr" runat="server">
                                        <td class="bluecol" style="width: 20%">訓練機構
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                            <input id="RIDValue" type="hidden" name="Hidden2" runat="server" />
                                            <input id="Button2" type="button" value="..." name="Button2" runat="server" class="asp_button_Mini" />
                                            <asp:Button ID="Button7" Style="display: none" runat="server"></asp:Button>
                                            <asp:Button ID="Button6" Style="display: none" runat="server"></asp:Button>
                                            <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                                <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                            </span>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">職類/班別
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                            <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                            <input onclick="choose_class()" type="button" value="..." class="asp_button_Mini" />
                                            <input id="Button4" type="button" value="清除" name="Button4" runat="server" class="asp_button_S" />
                                            <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                            <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                            <span id="HistoryList" style="position: absolute; display: none; left: 270px">
                                                <asp:Table ID="Historytable" runat="server" Width="100%">
                                                </asp:Table>
                                            </span>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" style="width: 20%">身分證號碼
                                        </td>
                                        <td class="whitecol" style="width: 30%">
                                            <asp:TextBox ID="txtQIDNO" runat="server" Width="60%"></asp:TextBox>
                                        </td>
                                        <td class="bluecol" style="width: 20%">學員姓名
                                        </td>
                                        <td class="whitecol" style="width: 30%">
                                            <asp:TextBox ID="txtQName" runat="server" Width="60%"></asp:TextBox>
                                        </td>
                                    </tr>
                                </table>
                                <table width="100%">
                                    <tr>
                                        <td class="whitecol" align="center" colspan="4">
                                            <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                            <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                            <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol" align="center" colspan="4">
                                            <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                                        </td>
                                    </tr>
                                </table>
                                <table id="tbList" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                    <tr>
                                        <td>
                                            <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AllowPaging="true" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                <AlternatingItemStyle BackColor="WhiteSmoke" />
                                                <HeaderStyle CssClass="head_navy" HorizontalAlign="Center" />
                                                <Columns>
                                                    <asp:BoundColumn HeaderText="序號">
                                                        <HeaderStyle Width="4%" HorizontalAlign="Center"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="Name" HeaderText="姓名">
                                                        <HeaderStyle Width="12%" HorizontalAlign="Center"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:TemplateColumn HeaderText="身分證號碼">
                                                        <HeaderStyle Width="12%" HorizontalAlign="Center" />
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        <ItemTemplate>
                                                            <asp:Label ID="labIDNO" runat="server"></asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:BoundColumn DataField="ORGNAME" HeaderText="報名機構">
                                                        <HeaderStyle Width="12%" HorizontalAlign="Center"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="CLASSCNAME" HeaderText="報名班級">
                                                        <HeaderStyle Width="12%" HorizontalAlign="Center"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="TRound" HeaderText="訓練期間">
                                                        <HeaderStyle Width="12%" HorizontalAlign="Center"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="ENTERDATE" HeaderText="報名日期">
                                                        <HeaderStyle Width="12%" HorizontalAlign="Center"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="MODIFYDATE" HeaderText="民眾自行取消時間">
                                                        <HeaderStyle Width="12%" HorizontalAlign="Center"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:TemplateColumn HeaderText="功能">
                                                        <HeaderStyle Width="12%" HorizontalAlign="Center"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                                                        <ItemTemplate>
                                                            <asp:LinkButton ID="BtnView" runat="server" Text="檢視" CssClass="linkbutton" CommandName="vie"></asp:LinkButton>
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
                            </td>
                        </tr>
                    </table>
                    <table id="tbView" class="table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td align="center" class="whitecol">
                                <table id="Table1" class="Table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                    <tr>
                                        <td class="bluecol" style="width: 20%">姓名
                                        </td>
                                        <td class="whitecol" style="width: 30%">
                                            <asp:Label ID="LabNAME" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol" style="width: 20%">出生日期
                                        </td>
                                        <td class="whitecol" style="width: 30%">
                                            <asp:Label ID="Birthday" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">身分別
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="PassPortNO" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol">身分證號碼
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="LabIDNO" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">性別
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Sex" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol">最高學歷
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="DegreeID" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">婚姻狀況
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="MaritalStatus" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol">畢業狀況
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="GradID" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">學校名稱
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="School" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol">科系
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Department" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">兵役
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:Label ID="MilitaryID" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">通訊地址
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:Label ID="Address" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">聯絡電話(日)
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Phone1" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol">聯絡電話(夜)
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Phone2" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">電子信箱
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Email" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol">行動電話
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="CellPhone" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr id="TRIdentityID" runat="server">
                                        <td class="bluecol">參訓身分別
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:Label ID="IdentityID" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr id="TRHandTypeID" runat="server">
                                        <td class="bluecol">障礙類別
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="labHandTypeID" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol">障礙等級
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="labHandLevelID" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">報名日期
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:Label ID="RelEnterDate" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">報名志願
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <table id="Table4" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                                                <tr>
                                                    <td style="width: 9%">第一志願：
                                                    </td>
                                                    <td>
                                                        <asp:Label ID="LabOCID1" runat="server"></asp:Label>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td style="width: 9%">第二志願：
                                                    </td>
                                                    <td>
                                                        <asp:Label ID="OCID2" runat="server"></asp:Label>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td style="width: 9%">第三志願：
                                                    </td>
                                                    <td>
                                                        <asp:Label ID="OCID3" runat="server"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                </table>
                                <table id="Table11" class="Table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                    <tr>
                                        <td class="bluecol" style="width: 20%">報名班級
                                        </td>
                                        <td class="whitecol" style="width: 30%">
                                            <asp:Label ID="LabClassCname" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol" style="width: 20%">報名日期
                                        </td>
                                        <td class="whitecol" style="width: 30%">
                                            <asp:Label ID="LabEnterDate" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">姓名
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="LabName2" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol">出生日期
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="LabBirthDay" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">身分別
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Label5" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol">身分證號碼
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Label6" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">性別
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Label7" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol">最高學歷
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Label9" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">通訊地址
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:Label ID="Label14" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">戶籍地址
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:Label ID="Label15" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">聯絡電話(日)
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Label16" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol">聯絡電話(夜)
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Label17" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">電子信箱
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Label18" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol">行動電話
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Label19" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">主要參訓身分別
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:Label ID="Label20" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">受訓前薪資
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:Label ID="Label30" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">郵政/銀行帳號
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:Label ID="Label33" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                </table>
                                <table id="Table22" class="Table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                    <tr>
                                        <td class="bluecol" style="width: 20%">局號
                                        </td>
                                        <td class="whitecol" style="width: 30%">
                                            <asp:Label ID="Label35" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol" style="width: 20%">帳號
                                        </td>
                                        <td class="whitecol" style="width: 30%">
                                            <asp:Label ID="Label36" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                </table>
                                <table id="Table23" class="Table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                    <tr>
                                        <td class="bluecol" style="width: 20%">總行名稱
                                        </td>
                                        <td class="whitecol" style="width: 30%">
                                            <asp:Label ID="Label37" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol" style="width: 20%">總行代號
                                        </td>
                                        <td class="whitecol" style="width: 30%">
                                            <asp:Label ID="Label38" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">分行名稱
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Label61" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol">分行代號
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Label62" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">帳號
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:Label ID="Label34" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                </table>
                                <table id="Table12" class="Table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                    <tr>
                                        <td class="bluecol" style="width: 20%">投保單位名稱
                                        </td>
                                        <td class="whitecol" style="width: 30%">
                                            <asp:Label ID="Label59" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol" style="width: 20%">投保單位<br />
                                            保險證號
                                        </td>
                                        <td class="whitecol" style="width: 30%">
                                            <asp:Label ID="Label60" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">投保單位類別
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Label63" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol">投保單位電話
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Label64" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <%--	
												    <asp:BoundColumn DataField="IDNO" HeaderText="身分證號碼">
													    <HeaderStyle Width="66px" HorizontalAlign="Center"></HeaderStyle>
													    <ItemStyle HorizontalAlign="Center"></ItemStyle>
												    </asp:BoundColumn>
                                    --%>
                                    <tr>
                                        <td class="bluecol">投保單位地址
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:Label ID="Label65" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <%--	
												    <asp:BoundColumn DataField="IDNO" HeaderText="身分證號碼">
													    <HeaderStyle Width="66px" HorizontalAlign="Center"></HeaderStyle>
													    <ItemStyle HorizontalAlign="Center"></ItemStyle>
												    </asp:BoundColumn>
                                    --%>
                                    <tr>
                                        <td class="bluecol">目前公司名稱
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Label40" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol">統一編號
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Label41" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <%--<tr>
								    <td class="bluecol">投保單位<br />統一編號</td>
								    <td class="whitecol" colspan="3"><asp:Label id="actcomIDNO" runat="server"></asp:Label></td>
							    </tr>--%>
                                    <tr>
                                        <td class="bluecol">目前任職部門
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Label45" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol">職稱
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Label46" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <%--<tr>
								    <td class="bluecol">第一次投保日</td>
								    <td class="whitecol" colspan="3"><asp:Label id="Label39" runat="server" width="136px"></asp:Label></td>
							    </tr>--%>
                                    <tr>
                                        <td class="bluecol">是否由公司推薦參訓
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:Label ID="Label50" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">參訓動機
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:Label ID="Label51" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">服務單位行業別
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:Label ID="Label58" runat="server" Width="448px"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">訓後動向
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Label52" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol">服務單位是否屬於中小企業
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Label53" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">個人工作年資
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Label54" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol">在這家公司的年資
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Label55" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">在這職位的年資
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Label56" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol">最近升遷離本職幾年
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="Label57" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">預算別
                                        </td>
                                        <td class="whitecol">
                                            <asp:DropDownList ID="ddlBudID" runat="server">
                                                <asp:ListItem Value="01">公務</asp:ListItem>
                                                <asp:ListItem Value="02">就安</asp:ListItem>
                                                <asp:ListItem Value="03">就保</asp:ListItem>
                                                <asp:ListItem Value="97">協助</asp:ListItem>
                                                <asp:ListItem Value="99">不補助</asp:ListItem>
                                                <asp:ListItem Value="請選擇">請選擇</asp:ListItem>
                                            </asp:DropDownList>
                                        </td>
                                        <td class="bluecol">補助比例
                                        </td>
                                        <td class="whitecol">
                                            <asp:DropDownList ID="ddlSupplyID" runat="server">
                                                <asp:ListItem Value="1">一般80%</asp:ListItem>
                                                <asp:ListItem Value="2">特定100%</asp:ListItem>
                                                <asp:ListItem Value="9">0%</asp:ListItem>
                                                <asp:ListItem>請選擇</asp:ListItem>
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                </table>
                                <table id="TablePWTYPE" class="Table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                    <tr>
                                        <td class="bluecol" style="width: 20%">受訓前任職狀況
                                        </td>
                                        <td class="whitecol" style="width: 30%">
                                            <asp:Label ID="PriorWorkType1" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol" style="width: 20%">最後一次任職<br />
                                            單位名稱
                                        </td>
                                        <td class="whitecol" style="width: 30%">
                                            <asp:Label ID="PriorWorkOrg1" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">最後投保單位<br />
                                            起迄日
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="OfficeDate" runat="server"></asp:Label>
                                        </td>
                                        <td class="bluecol">最後投保單位<br />
                                            保險證號
                                        </td>
                                        <td class="whitecol">
                                            <asp:Label ID="ActNo" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol" align="center">
                                <table width="100%">
                                    <tr>
                                        <td align="center">
                                            <asp:Button ID="BtnBack1" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <input id="HidIDNO" type="hidden" runat="server" />
        <input id="HidSTDate" type="hidden" runat="server" />
        <input id="HidFTDate" type="hidden" runat="server" />
        <input id="HideSerNum" type="hidden" runat="server" />
        <input id="HideSETID" type="hidden" runat="server" />
    </form>
</body>
</html>
