<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_13_003_15.aspx.vb" Inherits="WDAIIP.SD_13_003_15" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_13_003</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script language="javascript" src="../../js/common.js"></script>
    <script language="javascript">
        function HistoryShearch() {
            //查詢重複參訓 HistoryShearch
            window.open('SD_13_History.aspx?OCID=' + document.getElementById('OCIDValue1').value, 'history', 'width=770,height=600,scrollbars=1')
        }

        function CheckDate() {
            var msg = ""
            if (document.form1.Dclass.value == 1) {
                //if(!confirm('身分證號碼錯誤，是否要繼續儲存?')) msg=msg+'身分證號碼錯誤\n';
                if (!confirm('此班級的學員在其他計畫有參訓紀錄,按查詢重複參訓可以查看,您是否確定要經費審核確認?')) { msg = msg + '此班級的學員在其他計畫有參訓紀錄\n'; }
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
            else {
                return chkMoney();
            }
        }

        function SelectAll() {
            var tb = document.getElementById("Datagrid2");
            for (i = 1; i < tb.rows.length; i++) {
                if (tb.rows(i).cells(11).children(0).disabled == false) {
                    tb.rows(i).cells(11).children(0).value = tb.rows(0).cells(11).children(0).value;
                }
            }
        }

        function chkMoney() {
            var cst_name = 1;
            var cst_Sum = 7;    //補助金額(本次補助金額)
            var cst_Remain = 9; //剩餘可用
            var cst_Verify = 11; //審核狀態

            var MyTable = document.getElementById('Datagrid2');
            var msg = '';
            for (i = 1; i < MyTable.rows.length; i++) {
                var SumOfMoney = parseInt(MyTable.rows(i).cells(cst_Sum).children(0).value);
                var RemainSub = parseInt(MyTable.rows(i).cells(cst_Remain).children(0).value);
                //alert(MyTable.rows(i).cells(11).children(1).value);
                //alert(MyTable.rows(i).cells(11).children(0).value);
                if (MyTable.rows(i).cells(cst_Verify).children(1).value == '1') {
                    if (MyTable.rows(i).cells(cst_Verify).children(0).value == 'Y') {
                        if (SumOfMoney > RemainSub) {
                            //debugger;	 
                            msg += '補助金額' + SumOfMoney + '不能超過剩餘可用餘額' + RemainSub + '(學員:' + MyTable.rows(i).cells(cst_name).children(1).innerHTML + ')\n';
                        }
                    }
                }
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function GETvalue() {
            document.getElementById('Button3').click();
        }
        function SetOneOCID() {
            document.getElementById('Button4').click();
        }
        function choose_class() {
            if (document.getElementById('OCID1').values == '') {
                document.getElementById('Button4').click();
            }
            document.getElementById('DataGridTable').style.display = 'none';
            openClass('../02/SD_02_ch.aspx?RID=' + document.getElementById('RIDValue').value);
        }
        function CheckSearch() {
            if (document.getElementById('OCIDValue1').value == '') {
                alert('請選擇職類班別');
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <asp:Button ID="Button4" Style="display: none" runat="server"></asp:Button><asp:Button ID="Button3" Style="display: none" runat="server"></asp:Button>
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label><asp:Label ID="TitleLab2" runat="server">
									首頁&gt;&gt;學員動態管理&gt;&gt;補助金請領&gt;&gt;<FONT color="#990000">補助審核</FONT>
                                </asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="table_nw" id="Table2" cellspacing="1" cellpadding="1" width="600">
                        <tr>
                            <td class="bluecol" width="100">訓練機構
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" Width="410px"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="asp_button_Mini" />
                                <span id="HistoryList2" style="display: none; z-index: 100; position: absolute" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="310px">
                                    </asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">職類/班別
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="210px"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="200px"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="asp_button_Mini" />
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                <span id="HistoryList" style="display: none; z-index: 101; left: 270px; position: absolute">
                                    <asp:Table ID="HistoryTable" runat="server" Width="310">
                                    </asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="2" class="whitecol">
                                <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="2" class="whitecol">
                                <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                    </table>
                    <asp:Panel ID="AuditNumPanel" runat="server">
                        <table class="font" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <tr>
                                <td>
                                    <asp:Label ID="Label1" runat="server">審核成功學員數：</asp:Label><asp:Label ID="SNum" runat="server"></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								<asp:Label ID="Label2" runat="server">審核失敗學員數：</asp:Label><asp:Label ID="FNum" runat="server"></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								<asp:Label ID="Label3" runat="server">未審核學員數：</asp:Label><asp:Label ID="ANum" runat="server"></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								<asp:Label ID="Label4" runat="server">退件修正學員數：</asp:Label><asp:Label ID="RNum" runat="server"></asp:Label>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                    <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>滑鼠移至姓名可以顯示此學員前三年的申請紀錄
							<asp:Label ID="Label6" runat="server" ForeColor="Red">　保險證號如果是"09"開頭者，用紅色顯示提醒不予補助</asp:Label><br>
                                <asp:Label ID="Label5" runat="server" ForeColor="Blue">*表學員參加此班的開訓日，未在勞保投保期間（點選學員姓名，可查看勞保明細）</asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:DataGrid ID="Datagrid2" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" AllowSorting="True">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    <Columns>
                                        <asp:TemplateColumn SortExpression="StudentID" HeaderText="學號">
                                            <HeaderStyle ForeColor="#B0E2FF"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_StudentID" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="姓名">
                                            <ItemTemplate>
                                                <input id="hid_Name" type="hidden" name="hid_Name" runat="server" />
                                                <asp:Label ID="lab_Star" runat="server" ForeColor="Blue">*</asp:Label>
                                                <asp:LinkButton ID="link_Name" runat="server" CssClass="l" CommandName="Link"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn SortExpression="IDNO" HeaderText="身分證號碼">
                                            <HeaderStyle ForeColor="#B0E2FF"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_IDNO" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="是否取得結訓資格">
                                            <ItemTemplate>
                                                <asp:Label ID="lab_EndClass" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="出席達3/4">
                                            <ItemTemplate>
                                                <asp:Label ID="lab_OnClassRate" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="是否補助">
                                            <ItemTemplate>
                                                <asp:Label ID="lab_IsSubSidy" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="總費用">
                                            <ItemTemplate>
                                                <asp:Label ID="lab_Total" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="補助費用">
                                            <ItemTemplate>
                                                <input id="hid_SubSidyCost" type="hidden" runat="server" name="hid_SubSidyCost">
                                                <asp:Label ID="lab_SubSidyCost" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="個人支付">
                                            <ItemTemplate>
                                                <asp:Label ID="lab_PersonalCost" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="剩餘可用餘額">
                                            <ItemTemplate>
                                                <asp:Label ID="lab_Balance" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="其他申請中金額">
                                            <ItemTemplate>
                                                <asp:Label ID="lab_OtherGovApply" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="審核">
                                            <HeaderTemplate>
                                                審核
											<asp:DropDownList ID="list_VerifyAll" runat="server">
                                                <asp:ListItem Value="Null" Selected="true">請選擇</asp:ListItem>
                                                <asp:ListItem Value="Y">審核成功</asp:ListItem>
                                                <asp:ListItem Value="N">審核失敗</asp:ListItem>
                                                <asp:ListItem Value="R">退件修正</asp:ListItem>
                                            </asp:DropDownList>
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:DropDownList ID="list_Verify" runat="server">
                                                    <asp:ListItem Value="Null">請選擇</asp:ListItem>
                                                    <asp:ListItem Value="Y">審核成功</asp:ListItem>
                                                    <asp:ListItem Value="N">審核失敗</asp:ListItem>
                                                    <asp:ListItem Value="R">退件修正</asp:ListItem>
                                                </asp:DropDownList>
                                                <input id="hid_vstatus" type="hidden" name="hid_vstatus" runat="server" />
                                                <asp:LinkButton ID="btn_BackVerify" runat="server" Text="還原" CommandName="back" CssClass="linkbutton"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="審核備註">
                                            <ItemTemplate>
                                                <asp:TextBox ID="txt_VerifyNote" runat="server" TextMode="MultiLine"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="保險證號">
                                            <ItemTemplate>
                                                <asp:Label ID="lab_ActNO" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="預算別">
                                            <ItemTemplate>
                                                <asp:DropDownList ID="list_BudID" runat="server">
                                                </asp:DropDownList>
                                                <input id="hid_OverPay" type="hidden" runat="server" name="hid_OverPay">
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="AuditCheck" runat="server" Text="經費審核確認" ToolTip="整班經費審核通過及不補助" CssClass="asp_button_M"></asp:Button>
                                <input id="HistorySh" onclick="HistoryShearch();" type="button" value="查詢重複參訓" name="HistorySh" runat="server" class="asp_button_M" />
                                <asp:Button ID="AuditCheckR" runat="server" Text="還原經費審核確認" ToolTip="整班經費審核通過及不補助" CssClass="asp_button_M" Visible="False"></asp:Button>
                                <input id="Dclass" type="hidden" name="Dclass" runat="server" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
