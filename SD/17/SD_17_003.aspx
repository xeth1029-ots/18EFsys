<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_17_003.aspx.vb" Inherits="WDAIIP.SD_17_003" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>補助審核</title>
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
        //'查詢鈕檢查
        function CheckSearch() {
            if (document.getElementById('OCIDValue1').value == '') {
                alert('請選擇職類班別');
                return false;
            }
        }

        function CheckDate() {
            //var msg = ""
            /*,if (document.form1.Dclass.value ==1){,//if(!confirm('身分證號碼錯誤，是否要繼續儲存?')) msg=msg+'身分證號碼錯誤\n';
             * ,if(!confirm('此班級的學員在其他計畫有參訓紀錄,按查詢重複參訓可以查看,您是否確定要經費審核確認?')) 
             * {msg=msg+'此班級的學員在其他計畫有參訓紀錄\n';},},*/
            //經費審核確認
            if (confirm('經費審核確認')) {
                return chkMoney();
            }
            else {
                return false;
            }
            /*,if(msg!=''){,alert(msg);,return false;,},else{,return chkMoney();,},*/
        }

        var cst_name = 1;   //學員姓名
        var cst_Sum = 7;    //補助金額
        var cst_Remain = 9; //剩餘可用
        var cst_hid_vstatus = 12; //審核狀態

        function SelectAll() {
            var tb = document.getElementById("Datagrid2");
            for (i = 1; i < tb.rows.length; i++) {
                if (tb.rows(i).cells(cst_hid_vstatus).children(0).disabled == false) {
                    tb.rows(i).cells(cst_hid_vstatus).children(0).value = tb.rows(0).cells(cst_hid_vstatus).children(0).value;
                }
            }
        }

        function chkMoney() {
            var msg = '';
            var MyTable = document.getElementById('DataGrid2');
            for (i = 1; i < MyTable.rows.length; i++) {
                var SumOfMoney = parseInt(MyTable.rows(i).cells(cst_Sum).children(0).value);
                var RemainSub = parseInt(MyTable.rows(i).cells(cst_Remain).children(0).value);
                //alert(MyTable.rows(i).cells(cst_hid_vstatus).children(1).value);
                //alert(MyTable.rows(i).cells(cst_hid_vstatus).children(0).value);
                if (MyTable.rows(i).cells(cst_hid_vstatus).children(1).value == '1') {
                    if (MyTable.rows(i).cells(cst_hid_vstatus).children(0).value == 'Y') {
                        if (SumOfMoney > RemainSub) {
                            msg += '補助金額' + SumOfMoney + '不能超過剩餘可用餘額' + RemainSub + '(學員:' + MyTable.rows(i).cells(cst_name).children(0).value + ')\n';
                        }
                    }
                }
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
        }
        /*
        function HistoryShearch() {
        window.open('SD_13_History.aspx?OCID='+document.getElementById('OCIDValue1').value,'history','width=700,height=600,scrollbars=1')
        }
        */
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
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">
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
                                <asp:TextBox ID="center" runat="server" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="Hidden2" runat="server" />
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="asp_button_Mini" />
                                <asp:Button ID="Button4" Style="display: none" runat="server"></asp:Button>
                                <asp:Button ID="Button3" Style="display: none" runat="server" Text="Button3"></asp:Button>
                                <span id="HistoryList2" style="z-index: 100; position: absolute; display: none" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="310px">
                                    </asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">職類/班別
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="asp_button_Mini" />
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                <span id="HistoryList" style="z-index: 101; position: absolute; display: none; left: 270px">
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
                        <table id="Table3" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td>
                                    <asp:Label ID="Label1" runat="server">審核成功學員數：</asp:Label>
                                    <asp:Label ID="SNum" runat="server"></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                <asp:Label ID="Label2" runat="server">審核失敗學員數：</asp:Label>
                                    <asp:Label ID="FNum" runat="server"></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                <asp:Label ID="Label3" runat="server">未審核學員數：</asp:Label>
                                    <asp:Label ID="ANum" runat="server"></asp:Label>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                <asp:Label ID="Label4" runat="server">退件修正學員數：</asp:Label>
                                    <asp:Label ID="RNum" runat="server"></asp:Label>
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
                                <asp:DataGrid ID="Datagrid2" runat="server" Width="100%" AllowSorting="True" CssClass="font" AutoGenerateColumns="False">
                                    <AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn SortExpression="StudentID" HeaderText="學號">
                                            <HeaderStyle ForeColor="Blue"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_StudentID" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="姓名">
                                            <HeaderStyle Width="25px"></HeaderStyle>
                                            <ItemTemplate>
                                                <input id="hid_Name" type="hidden" runat="server" name="hid_Name" />
                                                <asp:Label ID="lab_Star" runat="server" ForeColor="#B0E2FF">*</asp:Label>
                                                <asp:LinkButton ID="link_Name" runat="server" CssClass="l" CommandName="Link"></asp:LinkButton>
                                                <input type="hidden" id="hSOCID" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn SortExpression="IDNO" HeaderText="身分證號碼">
                                            <HeaderStyle ForeColor="#B0E2FF"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_IDNO" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="是否取得結訓資格">
                                            <HeaderStyle Width="25px"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_EndClass" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="出席達&lt;BR&gt;80%">
                                            <HeaderStyle Width="20px"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_OnClassRate" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="是否為&lt;BR&gt;在職者">
                                            <HeaderStyle Width="20px"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_WorkSuppIdent" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="是否補助">
                                            <HeaderStyle Width="16px"></HeaderStyle>
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
                                                <input id="hid_Balance" type="hidden" runat="server" name="hid_Balance">
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
                                                <input id="hid_vstatus" type="hidden" runat="server" name="hid_vstatus" />
                                                <asp:LinkButton ID="btn_BackVerify" runat="server" Text="還原" CommandName="back" CssClass="linkbutton"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="審核備註">
                                            <ItemTemplate>
                                                <asp:TextBox ID="txt_VerifyNote" runat="server" Width="134px" TextMode="MultiLine"></asp:TextBox>
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
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <asp:Button ID="AuditCheck" runat="server" Text="經費審核確認" ToolTip="整班經費審核通過及不補助" CssClass="asp_button_M"></asp:Button>&nbsp;
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
