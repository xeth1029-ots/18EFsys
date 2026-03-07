<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_17_001.aspx.vb" Inherits="WDAIIP.SD_17_001" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>補助申請</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        /*
        function HistoryShearch()
        {
        window.open('SD_13_History.aspx?OCID='+document.getElementById('OCIDValue1').value,'history','width=700,height=600,scrollbars=1')
        }
			
        */
        var cst_studentid = 0; //學號
        var cst_name = 1; //姓名
        var cst_idno = 2; //身分證號碼
        var cst_creditpoints = 3; //是否獲得學分
        var cst_ass80percent = 4; //出席達80%
        var cst_worksuppident = 5; //是否為在職者
        var cst_bonus = 6; //是否補助
        var cst_total = 7; //總費用
        var cst_sumofmoney = 8; //補助費用

        var cst_personalpay = 9;   //個人支付

        var cst_balancemoney = 10; //剩餘可用餘額
        var cst_GovAppl = 11;  //其他申請中金額
        var cst_petition = 12; //是否提出申請
        var cst_petitionstate = 13; //申請狀態
        var cst_budid = 14; //預算別

        function GETvalue() {
            document.getElementById('Button4').click();
        }
        function SetOneOCID() {
            document.getElementById('Button5').click();
        }
        function choose_class() {
            if (document.getElementById('OCID1').values == '') {
                document.getElementById('Button5').click();
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
        function ChangeMoney(num, obj, obj2, obj3, obj4) {
            var MyTable = document.getElementById('DataGrid1');
            //debugger;
            if (isUnsignedInt(document.getElementById(obj).value)) {
                if (parseInt(document.getElementById(obj).value) > parseInt(document.getElementById(obj2).value)) {
                    document.getElementById(obj).select();
                    alert('超過剩餘補助金額,此學員的剩餘補助金額為' + document.getElementById(obj2).value);
                    document.getElementById(obj).value = document.getElementById(obj2).value;
                }
                else if (parseInt(document.getElementById(obj).value) > parseInt(document.getElementById(obj3).value)) {
                    document.getElementById(obj).select();
                    alert('超過最大補助金額,最大補助金額為' + document.getElementById(obj3).value);
                    document.getElementById(obj).value = document.getElementById(obj3).value;
                }
                else {
                    if (MyTable.rows(num).cells(cst_total).innerHTML - parseInt(document.getElementById(obj).value) >= 0) {

                        MyTable.rows(num).cells(cst_personalpay).innerHTML = MyTable.rows(num).cells(cst_total).innerHTML - parseInt(document.getElementById(obj).value);
                        document.getElementById(obj4).value = MyTable.rows(num).cells(cst_personalpay).innerHTML;

                        var Balancemoney = parseInt(document.getElementById(obj2).value);
                        var Total = parseInt(document.getElementById(obj2).value) - parseInt(document.getElementById(obj).value);
                        if (Total >= 0)
                            //MyTable.rows(num).cells(cst_balancemoney).innerHTML=Total;
                            MyTable.rows(num).cells(cst_balancemoney).innerHTML = Balancemoney;
                        else
                            //MyTable.rows(num).cells(cst_balancemoney).innerHTML='<font color=Red>'+Total+'</font>';
                            MyTable.rows(num).cells(cst_balancemoney).innerHTML = '<font color=Red>' + Balancemoney + '</font>';
                    }
                    else {
                        document.getElementById(obj).focus();
                        document.getElementById(obj).value = 0;
                        alert('補助費用大於總費用,請輸入合理數字');
                    }
                }
            }
            else {
                document.getElementById(obj).focus();
                document.getElementById(obj).value = 0;
                //alert('請輸入數字');
            }
        }


        function CheckData() {
            //alert(cst_balancemoney);
            var MyTable = document.getElementById('DataGrid1');
            var msg = '';

            for (i = 1; i < MyTable.rows.length; i++) {
                var Sign = MyTable.rows(i).cells(cst_studentid).children(0).value;
                var SumOfMoney = parseInt(MyTable.rows(i).cells(cst_sumofmoney).children(0).value);
                var RemainSub = parseInt(MyTable.rows(i).cells(cst_sumofmoney).children(1).value);
                var MaxSub = parseInt(MyTable.rows(i).cells(cst_sumofmoney).children(2).value);
                var MyCehck = MyTable.rows(i).cells(cst_petition).children(0); //是否提出申請
                var balancemoney = MyTable.rows(i).cells(cst_balancemoney).innerHTML;
                var personalpay = MyTable.rows(i).cells(cst_personalpay).innerHTML; //個人支付
                //alert((balancemoney));
                //alert(isNegativeInt(balancemoney));
                //msg+='(學員:'+MyTable.rows(i).cells(cst_name).innerHTML+')\n申請補助總額超過剩餘可用餘額，請確認資料正確性\n';

                if (isNegativeInt(personalpay)) {
                    msg += '(學員:' + MyTable.rows(i).cells(cst_name).innerHTML + ')\n個人支付額不可為負，請確認資料正確性\n';
                }

                if (isNegativeInt(balancemoney)) {
                    msg += '(學員:' + MyTable.rows(i).cells(cst_name).innerHTML + ')\n申請補助總額超過剩餘可用餘額，請確認資料正確性\n';
                }
                if (!MyCehck.disabled && MyCehck.checked) {
                    if (!isUnsignedInt(SumOfMoney)) {
                        msg += '補助金額必須為數字(學員:' + MyTable.rows(i).cells(cst_name).innerHTML + ')\n';
                    }
                    else {
                        if (SumOfMoney > RemainSub)
                            msg += '補助金額不能超過剩餘補助金額(學員:' + MyTable.rows(i).cells(cst_name).innerHTML + ',剩餘金額' + RemainSub + ')\n';
                        else if (SumOfMoney > MaxSub)
                            msg += '補助金額不能超此班最大補助金額(學員:' + MyTable.rows(i).cells(cst_name).innerHTML + ')\n';
                    }
                    if (Sign != '') {
                        msg += '未填寫調查表無法申請補助金(學員:' + MyTable.rows(i).cells(cst_name).innerHTML + ')\n';
                    }
                }
            }
            if (document.form1.Dclass.value == 1) {
                if (!confirm('此班級的學員在其他計畫有參訓紀錄,按查詢重複參訓可以查看,您是否確定要儲存?')) { msg = msg + '學員在其他計畫有參訓\n'; }
            }

            if (msg != '') {
                alert(msg);
                return false;
            }

            if (msg == '') {
                return Chk_Blacklist();
            }
        }

        function Chk_Blacklist() {
            var msg = "";
            msg += document.getElementById('hidBlackMsg').value;
            if (msg != "") {
                msg += "\n詳情請至教務管理-學員黑名單查詢\n是否續繼儲存?";
            }
            if (msg != "") {
                if (!confirm(msg)) {
                    return false;
                }
            }
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
										首頁&gt;&gt;學員動態管理&gt;&gt;補助金請領&gt;&gt;<FONT color="#990000">補助申請</FONT>
                                </asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="table_nw" id="Table2" cellspacing="1" cellpadding="1" width="584">
                        <tr>
                            <td class="bluecol" width="100">訓練機構
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="Hidden2" runat="server" />
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="asp_button_Mini" />
                                <asp:Button ID="Button5" Style="display: none" runat="server"></asp:Button>
                                <asp:Button ID="Button4" Style="display: none" runat="server" Text="Button4"></asp:Button>
                                <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
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
                                <span id="HistoryList" style="position: absolute; display: none; left: 270px">
                                    <asp:Table ID="HistoryTable" runat="server" Width="310">
                                    </asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">是否為在職者
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="rblWork" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="Y" Selected="True">是</asp:ListItem>
                                    <asp:ListItem Value="N">否</asp:ListItem>
                                    <asp:ListItem Value="A">不區分</asp:ListItem>
                                </asp:RadioButtonList>
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
                    <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>滑鼠移至姓名可以顯示此學員前三年的申請紀錄
                                <%--
										<br>
										<FONT color="#ff0000" size="2">*表示為該學員未填寫調查表</FONT>
                                --%>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AllowSorting="True" CssClass="font" AutoGenerateColumns="False">
                                    <AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn SortExpression="StudentID" HeaderText="學號">
                                            <HeaderStyle ForeColor="#B0E2FF" Width="50px"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:TextBox ID="star1" runat="server" Width="10px" onfocus="this.blur()" ForeColor="Red" Columns="8">*</asp:TextBox>
                                                <asp:TextBox ID="stud1" runat="server" Width="25px" onfocus="this.blur()" Columns="8" Enabled="False" MaxLength="3"></asp:TextBox>
                                                <input id="setid" style="width: 22px; height: 22px" type="hidden" runat="server" />
                                                <input id="ocid" style="width: 22px; height: 22px" type="hidden" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="Name" HeaderText="姓名">
                                            <HeaderStyle Width="50px"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="IDNO" SortExpression="IDNO" HeaderText="身分證號碼">
                                            <HeaderStyle ForeColor="#B0E2FF"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="是否取得&lt;BR&gt;結訓資格">
                                            <ItemTemplate>
                                                <asp:DataGrid ID="DataGrid2" runat="server" Width="300px" CssClass="font" AutoGenerateColumns="False" BorderWidth="1px" BorderColor="Black" BackColor="LemonChiffon">
                                                    <Columns>
                                                        <asp:BoundColumn DataField="ClassCName" HeaderText="班別名稱"></asp:BoundColumn>
                                                        <asp:BoundColumn DataField="SumOfMoney" HeaderText="金額">
                                                            <HeaderStyle Width="40px"></HeaderStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="AppliedStatus" HeaderText="申請狀態">
                                                            <HeaderStyle Width="60px"></HeaderStyle>
                                                        </asp:BoundColumn>
                                                    </Columns>
                                                </asp:DataGrid>
                                                <asp:Label ID="CreditPoints" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn HeaderText="出席達&lt;BR&gt;80%"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="是否為&lt;BR&gt;在職者"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="是否&lt;BR&gt;補助"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="Total" HeaderText="總費用"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="補助費用">
                                            <ItemTemplate>
                                                <asp:TextBox ID="SumOfMoney" runat="server" Width="62px" Columns="8"></asp:TextBox>
                                                <input id="RemainSub" type="hidden" runat="server">
                                                <input id="MaxSub" type="hidden" runat="server">
                                                <input id="PayMoney" type="hidden" runat="server">
                                                <input id="balancemoney" type="hidden" runat="server">
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="個人支付">
                                            <ItemTemplate>
                                                <asp:Label ID="personalpay" runat="server"></asp:Label>
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn HeaderText="剩餘可&lt;BR&gt;用餘額">
                                            <HeaderStyle Width="40px"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="GovAppl2" HeaderText="其他申請中金額">
                                            <HeaderStyle Width="50px"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="是否提&lt;BR&gt;出申請">
                                            <ItemTemplate>
                                                <input id="Checkbox1" type="checkbox" runat="server">&nbsp;
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="AppliedStatusM" HeaderText="申請狀態"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="AppliedStatus" HeaderText="撥款狀態"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="預算別">
                                            <ItemTemplate>
                                                <asp:DropDownList ID="BudID" runat="server">
                                                </asp:DropDownList>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <asp:Button ID="Button3" runat="server" Text="儲存" CssClass="asp_button_S"></asp:Button>
                                <input id="Dclass" type="hidden" name="Dclass" runat="server" />
                                <input id="hidBlackMsg" type="hidden" name="hidBlackMsg" runat="server" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
