<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_09_001_R.aspx.vb" Inherits="WDAIIP.SD_09_001_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>列印學員名冊</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
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
        function GETvalue() { document.getElementById('Button12').click(); }

        function choose_class() {
            var RID = document.form1.RIDValue.value;
            openClass('../02/SD_02_ch.aspx?RID=' + RID);
        }
        /*function ReportPrint(){var msg='';if(document.form1.OCIDValue1.value=='') {msg+='請選擇班級職類\n';}if (msg!=''){alert(msg);return false;}return true;}*/
        function ChangeAll(obj) {
            var objLen = document.form1.length;
            for (var iCount = 0; iCount < objLen; iCount++) {
                if (obj.checked == true) {
                    if (document.form1.elements[iCount].type == "checkbox") {
                        if (document.form1.elements[iCount].disabled == false) {
                            document.form1.elements[iCount].checked = true;
                        }
                    }
                }
                else {
                    if (document.form1.elements[iCount].type == "checkbox") {
                        document.form1.elements[iCount].checked = false;
                    }
                }
            }
        }

        function CheckPrint() {
            var flag = false;
            var MyTable = document.getElementById('DataGrid1');
            var OCID = '';
            var PrintValue = document.getElementById('PrintValue');

            //var OCID=document.getElementById('OCIDValue1').value;
            for (var i = 1; i < MyTable.rows.length; i++) {
                var MyCheck = false;

                var Mycells0 = MyTable.rows[i].cells[0];
                if (Mycells0.children[2] == null) {
                    MyCheck = (Mycells0.children[1] == null) ? Mycells0.children[0] : Mycells0.children[1];
                }
                else {
                    MyCheck = Mycells0.children[2];
                }
                if (MyCheck.checked) {
                    flag = true;
                    if (OCID != '') OCID += ',';
                    OCID += '\'' + MyCheck.value + '\'';
                }
            }
            PrintValue.value = OCID;
            if (PrintValue.value == '') {
                alert('請勾選要列印的學員班級!');
                return false;
            }
            //else{
            //alert(document.getElementById('PrintValue').value);
            //Student_Report
            //url='../../SQControl.aspx?&SQ_AutoLogout=true&sys=Member&filename=Student_Report&path='+SMpath+'&'
            //window.open(url+'OCID='+document.getElementById('PrintValue').value+'&RID='+document.getElementById('RIDValue').value,'print','toolbar=0,location=0,status=0,menubar=0,resizable=1');
            //}
        }

        function InsertValue(Flag, MyValue) {
            //alert(Flag);
            //alert(MyValue);
            var PrintValue = document.getElementById('PrintValue');
            if (Flag) {
                if (PrintValue.value.indexOf('\'' + MyValue + '\'') == -1) {
                    if (PrintValue.value != '') PrintValue.value += ',';
                    PrintValue.value += '\'' + MyValue + '\'';
                }
            }
            else {
                if (PrintValue.value.indexOf('\'' + MyValue + '\'') != -1) {
                    if (PrintValue.value.indexOf(',\'' + MyValue + '\'') != -1) {
                        PrintValue.value = PrintValue.value.replace(',\'' + MyValue + '\'', '');
                    }
                    if (PrintValue.value.indexOf('\'' + MyValue + '\',') != -1) {
                        PrintValue.value = PrintValue.value.replace('\'' + MyValue + '\',', '');
                    }
                    if (PrintValue.value.indexOf('\'' + MyValue + '\'') != -1) {
                        PrintValue.value = PrintValue.value.replace('\'' + MyValue + '\'', '');
                    }
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
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務報表管理&gt;&gt;列印學員名冊</asp:Label>
                </td>
            </tr>
        </table>
        <table cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" style="width: 20%">訓練機構
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini" />
                                <asp:Button ID="Button12" Style="display: none" runat="server" Text="Button12"></asp:Button>
                                <span id="HistoryList2" style="z-index: 100; position: absolute; display: none" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">職類/班別
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="button_b_Mini" />
                                <input id="OCIDValue1" type="hidden" name="Hidden2" runat="server" />
                                <input id="TMIDValue1" type="hidden" name="Hidden1" runat="server" />
                                <span id="HistoryList" style="z-index: 101; position: absolute; display: none; left: 270px">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label>
                            </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="30%"></asp:TextBox>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="button_b_Mini" />
                                <input id="cjobValue" type="hidden" name="cjobValue" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">期別
                            </td>
                            <td class="whitecol">
                                <input id="CyclType" name="CyclType" runat="server" style="width: 15%" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">班級範圍
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="ClassRound" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="開訓二週前">開訓二週前</asp:ListItem>
                                    <asp:ListItem Value="已開訓">已開訓</asp:ListItem>
                                    <asp:ListItem Value="已結訓">已結訓</asp:ListItem>
                                    <asp:ListItem Value="未開訓">未開訓</asp:ListItem>
                                    <asp:ListItem Value="全部" Selected="True">全部</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <div align="center" class="whitecol">
                        <asp:Button ID="Query" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button><br />
                        <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                    </div>
                    <table id="Table4" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td class="font">
                                <font color="red">(有 * 號的班級代表至少有一個學員，於學員資料維護的必填資料未填)</font>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn HeaderText="選取">
                                            <HeaderStyle Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <HeaderTemplate>
                                                選取<input id="CheckboxAll" type="checkbox" runat="server" name="CheckboxAll" />
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:Label ID="LabelStar" runat="server" ForeColor="Red">*</asp:Label>
                                                <input id="Checkbox1" type="checkbox" value='<%# DataBinder.Eval(Container.DataItem,"OCID")%>' runat="server" name="Checkbox1" />
                                                <asp:Label ID="labOCID" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="TrainName2" HeaderText="訓練職類">
                                            <HeaderStyle HorizontalAlign="Center" Width="19%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassID" HeaderText="班級代碼">
                                            <HeaderStyle HorizontalAlign="Center" Width="19%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="CLASSCNAME2" HeaderText="班級名稱">
                                            <HeaderStyle HorizontalAlign="Center" Width="19%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDate" HeaderText="開訓日期" DataFormatString="{0:d}">
                                            <HeaderStyle HorizontalAlign="Center" Width="19%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="IsApplic" HeaderText="志願班別">
                                            <HeaderStyle HorizontalAlign="Center" Width="19%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
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
            <tr>
                <td>
                    <div align="center" class="whitecol">
                        <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                    </div>
                </td>
            </tr>
        </table>
        <input id="PrintValue" type="hidden" runat="server" />
    </form>
</body>
</html>
