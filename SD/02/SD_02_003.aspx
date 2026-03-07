<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_02_003.aspx.vb" Inherits="WDAIIP.SD_02_003" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD html 4.0 Transitional//EN">
<html>
<head>
    <title>錄訓作業</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function savedataCHK1() {
            var rst1 = true; //正常再次檢核。
            var vMsgchk1 = "是否已完成錄取作業，送出後，將不可再進行錄取作業修改!";
            rst1 = confirm(vMsgchk1);
            return rst1;
        }

        function GETvalue() {
            document.getElementById('Button4').click();
        }

        function choose_class() {
            var RIDValue = document.getElementById('RIDValue');
            openClass('../02/SD_02_ch.aspx?RID=' + RIDValue.value);
        }

        //function choose_other() {
        //    var OCID = document.form1.OCIDValue1.value; 挑選其他志願
        //    var OCIDValue1 = document.getElementById('OCIDValue1');
        //    window.open('SD_02_003_other.aspx?OCID=' + OCIDValue1.value, '', 'width=550,height=250,location=0,status=0,menubar=0,scrollbars=1,resizable=0');
        //}

        function search() {
            var OCIDValue1 = document.getElementById('OCIDValue1');
            if (OCIDValue1.value == '') {
                alert('請選擇職類班別!');
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;錄訓作業</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Panel ID="Table_3" runat="server">
                        <table class="table_nw" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol_need" style="width: 20%">訓練機構 </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                    <input id="button8" type="button" value="..." runat="server" class="asp_button_Mini" />
                                    <asp:Button ID="Button4" Style="display: none" runat="server" Text="Button4"></asp:Button>
                                    <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                        <asp:Table ID="historyrid" runat="server" Width="100%"></asp:Table>
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">職類/班別 </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                    <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                    <input onclick="choose_class()" type="button" value="..." class="asp_button_Mini" />
                                    <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                    <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                    <span id="HistoryList" style="position: absolute; display: none; left: 28%">
                                        <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                    </span>
                                </td>
                            </tr>
                            <tr id="tr_rblsortmode" runat="server">
                                <td class="bluecol">排序方式 </td>
                                <td class="whitecol">
                                    <asp:RadioButtonList ID="rblsortmode" runat="server" RepeatLayout="flow" RepeatDirection="horizontal" CssClass="font">
                                        <asp:ListItem Value="1" Selected="true">成績排序</asp:ListItem>
                                        <asp:ListItem Value="2">准考證號排序</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                        </table>
                        <table width="100%">
                            <tr>
                                <td class="whitecol" align="center">
                                    <asp:Button ID="button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button></td>
                            </tr>
                        </table>
                    </asp:Panel>
                    <div align="center">&nbsp;<asp:Label ID="lab_msg1" runat="server" ForeColor="red" Font-Size="10"></asp:Label></div>
                    <table id="table4" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server" class="font">
                        <tr>
                            <td>
                                <asp:Label ID="classname" runat="server"></asp:Label>
                                <asp:Label ID="argrole" runat="server"></asp:Label><br />
                                <asp:Label ID="labmsg219" runat="server" ForeColor="Blue"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="false" AllowSorting="true" Style="z-index: 0" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:BoundColumn DataField="examno" SortExpression="examno" HeaderText="准考證序號">
                                            <HeaderStyle HorizontalAlign="Center" ForeColor="#B0E2FF" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="SIGNNO" SortExpression="SIGNNO" HeaderText="e網報名序號">
                                            <HeaderStyle HorizontalAlign="Center" ForeColor="#B0E2FF" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="姓名">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" Wrap="false"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="labstarSOLDER" runat="server"></asp:Label>
                                                <asp:Label ID="labname" runat="server"></asp:Label>
                                                <asp:HiddenField ID="Hid_SETID" runat="server" />
                                                <asp:HiddenField ID="Hid_EnterDate" runat="server" />
                                                <asp:HiddenField ID="Hid_SerNum" runat="server" />
                                                <asp:HiddenField ID="Hid_rsort" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="writeresult" HeaderText="筆試成績">
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="oralresult" HeaderText="口試成績">
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="totalresult" SortExpression="totalresult" HeaderText="總成績">
                                            <HeaderStyle HorizontalAlign="Center" ForeColor="#B0E2FF" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ENTERDATE" HeaderText="報名日期">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="rsort" HeaderText="名次" SortExpression="rsort">
                                            <HeaderStyle HorizontalAlign="Center" ForeColor="#B0E2FF" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="甄試結果">
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:DropDownList ID="ddlResult" runat="server"></asp:DropDownList>
                                                <asp:Label ID="labPathW" runat="server"></asp:Label>
                                                <input id="hidIDNO" type="hidden" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="selsort" HeaderText="備取名次">
                                            <HeaderStyle Width="5%" />
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="備取或未錄取原因">
                                            <HeaderStyle HorizontalAlign="Center" Width="15%" VerticalAlign="middle"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:TextBox ID="notes2" runat="server" Width="90%" MaxLength="150" TextMode="multiline"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                                <div align="center" class="whitecol">
                                    <asp:Button ID="btnSEND1" runat="server" Text="完成錄取" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="btnPrint1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                                    <asp:Button ID="btnSAVE2" runat="server" Text="送出-解鎖" CssClass="asp_Export_M"></asp:Button>
                                </div>

                            </td>
                        </tr>
                    </table>
                    <div align="center">&nbsp;<asp:Label ID="lab_msg_r1" runat="server" ForeColor="red"></asp:Label></div>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:DataGrid ID="Datagrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="false">
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <HeaderStyle CssClass="head_navy" />
                        <Columns>
                            <asp:BoundColumn DataField="name" HeaderText="姓名"></asp:BoundColumn>
                            <asp:BoundColumn DataField="selsort" HeaderText="原備取排名"></asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="備取排名">
                                <ItemTemplate>
                                    <asp:DropDownList ID="DrpWatingS" runat="server" Width="100px" Height="22px">
                                    </asp:DropDownList>
                                    <asp:HiddenField ID="Hid_SETID" runat="server" />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <%--<asp:BoundColumn Visible="false" DataField="setid" HeaderText="setid"></asp:BoundColumn>--%>
                        </Columns>
                    </asp:DataGrid>
                </td>
            </tr>
            <tr align="center" class="whitecol">
                <td>
                    <asp:Button ID="btnsave" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                    <asp:Button ID="btncancel" runat="server" Text="取消" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
        </table>
        <input id="isBlack" type="hidden" name="isBlack" runat="server">
        <input id="Blackorgname" type="hidden" name="Blackorgname" runat="server">
        <asp:HiddenField ID="Hid_ChangeWating" runat="server" />
        <asp:HiddenField ID="Hid_StudTNum" runat="server" />
        <asp:HiddenField ID="Hid_OCID1" runat="server" />
        <asp:HiddenField ID="Hid_CFGUID" runat="server" />
        <asp:HiddenField ID="Hid_CFSEQNO" runat="server" />
        <asp:HiddenField ID="Hid_DG_SORT1" runat="server" />
        <input id="ItemVar1" type="hidden" name="ItemVar1" runat="server" />
        <input id="ItemVar2" type="hidden" name="ItemVar2" runat="server" />
        <asp:HiddenField ID="Hid_CAN_IGNORE_RULE1_CNT" runat="server" />
    </form>
</body>
</html>
