<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_01_001_ch.aspx.vb" Inherits="WDAIIP.SD_01_001_ch" %>


<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>職類班級選擇</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css" />
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
        function chkdata() {
            var msg = '';
            var start_date = document.getElementById('start_date');
            var end_date = document.getElementById('end_date');
            if (start_date.value != '') {
                if (!checkDate(start_date.value))
                    msg += '[開訓日期起始]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }
            if (end_date.value != '') {
                if (!checkDate(end_date.value))
                    msg += '[開訓日期迄止]不是正確的日期格式,請輸入正確的日期格式,YYYY/MM/DD!!\n';
            }
            if (msg != '') {
                window.alert(msg);
                return false;
            }
        }

        function returnNum() {
            window.opener.form1.TMID1.value = document.form1.class1.value
        }

        function SelectItem(num) {
            var MyTable = document.getElementById('DataGrid2');
            for (i = 1; i < MyTable.rows.length; i++) {
                MyTable.rows[i].cells[0].children[0].checked = false;
            }
            MyTable.rows(num).cells(0).children(0).checked = true;
        }

        //送出檢查是否有選擇班級
        function CheckData(num) {
            var MyTable;
            if (num == 1)
                MyTable = document.getElementById('DataGrid1');
            else
                MyTable = document.getElementById('DataGrid2');
            var Flag = false;
            for (i = 1; i < MyTable.rows.length; i++) {
                if (MyTable.rows[i].cells[0].children[0].checked)
                    Flag = true;
            }
            if (!Flag) {
                alert('請選擇班級');
                return false;
            }
        }

        //清除日期欄位內容
        function clearDate(objId) {
            var myObj = document.getElementById(objId);
            myObj.value = "";
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <div style="overflow-y: auto; height: 640px;">
            <table class="table_nw" id="Table1" cellspacing="1" cellpadding="1" width="100%">
                <tr id="trCenter" runat="server">
                    <td class="bluecol" width="20%">訓練機構 </td>
                    <td class="whitecol" colspan="3" width="80%">
                        <asp:TextBox ID="center" runat="server" Width="60%"></asp:TextBox>
                        <input id="RIDValue" type="hidden" runat="server" />
                        <input id="Button8" type="button" value="..." runat="server" class="asp_button_Mini" />
                    </td>
                </tr>
                <tr>
                    <td class="bluecol" width="20%">班級名稱 </td>
                    <td class="whitecol" width="30%">
                        <asp:TextBox ID="ClassCName" runat="server" MaxLength="30" Width="80%"></asp:TextBox></td>
                    <td class="bluecol" width="20%">期別 </td>
                    <td class="whitecol" width="30%">
                        <asp:TextBox ID="CyclType" runat="server" Columns="5" MaxLength="2" Width="30%"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol" width="20%">訓練職類 </td>
                    <td class="whitecol" colspan="3" width="80%">
                        <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Width="40%"></asp:TextBox>
                        <input id="trainValue" type="hidden" name="trainValue" runat="server" />
                        <input onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." class="asp_button_Mini" />&nbsp;
				<input onclick="document.getElementById('trainValue').value = ''; document.getElementById('TB_career_id').value = '';" type="button" value="清除" class="asp_button_S" />
                    </td>
                </tr>
                <tr>
                    <td class="bluecol" width="20%">開訓日期 </td>
                    <td class="whitecol" colspan="3" width="80%">
                        <asp:TextBox ID="start_date" runat="server" onfocus="this.blur()" Width="26%" MaxLength="10"></asp:TextBox>
                        <span id="span1" runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span>
                        <span id="span3" runat="server">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= start_date.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span> ～
                <asp:TextBox ID="end_date" runat="server" onfocus="this.blur()" Width="26%" MaxLength="10"></asp:TextBox>
                        <span id="span2" runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span>
                        <span id="span4" runat="server">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= end_date.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol" width="20%">班級種類 </td>
                    <td class="whitecol" colspan="3" width="80%">
                        <asp:RadioButtonList ID="ClassSort" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                            <asp:ListItem Value="1" Selected="True">報名尚未結束班級</asp:ListItem>
                            <asp:ListItem Value="2">報名結束班級</asp:ListItem>
                            <asp:ListItem Value="4">尚未甄試班級</asp:ListItem>
                            <asp:ListItem Value="5">未結訓班級</asp:ListItem>
                            <asp:ListItem Value="3">所有的班級</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol" width="20%">階段班級 </td>
                    <td class="whitecol" colspan="3" width="80%">
                        <asp:DropDownList ID="LevelFlag" runat="server">
                            <asp:ListItem Value="否" Selected="True">否</asp:ListItem>
                            <asp:ListItem Value="是">是</asp:ListItem>
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td colspan="4" class="whitecol">
                        <p align="center">
                            <asp:Button ID="search_but" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>
                        </p>
                    </td>
                </tr>
                <tr>
                    <td colspan="4" class="whitecol" align="center">
                        <table id="DataGridTable1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td>
                                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False">
                                        <AlternatingItemStyle BackColor="#EEEEEE" />
                                        <HeaderStyle CssClass="head_navy" />
                                        <Columns>
                                            <asp:TemplateColumn>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <input id="radio1" value='<%# DataBinder.Eval(Container.DataItem, "OCID")%>' type="radio" name="class1" />
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <input id="chkbox1" value='<%# DataBinder.Eval(Container.DataItem, "OCID")%>' type="checkbox" name="class2">
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:BoundColumn DataField="TrainName" HeaderText="訓練職類">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="ClassCName" HeaderText="班級名稱">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="STDate" HeaderText="開訓日期" DataFormatString="{0:d}">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="IsApplic" HeaderText="志願班別" Visible="false">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn Visible="False" DataField="OCID"></asp:BoundColumn>
                                            <asp:BoundColumn Visible="False" DataField="TrainID"></asp:BoundColumn>
                                            <asp:BoundColumn Visible="False" DataField="CyclType"></asp:BoundColumn>
                                            <asp:BoundColumn Visible="False" DataField="LevelType"></asp:BoundColumn>
                                        </Columns>
                                    </asp:DataGrid>
                                </td>
                            </tr>
                            <tr>
                                <td align="center">
                                    <asp:Button ID="send" runat="server" Text="送出" CssClass="asp_button_S"></asp:Button></td>
                            </tr>
                        </table>
                        <table id="DataGridTable2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td>
                                    <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False">
                                        <AlternatingItemStyle BackColor="#EEEEEE" />
                                        <HeaderStyle CssClass="head_navy" />
                                        <Columns>
                                            <asp:TemplateColumn>
                                                <HeaderStyle Width="12%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <input id="CCLID" type="radio" value="Radio2" runat="server" />
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:BoundColumn DataField="ClassCName" HeaderText="班級階段名稱">
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="LevelName" HeaderText="課程階段">
                                                <HeaderStyle Width="20%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="LevelSDate" HeaderText="階段開訓日期" DataFormatString="{0:d}">
                                                <HeaderStyle Width="20%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                        </Columns>
                                    </asp:DataGrid>
                                </td>
                            </tr>
                            <tr>
                                <td align="center">
                                    <asp:Button ID="Button1" runat="server" Text="送出" CssClass="asp_button_S"></asp:Button></td>
                            </tr>
                        </table>
                        <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                    </td>
                </tr>
            </table>
            <input id="StudIDNO" type="hidden" name="StudIDNO" runat="server" />
            <input id="Hid_IJC" runat="server" name="Hid_IJC" type="hidden" size="1">
        </div>
    </form>
</body>
</html>
