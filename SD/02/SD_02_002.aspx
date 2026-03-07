<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_02_002.aspx.vb" Inherits="WDAIIP.SD_02_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>甄試結果試算</title>
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
        function GETvalue() {
            document.getElementById('Button4').click();
        }

        function chk() {
            /*if (document.form1.start_date.value=='' || document.form1.end_date.value==''){
			window.alert('請選擇日期範圍');
			return false;
			}*/
        }

        function chall(num) {
            if (num == 1) {
                document.form1.OCID_Grade.checked = document.form1.Choose1.checked
                for (var i = 0; i < document.form1.OCID_Grade.length; i++)
                    document.form1.OCID_Grade[i].checked = document.form1.Choose1.checked
            }
            else {
                document.form1.OCID_Sort.checked = document.form1.Choose2.checked
                for (var i = 0; i < document.form1.OCID_Sort.length; i++)
                    document.form1.OCID_Sort[i].checked = document.form1.Choose2.checked
            }
        }

        function choose_class() {
            openClass('../02/SD_02_ch.aspx?RID=' + document.form1.RIDValue.value);
        }

        function search() {
            if (document.form1.OCIDValue1.value == '') {
                alert('請先選擇班別\n');
                return false;
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;甄試結果試算</asp:Label>
                </td>
            </tr>
        </table>
        <table id="table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_nw" id="table3" cellspacing="1" cellpadding="1" width="100%">
                        <!--test start-->
                        <tr>
                            <td class="bluecol" style="width: 20%">訓練機構</td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="Button8" type="button" value="..." name="Button5" runat="server" class="asp_button_Mini" />
                                <asp:Button ID="Button4" Style="display: none" runat="server" Text="Button5" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="Button5" Style="display: none" runat="server" Text="Button5" CssClass="asp_button_M"></asp:Button>
                                <span id="HistoryList2" style="z-index: 100; position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">職類/班級</td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="asp_button_Mini" />
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                <span id="HistoryList" style="z-index: 102; position: absolute; display: none; left: 28%">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <!-- test end -->
                        <tr>
                            <td class="bluecol">通俗職類</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="30%"></asp:TextBox>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="asp_button_Mini" />
                                <input id="cjobValue" type="hidden" name="cjobValue" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">班級名稱</td>
                            <td class="whitecol">
                                <asp:TextBox ID="classname" runat="server" Width="20%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">開訓日期</td>
                            <td class="whitecol" colspan="3">
                                <span id="span1" runat="server">
                                    <asp:TextBox ID="start_date" runat="server" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.Clientid %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                                    ～
							    <asp:TextBox ID="end_date" runat="server" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.Clientid %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                                    &nbsp;
                                </span>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button></td>
                        </tr>
                    </table>
                    <table class="table_nw" id="Table4" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td width="50%" class="bluecol">依成績</td>
                            <td width="50%" class="bluecol">依報名先後</td>
                        </tr>
                        <tr>
                            <td valign="top" width="50%" class="whitecol">
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="false" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <HeaderTemplate>
                                                <input onclick="chall(1)" type="checkbox" name="Choose1">
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <input type="checkbox" value='<%#convert.tostring(databinder.eval(container.dataitem, "ocid"))%>' name="OCID_Grade">
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="classcname" HeaderText="班級名稱">
                                            <HeaderStyle HorizontalAlign="Center" Width="75%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="stdate" HeaderText="開訓日期" DataFormatString="{0:d}">
                                            <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn Visible="false" DataField="cycltype"></asp:BoundColumn>
                                        <asp:BoundColumn Visible="false" DataField="leveltype"></asp:BoundColumn>
                                    </Columns>
                                </asp:DataGrid>
                                <div align="center" class="whitecol">
                                    <asp:Button ID="Button2" runat="server" Text="計算名次" CssClass="asp_button_M"></asp:Button>&nbsp;
								<asp:Button ID="Button6" runat="server" Text="清除試算" Visible="false" CssClass="asp_button_M"></asp:Button>
                                </div>
                                <div style="margin-top: 3px; margin-bottom: 3px" align="center">
                                    <asp:Label ID="msg1" runat="server" ForeColor="red"></asp:Label></div>
                            </td>
                            <td valign="top" width="50%" class="whitecol">
                                <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="false" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <HeaderTemplate>
                                                <input type="checkbox" name="Choose2" onclick="chall(2)">
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <input type="checkbox" value='<%# convert.tostring(databinder.eval(container.dataitem, "OCID1"))%>' name="OCID_Sort">
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="classcname" HeaderText="班級名稱">
                                            <HeaderStyle HorizontalAlign="Center" Width="75%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="stdate" HeaderText="開訓日期" DataFormatString="{0:d}">
                                            <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn Visible="false" DataField="cycltype"></asp:BoundColumn>
                                        <asp:BoundColumn Visible="false" DataField="leveltype"></asp:BoundColumn>
                                    </Columns>
                                </asp:DataGrid>
                                <div align="center" class="whitecol">
                                    <asp:Button ID="Button3" runat="server" Text="計算名次" CssClass="asp_button_M"></asp:Button>&nbsp;
								<asp:Button ID="Button7" runat="server" Text="清除試算" Visible="false" CssClass="asp_button_M"></asp:Button>
                                </div>
                                <div align="center" class="whitecol">
                                    <asp:Label ID="msg2" runat="server" ForeColor="red"></asp:Label></div>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
