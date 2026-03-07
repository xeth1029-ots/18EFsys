<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_11_004.aspx.vb" Inherits="WDAIIP.SD_11_004" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>受訓學員意見調查表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function GETvalue() {
            var Button3 = document.getElementById('Button3');
            Button3.click();
        }

        function SetOneOCID() {
            var Button7 = document.getElementById('Button7');
            Button7.click();
        }

        function choose_class() {
            var OCID1 = document.getElementById('OCID1');
            var RIDValue = document.getElementById('RIDValue');
            var Button7 = document.getElementById('Button7');
            if (OCID1.value == '') {
                Button7.click();
            }
            openClass('../02/SD_02_ch.aspx?RID=' + RIDValue.value);
        }

        function closeDiv() {
            var eMeng = document.getElementById('eMeng');
            eMeng.style.visibility = 'hidden';
        }
    </script>
    <style type="text/css">
        .auto-style1 { width: 109%; }
        .auto-style2 { height: 100%; width: 109%; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <%--<input id="check_search" type="hidden" name="check_search" runat="server">
	    <input id="check_add" type="hidden" name="check_add" runat="server">
	    <input id="check_mod" type="hidden" name="check_mod" runat="server">
	    <input id="check_del" type="hidden" name="check_del" runat="server">--%>
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;受訓學員意見調查表</asp:Label>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <table class="table_nw" width="100%">
            <tr>
                <td class="bluecol" width="20%">訓練機構</td>
                <td class="whitecol">
                    <asp:TextBox ID="center" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox>
                    <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                    <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                        <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                    </span>
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">職類/班別</td>
                <td class="whitecol">
                    <div>
                        <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                        <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                        <input onclick="choose_class()" type="button" value="..." class="button_b_Mini">
                        <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                        <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                        <span id="HistoryList" style="position: absolute; display: none; left: 28%">
                            <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                        </span>
                        <asp:Button ID="Button3" Style="display: none" runat="server" Text="Button"></asp:Button>
                        <asp:Button ID="Button7" Style="display: none" runat="server"></asp:Button><br />
                        <asp:Button ID="PrintBlank" runat="server" Text="列印空白表單(產業人才)" CssClass="asp_button_L"></asp:Button>&nbsp;
					    <asp:Button ID="PrintBlank2" runat="server" Text="列印空白表單(在職勞工)" CssClass="asp_button_L"></asp:Button>
                    </div>
                </td>
            </tr>
            <%--
            <tr>
			    <td class="bluecol">匯入調查資料</td>
			    <td class="whitecol">
				    <div>
					    <input id="File1" type="file" name="File1" runat="server" size="36" accept=".csv,.xls" />
					    <asp:Button ID="Button13" runat="server" Text="匯入調查表"></asp:Button>(必須為csv格式)
					    <asp:HyperLink ID="HyperLink1" runat="server" CssClass="font" NavigateUrl="../../Doc/Stud_QuestionFac.zip" ForeColor="#8080FF">下載整批上載格式檔</asp:HyperLink>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					    <asp:HyperLink ID="Hyperlink2" runat="server" CssClass="font" NavigateUrl="../../Doc/Stud_QuestionFac08.zip" ForeColor="#8080FF">下載整批上載格式檔</asp:HyperLink>
                    </div>
			    </td>
		    </tr>
            --%>
            <tr>
                <td align="center" colspan="4" class="whitecol">
                    <asp:Button ID="search" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button></td>
            </tr>
        </table>
        <br>

        <asp:Panel ID="PanelDataGrid1" runat="server" Width="100%">
            <div align="center">
                <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
            </div>
            <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                <AlternatingItemStyle BackColor="WhiteSmoke"></AlternatingItemStyle>
                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                <Columns>
                    <asp:BoundColumn HeaderText="序號">
                        <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ClassCName2" HeaderText="班別">
                        <HeaderStyle HorizontalAlign="Center" Width="30%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="total" HeaderText="結訓人數">
                        <HeaderStyle HorizontalAlign="Center" Width="25%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="num1" HeaderText="填寫人數">
                        <HeaderStyle HorizontalAlign="Center" Width="25%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:TemplateColumn HeaderText="功能">
                        <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        <ItemTemplate>
                            <asp:Button ID="Button1" runat="server" Text="查詢" CommandName="View" CssClass="asp_button_M"></asp:Button>
                            <asp:Button ID="PrintBlank3" runat="server" Text="列印空白調查表" CommandName="Print" CssClass="asp_Export_M"></asp:Button>
                        </ItemTemplate>
                    </asp:TemplateColumn>
                    <%--<asp:BoundColumn Visible="False" DataField="FTDate" HeaderText="結訓日期"></asp:BoundColumn>
				<asp:BoundColumn Visible="False" DataField="CyclType" HeaderText="CyclType"></asp:BoundColumn>
				<asp:BoundColumn Visible="False" DataField="LevelType" HeaderText="LevelType"></asp:BoundColumn>--%>
                </Columns>
            </asp:DataGrid>
        </asp:Panel>
        <asp:Panel ID="PanelDG_stud" runat="server" Width="100%">
            <div>
                <asp:Label ID="Label1" runat="server" CssClass="font"></asp:Label>
            </div>
            <div style="margin-top: 3px; margin-bottom: 3px" align="center">
                <asp:Label ID="msg2" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
            </div>
            <table id="StudentTable" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                <tr>
                    <td>
                        <asp:DataGrid ID="DG_stud" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                            <AlternatingItemStyle BackColor="WhiteSmoke"></AlternatingItemStyle>
                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                            <Columns>
                                <asp:BoundColumn DataField="STUDID2" HeaderText="學號">
                                    <HeaderStyle HorizontalAlign="Center" Width="25%"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="Name" HeaderText="姓名(離退訓日期)">
                                    <HeaderStyle HorizontalAlign="Center" Width="25%"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn HeaderText="填寫狀態">
                                    <HeaderStyle HorizontalAlign="Center" Width="25%"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:TemplateColumn HeaderText="功能">
                                    <HeaderStyle HorizontalAlign="Center" Width="25%"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    <ItemTemplate>
                                        <asp:Button ID="Button4" runat="server" Text="新增" CommandName="Insert" CssClass="asp_button_M"></asp:Button>
                                        <asp:Button ID="Edit" runat="server" Text="修改" CommandName="Edit" CssClass="asp_button_M"></asp:Button>
                                        <asp:Button ID="Button5" runat="server" Text="查詢" CommandName="Check" CssClass="asp_button_M"></asp:Button>
                                        <asp:Button ID="Print" runat="server" Text="列印" CommandName="print" ToolTip="填寫狀態為「是」，才可列印" CssClass="asp_Export_M"></asp:Button>
                                        <asp:Button ID="Button6" runat="server" Text="清除重填" CommandName="Clear" CssClass="asp_button_M"></asp:Button>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:BoundColumn Visible="False" DataField="OCID" HeaderText="OCID"></asp:BoundColumn>
                                <asp:BoundColumn Visible="False" DataField="StudentID" HeaderText="StudentID"></asp:BoundColumn>
                                <asp:BoundColumn Visible="False" DataField="SOCID" HeaderText="SOCID"></asp:BoundColumn>
                            </Columns>
                        </asp:DataGrid>
                    </td>
                </tr>
                <tr>
                    <td align="center"></td>
                </tr>
            </table>
        </asp:Panel>

        <table class="font" id="eMeng" style="z-index: 99999; border-bottom: #455690 1px solid; position: absolute; border-left: #a6b4cf 1px solid; background-color: #c9d3f3; width: 50%; height: 248px; visibility: visible; border-top: #a6b4cf 1px solid; top: 0px; border-right: #455690 1px solid; left: 0px" cellspacing="1" cellpadding="1" border="0" runat="server">
            <tr>
                <td background="../../images/MSNTitle.gif" class="auto-style1"></td>
            </tr>
            <tr>
                <td style="border-bottom: #b9c9ef 1px solid; border-left: #728eb8 1px solid; padding-bottom: 10px; padding-left: 10px; padding-right: 10px; color: #1f336b; font-size: 12px; border-top: #728eb8 1px solid; border-right: #b9c9ef 1px solid; padding-top: 15px" align="center" background="../../images/MsnBack.gif" height="100" class="auto-style2">
                    <table class="font" id="Table4" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td><strong><font color="#0000ff">問題轉入資料訊息：</font></strong></td>
                            <td style="cursor: pointer" onclick="closeDiv();" align="center" width="15">
                                <img src="../../images/CloseMsn.gif" width="13" height="13" alt=""></td>
                        </tr>
                    </table>
                    <asp:DataGrid ID="Datagrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                        <AlternatingItemStyle BackColor="WhiteSmoke"></AlternatingItemStyle>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn DataField="Index" HeaderText="第幾筆">
                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="FillFormDate" HeaderText="填寫日期">
                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="StudID" HeaderText="學號">
                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Status" HeaderText="轉入狀態">
                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Reason" HeaderText="原因">
                                <HeaderStyle HorizontalAlign="Center" Width="60%"></HeaderStyle>
                            </asp:BoundColumn>
                        </Columns>
                    </asp:DataGrid>
                </td>
            </tr>
        </table>
        <%--
	    <table class="font" id="eMeng" runat="server"></table>
	    <asp:DataGrid ID="Datagrid2" runat="server" Width="100%" CssClass="font" Height="208px" AutoGenerateColumns="False"></asp:DataGrid>
        --%>
        <input id="Years" type="hidden" name="Years" runat="server">
        <asp:Label ID="VeMeng" runat="server" Visible="False">none</asp:Label>
    </form>
</body>
</html>
