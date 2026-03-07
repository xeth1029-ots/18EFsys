<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_005_import.aspx.vb" Inherits="WDAIIP.TC_01_005_import" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>TC_01_005_import</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        function show_reason(myID) {
            var mytable = document.getElementById(myID);
            mytable.style.display = '';//'inline';
        }

        function dis_reason(myID) {
            var mytable = document.getElementById(myID);
            mytable.style.display = 'none';
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <%--<tr><td>
				<table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>
							<asp:Label ID="TitleLab1" runat="server"></asp:Label>
							<asp:Label ID="TitleLab2" runat="server">
                                首頁&gt;&gt;訓練機構管理&gt;&gt;基本資料設定&gt;&gt;課程資料設定
							</asp:Label>
							<font color="#990000">-
                                    <asp:Label ID="lblProecessType" runat="server" Width="24px"></asp:Label></font> (<font color="#ff0000">*</font>為必填欄位)
						</td>
					</tr>
				</table>
			</td>
		</tr>--%>
            <tr>
                <td>
                    <%--<asp:Label ID="lblProecessType" runat="server" Width="24px" Visible="false"></asp:Label>--%>
                    <table class="table_nw" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="bluecol" style="width: 20%">檔案位置(限XLS、ODS)
                            </td>
                            <td class="whitecol">
                                <input id="File2" type="file" name="File1" runat="server"  size="60" accept=".xls,.ods" />
                                <asp:HyperLink ID="Hyperlink2" runat="server" NavigateUrl="../../Doc/CourseInfo_v14b.zip" ForeColor="#8080FF" CssClass="font">下載整批上載格式檔</asp:HyperLink>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="2" class="whitecol">
                                <asp:Button ID="Btn_XlsImport" runat="server" Text="匯入" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="Button2" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="Table4" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <table class="font" id="Table5" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td align="left">&nbsp;&nbsp;&nbsp; <font color="red">移到課程代碼可以看見錯誤訊息</font>
                                        </td>
                                        <td>
                                            <p align="right">
                                                一共有
											<asp:Label ID="Label1" runat="server"></asp:Label>筆資料無法轉入
                                            </p>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" AllowPaging="True" Width="100%" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn DataField="serial" HeaderText="序號">
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="CourseID" HeaderText="課程代碼">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="CourseName" HeaderText="課程名稱">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Hours" HeaderText="小時數">
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Classification1" HeaderText="學/術科">
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Classification2" HeaderText="共同/一般/專業">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="MainCourID" HeaderText="主課程代碼">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Years" HeaderText="計畫年度">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="CLSID" HeaderText="歸屬班別代碼">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="BusID" HeaderText="行業別">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TMID" HeaderText="訓練職類代碼">
                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Valid" HeaderText="是否有效">
                                            <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <p align="center">
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                </p>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
