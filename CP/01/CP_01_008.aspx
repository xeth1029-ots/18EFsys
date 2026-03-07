<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="CP_01_008.aspx.vb" Inherits="WDAIIP.CP_01_008" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訪視計畫表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <%-- <script type="text/javascript"src="../../js/date-picker.js"></script>--%>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        //決定date-picker元件使用的是西元年or民國年，by:20181018
        var bl_rocYear = "<%=ConfigurationManager.AppSettings("REPLACE2ROC_YEARS") %>";
        var scriptTag = document.createElement('script');
        var jsPath = (bl_rocYear == "Y" ? "../../js/date-picker2.js" : "../../js/date-picker.js");
        scriptTag.src = jsPath;
        document.head.appendChild(scriptTag);

        function GETvalue() {
            document.getElementById('Button13').click();
        }

        var cst_VisitorName = 14;
        var cst_ExpectDate = 15;

        function CheckAdd() {
            if (document.form1.OCIDValue1.value == '') {
                alert('請選擇職類班別!')
                return false;
            }
        }

        function CheckSearch() {
            var STDate1 = document.getElementById('STDate1').value;
            var STDate2 = document.getElementById('STDate2').value;
            var FTDate1 = document.getElementById('FTDate1').value;
            var FTDate2 = document.getElementById('FTDate2').value;
            var msg = '';
            if (document.form1.STDate1.value == '') msg += '開訓起始日期不可空白!\n';
            if (document.form1.STDate2.value == '') msg += '開訓結束日期不可空白!\n';
            if (!checkDate(STDate1) && STDate1 != '') msg += '開訓起始日期必須為正確日期格式\n';
            if (!checkDate(STDate2) && STDate2 != '') msg += '開訓結束日期必須為正確日期格式\n';
            if (!checkDate(FTDate1) && FTDate1 != '') msg += '結訓起始日期必須為正確日期格式\n';
            if (!checkDate(FTDate2) && FTDate2 != '') msg += '結訓結束日期必須為正確日期格式\n';
            if (msg != '') {
                alert(msg);
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;訪視計畫表</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tbody>
                <tr>
                    <td align="center">
                        <table class="table_sch" id="Table3">
                            <tr>
                                <td class="bluecol" style="width: 20%">機構</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                    <input id="Button5" type="button" value="..." name="Button5" runat="server">
                                    <input id="RIDValue" type="hidden" name="Hidden1" runat="server"><br>
                                    <asp:Button ID="Button13" Style="display: none" runat="server" Text="Button13"></asp:Button>
                                    <span id="HistoryList2" style="display: none" onclick="GETvalue()">
                                        <asp:Table ID="HistoryRID" runat="server" Width="50%"></asp:Table>
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">職類/班別 </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                    <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                    <input id="Button6" onclick="javascript: openClass('../CP_01_ch.aspx?RID=' + document.form1.RIDValue.value);" type="button" value="..." name="Button6" runat="server">
                                    <input id="TMIDValue1" type="hidden" name="Hidden2" runat="server"><input id="OCIDValue1" type="hidden" name="Hidden1" runat="server"><br>
                                    <span id="HistoryList" style="display: none; left: 270px; position: absolute">
                                        <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need"><font>&nbsp;開訓期間</font> </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="STDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">～
								    <asp:TextBox ID="STDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol"><font>結訓期間</font> </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="FTDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">～
								    <asp:TextBox ID="FTDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2" class="whitecol">
                                    <p align="center">
                                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue" DESIGNTIMEDRAGDROP="30">顯示列數</asp:Label>
                                        <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                        <asp:Button ID="Query" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                    </p>
                                </td>
                            </tr>
                        </table>
                        </font>
					    <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
                    </td>
                </tr>
            </tbody>
        </table>
        <table id="DataGridTable" cellspacing="1" cellpadding="1" border="0" runat="server">
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                        <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                        <HeaderStyle CssClass="head_navy" HorizontalAlign="Center"></HeaderStyle>
                        <Columns>
                            <asp:TemplateColumn HeaderText="序號">
                                <HeaderStyle></HeaderStyle>
                                <ItemStyle CssClass="whitecol" HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <asp:Label ID="labSEQNO" runat="server"></asp:Label>
                                    <input id="hid_OCID" type="hidden" runat="server" /><input id="hid_ceSEQNO" type="hidden" runat="server" />
                                </ItemTemplate>
                            </asp:TemplateColumn>

                     <%--<asp:BoundColumn HeaderText="序號"><HeaderStyle /><ItemStyle HorizontalAlign="Center"></ItemStyle></asp:BoundColumn>--%>
                            <asp:BoundColumn DataField="DistName" HeaderText="分署名稱"></asp:BoundColumn>
                            <asp:BoundColumn DataField="OrgName" HeaderText="訓練單位名稱"></asp:BoundColumn>
                            <asp:BoundColumn DataField="OrgTypeName" HeaderText="單位類型"></asp:BoundColumn>
                            <asp:BoundColumn DataField="ClassCName" HeaderText="課稱名稱"></asp:BoundColumn>
                            <asp:BoundColumn DataField="IsBusiness" HeaderText="是否為企業包班"></asp:BoundColumn>
                            <asp:BoundColumn DataField="PointYN" HeaderText="是否為學分班"></asp:BoundColumn>
                             <asp:TemplateColumn HeaderText="開訓日期">
                                <HeaderStyle></HeaderStyle>
                                <ItemStyle CssClass="whitecol" HorizontalAlign="Center" />
                                <ItemTemplate> <asp:Label ID="labSTDate" runat="server"></asp:Label> </ItemTemplate>
                            </asp:TemplateColumn>
                             <asp:TemplateColumn HeaderText="結訓日期">
                                <HeaderStyle></HeaderStyle>
                                <ItemStyle CssClass="whitecol" HorizontalAlign="Center" />
                                <ItemTemplate> <asp:Label ID="labFTDate" runat="server"></asp:Label> </ItemTemplate>
                            </asp:TemplateColumn>

                            <%--<asp:BoundColumn DataField="STDate" HeaderText="開訓日期" DataFormatString="{0:d}"><HeaderStyle></HeaderStyle><ItemStyle HorizontalAlign="Center"></ItemStyle></asp:BoundColumn>
                                <asp:BoundColumn DataField="FTDate" HeaderText="結訓日期" DataFormatString="{0:d}"></asp:BoundColumn>--%>

                            <asp:BoundColumn DataField="WEEKS" HeaderText="上課時間"></asp:BoundColumn>
                            <asp:BoundColumn DataField="Address" HeaderText="上課地點"></asp:BoundColumn>
                            <asp:BoundColumn DataField="SeqNo" HeaderText="已抽訪次數"></asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="檢閱抽訪紀錄">
                                <ItemTemplate>
                                    <asp:Label ID="ViewRecord" runat="server" ForeColor="#8080FF" ToolTip="點選可查看抽訪紀錄">Label</asp:Label>
                                    <asp:LinkButton ID="vRecord" runat="server" ForeColor="Blue"></asp:LinkButton>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn DataField="UnexpectTimes" HeaderText="異常次數"></asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="預計抽訪人員">
                                <HeaderStyle></HeaderStyle>
                                <ItemStyle CssClass="whitecol" />
                                <ItemTemplate>
                                    &nbsp;<asp:TextBox ID="VisitorName" runat="server" MaxLength="8"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="預計抽訪日期">
                                <HeaderStyle></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center" CssClass="whitecol"></ItemStyle>
                                <ItemTemplate>
                                    <asp:TextBox ID="ExpectDate" runat="server" onfocus="this.blur()" Columns="10"></asp:TextBox>
                                    <img id="Img1" style="cursor: pointer" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30">
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="新增抽訪紀錄">
                                <ItemStyle HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <asp:Button ID="AddUnexpectVisitor" runat="server" Text="實地抽訪" CommandName="AddUV" class="asp_button_M"></asp:Button>
                                    <asp:Button ID="AddUnExpectTel" runat="server" Text="電話抽訪" CommandName="AddUT" class="asp_button_M"></asp:Button>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="列印空白抽訪紀錄表">
                                <ItemStyle HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <asp:Button ID="PrintUnexpectVisitor" runat="server" Text="實地抽訪" class="asp_button_M"></asp:Button>
                                    <asp:Button ID="PrintUnExpectTel" runat="server" Text="電話抽訪" class="asp_button_M"></asp:Button>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <%--<asp:BoundColumn Visible="False" DataField="OCID" HeaderText=" OCID"></asp:BoundColumn>
                            <asp:BoundColumn Visible="False" DataField="ceSeqNo" HeaderText="ceSeqNo"></asp:BoundColumn>--%>
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
            <tr>
                <td align="center" class="whitecol">
                    <asp:Button ID="Save" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                    <asp:Button ID="Print" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
