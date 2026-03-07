<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="OB_01_009.aspx.vb" Inherits="WDAIIP.OB_01_009" %>

 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>OB_01_009</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/common.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script language="JavaScript">
        function chkdata() {
            var msg = '';

            var mytable = document.getElementById('DataGrid2');
            for (var i = 1; i < mytable.rows.length; i++) {
                for (var j = 3; j < mytable.rows(i).cells.length; j++) {
                    var txt = mytable.rows(i).cells(j).children(0);
                    if (txt) {
                        if (!isPositiveInt(txt.value) || parseInt(txt.value) > 100) {
                            msg = "請輸入介於0~100的評分！";
                            j = 999;
                        }
                        //alert(txt.value);
                    }
                }
                if (msg != "") {
                    i = 999;
                }
            }
            if (msg != "") {
                alert(msg);
                return false;
            }
        }

        function sum(num) {
            var mytable = document.getElementById('DataGrid2');
            var total = mytable.rows(num).cells(2).children(0);
            //var total =mytable.rows(num).cells(2).childNodes[0].outerText;
            //var total =mytable.rows(num).cells(2).outerText;
            total.value = 0;

            for (var j = 3; j < mytable.rows(num).cells.length; j++) {
                //debugger;
                var txt = mytable.rows(num).cells(j).children(0);
                if (txt) {
                    //total.value=parseInt(total.value)+parseInt(txt.value);
                    //mytable.rows(num).cells(2).childNodes[0].outerText=parseInt(total)+parseInt(txt.value);
                    //debugger;
                    if (!isPositiveInt(txt.value) || parseInt(txt.value) > 100) {
                        total.value = parseInt(total.value) + 0;
                        //mytable.rows(num).cells(2).innerHTML ='<SPAN id=DataGrid2__ctl5_labScore>'+(parseInt(total)+0)+'"';
                    }
                    else {
                        total.value = parseInt(total.value) + parseInt(txt.value);
                        //mytable.rows(num).cells(2).innerHTML ='"'+(parseInt(total)+parseInt(txt.value))+'"';
                    }
                }
            }
        }
			
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table class="font" cellspacing="1" cellpadding="1" width="740" border="0">
        <tr>
            <td>
                <asp:Label ID="lblTitle1" runat="server"></asp:Label><asp:Label ID="lblTitle2" runat="server">
							<FONT face="新細明體">首頁&gt;&gt;委外訓練管理&gt;&gt;<font color="#990000">投標單位評分資料維護</font></FONT>
                </asp:Label><font color="#000000">(<font face="新細明體"><font color="#ff0000">*</font>為必填欄位</font>)</font>
            </td>
        </tr>
    </table>
    <asp:Panel ID="panelSch" runat="server">
        <table class="font" border="0" cellspacing="1" cellpadding="1" width="740">
            <tr>
                <td>
                    <table class="table_sch" cellspacing="1" cellpadding="1">
                        <tr>
                            <td width="15%" class="bluecol">
                                年度
                            </td>
                            <td style="height: 19px" width="25%" class="whitecol">
                                <asp:DropDownList ID="ddlyears" runat="server">
                                </asp:DropDownList>
                            </td>
                            <td width="15%" class="bluecol">
                                序號
                            </td>
                            <td style="height: 19px" width="45%" class="whitecol">
                                <asp:TextBox ID="txttsn" runat="server"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td width="15%" class="bluecol">
                                訓練計畫名稱
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:DropDownList ID="TPlanID" runat="server">
                                </asp:DropDownList>
                                <asp:TextBox ID="PlanName" runat="server" MaxLength="20"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td width="15%" class="bluecol">
                                標案名稱
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="TenderName" runat="server" MaxLength="20" Width="400px"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td width="15%" class="bluecol">
                                主辦單位
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="Sponsor" runat="server" MaxLength="20" Width="400px"></asp:TextBox>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <p align="center">
                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" MaxLength="2" Width="23px">10</asp:TextBox>
                        <asp:Button ID="btnQuery" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button></p>
                </td>
            </tr>
            <tr>
                <td>
                    <table border="0" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td align="center">
                                <p>
                                    <table id="DataGridTable" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                        <tr>
                                            <td>
                                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" PagerStyle-Mode="NumericPages" PagerStyle-HorizontalAlign="Left" AllowSorting="True">
                                                    <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <Columns>
                                                        <asp:BoundColumn HeaderText="序號">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="tsn" HeaderText="委外序號">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="PlanName" HeaderText="訓練計畫名稱">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="TenderName" HeaderText="標案名稱">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="Sponsor" HeaderText="主辦單位">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="TenderSDate" HeaderText="投標日期">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:TemplateColumn HeaderText="功能">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Button ID="btnSet" runat="server" Text="評分" CommandName="set"></asp:Button>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
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
                                </p>
                                <p>
                                    <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label></p>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </asp:Panel>
    <asp:Panel ID="panelSet" runat="server">
        <table class="font" border="0" cellspacing="1" cellpadding="1" width="740">
            <tr>
                <td>
                    <table class="table_sch" cellspacing="1" cellpadding="1">
                        <tr>
                            <td width="15%" class="bluecol">
                                評審委員人數
                            </td>
                            <td style="height: 19px" class="whitecol">
                                <asp:TextBox ID="txtJudgeNum" runat="server" MaxLength="2" Width="24px"></asp:TextBox>
                                <asp:Button ID="btnSend" runat="server" Text="產生"></asp:Button>評審委員人數介於1~10人
                                <input id="hidTsn" type="hidden" name="hidTsn" runat="server">
                            </td>
                        </tr>
                    </table>
                    <table id="DataGridTable2" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" PagerStyle-Mode="NumericPages" PagerStyle-HorizontalAlign="Left" AllowSorting="True">
                                    <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號">
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="投標單位">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="總分">
                                            <HeaderStyle HorizontalAlign="Center" Width="9%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtScore" runat="server" MaxLength="3" Width="26px" onfocus="this.blur()" BorderStyle="None"></asp:TextBox>
                                                <asp:Label ID="lblTCsn" runat="server" Visible="False"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="評分1">
                                            <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtScore1" runat="server" MaxLength="3" Width="26px"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn Visible="False" HeaderText="評分2">
                                            <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtScore2" runat="server" MaxLength="3" Width="26px"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn Visible="False" HeaderText="評分3">
                                            <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtScore3" runat="server" MaxLength="3" Width="26px"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn Visible="False" HeaderText="評分4">
                                            <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtScore4" runat="server" MaxLength="3" Width="26px"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn Visible="False" HeaderText="評分5">
                                            <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtScore5" runat="server" MaxLength="3" Width="26px"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn Visible="False" HeaderText="評分6">
                                            <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtScore6" runat="server" MaxLength="3" Width="26px"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn Visible="False" HeaderText="評分7">
                                            <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtScore7" runat="server" MaxLength="3" Width="26px"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn Visible="False" HeaderText="評分8">
                                            <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtScore8" runat="server" MaxLength="3" Width="26px"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn Visible="False" HeaderText="評分9">
                                            <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtScore9" runat="server" MaxLength="3" Width="26px"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn Visible="False" HeaderText="評分10">
                                            <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtScore10" runat="server" MaxLength="3" Width="26px"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False" HorizontalAlign="Left" Mode="NumericPages"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <p align="center">
                                    <asp:Button ID="btnSave" runat="server" Text="儲存" CssClass="asp_button_S"></asp:Button>
                                    <asp:Button ID="btnExit" runat="server" Text="離開" CssClass="asp_button_S"></asp:Button></p>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </asp:Panel>
    <%--
			<asp:textbox id="txtScore" Runat="server" MaxLength="3" Width="26px" onfocus="this.blur()"></asp:textbox>
    --%>
    </form>
</body>
</html>
