<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_032.aspx.vb" Inherits="WDAIIP.SD_05_032" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <%--<title>屆退官兵荐訓名冊維護</title>--%>
    <title>送訓官兵名冊</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        //if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;送訓官兵名冊</asp:Label>
                </td>
            </tr>
        </table>
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr>
                <td>

                    <table id="tbSch" runat="server" width="100%" cellspacing="1" cellpadding="1">
                        <tr>
                            <td>
                                <table width="100%" class="table_nw" cellpadding="1" cellspacing="1">
                                    <tr>
                                        <td class="bluecol" width="20%">姓名</td>
                                        <td class="whitecol" width="80%">
                                            <asp:TextBox ID="sCNAME" runat="server" MaxLength="22" Width="20%"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">身分證號碼</td>
                                        <td class="whitecol" width="80%">
                                            <asp:TextBox ID="sIDNO" runat="server" MaxLength="12" Width="20%"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">出生年月日</td>
                                        <td class="whitecol" width="80%">
                                            <asp:TextBox ID="sBIRTHDAY1" runat="server" Columns="10" MaxLength="10" Width="15%"></asp:TextBox>
                                            <span id="span1" runat="server">
                                                <img style="cursor: pointer" onclick="javascript:show_calendar('<%= sBIRTHDAY1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span> &nbsp;~
                                            <asp:TextBox ID="sBIRTHDAY2" runat="server" Columns="10" MaxLength="10" Width="15%"></asp:TextBox>
                                            <span id="span2" runat="server">
                                                <img style="cursor: pointer" onclick="javascript:show_calendar('<%= sBIRTHDAY2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                            <%--<asp:label id="Note" runat="server" ForeColor="Red">搜尋條件【身分證號碼】與【姓名】，請擇一輸入</asp:label>--%>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">預定退伍日</td>
                                        <td class="whitecol" width="80%">
                                            <asp:TextBox ID="sPREEXDATE1" runat="server" Columns="10" MaxLength="10" Width="15%"></asp:TextBox>
                                            <span id="span3" runat="server">
                                                <img style="cursor: pointer" onclick="javascript:show_calendar('<%= sPREEXDATE1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span> &nbsp;~
                                            <asp:TextBox ID="sPREEXDATE2" runat="server" Columns="10" MaxLength="10" Width="15%"></asp:TextBox>
                                            <span id="span4" runat="server">
                                                <img style="cursor: pointer" onclick="javascript:show_calendar('<%= sPREEXDATE2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">送訓至分署</td>
                                        <td class="whitecol" width="80%">
                                            <asp:DropDownList ID="sRECOMMDISTID" runat="server"></asp:DropDownList></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">匯入日期</td>
                                        <td class="whitecol" width="80%">
                                            <asp:TextBox ID="sCREATEDATE1" runat="server" Columns="10" MaxLength="10" Width="15%"></asp:TextBox>
                                            <span id="span7" runat="server">
                                                <img style="cursor: pointer" onclick="javascript:show_calendar('<%= sCREATEDATE1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span> &nbsp;~
                                            <asp:TextBox ID="sCREATEDATE2" runat="server" Columns="10" MaxLength="10" Width="15%"></asp:TextBox>
                                            <span id="span8" runat="server">
                                                <img style="cursor: pointer" onclick="javascript:show_calendar('<%= sCREATEDATE2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                        </td>
                                    </tr>
                                    <tr id="trfile1" runat="server">
                                        <td class="bluecol" width="20%">匯入名冊</td>
                                        <td class="whitecol" width="80%">
                                            <input id="File1" type="file" name="File1" runat="server" size="50" accept=".xls,.ods" />
                                            <asp:Button ID="btnImport" runat="server" Text="名冊匯入" CssClass="asp_button_M"></asp:Button>(必須為ods或xls格式)
                                            <br />
                                            <asp:HyperLink ID="HyperLink1" runat="server" CssClass="font" ForeColor="#8080FF">下載整批上載格式檔</asp:HyperLink>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">匯出檔案格式</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                                                <asp:ListItem Value="ODS">ODS</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <%--<tr>
                                        <td class="bluecol" width="20%">停止結束日期</td>
                                        <td class="whitecol" width="80%"><asp:TextBox ID="StopEDate1" runat="server" Columns="10" MaxLength="10" Height="19px"></asp:TextBox>&nbsp;<img style="cursor: pointer" onclick="javascript:show_calendar('<%= StopEDate1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">&nbsp; ~<asp:TextBox ID="StopEDate2" runat="server" Columns="10" MaxLength="10" Height="19px"></asp:TextBox>&nbsp;<img style="cursor: pointer" onclick="javascript:show_calendar('<%= StopEDate2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">&nbsp;</td>
                                    </tr>--%>
                                    <tr>
                                        <td align="center" colspan="2" class="whitecol">
                                            <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue" DESIGNTIMEDRAGDROP="30">顯示列數</asp:Label>
                                            <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                            <asp:Button ID="btnSearch1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                            &nbsp;<asp:Button ID="btnAdd1" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                                            &nbsp;<asp:Button ID="btnExport" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <div align="center">
                                    <asp:Label ID="labmsg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                                </div>
                                <table id="tbList" runat="server" cellspacing="0" bordercolordark="#ffffff" cellpadding="0" width="100%" align="left" bordercolorlight="#666666" border="0">
                                    <tr>
                                        <td>
                                            <div id="Div2" runat="server">
                                                <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" Visible="false" CellPadding="8">
                                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <Columns>
                                                        <asp:BoundColumn DataField="CNAME" HeaderText="姓名">
                                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="IDNO" HeaderText="身分證字號">
                                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:TemplateColumn HeaderText="出生年月日">
                                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="BIRTHDAY" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.BIRTHDAY") %>'>
                                                                </asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="任職單位全銜">
                                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="POSITION" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.POSITION") %>'>
                                                                </asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="預定退伍日">
                                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="PREEXDATE" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.PREEXDATE") %>'>
                                                                </asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:BoundColumn DataField="RCDISTNAME" HeaderText="送訓至分署">
                                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:TemplateColumn HeaderText="匯入日期">
                                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="CREATEDTE" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.CREATEDTE") %>'>
                                                                </asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                    </Columns>
                                                    <PagerStyle Visible="False"></PagerStyle>
                                                </asp:DataGrid>
                                            </div>
                                            <div align="center">
                                                <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Width="100%" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <Columns>
                                                        <asp:BoundColumn HeaderText="序號">
                                                            <HeaderStyle HorizontalAlign="Center" Width="6%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="CNAME" HeaderText="姓名">
                                                            <HeaderStyle HorizontalAlign="Center" Width="14%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="IDNO" HeaderText="身分證字號">
                                                            <HeaderStyle HorizontalAlign="Center" Width="16%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="RCDISTNAME" HeaderText="送訓至分署">
                                                            <HeaderStyle HorizontalAlign="Center" Width="28%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:TemplateColumn HeaderText="出生年月日">
                                                            <HeaderStyle HorizontalAlign="Center" Width="14%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="BIRTHDAY" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.BIRTHDAY") %>'>
                                                                </asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="預定退伍日">
                                                            <HeaderStyle HorizontalAlign="Center" Width="14%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="PREEXDATE" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.PREEXDATE") %>'>
                                                                </asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <%--<asp:TemplateColumn HeaderText="姓名">
                                                            <ItemTemplate>
                                                                <asp:Label ID="CNAME" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.CNAME") %>'></asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn> 
                                                        <asp:BoundColumn DataField="PostDate" HeaderText="發佈日期">
                                                            <HeaderStyle HorizontalAlign="Center" Width="100px"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>--%>
                                                        <asp:TemplateColumn HeaderText="功能" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="8%">
                                                            <ItemStyle HorizontalAlign="Center" Font-Size="Small" />
                                                            <ItemTemplate>
                                                                <asp:LinkButton ID="lbtUpdate" Text="修改" runat="server" CssClass="linkbutton" CommandName="UPD"></asp:LinkButton>&nbsp;&nbsp;
															    <%--<asp:LinkButton ID="lbtDelete" Text="刪除" runat="server" CssClass="linkbutton" CommandName="DEL"></asp:LinkButton>--%>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                    </Columns>
                                                    <PagerStyle Visible="False"></PagerStyle>
                                                </asp:DataGrid>
                                                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                            </div>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table id="tbEdit" runat="server" class="table_nw" width="100%" border="0" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol_need" width="20%">姓名</td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="tCNAME" runat="server" MaxLength="10" Width="20%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="20%">身分證號碼</td>
                            <td class="whitecol">
                                <asp:TextBox ID="tIDNO" runat="server" MaxLength="10" Width="20%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="20%">出生年月日</td>
                            <td class="whitecol">
                                <asp:TextBox ID="tBIRTHDAY" runat="server" Columns="10" MaxLength="10" Width="15%"></asp:TextBox>
                                <span id="span5" runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= tBIRTHDAY.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="20%">任職單位(全銜)</td>
                            <td class="whitecol">
                                <asp:TextBox ID="tPOSITION" runat="server" MaxLength="100" Width="80%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="20%">預定退伍日</td>
                            <td class="whitecol">
                                <asp:TextBox ID="tPREEXDATE" runat="server" Columns="10" MaxLength="10" Width="15%"></asp:TextBox>
                                <span id="span6" runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= tPREEXDATE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="20%">送訓至分署</td>
                            <td class="whitecol" width="80%">
                                <asp:DropDownList ID="ddlRECOMMDISTID" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td colspan="2" class="whitecol" align="center">
                                <asp:Button ID="btnSave1" Text="儲存" runat="server" CssClass="asp_Export_M"></asp:Button>&nbsp;
                                <asp:Button ID="btnBack1" runat="server" Text="回上頁" CssClass="asp_button_M"></asp:Button>&nbsp;
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="Hid_ARSID" runat="server" />
    </form>
</body>
</html>
