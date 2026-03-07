<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_16_002.aspx.vb" Inherits="WDAIIP.SD_16_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>補登申請</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
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
        function CheckSave1() {
            //debugger;
            var strMsg = "";
            var ddlYears = document.getElementById("ddlYears");
            var ddlDistID = document.getElementById("ddlDistID");
            var ddlPlan = document.getElementById("ddlPlan");
            var ddlOrgName = document.getElementById("ddlOrgName");
            var ddlClassCName = document.getElementById("ddlClassCName");
            var ddlAccount = document.getElementById("ddlAccount");
            var ddlReasonID = document.getElementById("ddlReasonID");
            var txtReason = document.getElementById("txtReason");
            var cb_SelFunID = document.getElementById("cb_SelFunID");
            var EndDate = document.getElementById("EndDate");
            //var rblBlameUnit = document.getElementById("rblBlameUnit");

            strMsg += chkValue('select', '年度', ddlYears);
            strMsg += chkValue('select', '轄區', ddlDistID);
            strMsg += chkValue('select', '訓練計畫', ddlPlan);
            strMsg += chkValue('select', '訓練機構', ddlOrgName);
            strMsg += chkValue('select', '班級名稱', ddlClassCName);
            strMsg += chkValue('select', '承辦人員', ddlAccount);
            strMsg += chkValue('select', '補登資料原因', ddlReasonID);
            strMsg += chkValue('empty', '補登原因及檢討改善作法', txtReason);
            strMsg += chkValue('checkboxlist_must', '開放功能', cb_SelFunID);
            strMsg += chkValue('date_must', '結束日期', EndDate);
            //strMsg += chkValue('radio_must', '是否歸責單位', rblBlameUnit);

            if (strMsg != "") {
                alert(strMsg);
                return false;
            }
            else {
                strMsg = '';
                strMsg += '\n請確認資料是否無誤,儲存後資料將不可修改\n';
                strMsg += '\n如確認資料無誤後,請按下確定,謝謝!!\n';
                return confirm(strMsg);
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td align="center">
                    <table class="font" width="100%">
                        <tr>
                            <td class="font">
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">
				                    首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;補登申請
                                </asp:Label>
                            </td>
                        </tr>
                    </table>
                    <asp:Panel ID="Panelsch1" runat="server" Visible="True">
                        <table class="font" width="100%">
                            <tr>
                                <td align="center">
                                    <table class="table_sch" cellspacing="1" cellpadding="1" width="100%">
                                        <tr>
                                            <td class="bluecol_need" style="width: 20%;">年度： </td>
                                            <td class="whitecol" style="width: 30%;">
                                                <asp:DropDownList ID="sddlYears" runat="server" AutoPostBack="true">
                                                </asp:DropDownList>
                                            </td>
                                            <td class="bluecol_need" style="width: 20%;">轄區： </td>
                                            <td class="whitecol" style="width: 30%;">
                                                <asp:DropDownList ID="sddlDistID" runat="server" AutoPostBack="true">
                                                </asp:DropDownList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol_need">訓練計畫： </td>
                                            <td class="whitecol" colspan="3">
                                                <asp:DropDownList ID="sddlPlan" runat="server" AutoPostBack="true">
                                                </asp:DropDownList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol_need">訓練機構： </td>
                                            <td class="whitecol" colspan="3">
                                                <asp:DropDownList ID="sddlOrgName" runat="server">
                                                </asp:DropDownList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">班級名稱： </td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="sClassName" runat="server" Columns="40" MaxLength="40" Width="40%"></asp:TextBox>
                                            </td>
                                            <td class="bluecol">期別： </td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="sCyclType" runat="server" Columns="5" MaxLength="3" Width="30%"></asp:TextBox>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">申請日期區間 </td>
                                            <td class="whitecol" colspan="3">
                                                <asp:TextBox ID="sAPPLYDATE1" runat="server" MaxLength="10" Columns="10" Width="15%"></asp:TextBox>
                                                <img style="cursor: pointer" onclick="javascript:show_calendar('sAPPLYDATE1','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">～
											<asp:TextBox ID="sAPPLYDATE2" runat="server" MaxLength="10" Columns="10" Width="15%"></asp:TextBox>
                                                <img style="cursor: pointer" onclick="javascript:show_calendar('sAPPLYDATE2','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">結束日期區間 </td>
                                            <td class="whitecol" colspan="3">
                                                <asp:TextBox ID="sENDDATE1" runat="server" MaxLength="10" Columns="10" Width="15%"></asp:TextBox>
                                                <img style="cursor: pointer" onclick="javascript:show_calendar('sENDDATE1','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">～
											<asp:TextBox ID="sENDDATE2" runat="server" MaxLength="10" Columns="10" Width="15%"></asp:TextBox>
                                                <img style="cursor: pointer" onclick="javascript:show_calendar('sENDDATE2','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">
                                            </td>
                                        </tr>
                                        <tr>
                                            <td align="center" class="whitecol" colspan="4">
                                                <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                                <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                                <asp:Button ID="btnSch" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                                <asp:Button ID="btnAdd" runat="server" Text="申請" CssClass="asp_button_M"></asp:Button>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td align="center">
                                    <asp:Label ID="labmsg" runat="server" ForeColor="Red"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td align="center">
                                    <table class="font" id="tbSearch1" style="width: 100%" cellspacing="1" cellpadding="1" width="900" align="center" border="0" runat="server">
                                        <tr>
                                            <td align="center">
                                                <asp:DataGrid ID="DG_ClassInfo" runat="server" Width="100%" CssClass="font" AllowPaging="true" AutoGenerateColumns="False" PageSize="15" CellPadding="8">
                                                    <AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
                                                    <HeaderStyle CssClass="head_navy" />
                                                    <Columns>
                                                        <asp:TemplateColumn HeaderText="序號">
                                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <%--<input id="checkbox1" type="checkbox"  runat="server" />--%>
                                                                <asp:Label ID="labseqno" runat="server"></asp:Label>
                                                                <asp:HiddenField ID="hRETID" runat="server" />
                                                                <asp:HiddenField ID="hOCID" runat="server" />
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:BoundColumn DataField="ORGNAME" HeaderText="訓練機構">
                                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="CLASSID2" HeaderText="班別代碼">
                                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="STDATE" HeaderText="開訓日">
                                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="FTDATE" HeaderText="結訓日">
                                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="CLASSCNAME" HeaderText="班別名稱">
                                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="APPLYDATE" HeaderText="申請日期">
                                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="ACCTNAME" HeaderText="已授權者">
                                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundColumn>
                                                        <asp:TemplateColumn HeaderText="補登開放功能">
                                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" />
                                                            <ItemTemplate>
                                                                <asp:Label ID="labFUNIDN" runat="server"></asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:BoundColumn DataField="ENDDATE" HeaderText="結束日期">
                                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="APPLIEDRESULT2" HeaderText="審核結果">
                                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundColumn>
                                                        <asp:TemplateColumn HeaderText="功能">
                                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" Font-Size="Small" />
                                                            <ItemTemplate>
                                                                <asp:LinkButton ID="lbupdate1" runat="server" Text="修改" CommandName="UPD1" CssClass="asp_Export_M"></asp:LinkButton>
                                                                <asp:LinkButton ID="lbview1" runat="server" Text="檢視" CommandName="VIE1" CssClass="linkbutton"></asp:LinkButton>
                                                                <asp:LinkButton ID="lbprint1" runat="server" Text="列印" CommandName="PRT1" CssClass="asp_Export_M"></asp:LinkButton>
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
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                    <asp:Panel ID="Paneledit1" runat="server" Visible="False">
                        <table class="table_sch" cellpadding="1" cellspacing="1" width="100%">
                            <tr>
                                <td class="bluecol" style="width: 20%;">申請日期： </td>
                                <td class="whitecol" style="width: 30%;">
                                    <asp:Label ID="labAPPLYDATE" runat="server"></asp:Label>
                                </td>
                                <td class="bluecol" style="width: 20%;">審核結果： </td>
                                <td class="whitecol" style="width: 30%;">
                                    <asp:Label ID="labAPPLIEDRESULT2" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">年度： </td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="ddlYears" runat="server" AutoPostBack="true">
                                    </asp:DropDownList>
                                </td>
                                <td class="bluecol_need">轄區： </td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="ddlDistID" runat="server" AutoPostBack="true">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">訓練計畫： </td>
                                <td class="whitecol" colspan="3">
                                    <asp:DropDownList ID="ddlPlan" runat="server" AutoPostBack="true">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">訓練機構： </td>
                                <td class="whitecol" colspan="3">
                                    <asp:DropDownList ID="ddlOrgName" runat="server" AutoPostBack="True">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">班級名稱： </td>
                                <td class="whitecol" colspan="3">
                                    <asp:DropDownList ID="ddlClassCName" runat="server">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">承辦人員： </td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="ddlAccount" runat="server">
                                    </asp:DropDownList>
                                </td>
                                <td class="bluecol_need">補登資料<br />
                                    原因： </td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="ddlReasonID" runat="server">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">補登原因<br />
                                    及檢討改善作法 </td>
                                <td class="whitecol" colspan="3">
                                    <asp:TextBox ID="txtReason" runat="server" Columns="5" Width="60%"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">&nbsp; 開放功能 ： </td>
                                <td class="whitecol">
                                    <asp:CheckBoxList ID="cb_SelFunID" runat="server" RepeatLayout="Flow">
                                    </asp:CheckBoxList>
                                </td>
                                <td class="bluecol_need">&nbsp; 結束日期 ： </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="EndDate" Width="40%" onfocus="this.blur()" runat="server"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('EndDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                </td>
                            </tr>
                            <%--<tr>
							<td class="bluecol_need">&nbsp; 是否歸責單位 ：</td>
							<td class="whitecol" colspan="3">
								<asp:RadioButtonList ID="rblBlameUnit" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
									<asp:ListItem Selected="True" Value="Y">是</asp:ListItem>
									<asp:ListItem Value="N">否</asp:ListItem>
								</asp:RadioButtonList>
							</td>
						</tr>--%>
                            <tr>
                                <td align="center" class="whitecol" colspan="4">
                                    <asp:Button ID="btnSave1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="btnQuit1" runat="server" Text="離開" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="Hid_RETID" runat="server" />
    </form>
</body>
</html>
