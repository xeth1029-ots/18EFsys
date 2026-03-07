<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_02_001.aspx.vb" Inherits="WDAIIP.TC_02_001" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>班級查詢作業</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function checkaudit1() {
            var myaudit1 = document.getElementById('audit1');
            var myIsApprPaper = document.form1.IsApprPaper;
            myaudit1.style.display = 'none';
            if (myIsApprPaper.item(0).checked == true) { myaudit1.style.display = ''; }
            /*
            var valIsApprPaper = $("input:radio[name='IsApprPaper']:checked").val();
            $("#audit1").removeAttr("style").hide();
            if (valIsApprPaper == 'Y') { $("#audit1").show(); }
            */
            //$('#<=IsApprPaper.ClientID >').find("input[value='Y']").prop("checked", true);
        }

        //[全選／全不選]
        function doSelectAll(obj) {
            //chkItem
            $("input[type=checkbox][data-role='chkItem']:enabled").prop("checked", obj.checked);
        }

    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" cellspacing="1" id="FrameTable" cellpadding="1" width="100%" border="0">
            <tr>
                <td class="font">
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;班級查詢</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol" width="16%">訓練機構</td>
                            <td colspan="3" class="whitecol" width="84%">
                                <asp:TextBox ID="center" runat="server" Width="60%"></asp:TextBox><%--onfocus="this.blur()"--%>
                                <input id="Org" type="button" value="..." name="Org" runat="server">
                                <input id="RIDValue" style="width: 10%;" type="hidden" name="RIDValue" runat="server">
                                <input id="Orgidvalue" style="width: 10%;" type="hidden" name="Orgidvalue" runat="server">
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                <asp:Label ID="LabTMID" runat="server">訓練職類</asp:Label></td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="TB_career_id" runat="server" Columns="30" Width="40%"></asp:TextBox><%--onfocus="this.blur()"--%>
                                <input id="btu_sel" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." name="btu_sel" runat="server">&nbsp;
                                <input id="trainValue" type="hidden" name="trainValue" runat="server">
                                <input id="jobValue" type="hidden" name="jobValue" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label></td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" Columns="30" Width="40%"></asp:TextBox><%--onfocus="this.blur()"--%>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server">
                                <input id="cjobValue" type="hidden" name="cjobValue" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">班級名稱</td>
                            <td class="whitecol">
                                <asp:TextBox ID="ClassName" runat="server" Columns="44" MaxLength="55" Width="88%"></asp:TextBox></td>
                            <td class="bluecol" width="16%">期別</td>
                            <td class="whitecol" width="34%">
                                <asp:TextBox ID="CyclType" runat="server" Columns="10" MaxLength="3" Width="33%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">課程代碼</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="s_OCID" runat="server" Columns="20" MaxLength="10" Width="22%"></asp:TextBox></td>
                        </tr>
                        <tr id="tr_AppStage_TP28" runat="server">
                            <td class="bluecol">申請階段</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="AppStage" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol">資料類型</td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="IsApprPaper" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                                    <asp:ListItem Value="Y" Selected="True">正式</asp:ListItem>
                                    <asp:ListItem Value="N">草稿</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="tr_audit1" runat="server">
                            <td class="bluecol">審核狀態</td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="audit" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                                    <asp:ListItem Value="A" Selected="True">不區分</asp:ListItem>
                                    <asp:ListItem Value="Y">已審核</asp:ListItem>
                                    <asp:ListItem Value="N">審核中</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <%--增加【轉班上架】欄位，選項：不區分、未轉班、已轉班--%>
                        <tr id="tr_TransFlag" runat="server">
                            <td class="bluecol">轉班上架</td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="rbl_TransFlagS" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                                    <asp:ListItem Value="A" Selected="True">不區分</asp:ListItem>
                                    <asp:ListItem Value="N">未轉班</asp:ListItem>
                                    <asp:ListItem Value="Y">已轉班</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="trSearchYear" runat="server">
                            <td class="bluecol">查詢年度</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="yearlist" runat="server"></asp:DropDownList>
                                <font color="red">(若未選擇年度則依登入計畫，帶入訓練機構)</font>
                                <input id="TPlanid" type="hidden" name="TPlanid" runat="server" />
                                <input id="Re_ID" type="hidden" name="Re_ID" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">匯出檔案格式</td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                                    <asp:ListItem Value="ODS">ODS</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                    <table cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="whitecol">
                                <div align="center">
                                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                    <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                    <asp:Button ID="btnQuery" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="btnExport1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                                    <asp:Button ID="btnExport2" runat="server" Text="匯出開班預定表" CssClass="asp_Export_M"></asp:Button>
                                    <asp:Button ID="btnEnter1" runat="server" Text="批次轉班上架" CssClass="asp_Export_M"></asp:Button>
                                </div>
                            </td>
                        </tr>
                    </table>
                    <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td align="center">
                                <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                    <tr>
                                        <td>訓練計畫：<asp:Label ID="TPlanName" runat="server" CssClass="font"></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:DataGrid ID="dtPlan" runat="server" Width="100%" CssClass="font" AllowSorting="True" PagerStyle-HorizontalAlign="Left"
                                                PagerStyle-Mode="NumericPages" AllowPaging="True" OnItemCommand="dtPlan_ItemCommand" AutoGenerateColumns="False">
                                                <AlternatingItemStyle BackColor="#F5F5F5" />
                                                <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                                                <Columns>
                                                    <asp:BoundColumn HeaderText="序號" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                                    <asp:TemplateColumn>
                                                        <ItemStyle HorizontalAlign="Center" />
                                                        <HeaderTemplate>
                                                            <label>選取</label>
                                                            <br />
                                                            <input type='checkbox' onclick='doSelectAll(this)' />
                                                        </HeaderTemplate>
                                                        <ItemTemplate>
                                                            <input type="checkbox" id="chkItem" data-role="chkItem" runat="server" />
                                                            <asp:HiddenField ID="Hid_PCS" runat="server" />
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:BoundColumn DataField="PlanYearROCAG" HeaderText="計畫年度" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                                    <asp:BoundColumn DataField="AppliedDate" SortExpression="AppliedDate" HeaderText="申請日期">
                                                        <HeaderStyle ForeColor="#00ffff"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center" />
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="STDate" SortExpression="STDate" HeaderText="訓練起日">
                                                        <HeaderStyle ForeColor="#00ffff"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="FDDate" SortExpression="FDDate" HeaderText="訓練迄日">
                                                        <HeaderStyle ForeColor="#00ffff"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="OrgName2" HeaderText="管控單位"></asp:BoundColumn>
                                                    <asp:BoundColumn DataField="OrgName" SortExpression="OrgName" HeaderText="機構名稱">
                                                        <HeaderStyle ForeColor="#00ffff"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="OCID" HeaderText="課程代碼"></asp:BoundColumn>
                                                    <asp:BoundColumn DataField="ClassName" SortExpression="ClassName" HeaderText="班名">
                                                        <HeaderStyle ForeColor="#00ffff"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn HeaderText="審核狀態" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                                    <asp:BoundColumn DataField="VerReason" HeaderText="未通過原因"></asp:BoundColumn>
                                                    <asp:BoundColumn HeaderText="已轉班" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                                    <asp:TemplateColumn HeaderText="功能">
                                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        <ItemTemplate>
                                                            <asp:LinkButton ID="lbtUpdate" runat="server" Text="修改" CommandName="update" CssClass="linkbutton"></asp:LinkButton>
                                                            <asp:LinkButton ID="lbtDel" runat="server" Text="刪除" CommandName="Del" CssClass="linkbutton"></asp:LinkButton>
                                                            <asp:LinkButton ID="lbtPrint" runat="server" Text="列印" CommandName="Print" CssClass="asp_Export_M"></asp:LinkButton>
                                                            <asp:LinkButton ID="lbtSend" runat="server" Text="送出" CommandName="Send" CssClass="linkbutton"></asp:LinkButton>
                                                            <asp:LinkButton ID="lbtReturn" runat="server" Text="還原" CommandName="Return" CssClass="linkbutton"></asp:LinkButton>
                                                            <asp:LinkButton ID="lbtDef" runat="server" Text="經費明細" CommandName="Def" CssClass="linkbutton"></asp:LinkButton>
                                                            <asp:LinkButton ID="lbtShelf" runat="server" Text="轉班上架" CommandName="Shelf" CssClass="asp_Export_M"></asp:LinkButton>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="功能">
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        <ItemTemplate>
                                                            <asp:LinkButton ID="lbtEdit" runat="server" Text="修改" CommandName="btnEdit" CssClass="linkbutton"></asp:LinkButton>
                                                            <asp:LinkButton ID="lbtDel1" runat="server" Text="刪除" CommandName="btnDel" CssClass="linkbutton"></asp:LinkButton>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <%--<asp:BoundColumn Visible="False" DataField="FirResult" HeaderText="FirResult"></asp:BoundColumn>--%>
                                                </Columns>
                                                <PagerStyle Visible="False" HorizontalAlign="Left" ForeColor="Blue" Position="Top" Mode="NumericPages"></PagerStyle>
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
                            <td align="center">
                                <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label></td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <input id="isBlack" type="hidden" name="isBlack" runat="server" />
        <input id="orgname" type="hidden" name="orgname" runat="server" />
        <asp:HiddenField ID="hid_PPINFOtable_guid1" runat="server" />
    </form>
</body>
</html>
