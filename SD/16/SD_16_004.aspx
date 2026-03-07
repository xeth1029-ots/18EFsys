<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_16_004.aspx.vb" Inherits="WDAIIP.SD_16_004" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>行政管理疏失/重大異常狀況</title>
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
        function GETvalue() {
            document.getElementById('Button5').click();
        }
        function choose_class() {
            var RIDValue = document.getElementById('RIDValue');
            openClass('../02/SD_02_ch.aspx?RID=' + RIDValue.value);
        }
        function checkSave() {
            //debugger;
            var msg = '';
            var vHid_LID = document.getElementById('Hid_LID').value;
            if (document.getElementById('RIDValue').value == "") { msg += '請選擇【訓練機構】\n'; }
            if (vHid_LID != '0' && document.getElementById('OCIDValue1').value == "") { msg += '請選擇【職類/班別】\n'; }
            if (document.getElementById('txt_idno').value == "") { msg += '請填寫【身分證號碼】\n'; }
            if (document.getElementById('txt_No').value == "") { msg += '請選擇【處分文號】\n'; }
            if (isEmpty(document.form1.ddlSBTERMS)) { msg += '請選擇處分緣由\n'; }

            if (document.getElementById('txt_SBSdate').value == "") { msg += '請選擇【處分日期】\n'; }
            if (document.getElementById('ddl_SBYears').selectedIndex == 0) { msg += '請選擇【處分年限】\n'; }
            if (document.getElementById('txt_SBComment').value == "") { msg += '請填寫【處分事由】\n'; }
            else {
                if (checkMaxLen(document.getElementById('txt_SBComment').value, 300 * 2)) { msg += '【處分事由】長度不可超過300字元\n'; }
            }
            if (msg != '') {
                window.alert(msg);
                return false;
            } else {
                msg = '';
                msg += '\n請確認資料是否無誤,儲存後資料將不可修改\n\n';
                msg += '如確認資料無誤後,請按下確定,謝謝!!\n';
                return confirm(msg);
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" id="Table1" width="100%">
                        <tr>
                            <td class="font">
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;行政管理疏失/重大異常狀況</asp:Label>
                            </td>
                        </tr>
                    </table>
                    <asp:Panel ID="Panel1" runat="server" Visible="True">
                        <table class="table_sch" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol" width="20%">訓練機構 </td>
                                <td class="whitecol" width="80%">
                                    <asp:TextBox ID="center" runat="server" Width="70%"></asp:TextBox>
                                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                    <input id="Button6" type="button" value="..." name="Button6" runat="server" class="asp_button_Mini" />
                                    <asp:Button ID="Button5" Style="display: none" runat="server" Text="Button5"></asp:Button>
                                    <span id="HistoryList2" style="display: none; z-index: 100; position: absolute" onclick="GETvalue()">
                                        <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">職類/班別 </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="35%"></asp:TextBox>
                                    <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="35%"></asp:TextBox>
                                    <input onclick="choose_class()" value="..." type="button" />
                                    <input id="TMIDValue1" type="hidden" runat="server" />
                                    <input id="OCIDValue1" type="hidden" runat="server" />
                                    <span style="z-index: 1; position: absolute; display: none; left: 270px" id="HistoryList">
                                        <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                    </span>
                                    <iframe style="position: absolute; display: none; left: 270px" id="FrameObj" height="0" frameborder="0" width="100%"></iframe>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">查核確認日期 </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="sch_VERIFYDATE1" runat="server" MaxLength="10" Width="15%"></asp:TextBox>
                                    <span id="span1" runat="server">
                                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= sch_VERIFYDATE1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span> ~
                                    <asp:TextBox ID="sch_VERIFYDATE2" runat="server" MaxLength="10" Width="15%"></asp:TextBox>
                                    <span id="span2" runat="server">
                                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= sch_VERIFYDATE2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span>
                                </td>
                            </tr>
                            <tr>
                                <td align="center" class="whitecol" colspan="4">
                                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                    <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                    <asp:Button ID="btn_Sch" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                    &nbsp;<asp:Button ID="btn_Add" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                        </table>
                        <div style="width: 100%; text-align: center">
                            <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
                        </div>
                        <table id="tb_DataGrid1" border="0" cellspacing="0" cellpadding="0" width="100%" runat="server" class="font">
                            <tr>
                                <td align="center">
                                    <div id="divDataGrid1" runat="server">
                                        <%--序號、訓練機構、班別名稱、開訓日期、結訓日期、查核確認日期、功能 (修改)--%>
                                        <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                                            <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                            <Columns>
                                                <asp:BoundColumn HeaderText="序號">
                                                    <HeaderStyle HorizontalAlign="Center" Width="3%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="ORGNAME" HeaderText="訓練機構">
                                                    <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="CLASSCNAME2" HeaderText="班別名稱">
                                                    <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="STDATE_ROC" HeaderText="開訓日期">
                                                    <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="FTDATE_ROC" HeaderText="結訓日期">
                                                    <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="VERIFYDATE_ROC" HeaderText="查核確認日期">
                                                    <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:TemplateColumn HeaderText="功能">
                                                    <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:HiddenField ID="H_SEQNO" runat="server" />
                                                        <asp:HiddenField ID="H_OCID" runat="server" />
                                                        <asp:HiddenField ID="H_CMID" runat="server" />
                                                        <%--<asp:LinkButton ID="lbtView" runat="server" CommandName="view" CssClass="linkbutton">檢視</asp:LinkButton>--%>
                                                        <asp:LinkButton ID="lbtEdit" runat="server" CommandName="edit" CssClass="linkbutton">修改</asp:LinkButton>
                                                        <asp:LinkButton ID="lbtDelt" runat="server" CommandName="delt" CssClass="linkbutton">刪除</asp:LinkButton>
                                                        <%--<asp:LinkButton ID="lbtDel" runat="server" CommandName="del" CssClass="asp_Export_M">刪除</asp:LinkButton>--%>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                            </Columns>
                                            <PagerStyle Visible="False"></PagerStyle>
                                        </asp:DataGrid>
                                    </div>
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                    <%--<asp:BoundColumn Visible="False" DataField="SBSN" HeaderText="流水號"></asp:BoundColumn>--%>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                    <asp:Panel ID="Panel2" runat="server" Visible="False">
                        <!--訓練期間 登錄日期 經查核確認日期 查核結果說明
重要工作事項未依核定課程施訓: 課程異常狀況: 其他未依核定課程施訓: 其他重大異常狀況:-->
                        <table class="table_sch" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol" style="width: 20%">訓練機構 </td>
                                <td class="whitecol" style="width: 80%"> <asp:Label ID="labORGNAME" runat="server" Text=""></asp:Label> </td>
                            </tr>
                            <tr>
                                <td class="bluecol">班別名稱 </td>
                                <td class="whitecol"> <asp:Label ID="labCLASSCNAME2" runat="server" Text=""></asp:Label> </td>
                            </tr>
                            <tr>
                                <td class="bluecol">訓練期間 </td>
                                <td class="whitecol">
                                    <asp:Label ID="labSFTDATE" runat="server" Text=""></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">登錄日期 </td>
                                <td class="whitecol">
                                    <asp:Label ID="labCREATEDATE" runat="server" Text=""></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">經查核確認日期 </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="VERIFYDATE" runat="server" MaxLength="10" Width="15%"></asp:TextBox>
                                    <span id="span3" runat="server"> <img style="cursor: pointer" onclick="javascript:show_calendar('<%= VERIFYDATE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">查核結果說明 </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="RESULT" runat="server" TextMode="MultiLine" Width="98%" Rows="5"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">重要工作事項未依核定課程施訓 </td>
                                <td class="whitecol">
                                    <asp:CheckBoxList ID="CBL_NAPPROV" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow"></asp:CheckBoxList>
                                    <asp:TextBox ID="NAPPROV" runat="server" TextMode="MultiLine" Width="98%"  Rows="3"></asp:TextBox>
                                </td>                                                                          
                            </tr>                                                                              
                            <tr>                                                                               
                                <td class="bluecol">課程異常狀況 </td>                                         
                                <td class="whitecol">                                                          
                                    <asp:CheckBoxList ID="CBL_CEXCEP" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow"></asp:CheckBoxList>
                                    <asp:TextBox ID="CEXCEP" runat="server" TextMode="MultiLine" Width="98%"  Rows="3"></asp:TextBox>
                                </td>                                                                          
                            </tr>                                                                              
                            <tr>                                                                               
                                <td class="bluecol">其他未依核定課程施訓 </td>                                 
                                <td class="whitecol">                                                          
                                    <asp:CheckBoxList ID="CBL_OTHNAPPROV" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow"></asp:CheckBoxList>
                                    <asp:TextBox ID="OTHNAPPROV" runat="server" TextMode="MultiLine" Width="98%"  Rows="3"></asp:TextBox>
                                </td>                                                                          
                            </tr>                                                                              
                            <tr>                                                                               
                                <td class="bluecol">其他重大異常狀況 </td>                                     
                                <td class="whitecol">                                                          
                                    <asp:CheckBoxList ID="CBL_OTHMAJOR" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow"></asp:CheckBoxList>
                                    <asp:TextBox ID="OTHMAJOR" runat="server" TextMode="MultiLine" Width="98%"  Rows="3"></asp:TextBox>
                                </td>
                            </tr>
                        </table>
                        <table width="100%">
                            <tr>
                                <td align="center" class="whitecol">
                                    <asp:Button ID="btn_Save1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                                    &nbsp;<asp:Button ID="btn_Leave1" runat="server" Text="離開" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>                    
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="Hid_CMID" runat="server" />
        <asp:HiddenField ID="Hid_OCID" runat="server" />
        <asp:HiddenField ID="Hid_SEQNO" runat="server" />
    </form>
</body>
</html>
