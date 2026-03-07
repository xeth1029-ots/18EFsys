<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_19_001.aspx.vb" Inherits="WDAIIP.SD_19_001" %>

<!DOCTYPE html PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>開訓資料線上申辦</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link type="text/css" href="../../css/style.css" rel="stylesheet" />
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
        function GETvalue() {
            document.getElementById('Button4').click();
        }
        function choose_class() {
            openClass('../02/SD_02_ch.aspx?RID=' + document.form1.RIDValue.value);
        }
        function checkFile1(sizeLimit) {
            //sizeLimit單位:byte   
            const fileInput = document.querySelector('input[type="file"]');
            const file = fileInput.files[0];
            const fileType = file.type;
            const fileType2 = file.type.split('/')[1];
            //console.log('fileType : ' + fileType); //console.log('fileType2 : ' + fileType2);
            if (fileType !== 'application/pdf' || fileType2 !== 'pdf') {
                alert('只允許上傳 PDF 檔！');
                return false;
            }
            const fileSize = Math.round(file.size / 1024 / 1024);
            const fileSizeLimit = Math.round(sizeLimit / 1024 / 1024);
            //console.log('fileSize : ' + fileSize); //console.log('fileSizeLimit : ' + fileSizeLimit);
            if (fileSize > fileSizeLimit) {
                alert('您所選擇的檔案大小為 ' + fileSize + 'MB，超過了上傳上限! (檔案大小限制' + fileSizeLimit + 'MB以下)\n不允許上傳！');
                document.getElementById("File1").outerHTML = '<input name="File1" type="file" id="File1" size="66" accept=".pdf">';
                return false;
            }
            return true;
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;學員資料管理&gt;&gt;開訓資料線上申辦</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTableSch1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td>
                    <table class="table_sch">
                        <tr>
                            <td class="bluecol" style="width: 20%">訓練機構</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="center" runat="server" Width="60%"></asp:TextBox>
                                <input id="Button8" type="button" value="..." name="Button8" runat="server" class="asp_button_Mini" />
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <asp:Button ID="Button4" Style="display: none" runat="server" Text="Button4" CssClass="asp_button_S"></asp:Button>
                                <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">職類/班別</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="asp_button_Mini" />
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                <span id="HistoryList" style="position: absolute; display: none; left: 28%">
                                    <asp:Table ID="Historytable" runat="server" Width="100%"></asp:Table>
                                </span><span id="LAB_ADDREQUIRED_MSG">(新增必選)</span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">申請階段</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="SCH_DDLAPPSTAGE" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol">案件編號</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="SCH_TBCASENO" runat="server" Width="40%" MaxLength="30"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">開訓日期</td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="SCH_STDATE1" runat="server" MaxLength="10" Width="15%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('SCH_STDATE1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span>
                                ~
                                    <asp:TextBox ID="SCH_STDATE2" runat="server" MaxLength="10" Width="15%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('SCH_STDATE2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">結訓期間</td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="SCH_FTDATE1" runat="server" MaxLength="10" Width="15%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('SCH_FTDATE1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span>
                                ~
                                    <asp:TextBox ID="SCH_FTDATE2" runat="server" MaxLength="10" Width="15%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('SCH_FTDATE2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span>
                            </td>

                        </tr>
                        <tr>
                            <td class="bluecol">線上申辦日期 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="SCH_BIDATE1" runat="server" Columns="10" MaxLength="10" Width="15%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('SCH_BIDATE1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                ~
                                    <asp:TextBox ID="SCH_BIDATE2" runat="server" Columns="10" MaxLength="10" Width="15%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('SCH_BIDATE2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol" colspan="4">
                                <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                <asp:Button ID="BTN_SEARCH1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="BTN_ADDNEW1" runat="server" Text="新增申辦案件" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <table id="TB_DataGrid1" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td>
                                <!--序號	案件編號	申請階段	訓練機構	班級名稱	申辦人姓名	申辦日期	申辦狀態	審查狀態	功能-->
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" AllowCustomPaging="True" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號">
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TBCASENO" HeaderText="案件編號">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="APPSTAGE_N" HeaderText="申請階段">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ORGNAME" HeaderText="訓練機構">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="CLASSCNAME2" HeaderText="班級名稱">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TBCNAME" HeaderText="申辦人姓名">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TBCDATE_ROC" HeaderText="申辦日期">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TBCSTATUS_N" HeaderText="申辦狀態">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="APPLIEDRESULT_N" HeaderText="審查狀態">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="lBTN_DELETE1" runat="server" Text="刪除" CommandName="DELETE1" CssClass="linkbutton"></asp:LinkButton>&nbsp;
                                                <asp:LinkButton ID="lBTN_VIEW1" runat="server" Text="查看" CommandName="VIEW1" CssClass="linkbutton"></asp:LinkButton>&nbsp;
                                                <asp:LinkButton ID="lBTN_EDIT1" runat="server" Text="修改" CommandName="EDIT1" CssClass="linkbutton"></asp:LinkButton>&nbsp;
                                                <asp:LinkButton ID="lBTN_SENDOUT1" runat="server" Text="送出" CommandName="SENDOUT1" CssClass="linkbutton"></asp:LinkButton>&nbsp;
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
            <tr>
                <td align="center">
                    <asp:Label ID="labmsg1" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTableEdt1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td align="center">
                    <table id="Table4" class="table_sch" cellpadding="1" cellspacing="1">
                        <tr>
                            <td colspan="4" class="table_title" width="100%">申辦訓練機構資訊</td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">訓練機構</td>
                            <td class="whitecol" width="80%" colspan="3">
                                <asp:Label ID="labOrgNAME" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">案件編號</td>
                            <td class="whitecol" width="30%">
                                <asp:Label ID="labTBCASENO" runat="server"></asp:Label></td>
                            <td class="bluecol" width="20%">案件建立時間</td>
                            <td class="whitecol" width="30%">
                                <asp:Label ID="LabCREATEDATE" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">計畫年度 </td>
                            <td class="whitecol" width="30%">
                                <asp:Label ID="labBIYEARS" runat="server"></asp:Label></td>
                            <td class="bluecol" width="20%">申請階段 </td>
                            <td class="whitecol" width="30%">
                                <asp:Label ID="labAPPSTAGE" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">班級名稱</td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="labCLASSNAME2S" runat="server"></asp:Label></td>
                        </tr>

                        <tr id="tr_HISREVIEW" runat="server">
                            <td class="bluecol">歷程資訊</td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="labHISREVIEW" runat="server" Text=""></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">切換至</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="ddlSwitchTo" runat="server" AppendDataBoundItems="True" AutoPostBack="True"></asp:DropDownList>
                                <asp:Button ID="BTN_SEARCH2" runat="server" Text="重新查詢" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4" class="table_title" width="100%">線上申辦進度</td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">線上申辦進度</td>
                            <td class="whitecol" colspan="3" width="80%">
                                <asp:Label ID="labProgress" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol" colspan="4">
                                <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False">
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <AlternatingItemStyle />
                                    <Columns>
                                        <%--<asp:BoundColumn DataField="KBSID" HeaderText="項目"><HeaderStyle Width="10%" /></asp:BoundColumn>--%>
                                        <asp:BoundColumn DataField="KTNAME2" HeaderText="項目名稱">
                                            <HeaderStyle Width="20%" />
                                        </asp:BoundColumn>
                                        <%--<asp:BoundColumn DataField="BCFID" HeaderText="序號"><HeaderStyle Width="10%" /></asp:BoundColumn>--%>
                                        <%--<asp:BoundColumn DataField="FILENAME1" HeaderText="檔案名稱1"><HeaderStyle Width="10%" /></asp:BoundColumn>--%>
                                        <asp:BoundColumn DataField="SRCFILENAME1" HeaderText="上傳檔案">
                                            <HeaderStyle Width="14%" />
                                        </asp:BoundColumn>
                                        <%-- <asp:TemplateColumn HeaderText="檔案名稱">
                                            <HeaderStyle Width="14%" />
                                            <ItemTemplate>
                                                <asp:Label ID="LabFileName1" runat="server"></asp:Label>
                                                <input id="HFileName" type="hidden" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>--%>
                                        <asp:TemplateColumn HeaderText="退件原因">
                                            <HeaderStyle Width="14%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:Label ID="labRTUREASON" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:Button ID="BTN_DELFILE4" runat="server" Text="刪除檔案" CommandName="DELFILE4" CssClass="asp_Export_M" />
                                                <asp:Button ID="BTN_DOWNLOAD4" runat="server" Text="檔案下載" CommandName="DOWNLOAD4" CssClass="asp_Export_M" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4" class="table_title" width="100%">應備文件上傳</td>
                        </tr>
                        <tr>
                            <td colspan="4" class="class_title2_left">
                                <asp:Label ID="LabSwitchTo" runat="server" Text=""></asp:Label></td>
                        </tr>
                        <tr>
                            <td colspan="4">
                                <table cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr id="tr_LiteralSwitchTo" runat="server">
                                        <td class="bluecol" width="20%">文件說明</td>
                                        <td class="whitecol" colspan="3" width="80%">
                                            <%--1.持有本署TTQS訓練機構版評核等級有效期限證書之影本。如有效期限於所提訓練計畫開訓日前屆滿者，應檢附已申請TTQS評核，且須於課程開放報名日前補具TTQS評核證書之證明文件。	
                                            <br />2.勞動力發展署核可通知公函或由評核單位出具之函件<span class="Brown-style1">（需有評核單位之章戳）</span>。
                                            <asp:Literal ID="LiteralSwitchTo" runat="server"></asp:Literal>--%>
                                            <asp:Literal ID="LiteralSwitchTo" runat="server"></asp:Literal>
                                        </td>
                                    </tr>
                                    <tr id="tr_FILEDESC1" runat="server">
                                        <td class="bluecol">檔案格式說明</td>
                                        <td class="whitecol" colspan="3">
                                            <asp:Label ID="labFILEDESC1" runat="server" Text="PDF(掃瞄畫面需清楚，檔案大小2MB以下)"></asp:Label></td>
                                    </tr>
                                    <tr id="tr_DOWNLOADRPT1" runat="server">
                                        <td class="bluecol">下載報表 </td>
                                        <td colspan="3" class="whitecol">
                                            <asp:Button ID="BTN_DOWNLOADRPT1" runat="server" Text="下載報表1" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                                            <asp:Button ID="BTN_DOWNLOADRPT2" runat="server" Text="下載報表2" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                                        </td>
                                    </tr>

                                    <tr id="tr_DataGrid06" runat="server">
                                        <td colspan="4" class="whitecol">
                                            <asp:DataGrid ID="DataGrid06" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="6">
                                                <AlternatingItemStyle BackColor="#f5f5f5" />
                                                <HeaderStyle CssClass="head_navy" />
                                                <Columns>
                                                    <asp:BoundColumn DataField="ROWNUM1" HeaderText="序號">
                                                        <HeaderStyle HorizontalAlign="Center" Width="6%"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="SRCFILENAME1" HeaderText="上傳檔案">
                                                        <HeaderStyle Width="21%" />
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="MEMO1" HeaderText="備註說明">
                                                        <HeaderStyle Width="21%" />
                                                    </asp:BoundColumn>
                                                    <asp:TemplateColumn>
                                                        <HeaderStyle Width="11%" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <asp:HiddenField ID="Hid_CS14OFID" runat="server" />
                                                            <asp:Button ID="BTN_DOWNLOAD06" runat="server" Text="檔案下載" CommandName="DOWNLOAD06" CssClass="asp_Export_M" />
                                                            <asp:Button ID="BTN_DELFILE06" runat="server" Text="刪除檔案" CommandName="DELFILE06" CssClass="asp_button_M" />
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                </Columns>
                                                <PagerStyle Visible="False"></PagerStyle>
                                            </asp:DataGrid>
                                        </td>
                                    </tr>

                                    <tr id="tr_SENTBATVER" runat="server">
                                        <td class="bluecol">以目前版本批次送出</td>
                                        <td colspan="3" class="whitecol">
                                            <asp:Button ID="BTN_SENTBATVER" runat="server" Text="以目前版本批次送出" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                                        </td>
                                    </tr>
                                    <tr id="tr_SENDCURRVER" runat="server">
                                        <td class="bluecol">以目前版本送出</td>
                                        <td colspan="3" class="whitecol">
                                            <asp:Button ID="BTN_SENDCURRVER" runat="server" Text="以目前版本送出" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                                        </td>
                                    </tr>
                                    <tr id="tr_UPLOADFL1" runat="server">
                                        <td class="bluecol">檔案上傳 </td>
                                        <td colspan="3" class="whitecol">
                                            <%--<asp:DropDownList ID="depID" runat="server"></asp:DropDownList>--%>
                                            <input id="File1" type="file" size="66" name="File1" runat="server" accept=".pdf" />
                                            <asp:Button ID="But1" runat="server" Text="確定檔案上傳" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                                        </td>
                                    </tr>
                                    <tr id="tr_USELATESTVER" runat="server">
                                        <td class="bluecol">最近一次版本送件</td>
                                        <td class="whitecol" colspan="3">
                                            <asp:Button ID="bt_latestSend1" runat="server" Text="最近一次版本送件" CssClass="asp_Export_M" />
                                            <asp:Button ID="bt_latestDown1" runat="server" Text="下載" CssClass="asp_Export_M" />
                                        </td>
                                    </tr>
                                    <%--WAIVED 免附文件--%>
                                    <tr id="tr_WAIVED" runat="server">
                                        <td class="bluecol">免附文件</td>
                                        <td class="whitecol" colspan="3">
                                            <asp:CheckBox ID="CHKB_WAIVED" runat="server" />免附文件
                                        </td>
                                    </tr>
                                    <tr id="tr_USEMEMO1" runat="server">
                                        <td class="bluecol" id="td_USEMEMO1" runat="server">備註說明</td>
                                        <td class="whitecol" colspan="3">
                                            <asp:TextBox ID="txtMEMO1" runat="server" Width="60%"></asp:TextBox>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol" colspan="4">
                                <asp:Button ID="BTN_PREV1" runat="server" Text="回上一步" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="BTN_SAVETMP1" runat="server" Text="儲存(暫存)" CssClass="asp_button_M"></asp:Button>
                                <%--<asp:Button ID="BTN_SAVERC2" runat="server" Text="下載檢核表" CssClass="asp_button_M"></asp:Button>--%>
                                <asp:Button ID="BTN_SAVENEXT1" runat="server" Text="儲存後進下一步" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="BTN_BACK1" runat="server" Text="不儲存返回查詢" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <asp:Label ID="labmsg2" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                </td>
            </tr>
        </table>
        <%--申辦流水號：Hid_TBCID--%>
        <asp:HiddenField ID="Hid_TBCID" runat="server" />
        <asp:HiddenField ID="Hid_TBCASENO" runat="server" />
        <asp:HiddenField ID="Hid_ORGKINDGW" runat="server" />
        <asp:HiddenField ID="Hid_FirstKTSEQ" runat="server" />
        <asp:HiddenField ID="Hid_LastKTID" runat="server" />

        <%--目前項次(流水號)Hid_KTSEQ,(序號)Hid_KTID--%>
        <asp:HiddenField ID="Hid_KTSEQ" runat="server" />
        <asp:HiddenField ID="Hid_KTID" runat="server" />
        <asp:HiddenField ID="Hid_TBCFID" runat="server" />
        <%--<asp:HiddenField ID="HiddenField1" runat="server" />--%>
    </form>
</body>
</html>
