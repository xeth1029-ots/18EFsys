<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" EnableEventValidation="false" CodeBehind="TC_01_007.aspx.vb" Inherits="WDAIIP.TC_01_007" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>師資資料維護</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1" />
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1" />
    <meta name="vs_defaultClientScript" content="JavaScript" />
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5" />
    <link rel="stylesheet" type="text/css" href="../../css/style.css" />
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/OpenWin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-migrate-3.4.1.min.js"></script>
    <%--<script type="text/javascript" language="javascript" src="../../js/selectControl.js.aspx" charset="UTF-8"></script>--%>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        //顯示遮罩
        function aloader2on() {
            var construction2 = document.getElementById("construction2");
            var form1 = document.getElementById("form1");
            form1.style.display = "none";            //不顯示
            //construction2.style.display = "block"; //顯示     //(20180907 由於此遮罩屬於TIMS功能，因此先將此遮罩拿掉)
        }

        //關閉遮罩
        function aloader2off() {
            var construction2 = document.getElementById("construction2");
            //construction2.style.display = "none";  //不顯示
            //補顯示
            var form1 = document.getElementById("form1");
            form1.style.display = "";  //顯示
        }

        function ShowFrame() {
            document.getElementById('FrameObj').height = document.getElementById('HistoryRID').rows.length * 20;
            document.getElementById('FrameObj').style.display = document.getElementById('HistoryList2').style.display;
        }

        function closeDiv() {
            document.getElementById('eMeng').style.visibility = 'hidden';
        }

        /*個資法js*/
        function showLoginPwdDiv(num) {
            //num: 1:查詢 2:匯出 (記錄目前查詢按鈕)
            var hidSchBtnNum = document.getElementById('hidSchBtnNum');    //記錄目前查詢按鈕
            hidSchBtnNum.value = num;                                      //num: 1:查詢 2:匯出 (記錄目前查詢按鈕)
            var rblWorkMode_0 = document.getElementById('rblWorkMode_0');  //模糊顯示
            var rblWorkMode_1 = document.getElementById('rblWorkMode_1');  //正常顯示 
            var hidLockTime1 = document.getElementById('hidLockTime1');    //啟用鎖定
            var hidLockTime2 = document.getElementById('hidLockTime2');
            var divPwdFrame = document.getElementById('divPwdFrame');
            var txtdivPxssward = document.getElementById('txtdivPxssward');
            //document.getElementById('divFrame').style.display = 'none';
            //if (OCIDValue1.value == '') {
            //	alert('請選擇班級');
            //	return false;
            //}
            var blnPwdFrame = false; //不顯示密碼輸入
            if (rblWorkMode_1.checked != true) { hidLockTime1.value = '1'; }
            if (rblWorkMode_1.checked == true && hidLockTime1.value == '1' && hidLockTime2.value == '1') {
                blnPwdFrame = true; //顯示密碼輸入
            }
            //alert(hidLockTime1.value);
            if (blnPwdFrame) {
                divPwdFrame.style.display = 'inline'; //顯示
                if (txtdivPxssward != null) txtdivPxssward.focus();
                return false;
            }
            else {
                aloader2on();
                document.getElementById('divPwdFrame').style.display = 'none';
                return true;
            }
        }

        function chkTxtPxssward() {
            var txtdivPxssward = document.getElementById('txtdivPxssward');
            var msg = '';
            if (txtdivPxssward.value == '') msg = '請輸入您的個資安全密碼!';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
<%--    <style type="text/css">
        .auto-style1 { color: Black; text-align: right; padding: 4px 6px; background-color: #f1f9fc; border-right: 3px solid #49cbef; height: 45px; }
        .auto-style2 { color: #333333; padding: 4px; height: 45px; }
    </style>--%>
</head>
<body>
    <div id="construction2" onclick="aloader2off();">
        <table width="100%">
            <tr>
                <td align="center" valign="middle">
                    <img id="construction2-img" src="../../images/icon_construction-a.gif" alt="系統正在處理您的需求 請稍候.."></td>
            </tr>
        </table>
    </div>
    <form id="form1" method="post" runat="server">
        <div style="position: absolute; top: -333px">
            <input type="text" title="Chaff for Chrome Smart Lock" /><input type="password" title="Chaff for Chrome Smart Lock" /></div>
        <input id="HidVeMeng" type="hidden" value="none" runat="server" />
        <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%" class="font">
            <tr>
                <td align="center">
                    <table id="Table2" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                        <tbody>
                            <tr>
                                <td>
                                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;教師資料管理&gt;&gt;師資資料設定</asp:Label>
                                    <%--<asp:Label ID="VeMeng" runat="server" Visible="False">none</asp:Label>--%>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                    <table id="Table3" class="table_sch" cellpadding="1" cellspacing="1">
                        <tr>
                            <td class="bluecol" style="width: 20%">訓練機構 </td>
                            <td colspan="3" class="whitecol">
                                <input style="width: 40%;" id="center" onfocus="this.blur()" maxlength="50" size="16" name="center" runat="server">
                                <input id="Button5" value="..." type="button" name="Button5" runat="server" class="button_b_Mini">
                                <span style="z-index: 1; position: absolute; display: none" id="HistoryList2">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                                <iframe style="position: absolute; display: none" id="FrameObj" height="55" frameborder="0" width="312" scrolling="no"></iframe>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" style="width: 20%">講師姓名 </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="TextBox2" runat="server" Width="40%"></asp:TextBox></td>
                            <td class="bluecol" style="width: 20%">身分證號碼 </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="TextBox3" runat="server" Width="40%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">內外聘 </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="DropDownList4" runat="server" AutoPostBack="True">
                                    <asp:ListItem Value="0">--請選擇--</asp:ListItem>
                                    <asp:ListItem Value="1">內聘(專任)</asp:ListItem>
                                    <asp:ListItem Value="2">外聘(兼任)</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                            <td class="bluecol">師資別 </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="DropDownList1" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol">講師代碼 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TextBox4" runat="server" Width="20%"></asp:TextBox></td>
                            <td class="bluecol">排課使用 </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="DropDownList2" runat="server">
                                    <asp:ListItem Value="0">請選擇</asp:ListItem>
                                    <asp:ListItem Value="1">是</asp:ListItem>
                                    <asp:ListItem Value="2">否</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                        </tr>
                       
                        <tr>
                            <td class="bluecol">主要職類 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TB_career_id" runat="server" Columns="30" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input onclick="openTrain2(document.getElementById('trainValue').value);" value="..." type="button" class="button_b_Mini">
                                <input style="width: 48px; height: 22px" id="trainValue" size="2" type="hidden" name="trainValue" runat="server">
                                <input id="jobValue" type="hidden" name="jobValue" runat="server">
                            </td>
                            <td class="bluecol">職稱 </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="DropDownList3" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr id="tr_techtype12" runat="server">
                            <td class="bluecol">類別 </td>
                            <td colspan="3" class="whitecol">
                                <asp:CheckBox Style="z-index: 0" ID="cb_techtype1" runat="server" Text="講師" CssClass="font"></asp:CheckBox>
                                <asp:CheckBox Style="z-index: 0" ID="cb_techtype2" runat="server" Text="助教" CssClass="font"></asp:CheckBox>
                            </td>
                        </tr>
                        <%-- 2018 add For 自辦在職的類別(教師/第二教師) --%>
                        <tr id="tr_techtype34" runat="server">
                            <td class="bluecol">類別 </td>
                            <td colspan="3" class="whitecol">
                                <asp:CheckBox Style="z-index: 0" ID="cb_techtype3" runat="server" Text="教師" CssClass="font"></asp:CheckBox>
                                <asp:CheckBox Style="z-index: 0" ID="cb_techtype4" runat="server" Text="第二教師" CssClass="font"></asp:CheckBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">含排課<br />
                                匯入用代碼 </td>
                            <td colspan="3" class="whitecol">
                                <asp:CheckBox ID="cb_CourID" runat="server" CssClass="font"></asp:CheckBox></td>
                        </tr>
                         
                        <tr>
                            <td class="bluecol">匯入師資名冊 </td>
                            <td colspan="3" class="whitecol">
                                <input id="File2" size="40" type="file" name="File1" runat="server" accept=".xls,.ods" />
                                <asp:Button ID="Btn_XlsImport" runat="server" Text="匯入名冊" CssClass="asp_button_M"></asp:Button>&nbsp;(必須為ods或xls格式)
							    <asp:HyperLink ID="lnkImpSample1" runat="server" CssClass="font" ForeColor="#8080FF" NavigateUrl="../../Doc/ClassTeacher_v12.zip">下載整批上載格式檔</asp:HyperLink>
                                <asp:HyperLink ID="lnkImpSample2" runat="server" CssClass="font" ForeColor="#8080FF" NavigateUrl="../../Doc/ClassTeacher_v16.zip">下載整批上載格式檔</asp:HyperLink>
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
                        <tr>
                            <td class="bluecol">匯出師資名冊 </td>
                            <td colspan="3" class="whitecol">
                                <asp:Button ID="Btn_XlsEmport" runat="server" Text="匯出名冊" CssClass="asp_Export_M"></asp:Button></td>
                        </tr>
                        <tr>
                            <td class="bluecol">資料顯示模式 </td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="rblWorkMode" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                    <asp:ListItem Value="1" Selected="True">模糊顯示</asp:ListItem>
                                    <asp:ListItem Value="2">正常顯示</asp:ListItem>
                                </asp:RadioButtonList>
                                <asp:HiddenField ID="hidWorkMode" runat="server" />
                            </td>
                        </tr>
                        <tr id="tr_ddl_INQUIRY_S" runat="server">
                            <td class="bluecol_need">查詢原因</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="ddl_INQUIRY_Sch" runat="server"></asp:DropDownList>
                            </td>
                        </tr>
                    </table>
                    <table id="Table4B" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>&nbsp;
                                <asp:Button ID="Button2" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>&nbsp;
                                <asp:Button ID="Button7" runat="server" Text="列印排課匯入用的講師代碼" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <table id="DataGridTable" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" AllowCustomPaging="True" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TeacherID" HeaderText="講師代碼">
                                            <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TeachCName" HeaderText="講師名稱">
                                            <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="IDNO" HeaderText="身分證">
                                            <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="KindID" HeaderText="師資別">
                                            <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="KindEngage" HeaderText="內外聘">
                                            <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="lbtEdit" runat="server" Text="修改" CommandName="edit" CssClass="linkbutton"></asp:LinkButton><br />
                                                <asp:LinkButton ID="lbtPrt" runat="server" Text="列印師資資料" CommandName="print" CssClass="asp_Export_M"></asp:LinkButton><br />
                                                <asp:LinkButton ID="lbtDel" runat="server" Text="刪除" CommandName="del" CssClass="linkbutton"></asp:LinkButton>
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
                    <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                </td>
            </tr>
        </table>
        <table style="border-bottom: #455690 1px solid; border-left: #a6b4cf 1px solid; background-color: #c9d3f3; width: 376px; visibility: visible; border-top: #a6b4cf 1px solid; border-right: #455690 1px solid" id="eMeng" class="font" border="0" cellspacing="1" cellpadding="1" width="376" runat="server">
            <tr>
                <td background="../../images/MSNTitle.gif">
                    <table id="Table4" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td><strong><font color="#0000ff">問題轉入資料訊息：</font></strong></td>
                            <td style="cursor: pointer" onclick="closeDiv();" width="15" align="center">
                                <img src="../../images/CloseMsn.gif" width="13" height="13" alt="" /></td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td style="border-bottom: #b9c9ef 1px solid; border-left: #728eb8 1px solid; padding-bottom: 10px; padding-left: 10px; width: 100%; padding-right: 10px; height: 100%; color: #1f336b; border-top: #728eb8 1px solid; border-right: #b9c9ef 1px solid; padding-top: 15px" background="../../images/MsnBack.gif" align="center">
                    <asp:DataGrid ID="Datagrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <Columns>
                            <asp:BoundColumn DataField="Index" HeaderText="第幾筆錯誤"></asp:BoundColumn>
                            <asp:BoundColumn DataField="TeacherID" HeaderText="講師代碼"></asp:BoundColumn>
                            <asp:BoundColumn DataField="Name" HeaderText="講師姓名"></asp:BoundColumn>
                            <asp:BoundColumn DataField="IDNO" HeaderText="身分證號碼"></asp:BoundColumn>
                            <asp:BoundColumn DataField="Reason" HeaderText="原因"></asp:BoundColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                </td>
            </tr>
        </table>
        <div id="divPwdFrame" runat="server" style="position: absolute; border-width: 6px; border-style: double; border-color: #4682B4; display: none; width: 350px; height: 300px; left: 195px; top: 200px; background-color: #FFFAF0; padding-left: 30px; padding-top: 30px;">
            <table align="center">
                <tr>
                    <td>請輸入個資安全密碼 </td>
                </tr>
                <tr>
                    <td>
                        <asp:TextBox ID="txtdivPxssward" runat="server" TextMode="Password"></asp:TextBox></td>
                </tr>
                <tr>
                    <td></td>
                </tr>
                <tr>
                    <td align="center">
                        <asp:Button ID="btndivPwdSubmit" runat="server" Text="確定" OnClientClick="return chkTxtPxssward();" CssClass="asp_button_S" CommandName="btndivPwdSubmit" />&nbsp;
					    <input id="btn_close" type="button" value="關閉" onclick="document.getElementById('divPwdFrame').style.display = 'none'; document.getElementById('labChkMsg').text = '';" class="button_b_S" />
                    </td>
                </tr>
                <tr>
                    <td></td>
                </tr>
                <tr>
                    <td align="center">
                        <asp:Label ID="labChkMsg" runat="server" CssClass="needFont"></asp:Label></td>
                </tr>
            </table>
        </div>
        <input id="hidLockTime1" type="hidden" name="hidLockTime1" runat="server" value="1" />
        <input id="hidSchBtnNum" type="hidden" name="hidSchBtnNum" runat="server" value="1" />
        <input id="hidLockTime2" type="hidden" name="hidLockTime2" runat="server" value="1" />
    </form>
</body>
</html>
