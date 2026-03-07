<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="CP_01_009.aspx.vb" Inherits="WDAIIP.CP_01_009" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>CP_01_009</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <%--<script language="javascript" type="text/javascript" src="../../js/date-picker.js"></script>--%>
    <script language="javascript" type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script language="javascript" type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script language="javascript" type="text/javascript">
        //決定date-picker元件使用的是西元年or民國年，by:20181018
        var bl_rocYear = "<%=ConfigurationManager.AppSettings("REPLACE2ROC_YEARS") %>";
        var scriptTag = document.createElement('script');
        var jsPath = (bl_rocYear == "Y" ? "../../js/date-picker2.js" : "../../js/date-picker.js");
        scriptTag.src = jsPath;
        document.head.appendChild(scriptTag);

        function search() {
            var form = document.getElementById("form1");
            var msg = '';
            if (form.txtSDate.value == '') msg += '請選擇訪查日期起\n';
            if (form.txtEDate.value == '') msg += '請選擇訪查日期迄\n';
            //if (msg!=''){
            //	alert(msg);
            //	return false;
            //}
        }

        function Q1Chk() {
            var form = document.getElementById("form1");
            if (trim(form.txtQ1A.value) != "" && trim(form.txtQ1B.value) != "") {
                if (isPositiveInt(form.txtQ1A.value) && isPositiveInt(form.txtQ1B.value)) {
                    form.txtQ1Per.value = parseInt(form.txtQ1A.value) / parseInt(form.txtQ1B.value) * 100;
                }
            } else {
                form.txtQ1Per.value = "";
            }
        }

        function saveChk() {
            var form = document.getElementById("form1");
            var msg = "";
            if (form.hidEditOrgID.value == "" || form.hidEditPlanID.value == "") {
                msg += "請選擇機構\n";
            }
            if (form.txtVDate.value == "") {
                msg += "請選擇訪查日期\n";
            }
            if (!form.rdoQ11.checked && !form.rdoQ10.checked) {
                msg += "請選擇訪查項目一符合狀況\n";
            }
            if ((!isPositiveFloat(form.txtQ1Per.value) && !isPositiveInt(form.txtQ1Per.value)) || parseInt(form.txtQ1Per.value) > 100) {
                msg += "請輸入訪查項目一正確的訪查率\n";
            }
            if (!form.rdoQ20.checked && !form.rdoQ21.checked && !form.rdoQ22.checked) {
                msg += "請選擇訪查項目二填報狀況\n";
            }
            if (form.rdoQ22.checked && form.txtQ2Other.value == "") {
                msg += "請輸入訪查項目二其他說明\n";
            }
            if (!form.rdoQ30.checked && !form.rdoQ31.checked && !form.rdoQ32.checked) {
                msg += "請選擇訪查項目三確認狀況\n";
            }
            if (form.rdoQ32.checked && form.txtQ3Other.value == "") {
                msg += "請輸入訪查項目三其他說明\n";
            }
            if (!form.rdoQ40.checked && !form.rdoQ41.checked && !form.rdoQ42.checked) {
                msg += "請選擇訪查項目四確認狀況\n";
            }
            if (form.rdoQ42.checked && form.txtQ4Other.value == "") {
                msg += "請輸入訪查項目四其他說明\n";
            }
            if (!form.rdoQ50.checked && !form.rdoQ51.checked && !form.rdoQ52.checked && !form.rdoQ53.checked) {
                msg += "請選擇訪查項目五確認狀況\n";
            }
            if (form.rdoQ52.checked && form.txtQ5Other.value == "") {
                msg += "請輸入訪查項目五其他說明\n";
            }
            if (form.txtUnitName.value == "") {
                msg += "請輸入受訪單位人員姓名\n";
            }
            if (form.txtFillerName.value == "") {
                msg += "請輸入訪查人員姓名\n";
            }
            if (msg != "") {
                alert(msg);
                return false;
            }
        }

        function Clear() {
            var form = document.getElementById("form1");
            form.hidOrgID.value = "";
            form.hidRID.value = "";
            form.hidPlanID.value = "";
            form.txtTBplan.value = "";
        }

        function Clear2() {
            var form = document.getElementById("form1");
            form.hidEditOrgID.value = "";
            form.hidEditRID.value = "";
            form.hidEditPlanID.value = "";
            form.txtTBplan2.value = "";
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;查核績效管理&gt;&gt;訪查報告表</asp:Label>
                </td>
            </tr>
        </table>

        <asp:Panel ID="panelSch" runat="server">
            <table border="0" cellspacing="1" cellpadding="1" width="100%">
                <tr>
                    <td>
                        <table class="table_sch" cellspacing="1" cellpadding="1">
                            <tr>
                                <td class="bluecol" style="width: 20%">隸屬機構</td>
                                <td class="whitecol" colspan="4">
                                    <asp:TextBox ID="txtTBplan" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox>
                                    <input id="hidRID" type="hidden" name="RIDValue" runat="server">
                                    <input id="hidOrgID" type="hidden" name="hidOrgID" runat="server">
                                    <input id="hidPlanID" type="hidden" name="PlanIDValue" runat="server">
                                    <input id="choice_button" onclick="javascript: wopen('../../Common/LevPlan.aspx?fOrgID=hidOrgID&amp;fPlanID=hidPlanID&amp;fTBplan=txtTBplan', '計畫階段', 850, 400, 1);" value="選擇" type="button" name="choice_button" runat="server" class="asp_button_M">
                                    <input id="btn_clear" onclick="Clear();" value="清除" type="button" name="Button3" runat="server" class="asp_button_M">
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">訪查日期</td>
                                <td class="whitecol" colspan="4">
                                    <asp:TextBox ID="txtSDate" runat="server" Width="15%" onfocus="this.blur()"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= txtSDate.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">
                                    ～
                                    <asp:TextBox ID="txtEDate" runat="server" Width="15%" onfocus="this.blur()"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= txtEDate.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">
                                </td>
                            </tr>
                        </table>
                        <div align="center" class="whitecol">
                            <asp:Button ID="btnSearch" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                            &nbsp;<asp:Button ID="btnAdd" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>
                            &nbsp;<asp:Button ID="btnERpt" runat="server" Text="列印空白訪查報告表" CssClass="asp_Export_M"></asp:Button>
                        </div>
                        <div align="center">
                            <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label></div>
                    </td>
                </tr>
            </table>
            <table id="DataGridTable" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                <tr>
                    <td>
                        <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                            <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                            <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                            <Columns>
                                <asp:BoundColumn HeaderText="序號">
                                    <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="VisitorDate" HeaderText="訪查日期" DataFormatString="{0:d}">
                                    <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="OrgName" HeaderText="受訪單位"></asp:BoundColumn>
                                <asp:BoundColumn DataField="PlanName" HeaderText="訪查計畫"></asp:BoundColumn>
                                <asp:TemplateColumn HeaderText="功能">
                                    <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    <ItemTemplate>
                                        <asp:Button ID="btnEdit" runat="server" Text="修改" CommandName="edit" CssClass="asp_button_M"></asp:Button>
                                        <asp:Button ID="btnDel" runat="server" Text="刪除" CommandName="del" CssClass="asp_button_M"></asp:Button>
                                        <asp:Button ID="btnPrint" runat="server" Text="列印" CommandName="prt" CssClass="asp_Export_M"></asp:Button>
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
        </asp:Panel>
        <asp:Panel ID="panelEdit" runat="server">
            <table border="0" cellspacing="1" cellpadding="1" width="100%">
                <tr>
                    <td>
                        <table class="table_sch" cellspacing="1" cellpadding="1">
                            <tr>
                                <td class="bluecol" style="width: 20%">隸屬機構</td>
                                <td colspan="4" class="whitecol">
                                    <asp:TextBox ID="txtTBplan2" runat="server" Width="40%" onfocus="this.blur()"></asp:TextBox><input id="hidEditRID" type="hidden" name="RIDValue" runat="server">
                                    <input id="btnChoice" onclick="javascript: wopen('../../Common/LevPlan.aspx?fOrgID=hidEditOrgID&amp;fPlanID=hidEditPlanID&amp;fTBplan=txtTBplan2', '計畫階段', 850, 570, 1);" value="選擇" type="button" runat="server" class="asp_button_M">
                                    <input id="btnClear" onclick="Clear2();" value="清除" type="button" runat="server" class="asp_button_M">
                                    <input id="hidChkEdit" type="hidden" runat="server">
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" style="width: 20%">訪查日期</td>
                                <td class="whitecol">
                                    <div>
                                        <div style="float: left; width: 15%;">
                                            <asp:TextBox ID="txtVDate" runat="server" Width="100%" onfocus="this.blur()"></asp:TextBox></div>
                                        <div style="float: left; width: 15%;">
                                            <asp:Panel ID="panelImg" runat="server" Width="20%">
                                                <img style="cursor: pointer" id="imgVDate" onclick="javascript:show_calendar('<%= txtVDate.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30"></asp:Panel>
                                            <%--<div style="clear:both;"></div>--%>
                                            <%--這是用來還原float:left的--%>
                                        </div>
                                    </div>
                                </td>
                                <%--
                                <td class="whitecol"><asp:Panel ID="panelImg" runat="server"><img style="cursor: pointer" id="imgVDate" onclick="javascript:show_calendar('<%= txtVDate.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30"></asp:Panel></td>
                                --%>
                            </tr>
                        </table>
                        <br />
                        <table class="table_sch" cellspacing="1" cellpadding="1">
                            <tr>
                                <td class="head_navy">訪查項目</td>
                                <td class="head_navy">訪查實況</td>
                                <td class="head_navy">備註</td>
                            </tr>
                            <tr>
                                <td class="bluecol">一、各項訓練計畫訪查<br>
                                    紀錄是否確實登錄<br>
                                    OJT 系統及訪查<br>
                                    次數是否符合計畫<br>
                                    訪查比率
                                </td>
                                <td class="whitecol">
                                    <input id="rdoQ11" value="1" type="radio" name="rdoQ1" runat="server">符合
                                    <input id="rdoQ10" value="0" type="radio" name="rdoQ1" runat="server">
                                    不符合<br>
                                    訪查率
                                    <asp:TextBox ID="txtQ1Per" runat="server" onfocus="this.blur()" MaxLength="5" Columns="3"></asp:TextBox>％ (=
                                    <asp:TextBox ID="txtQ1A" runat="server" MaxLength="5" Columns="2"></asp:TextBox>/
                                    <asp:TextBox ID="txtQ1B" runat="server" MaxLength="5" Columns="2"></asp:TextBox>)
                                </td>
                                <td class="whitecol" align="center">訪查比率(次數)規定說明：<br>
                                    <asp:TextBox ID="txtQ1Note" runat="server" MaxLength="100" Rows="4" TextMode="MultiLine" Width="70%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol">二、訪查紀錄表是否確<br>
                                    實填報</td>
                                <td class="whitecol">
                                    <input id="rdoQ21" value="1" type="radio" name="rdoQ2" runat="server">是
                                    <input id="rdoQ20" value="0" type="radio" name="rdoQ2" runat="server">否<br>
                                    <input id="rdoQ22" value="2" type="radio" name="rdoQ2" runat="server">其他(請說明)：
                                    <asp:TextBox ID="txtQ2Other" runat="server" MaxLength="100"></asp:TextBox>
                                </td>
                                <td class="whitecol" align="center">
                                    <asp:TextBox ID="txtQ2Note" runat="server" MaxLength="100" Rows="4" TextMode="MultiLine" Width="70%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol">三、訪查項目未合格之<br>
                                    訓練單位是否依規<br>
                                    定進行後續追蹤訪<br>
                                    查作業並將結果確<br>
                                    實登入 OJT 系統
                                </td>
                                <td class="whitecol">
                                    <input id="rdoQ31" value="1" type="radio" name="rdoQ3" runat="server">是
                                    <input id="rdoQ30" value="0" type="radio" name="rdoQ3" runat="server">否<br>
                                    <input id="rdoQ32" value="2" type="radio" name="rdoQ3" runat="server">其他(請說明)：
                                    <asp:TextBox ID="txtQ3Other" runat="server" MaxLength="100"></asp:TextBox>
                                </td>
                                <td class="whitecol" align="center">
                                    <asp:TextBox ID="txtQ3Note" runat="server" MaxLength="100" Rows="4" TextMode="MultiLine" Width="70%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol">四、訪查成果及相關書<br>
                                    面資料是否依規定<br>
                                    按季陳核單位主管<br>
                                    ，並妥善歸類保管
                                </td>
                                <td class="whitecol">
                                    <input id="rdoQ41" value="1" type="radio" name="rdoQ4" runat="server">是
                                    <input id="rdoQ40" value="0" type="radio" name="rdoQ4" runat="server">否<br>
                                    <input id="rdoQ42" value="2" type="radio" name="rdoQ4" runat="server">其他(請說明)：
                                    <asp:TextBox ID="txtQ4Other" runat="server" MaxLength="100"></asp:TextBox>
                                </td>
                                <td class="whitecol" align="center">
                                    <asp:TextBox ID="txtQ4Note" runat="server" MaxLength="100" Rows="4" TextMode="MultiLine" Width="70%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol">五、辦理申領符合職業<br>
                                    訓練生活津貼補助<br>
                                    是否確實
                                </td>
                                <td class="whitecol">
                                    <input id="rdoQ51" value="1" type="radio" name="rdoQ5" runat="server">是
                                    <input id="rdoQ50" value="0" type="radio" name="rdoQ5" runat="server">否<br>
                                    <input id="rdoQ53" value="3" type="radio" name="rdoQ5" runat="server">本計畫無申請職訓生活津貼情事<br>
                                    <input id="rdoQ52" value="2" type="radio" name="rdoQ5" runat="server">其他(請說明)：
                                    <asp:TextBox ID="txtQ5Other" runat="server" MaxLength="100"></asp:TextBox>
                                </td>
                                <td class="whitecol" align="center">
                                    <asp:TextBox ID="txtQ5Note" runat="server" MaxLength="100" Rows="4" TextMode="MultiLine" Width="70%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol">六、其他<asp:TextBox ID="txtQ6" runat="server" MaxLength="100" Columns="10"></asp:TextBox></td>
                                <td colspan="2" class="whitecol">
                                    <asp:TextBox ID="txtQ6Other" runat="server" MaxLength="100" Columns="40" Rows="4" TextMode="MultiLine" Width="50%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol">七、建議事項</td>
                                <td colspan="2" class="whitecol">
                                    <asp:TextBox ID="txtQ7Note" runat="server" MaxLength="100" Columns="40" Rows="4" TextMode="MultiLine" Width="50%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol">綜合意見</td>
                                <td colspan="2" class="whitecol">
                                    <asp:TextBox ID="txtQ8Note" runat="server" MaxLength="100" Columns="40" Rows="4" TextMode="MultiLine" Width="50%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol">受訪單位人員姓名</td>
                                <td colspan="2" class="whitecol">
                                    <asp:TextBox ID="txtUnitName" runat="server" MaxLength="10" Columns="10" Width="20%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol">訪查人員姓名</td>
                                <td colspan="2" class="whitecol">
                                    <asp:TextBox ID="txtFillerName" runat="server" MaxLength="10" Columns="10" Width="20%"></asp:TextBox></td>
                            </tr>
                        </table>
                        <div align="center" class="whitecol">
                            <asp:Button ID="btnSave" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                            <input id="btnReset" value="重填" type="reset" class="asp_button_M" />
                            <asp:Button ID="btnExit" runat="server" Text="離開" CssClass="asp_button_M"></asp:Button>
                        </div>
                    </td>
                </tr>
            </table>
        </asp:Panel>
                                    <input id="hidEditOrgID" type="hidden" name="hidEditOrgID" runat="server">
                                    <input id="hidEditPlanID" type="hidden" name="PlanIDValue" runat="server">
    </form>
</body>
</html>
