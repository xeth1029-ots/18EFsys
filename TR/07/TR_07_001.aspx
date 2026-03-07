<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_07_001.aspx.vb" Inherits="WDAIIP.TR_07_001" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>年度執行成效</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        //選擇全部
        function SelectAll(obj, hidobj) {
            var num = getCheckBoxListValue(obj).length;
            var myallcheck = document.getElementById(obj + '_' + 0);
            //var smsg1 = "num:" + num + ", myallcheck:" + myallcheck; alert(smsg1);return false;
            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
            else {
                for (var i = 1; i < num; i++) {
                    if ('0' == getCheckBoxListValue(obj).charAt(i)) {
                        document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(i);
                        var mycheck = document.getElementById(obj + '_' + i);
                        myallcheck.checked = mycheck.checked;
                        break;
                    }
                }
            }
        }

 
        //$(document).ready(function () {

        //    var isDisabled_TB_Review90 = $('TB_Review90').di.is('[disabled="disabled"]');
        //    var isDisabled_TB_Review10 = $('TB_Review10').is('[disabled="disabled"]');
        //    var isDisabled_TB_ReviewNG = $('TB_ReviewNG').is('[disabled="disabled"]');
        //    console.log("isDisabled_TB_Review90:" + isDisabled_TB_Review90);
        //    console.log("isDisabled_TB_Review10:" + isDisabled_TB_Review10);
        //    console.log("isDisabled_TB_ReviewNG:" + isDisabled_TB_ReviewNG);
        //    //debugger;
        //    $('td_TB_Review90').css(isDisabled_TB_Review90 ? "bluecol" : "bluecol_need");
        //    $('td_TB_Review10').css(isDisabled_TB_Review10 ? "bluecol" : "bluecol_need");
        //    $('td_TB_ReviewNG').css(isDisabled_TB_ReviewNG ? "bluecol" : "bluecol_need");
        //});
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練需求管理&gt;&gt;年度執行成效</asp:Label>
                </td>
            </tr>
        </table>
        <table id="SearchTable" runat="server" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch" width="100%" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol_need" width="20%">年度</td>
                            <td class="whitecol" width="80%">
                                <asp:DropDownList ID="yearlist" runat="server" AutoPostBack="True">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">轄區</td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="cblDistid" runat="server" RepeatDirection="Horizontal" RepeatColumns="3" CssClass="whitecol">
                                </asp:CheckBoxList>
                                <input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">開訓期間</td>
                            <td class="whitecol">
                                <asp:TextBox ID="STDate1" runat="server" Columns="10" MaxLength="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>～
								<asp:TextBox ID="STDate2" runat="server" Columns="10" MaxLength="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">結訓期間</td>
                            <td class="whitecol">
                                <asp:TextBox ID="FTDate1" runat="server" Columns="10" MaxLength="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>～
								<asp:TextBox ID="FTDate2" runat="server" Columns="10" MaxLength="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>

                        <tr>
                            <td class="bluecol" width="20%">匯出檔案格式</td>
                            <td class="whitecol" width="80%">
                                <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                                    <asp:ListItem Value="ODS">ODS</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol" colspan="2">
                                <div align="center">
                                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                    <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                    <asp:Button ID="btnQuery" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="btnExport1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                                </div>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="DataGrid1Table" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td class="whitecol">
                                <asp:Label ID="labmsg_t1" runat="server" ForeColor="Red"></asp:Label></td>
                        </tr>
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AllowSorting="True" PagerStyle-HorizontalAlign="Left" PagerStyle-Mode="NumericPages" AllowPaging="True" AutoGenerateColumns="False">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <%--<asp:BoundColumn HeaderText="序號" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>--%>
                                        <asp:TemplateColumn HeaderText="序號">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="LabStar1" runat="server" ForeColor="Red">*</asp:Label><%--'*表示該班有以下情形之一：(1)開訓人數比率未達90%、(2)離退訓率超過10%(含)、(3)不開班--%>
                                                <asp:Label ID="LabSeqno" runat="server"></asp:Label><%--序號--%>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="DISTNAME" HeaderText="轄區" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="CLASSCNAME2" HeaderText="班別名稱" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="TRAINNAME" HeaderText="訓練職類" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDATE" HeaderText="開訓日期" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="FTDATE" HeaderText="結訓日期" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="TNUM" HeaderText="預訓人數" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="STUDETNUM" HeaderText="報名人數" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="STUDETNUM2" HeaderText="甄試人數" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="STUDETNUM3" HeaderText="錄訓人數" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="SNum1" HeaderText="開訓人數" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="ESNum1" HeaderText="結訓人數" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="JSNum1" HeaderText="離退訓人數" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="AVERAGE" HeaderText="滿意度(%)" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <ItemStyle HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="lbtEdit1" runat="server" Text="修改" CommandName="Edit1" CssClass="linkbutton"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
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
                <td align="center" class="whitecol">
                    <asp:Label ID="msg1" runat="server" CssClass="font" ForeColor="Red"></asp:Label></td>
            </tr>
        </table>
        <%--序號、班別名稱、訓練職類、政策性課程類型、開訓日期、結訓日期、預訓人數/訓練人數、報名人數、甄試人數、錄訓人數、
            錄訓率、開訓人數、開訓人數比率、結訓人數、離退訓人數、離退訓率、滿意度、開訓人數比率未達90%之檢討改善、
            離退訓率超過10%之檢討改善、不開班原因、不開班之檢討措施、其他。--%>
        <table id="DetailTable" runat="server" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="Table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td class="bluecol" width="20%">班別名稱 </td>
                            <td class="whitecol" width="80%" colspan="3">
                                <asp:Label ID="labCLASSCNAME2" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練職類 </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="labTRAINNAME" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">政策性課程類型 </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="labD20KNAME" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">開訓日期 </td>
                            <td class="whitecol">
                                <asp:Label ID="labSTDATE" runat="server"></asp:Label></td>
                            <td class="bluecol" width="20%">結訓日期 </td>
                            <td class="whitecol" width="30%">
                                <asp:Label ID="labFTDATE" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">預訓人數<%--訓練人數--%></td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="labTNUM" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">報名人數 </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="labSTUDETNUM" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">甄試人數 </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="labSTUDETNUM2" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">錄訓人數 </td>
                            <td class="whitecol">
                                <asp:Label ID="labSTUDETNUM3" runat="server"></asp:Label></td>
                            <td class="bluecol">錄訓率(%) </td>
                            <td class="whitecol">
                                <asp:Label ID="labACCEPRATE" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">開訓人數 </td>
                            <td class="whitecol">
                                <asp:Label ID="labSNum1" runat="server"></asp:Label></td>
                            <td class="bluecol">開訓人數比率(%) </td>
                            <td class="whitecol">
                                <asp:Label ID="labTRAINRATE" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" id="td_TB_Review90" runat="server">開訓人數比率未達90%之檢討改善 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="TB_Review90" runat="server" TextMode="MultiLine" Rows="5" Width="77%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">結訓人數 </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="labESNum1" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">離退訓人數 </td>
                            <td class="whitecol">
                                <asp:Label ID="labJSNum1" runat="server"></asp:Label></td>
                            <td class="bluecol">離退訓率(%) </td>
                            <td class="whitecol">
                                <asp:Label ID="labRTIRERATE" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" id="td_TB_Review10" runat="server">離退訓率超過10%之檢討改善 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="TB_Review10" runat="server" TextMode="MultiLine" Rows="5" Width="77%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">滿意度(%) </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="labAVERAGE" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol" >不開班原因</td>
                            <%--TC_01_004_add--%>
                            <td colspan="3" class="whitecol">
                                <asp:CheckBoxList ID="NORID" runat="server" RepeatDirection="Horizontal" CellSpacing="0" CellPadding="0" RepeatLayout="Flow"></asp:CheckBoxList>
                                <asp:TextBox ID="OtherReason" runat="server" Width="50%"></asp:TextBox>
                                <input id="NORIDValue" type="hidden" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" id="td_TB_ReviewNG" runat="server">不開班之檢討措施 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="TB_ReviewNG" runat="server" TextMode="MultiLine" Rows="5" Width="77%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">其他執行說明及檢討改善(非必填) </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="TB_ReviewOth" runat="server" TextMode="MultiLine" Rows="5" Width="77%"></asp:TextBox></td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <asp:Button ID="BtnSaveData1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                    <asp:Button ID="BtnBack1" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="Hid_PLANID" runat="server" />
        <asp:HiddenField ID="Hid_COMIDNO" runat="server" />
        <asp:HiddenField ID="Hid_SEQNO" runat="server" />
        <asp:HiddenField ID="Hid_OCID1" runat="server" />
        <asp:HiddenField ID="hid_TRAINRATE" runat="server" />
        <asp:HiddenField ID="hid_RTIRERATE" runat="server" />
    </form>
</body>
</html>
