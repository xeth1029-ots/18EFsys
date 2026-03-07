<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_01_001_add_06.aspx.vb" Inherits="WDAIIP.CP_01_001_add_06" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>實地訪查紀錄表</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <%--
                    <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
					    <tr>
						    <td>首頁&gt;&gt;訓練查核與績效管理&gt;&gt;統計分析&gt;&gt;實地訪查紀錄表</td>
					    </tr>
				    </table>
                    --%>
                    <table class="table_sch" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="bluecol" style="width: 20%">機構 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="center" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox>
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini" />
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">職類/班別 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input id="Button3" onclick="javascript: window.open('../CP_01_ch.aspx?RID=' + document.form1.RIDValue.value, '', 'width=540,height=520,location=0,status=0,menubar=0,scrollbars=1,resizable=0');" type="button" value="..." name="Button3" runat="server" class="button_b_Mini" />
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                <span id="HistoryList" style="position: absolute; display: none; left: 28%">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">訪查時間 </td>
                            <td class="whitecol" colspan="3">
                                <span id="span01" runat="server">
                                    <asp:TextBox ID="APPLYDATE" runat="server" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= APPLYDATE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                                </span>
                                <asp:TextBox ID="APPLYDATEHH1" runat="server" Width="10%" MaxLength="2"></asp:TextBox>時
                                <asp:TextBox ID="APPLYDATEMI1" runat="server" Width="10%" MaxLength="2"></asp:TextBox>分至
                                <asp:TextBox ID="APPLYDATEHH2" runat="server" Width="10%" MaxLength="2"></asp:TextBox>時
                                <asp:TextBox ID="APPLYDATEMI2" runat="server" Width="10%" MaxLength="2"></asp:TextBox>分
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" colspan="4">訪視次數至少<asp:Label ID="LabVISCOUNT" runat="server" Text="N"></asp:Label>次，本次為第
                                <asp:TextBox ID="VISTIMES" runat="server" MaxLength="2" Width="10%" CssClass="whitecol"></asp:TextBox>次訪視
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4">
                                <table class="table_sch" cellspacing="1" cellpadding="1" width="100%" border="1">
                                    <tr align="center">
                                        <td class="bluecol" width="20%">書面資料</td>
                                        <td class="bluecol" width="25%"></td>
                                        <td class="bluecol" width="25%">佐證資料及說明</td>
                                        <td class="bluecol" width="30%">備註</td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">1.教學(訓練)日誌</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="DATA1" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="2">
                                                <asp:ListItem Value="1">備齊</asp:ListItem>
                                                <asp:ListItem Value="3">部份備有</asp:ListItem>
                                                <asp:ListItem Value="2">未備</asp:ListItem>
                                                <asp:ListItem Value="4">免提供</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td class="whitecol">如附件<asp:TextBox ID="DATACOPY1" runat="server" MaxLength="50" Width="60%"></asp:TextBox>
                                        </td>
                                        <td class="whitecol">備齊<asp:TextBox ID="D1CMM" runat="server" MaxLength="2" Width="10%"></asp:TextBox>月<asp:TextBox ID="D1CDD" runat="server" MaxLength="9" Width="15%"></asp:TextBox>日資料 </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">2.學員簽到(退)表</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="DATA2" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="2">
                                                <asp:ListItem Value="1">備齊</asp:ListItem>
                                                <asp:ListItem Value="3">部份備有</asp:ListItem>
                                                <asp:ListItem Value="2">未備</asp:ListItem>
                                                <asp:ListItem Value="4">免提供</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td class="whitecol">如附件<asp:TextBox ID="DATACOPY2" runat="server" MaxLength="50" Width="60%"></asp:TextBox></td>
                                        <td class="whitecol">備齊<asp:TextBox ID="D2CMM" runat="server" MaxLength="2" Width="10%"></asp:TextBox>月<asp:TextBox ID="D2CDD" runat="server" MaxLength="9" Width="15%"></asp:TextBox>日資料 </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">3.請假單</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="DATA3" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="2">
                                                <asp:ListItem Value="1">備齊</asp:ListItem>
                                                <asp:ListItem Value="3">部份備有</asp:ListItem>
                                                <asp:ListItem Value="2">未備</asp:ListItem>
                                                <asp:ListItem Value="4">免提供</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td class="whitecol">如附件<asp:TextBox ID="DATACOPY3" runat="server" MaxLength="50" Width="60%"></asp:TextBox></td>
                                        <td class="whitecol">
                                            <asp:RadioButton ID="D3C1" runat="server" GroupName="D3C" />攜回
                                            <asp:TextBox ID="D3CMM" runat="server" MaxLength="2" Width="10%"></asp:TextBox>月
                                            <asp:TextBox ID="D3CDD" runat="server" MaxLength="9" Width="15%"></asp:TextBox>日課程請假單影本<br />
                                            <asp:RadioButton ID="D3C2" runat="server" GroupName="D3C" />無學員請假情形，故免提供<br />
                                            <asp:RadioButton ID="D3C3" runat="server" GroupName="D3C" />其他(請說明)：<br />
                                            <asp:TextBox ID="D3NOTE" runat="server" MaxLength="100" Width="50%"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">4.出缺勤狀況</td>
                                        <td id="notstudy3" colspan="3" runat="server" class="whitecol">核定：<asp:TextBox ID="APPROVEDCOUNT" runat="server" MaxLength="3" Width="10%"></asp:TextBox>人；
                                            開訓：<asp:TextBox ID="AUTHCOUNT" runat="server" MaxLength="3" Width="10%"></asp:TextBox>人； 
                                            實到：<asp:TextBox ID="TURTHCOUNT" runat="server" MaxLength="3" Width="10%"></asp:TextBox>人； 
                                            請假：<asp:TextBox ID="TURNOUTCOUNT" runat="server" MaxLength="3" Width="10%"></asp:TextBox>人；<br />
                                            缺(曠)課：<asp:TextBox ID="TRUANCYCOUNT" runat="server" MaxLength="3" Width="10%"></asp:TextBox>人；
                                            離訓：<asp:TextBox ID="LEAVECOUNT" runat="server" MaxLength="3" Width="10%"></asp:TextBox>人； 
                                            退訓：<asp:TextBox ID="REJECTCOUNT" runat="server" MaxLength="3" Width="10%"></asp:TextBox>人
                                            <%--；(含提前就業：<asp:TextBox ID="ADVJOBCOUNT" runat="server" MaxLength="3" Width="10%"></asp:TextBox>人)。--%>
                                            <%--<br />※點名未到課學員，應於訪查次日起三日內另以電話抽訪。 --%></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4">
                                <table class="table_sch" cellspacing="1" cellpadding="1" width="100%" border="1">
                                    <tr>
                                        <td colspan="2" class="bluecol">訪查項目 </td>
                                        <td class="bluecol">訪查實況 </td>
                                        <td class="bluecol">處理情形 </td>
                                        <td class="bluecol">備註 </td>
                                    </tr>
                                    <tr>
                                        <td rowspan="4" class="bluecol">課程(師資)實施狀況 </td>
                                        <td class="bluecol">1.有無週(月)課程表? </td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="ITEM1_1" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">有</asp:ListItem>
                                                <asp:ListItem Value="2">無</asp:ListItem>
                                                <asp:ListItem Value="3">免填</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td rowspan="4" class="whitecol">
                                            <asp:TextBox ID="ITEM1PROS" runat="server" TextMode="MultiLine" Height="100px"></asp:TextBox></td>
                                        <td rowspan="4" class="whitecol" align="center">
                                            <asp:TextBox ID="ITEM1NOTE" runat="server" TextMode="MultiLine" Height="100px"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">2.是否依課程表授課?</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="ITEM1_2" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">是</asp:ListItem>
                                                <asp:ListItem Value="2">否</asp:ListItem>
                                                <asp:ListItem Value="3">免填</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">3.課目或課題為何?</td>
                                        <td class="whitecol">課目：<asp:TextBox ID="ITEM1_COUR" runat="server" MaxLength="100" Width="60%"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">4.教師是否與計畫相符?</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="ITEM1_3" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">是</asp:ListItem>
                                                <asp:ListItem Value="2">否</asp:ListItem>
                                                <asp:ListItem Value="3">免填</asp:ListItem>
                                            </asp:RadioButtonList>
                                            教師1：<asp:TextBox ID="ITEM1_TEACHER" runat="server" MaxLength="100" Width="60%"></asp:TextBox><br />
                                            教師2：<asp:TextBox ID="ITEM1_ASSISTANT" runat="server" MaxLength="100" Width="60%"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" rowspan="2">教材設施運用狀況 </td>
                                        <%-- <td class="whitecol" colspan="2"><table cellspacing="1" cellpadding="1" width="100%" border="0"><tr></tr></table></td>--%>
                                        <td class="bluecol">1.有無書籍(講義)領用表? </td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="ITEM2_1" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">有</asp:ListItem>
                                                <asp:ListItem Value="2">無</asp:ListItem>
                                                <asp:ListItem Value="3">免填</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td class="whitecol" rowspan="2">
                                            <asp:TextBox ID="ITEM2PROS" runat="server" TextMode="MultiLine" Height="100px"></asp:TextBox></td>
                                        <td class="whitecol" align="center" rowspan="2">
                                            <asp:TextBox ID="ITEM2NOTE" runat="server" TextMode="MultiLine" Height="100px"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">2.有無材料領用表?</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="ITEM2_2" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">有</asp:ListItem>
                                                <asp:ListItem Value="2">無</asp:ListItem>
                                                <asp:ListItem Value="3">免填</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <%--<tr>
                                                <td class="whitecol">3.訓練設施設備是否依契約提供學員使用?</td>
                                                <td class="whitecol">
                                                <asp:RadioButtonList ID="ITEM2_3" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">是</asp:ListItem>
                                                <asp:ListItem Value="2">否</asp:ListItem>
                                                <asp:ListItem Value="3">免填</asp:ListItem>
                                                </asp:RadioButtonList>
                                                </td></tr>--%>

                                    <tr>
                                        <td rowspan="2" class="bluecol">教務管理狀況 </td>
                                        <td class="bluecol">1.教學(訓練)日誌是否確實填寫? </td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="ITEM3_1" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">是</asp:ListItem>
                                                <asp:ListItem Value="2">否</asp:ListItem>
                                                <asp:ListItem Value="3">免填</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td rowspan="2" class="whitecol">
                                            <asp:TextBox ID="ITEM3PROS" runat="server" TextMode="MultiLine" Height="100px"></asp:TextBox></td>
                                        <td rowspan="2" class="whitecol" align="center">
                                            <asp:TextBox ID="ITEM3NOTE" runat="server" TextMode="MultiLine" Height="100px"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">2.是否按時呈主管核閱?</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="ITEM3_2" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">是</asp:ListItem>
                                                <asp:ListItem Value="2">否</asp:ListItem>
                                                <asp:ListItem Value="3">免填</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">參訓學員反映意見及問題 </td>
                                        <td colspan="4" class="whitecol">
                                            <asp:TextBox ID="ITEM7NOTE" runat="server" Width="77%" TextMode="MultiLine" Height="80px"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">學員反映意見之訓練單位回應說明 </td>
                                        <td colspan="4" class="whitecol">
                                            <asp:TextBox ID="ITEM7NOTE2" runat="server" Width="77%" TextMode="MultiLine" Height="80px"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">訪查單位綜合建議</td>
                                        <td colspan="2" class="whitecol">
                                            <asp:TextBox ID="ITEM31NOTE" runat="server" Width="77%" TextMode="MultiLine" Height="80px"></asp:TextBox></td>
                                        <td class="whitecol" colspan="2">
                                            <table cellspacing="1" cellpadding="1" width="100%" border="0">
                                                <tr>
                                                    <td class="bluecol">訓練單位缺失處理</td>
                                                    <td class="whitecol">
                                                        <asp:RadioButtonList ID="ITEM32" runat="server" CssClass="font">
                                                            <asp:ListItem Value="4">無缺失</asp:ListItem>
                                                            <asp:ListItem Value="1">限期改善，研提檢討報告</asp:ListItem>
                                                            <asp:ListItem Value="2">擇期進行訪查</asp:ListItem>
                                                            <asp:ListItem Value="3">其他(請說明)：</asp:ListItem>
                                                        </asp:RadioButtonList>
                                                        <asp:TextBox ID="ITEM32NOTE" runat="server" Width="90%" MaxLength="100"></asp:TextBox>
                                                        <br />
                                                        (可輸入文字長度為100個中文字)
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4">
                                <table class="table_sch" cellspacing="1" cellpadding="1" width="100%" border="1">
                                    <tr>
                                        <%--<td class="bluecol" width="20%">培訓姓名 </td>
                                        <td class="whitecol" width="30%">
                                            <asp:TextBox ID="CURSENAME" runat="server" MaxLength="10" Width="70%"></asp:TextBox></td>--%>
                                        <td class="bluecol" width="20%">訪查人員 </td>
                                        <td class="whitecol" width="80%" colspan="3">
                                            <asp:TextBox ID="VISITORNAME" runat="server" MaxLength="10" Width="50%"></asp:TextBox></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                    <div align="center" class="whitecol">
                        <asp:Button ID="Button1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="Button4" runat="server" Text="回查詢頁面" CssClass="button_b_M"></asp:Button>
                    </div>
                </td>
            </tr>
        </table>
        <input id="HIDTPlanID" type="hidden" name="HIDTPlanID" runat="server" />
        <input id="EndDate" type="hidden" name="EndDate" runat="server" />
        <input id="StartDate" type="hidden" name="StartDate" runat="server" />
        <input id="NowDate" type="hidden" name="NowDate" runat="server" />
    </form>
</body>
</html>
