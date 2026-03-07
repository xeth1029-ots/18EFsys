<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_014_add.aspx.vb" Inherits="WDAIIP.TC_01_014_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>開班計畫表資料維護(產業人才專用)</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="Javascript">
        function GoReadOnly1() {
            var tPlanCause = document.getElementById('tPlanCause');
            var tPurScience = document.getElementById('tPurScience');
            var tPurTech = document.getElementById('tPurTech');
            var tPurMoral = document.getElementById('tPurMoral');
            INPUT_readOnly(tPlanCause);
            INPUT_readOnly(tPurScience);
            INPUT_readOnly(tPurTech);
            INPUT_readOnly(tPurMoral);
        }

        function setStyle(object, styleText) {
            if (object.style.setAttribute) {
                object.style.setAttribute("cssText", styleText);
            }
            else {
                object.setAttribute("style", styleText);
            }
        }

        function setDiv(object) {
            var result = object; // 取得div元素
            result.innerHTML = "";
        }

        function showPanel() {
            if (document.form1.Tnum2.value == '') {
                document.getElementById('TRA1').style.display = 'none';
                document.getElementById('TRA2').style.display = 'none';
            }
            if (document.form1.Tnum3.value == '') {
                document.getElementById('TRB1').style.display = 'none';
                document.getElementById('TRB2').style.display = 'none';
            }
            //if (!_isIE) { parent.setMainFrameHeight(); } //重新調整頁面高度
            if (parent && parent.setMainFrameHeight != undefined) { parent.setMainFrameHeight(); }
        }

        function lock1() {
            var Button6 = document.getElementById('Button6');
            var bt_addrow = document.getElementById('bt_addrow');
            if (!Button6 && !bt_addrow) {
                $('input[type=text],textarea').prop('readonly', true);
                $('input[type=checkbox]').attr("disabled", true);
                $('input[type=radio]').attr("disabled", true);
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <%-- <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td class="font">
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;開班計畫表資料維護作業</asp:Label>
                </td>
            </tr>
        </table>--%>
        <table class="table_nw" id="Table_T_D" width="100%" runat="server">
            <tr>
                <td id="Td13" colspan="2" runat="server" class="bluecol" width="20%">年度 </td>
                <td colspan="6" class="whitecol">
                    <asp:DropDownList ID="PlanYear" Enabled="False" runat="server"></asp:DropDownList>
                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                </td>
            </tr>
            <tr>
                <td id="Td15" colspan="2" runat="server" class="bluecol" width="20%">課程名稱 </td>
                <td colspan="2" class="whitecol" width="30%">
                    <asp:TextBox ID="ClassName" runat="server" onfocus="this.blur()" Columns="30" MaxLength="30" Width="80%"></asp:TextBox></td>
                <td id="Td16" colspan="2" runat="server" class="bluecol" width="20%">訓練職能 </td>
                <td colspan="2" class="whitecol" width="30%">
                    <asp:DropDownList ID="ClassCate" runat="server"></asp:DropDownList></td>
            </tr>
            <tr>
                <td id="Td17" colspan="2" runat="server" class="bluecol" width="20%">課程班別 </td>
                <td colspan="6" class="whitecol" width="30%">
                    <asp:DropDownList ID="ClassID" runat="server"></asp:DropDownList></td>
            </tr>
            <tr>
                <td id="Td1" colspan="2" runat="server" class="bluecol" width="20%">訓練人數 </td>
                <td id="Td2" align="left" colspan="2" runat="server" class="whitecol" width="30%">
                    <asp:TextBox ID="Tnum" runat="server" onfocus="this.blur()" MaxLength="5" Width="30%"></asp:TextBox></td>
                <td id="Td3" colspan="2" runat="server" class="bluecol" width="20%">&nbsp;&nbsp; 訓練時數 </td>
                <td id="Td4" align="left" runat="server" class="whitecol" width="30%">
                    <asp:TextBox ID="THours" runat="server" onfocus="this.blur()" MaxLength="5" Width="30%"></asp:TextBox></td>
            </tr>
            <tr>
                <td colspan="8" width="100%" class="whitecol">
                    <table class="font" id="Datagrid3Table" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td class="table_title" align="center" width="100%">教學方法 </td>
                        </tr>
                        <tr>
                            <td class="whitecol" width="100%">
                                <asp:CheckBoxList ID="cblTMethod" runat="server" CssClass="font" RepeatLayout="Flow">
                                    <asp:ListItem Value="01">講授教學法（運用敘述或講演的方式，傳遞教材知識的一種教學方法，提供相關教材或講義）</asp:ListItem>
                                    <asp:ListItem Value="02">討論教學法（指團體成員齊聚一起，經由說、聽和觀察的過程，彼此溝通意見，由講師帶領達成教學目標）</asp:ListItem>
                                    <asp:ListItem Value="03">演練教學法（由講師的帶領下透過設備或教材，進行練習、表現和實作，親自解說示範的技能或程序的一種教學方法）</asp:ListItem>
                                    <asp:ListItem Value="99">其他教學方法：</asp:ListItem>
                                </asp:CheckBoxList>
                                <asp:TextBox runat="server" ID="TMethodOth" Width="40%" MaxLength="100"></asp:TextBox><br />
                                <asp:Label ID="Label1" runat="server" ForeColor="red">(若選"其他教學方法"，需填寫輸入，上限100個字)</asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="table_title">課程大綱 </td>
                        </tr>
                        <tr>
                            <td class="whitecol">
                                <asp:DataGrid ID="Datagrid3" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" BackColor="#FFFFFF" CellPadding="8">
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <Columns>
                                        <asp:TemplateColumn HeaderText="日期">
                                            <ItemStyle HorizontalAlign="Center" Width="10%" />
                                            <ItemTemplate>
                                                <asp:Label ID="STrainDateLabel" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="授課時間">
                                            <ItemStyle HorizontalAlign="Center" Width="12%" />
                                            <ItemTemplate>
                                                <asp:Label ID="PNameLabel" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="時數">
                                            <ItemStyle HorizontalAlign="Center" Width="6%" />
                                            <ItemTemplate>
                                                <asp:Label ID="PHourLabel" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="課程進度／內容">
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:TextBox ID="PContText" runat="server" onfocus="this.blur()" Width="80%" Columns="50" TextMode="MultiLine" Rows="5" Enabled="False" Height="58px"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="學／術科">
                                            <ItemStyle HorizontalAlign="Center" Width="8%" />
                                            <ItemTemplate>
                                                <asp:DropDownList ID="drpClassification1" runat="server" Enabled="False" AutoPostBack="True">
                                                    <asp:ListItem Value="1">學科</asp:ListItem>
                                                    <asp:ListItem Value="2">術科</asp:ListItem>
                                                </asp:DropDownList>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="上課地點">
                                            <ItemStyle HorizontalAlign="Center" Width="18%" />
                                            <ItemTemplate>
                                                <asp:DropDownList ID="drpPTID" runat="server" Enabled="False"></asp:DropDownList>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="授課師資">
                                            <ItemStyle HorizontalAlign="Center" Width="12%" />
                                            <ItemTemplate>
                                                <input id="Tech1Value" type="hidden" name="Tech1Value" runat="server" />
                                                <asp:TextBox ID="Tech1Text" runat="server" onfocus="this.blur()" Columns="10" Enabled="False" Width="80%"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="助教">
                                            <ItemStyle HorizontalAlign="Center" Width="12%" />
                                            <ItemTemplate>
                                                <input id="Tech2Value" type="hidden" name="Tech2Value" runat="server">
                                                <asp:TextBox ID="Tech2Text" runat="server" onfocus="this.blur()" Columns="10" Enabled="False" Width="80%"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td colspan="8" class="whitecol">
                    <table class="font" id="DataGrid1Table" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td align="center" class="table_title">上課時間 </td>
                        </tr>
                        <tr>
                            <td colspan="7" class="whitecol">
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" BackColor="#FFFFFF" CellPadding="8">
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <Columns>
                                        <asp:TemplateColumn HeaderText="星期">
                                            <HeaderStyle Width="20%" />
                                            <ItemTemplate>
                                                <asp:Label ID="Weeks1" runat="server"></asp:Label>
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:DropDownList ID="Weeks2" runat="server">
                                                </asp:DropDownList>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="上課時段">
                                            <HeaderStyle Width="80%" />
                                            <ItemTemplate>
                                                <asp:Label ID="Times1" runat="server"></asp:Label>
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox ID="Times2" runat="server"></asp:TextBox>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn Visible="False" HeaderText="功能">
                                            <ItemTemplate>
                                                <asp:Button ID="Button2" runat="server" Text="修改" CausesValidation="False" CommandName="edit" Visible="False" CssClass="asp_button_M"></asp:Button>
                                                <asp:Button ID="Button3" runat="server" Text="刪除" CausesValidation="False" CommandName="del" Visible="False" CssClass="asp_button_M"></asp:Button>
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:Button ID="Button4" runat="server" Text="儲存" CausesValidation="False" CommandName="save" Visible="False" CssClass="asp_button_M"></asp:Button>
                                                <asp:Button ID="Button5" runat="server" Text="取消" CausesValidation="False" CommandName="cancel" Visible="False" CssClass="asp_button_M"></asp:Button>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td id="Td8" bgcolor="#CCD8EE" colspan="2" runat="server" class="bluecol">起迄日期 </td>
                <td colspan="6" class="whitecol">
                    <asp:TextBox ID="start_date" runat="server" onfocus="this.blur()" Width="16%"></asp:TextBox>
                    <span runat="server">
                        <img id="IMG1" style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30"></span>～
                    <asp:TextBox ID="end_date" runat="server" onfocus="this.blur()" Width="16%"></asp:TextBox>
                    <span runat="server">
                        <img id="IMG2" style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30"></span>
                    <asp:CompareValidator ID="comdate" runat="server" ControlToCompare="start_date" Operator="GreaterThan" Type="Date" ControlToValidate="end_date" Display="None" ErrorMessage="迄日不得小於起日或日期格式不正確"></asp:CompareValidator>
                    <asp:RequiredFieldValidator ID="mustsdate" runat="server" ControlToValidate="start_date" Display="None" ErrorMessage="請輸入起日"></asp:RequiredFieldValidator>
                    <asp:RequiredFieldValidator ID="mustedate" runat="server" ControlToValidate="end_date" Display="None" ErrorMessage="請輸入迄日"></asp:RequiredFieldValidator>
                    <asp:ValidationSummary ID="totalmsg" runat="server" DisplayMode="List" ShowMessageBox="True" ShowSummary="False"></asp:ValidationSummary>
                </td>
            </tr>
        </table>
        <table class="font" id="Table_1" style="width: 100%;" cellspacing="1" cellpadding="1" border="0">
            <tr>
                <td bgcolor="#FFCC00">一、規劃與執行能力 </td>
            </tr>
        </table>
        <table class="font" id="Table_1_D" style="width: 100%;" cellspacing="1" cellpadding="1" border="0">
            <tr>
                <td id="TD7" colspan="2" runat="server" class="bluecol" width="16%">訓練需求調查 </td>
                <td colspan="4" width="84%">
                    <table class="font" id="tbTrainDemain" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>產業人力需求調查：<asp:Label ID="Label9" runat="server" ForeColor="red">(填寫輸入，上限1000個字)</asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol">
                                <asp:TextBox ID="tPOWERNEED1" runat="server" Width="60%" Height="80px" TextMode="MultiLine" MaxLength="1000" placeholder="(應論述調查期間、區域範圍、調查對象、產業發展趨勢及該產業之訓練需求)"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td>區域人力需求調查：<asp:Label ID="Label2" runat="server" ForeColor="red">(填寫輸入，上限1000個字)</asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol">
                                <asp:TextBox ID="tPOWERNEED2" runat="server" Width="60%" Height="80px" TextMode="MultiLine" MaxLength="1000" placeholder="(依產業人力需求調查結果，進行區域性的人力需求調查，應論述調查期間、區域範圍、調查對象及該產業於該區域之訓練需求)"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td>訓練需求概述：<asp:Label ID="Label3" runat="server" ForeColor="red">(填寫輸入，上限200個字)</asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol">
                                <asp:TextBox ID="tPOWERNEED3" runat="server" Width="60%" Height="80px" TextMode="MultiLine" MaxLength="200"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td>
                                <asp:CheckBox ID="cbPOWERNEED4" runat="server" />課程須符合目的事業主管機關相關規定：
                                <asp:Label ID="Label4" runat="server" ForeColor="red">(填寫輸入，上限200個字)</asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol">
                                <asp:TextBox ID="tPOWERNEED4" runat="server" Width="60%" Height="80px" TextMode="MultiLine" MaxLength="200" placeholder="(如為目的事業主管機關已定有訓練課程、時數、參訓人員資格認定及程序等相關規定者，應依其規定辦理，並加以說明規定內容。)"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="whitecol">(是否瞭解區域產業需求) </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td id="TD10" colspan="2" runat="server" class="bluecol" width="16%">訓練目標 </td>
                <td colspan="4" width="84%">
                    <table class="font" id="tbTrainTarget" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>單位核心能力介紹： </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:TextBox ID="tPlanCause" runat="server" Width="60%" Height="80px" TextMode="MultiLine"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td>知識： </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:TextBox ID="tPurScience" runat="server" Width="60%" Height="80px" TextMode="MultiLine"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td>技能： </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:TextBox ID="tPurTech" runat="server" Width="60%" Height="80px" TextMode="MultiLine"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td>學習成效： </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:TextBox ID="tPurMoral" runat="server" Width="60%" Height="80px" TextMode="MultiLine"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="whitecol">(是否符合需求並配合訓練單位核心能力)</td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" style="text-align: center;">職能級別：(單選) </td>
                        </tr>
                        <tr>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="rblFuncLevel" runat="server" RepeatLayout="Flow">
                                    <asp:ListItem Value="01">級別1(能夠在可預計及有規律的情況中，在密切監督及清楚指示下，執行常規性及重複性的工作。且通常不需要特殊訓練、教育及專業知識與技術)</asp:ListItem>
                                    <asp:ListItem Value="02">級別2(能夠在大部分可預計及有規律的情況中，在經常性監督下，按指導進行需要某些判斷及理解性的工作。需具備基本知識、技術)</asp:ListItem>
                                    <asp:ListItem Value="03">級別3(能夠在部分變動及非常規性的情況中，在一般監督下，獨立完成工作。需要一定程度的專業知識與技術及少許的判斷能力)</asp:ListItem>
                                    <asp:ListItem Value="04">級別4(能夠在經常變動的情況中，在少許監督下，獨立執行涉及規劃設計且需要熟練技巧的工作。需要具備相當的專業知識與技術，及作判斷及決定的能力)</asp:ListItem>
                                    <asp:ListItem Value="05">級別5(能夠在複雜變動的情況中，在最少監督下，自主完成工作。需要具備應用、整合、系統化的專業知識與技術及策略思考與判斷能力)</asp:ListItem>
                                    <asp:ListItem Value="06">級別6(能夠在高度複雜變動的情況中，應用整合的專業知識與技術，獨立完成專業與創新的工作。需要具備策略思考、決策及原創能力)</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td colspan="6" class="whitecol">
                    <div>
                        <table class="font" id="Datagrid2Table" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td class="table_title" align="center">授課教師 </td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" CellPadding="8">
                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <Columns>
                                            <asp:TemplateColumn HeaderText="序號">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" Width="8%"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="seqno" runat="server"></asp:Label>
                                                    <input id="HidTechID" runat="server" type="hidden" />
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="教師姓名">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" Width="14%"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="TeachCName" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="學歷">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" Width="18%"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="DegreeName" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="專業領域">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" Width="20%"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="ProLicense" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="遴選辦法說明">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:TextBox ID="TeacherDesc" runat="server" Width="80%" Height="80px" TextMode="MultiLine"></asp:TextBox>
                                                    <input id="btn_TCTYPEA" type="button" value="..." runat="server" class="button_b_Mini" />
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                    </asp:DataGrid>
                                </td>
                            </tr>
                        </table>
                    </div>
                    <div>
                        <table class="font" id="Datagrid2Table2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td align="center" class="table_title">授課助教 </td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:DataGrid ID="DataGrid22" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" CellPadding="8">
                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <Columns>
                                            <asp:TemplateColumn HeaderText="序號">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" Width="8%"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="seqno" runat="server"></asp:Label>
                                                    <input id="HidTechID" runat="server" type="hidden" />
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="助教姓名">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" Width="14%"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="TeachCName" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="學歷">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" Width="18%"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="DegreeName" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="專業領域">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" Width="20%"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="ProLicense" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="遴選辦法說明">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:TextBox ID="TeacherDesc" runat="server" Width="80%" Height="80px" TextMode="MultiLine"></asp:TextBox>
                                                    <input id="btn_TCTYPEB" type="button" value="..." runat="server" class="button_b_Mini" />
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                    </asp:DataGrid>
                                </td>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
            <tr>
                <td colspan="2" class="bluecol" width="16%">學員學歷 </td>
                <td colspan="3" width="82%">
                    <asp:DropDownList ID="CapDegree" runat="server"></asp:DropDownList></td>
                <td width="2%"></td>
            </tr>
            <tr>
                <td colspan="2" class="bluecol_need" width="16%">學員資格 </td>
                <td colspan="3" width="82%">
                    <asp:TextBox ID="CapAll" runat="server" Width="80%" Height="80px" TextMode="MultiLine"></asp:TextBox></td>
                <td width="2%"></td>
            </tr>
            <tr>
                <td colspan="2" class="bluecol" width="16%">訓練費用編列說明 </td>
                <td colspan="3" width="82%">
                    <asp:TextBox ID="CostDesc" runat="server" Width="80%" Height="80px" TextMode="MultiLine"></asp:TextBox></td>
                <td width="2%"></td>
            </tr>
        </table>
        <table class="font" id="Table_2" style="width: 100%;" cellspacing="1" cellpadding="1" border="0">
            <tr>
                <td bgcolor="#FFCC00">二、裝備與設施 </td>
            </tr>
        </table>
        <table class="font" id="Table_2_D" style="width: 100%;" cellspacing="1" cellpadding="1" border="0">
            <tr id="TRA1" runat="server">
                <td bgcolor="#CCD8EE" rowspan="2" width="14%" class="bluecol">學科場地 </td>
                <td id="TD11" bgcolor="#CCD8EE" runat="server" width="14%" class="bluecol">容納人數 </td>
                <td width="72%">
                    <asp:TextBox ID="Tnum2" runat="server" Width="20%" onfocus="this.blur()"></asp:TextBox></td>
            </tr>
            <tr id="TRA2" runat="server">
                <td bgcolor="#CCD8EE" width="14%" class="bluecol">硬體設施說明 </td>
                <td width="72%">
                    <asp:TextBox ID="HwDesc2" runat="server" Width="60%" onfocus="this.blur()" Height="80px" TextMode="MultiLine"></asp:TextBox></td>
            </tr>
            <tr id="TRB1" runat="server">
                <td bgcolor="#CCD8EE" rowspan="2" width="14%" class="bluecol">術科場地&nbsp; </td>
                <td bgcolor="#CCD8EE" width="14%" class="bluecol">容納人數 </td>
                <td width="72%">
                    <asp:TextBox ID="Tnum3" runat="server" Width="20%" onfocus="this.blur()"></asp:TextBox></td>
            </tr>
            <tr id="TRB2" runat="server">
                <td bgcolor="#CCD8EE" width="14%" class="bluecol">硬體設施說明 </td>
                <td width="72%">
                    <asp:TextBox ID="HwDesc3" runat="server" Width="60%" onfocus="this.blur()" Height="78px" TextMode="MultiLine"></asp:TextBox></td>
            </tr>
            <tr>
                <td bgcolor="#CCD8EE" colspan="2" class="bluecol" width="28%">其他器材設備 </td>
                <td width="72%">
                    <asp:TextBox ID="OtherDesc23" runat="server" Width="60%" onfocus="this.blur()" Height="80px" TextMode="MultiLine"></asp:TextBox></td>
            </tr>
        </table>
        <table class="font" id="Table_3" style="width: 100%;" cellspacing="1" cellpadding="1" border="0">
            <tr>
                <td bgcolor="#FFCC00">三、訓練模式特色與創新性 </td>
            </tr>
        </table>
        <table class="font" id="Table_3_D" style="width: 100%;" cellspacing="1" cellpadding="1" border="0">
            <tr>
                <td width="16%" bgcolor="#ffcccc" class="bluecol">訓練方式 </td>
                <td width="84%" class="whitecol">
                    <asp:TextBox ID="TrainMode" runat="server" Width="60%" Height="80px" TextMode="MultiLine" Enabled="False"></asp:TextBox></td>
            </tr>
        </table>
        <table class="font" id="Table_4" style="width: 100%;" cellspacing="1" cellpadding="1" border="0">
            <tr>
                <td bgcolor="#FFCC00">四、訓練績效評估 </td>
            </tr>
        </table>
        <table class="font" id="Table_4_D" style="width: 100%;" cellspacing="1" cellpadding="1" border="0">
            <tr>
                <td bgcolor="#ffcccc" class="bluecol" width="16%">1.反應評估 </td>
                <td class="whitecol" width="84%">
                    <asp:CheckBox ID="chk_RecDesc" runat="server" />
                    （是評量學員對訓練的觀感，以量化評量的方式來設計課後評量表，衡量學員對於訓練的反應，例如:設計滿意度調查機制瞭解學員感受包括知識、學習後關聯性、行政作業、課程是否值得推薦等）<br />
                    <asp:TextBox ID="RecDesc" runat="server" Width="40%" MaxLength="500"></asp:TextBox>(滿意度調查機制)
                </td>
            </tr>
            <tr>
                <td bgcolor="#ffcccc" class="bluecol" width="16%">2.學習評估 </td>
                <td class="whitecol" width="84%">
                    <asp:CheckBox ID="chk_LearnDesc" runat="server" />
                    （是評量學員因為參與訓練而改變態度、增進知識技能的程度。在此階段是關於學員在課程加強知識或是技巧的延伸，學習的評量則可經由課前測驗與課後測驗來達成，即可判斷訓練課程的成效。例如:考試或報告機制）<br />
                    <asp:TextBox ID="LearnDesc" runat="server" Width="40%" MaxLength="500"></asp:TextBox>(考試或報告機制)
                </td>
            </tr>
            <tr>
                <td bgcolor="#ffcccc" class="bluecol" width="16%">3.行為評估 </td>
                <td class="whitecol" width="84%">
                    <asp:CheckBox ID="chk_ActDesc" runat="server" />
                    （是評量學員因參與訓練而產生工作行為上的改變程度。經過3到6個月的訓練後，可對學員與其主管以問卷、面談、直接觀察、360度績效考評、目標設定等調查方法來評量，評量學員是否真的依照訓練的結果改變工作的模式。例如:課後行動計畫調查機制）<br />
                    <asp:TextBox ID="ActDesc" runat="server" Width="40%" MaxLength="500"></asp:TextBox>(課後行動計畫調查機制)
                </td>
            </tr>
            <tr>
                <td bgcolor="#ffcccc" class="bluecol" width="16%">4.成果評估 </td>
                <td class="whitecol" width="84%">
                    <asp:CheckBox ID="chk_ResultDesc" runat="server" />
                    （是評量因為參與訓練而產生的最後結果，如銷售額提升、成本降低、績效提升等，同時也是回應到參與訓練的理由。例如:(1)提升較高的客戶滿意度、(2)提高產值、(3)提高銷售額、(4)增加更多的新客戶、(5)降低更多的成本、(6)提高利潤）<br />
                    <asp:TextBox ID="ResultDesc" runat="server" Width="40%" MaxLength="500"></asp:TextBox>(工作行動調查機制)
                </td>
            </tr>
            <tr>
                <td bgcolor="#ffcccc" class="bluecol" width="16%">5.其它機制 </td>
                <td class="whitecol" width="84%">
                    <asp:CheckBox ID="chk_OtherDesc" runat="server" /><br />
                    <asp:TextBox ID="OtherDesc" runat="server" Width="40%" MaxLength="500"></asp:TextBox>
                </td>
            </tr>
        </table>
        <table class="font" id="Table_5" style="width: 100%;" cellspacing="1" cellpadding="1" border="0">
            <tr>
                <td bgcolor="#FFCC00">五、促進學習機制 </td>
            </tr>
        </table>
        <table class="font" id="Table_5_D" style="width: 100%;" cellspacing="1" cellpadding="1" border="0">
            <%--<tr>
                <td width="16%" bgcolor="#ffcccc" class="bluecol">招訓及遴選方式 </td>
                <td width="84%" class="whitecol">
                    <asp:TextBox ID="Recruit" runat="server" Width="40%" Height="80px" TextMode="MultiLine"></asp:TextBox></td>
            </tr>
            <tr>
                <td bgcolor="#ffcccc" class="bluecol_need" width="16%">學員激勵辦法 </td>
                <td width="84%" class="whitecol">
                    <asp:TextBox ID="Inspire" runat="server" Width="40%" Height="80px" TextMode="MultiLine"></asp:TextBox></td>
            </tr>--%>
            <tr>
                <td width="16%" bgcolor="#ffcccc" class="bluecol_need">是否為iCAP課程</td>
                <td width="84%" class="whitecol">
                    <br />
                    <table width="100%">
                        <tr>
                            <td width="20%" class="whitecol">
                                <%--RB_ISiCAPCOUR--%>
                                <asp:RadioButton ID="RB_ISiCAPCOUR_Y" runat="server" Text="是,請填寫" GroupName="RB_ISiCAPCOUR" />
                                <br />
                                <br />
                                <asp:RadioButton ID="RB_ISiCAPCOUR_N" runat="server" Text="否" Checked="true" GroupName="RB_ISiCAPCOUR" />
                            </td>
                            <td class="whitecol" valign="top">
                                <%----%>
                                <asp:Label ID="lab_iCAPCOURDESC" runat="server" Text="iCAP課程相關說明"></asp:Label><br />
                                <asp:TextBox ID="iCAPCOURDESC" runat="server" Width="80%" Height="80px" TextMode="MultiLine" MaxLength="500"></asp:TextBox>

                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td width="16%" bgcolor="#ffcccc" class="bluecol_need">招訓方式</td>
                <td width="84%" class="whitecol">
                    <asp:TextBox ID="Recruit" runat="server" Width="80%" Height="80px" TextMode="MultiLine"></asp:TextBox></td>
            </tr>
            <tr>
                <td width="16%" bgcolor="#ffcccc" class="bluecol_need">遴選方式</td>
                <td width="84%" class="whitecol">
                    <asp:TextBox ID="Selmethod" runat="server" Width="80%" Height="80px" TextMode="MultiLine"></asp:TextBox></td>
            </tr>
            <tr>
                <td bgcolor="#ffcccc" class="bluecol_need" width="16%">學員激勵辦法 </td>
                <td width="84%" class="whitecol">
                    <asp:TextBox ID="Inspire" runat="server" Width="80%" Height="80px" TextMode="MultiLine"></asp:TextBox></td>
            </tr>


        </table>
        <table class="font" id="Table_6" style="width: 100%;" cellspacing="1" cellpadding="1" border="0">
            <tr>
                <td bgcolor="#FFCC00">六、學分費或訓練費 </td>
            </tr>
        </table>
        <table class="font" id="Table_6_D" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td class="bluecol" width="16%">政府補助 </td>
                <td class="whitecol" width="84%">
                    <asp:TextBox ID="DefGovCost" runat="server" Width="14%" onfocus="this.blur()"></asp:TextBox>&nbsp;元／每班
                    <asp:TextBox ID="DefGovCost_Tnum" runat="server" Width="14%" onfocus="this.blur()"></asp:TextBox>&nbsp;元／每人
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="16%">學員自付 </td>
                <td class="whitecol" width="84%">
                    <asp:TextBox ID="DefStdCost" runat="server" Width="14%" onfocus="this.blur()"></asp:TextBox>&nbsp;元／每班
                    <asp:TextBox ID="DefStdCost_Tnum" runat="server" Width="14%" onfocus="this.blur()"></asp:TextBox>&nbsp;元／每人
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="16%">總 計 </td>
                <td class="whitecol" width="84%">
                    <asp:TextBox ID="TotalCost" runat="server" Width="14%" onfocus="this.blur()"></asp:TextBox>&nbsp;元／每班
                    <asp:TextBox ID="TotalCost_Tnum" runat="server" Width="14%" onfocus="this.blur()"></asp:TextBox>&nbsp;元／每人
                </td>
            </tr>
        </table>
        <table class="font" id="Table_7" style="width: 100%;" cellspacing="1" cellpadding="1" border="0">
            <tr>
                <td bgcolor="#FFCC00">七、其他 </td>
            </tr>
        </table>
        <table class="font" id="Table_7b" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td class="bluecol" width="16%">是否輔導學員參加政府機關辦理相關證照考試或技能檢定 </td>
                <td class="whitecol" width="84%">
                    <asp:CheckBox ID="TGovExamCY" runat="server" />
                    是，證照或檢定名稱<asp:TextBox ID="TGovExamName" runat="server" MaxLength="50" Width="40%"></asp:TextBox><br />
                    <asp:CheckBox ID="TGovExamCN" runat="server" />
                    否 (包含非政府機關辦理相關證照或檢定) </td>
            </tr>
        </table>
        <table class="font" id="Table_8" style="width: 100%;" cellspacing="1" cellpadding="1" border="0">
            <tr>
                <td bgcolor="#FFCC00">八、備註 </td>
            </tr>
        </table>
        <table class="font" id="Table_8b" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td class="whitecol">&nbsp;&nbsp;&nbsp;<asp:CheckBox ID="chkMEMO8C1" runat="server" />
                    <asp:Label ID="lbMEMO8" runat="server"></asp:Label><br />
                    &nbsp;&nbsp;(課程內容類似職業安全衛生教育訓練且不報請主管機關核備者，應點選此項，避免民眾誤解可作為時數認列)。<br />
                    <br />
                    &nbsp;&nbsp;&nbsp;<asp:CheckBox ID="chkMEMO8C2" runat="server" />
                    <asp:TextBox ID="txtMemo8" runat="server" Width="40%" MaxLength="500"></asp:TextBox><br />
                </td>
            </tr>
            <tr>
                <td class="whitecol">&nbsp;</td>
            </tr>
        </table>
        <table class="font" id="Table_S" style="width: 100%;" cellspacing="1" cellpadding="1" border="0">
            <tr>
                <td>
                    <div align="center" class="whitecol">
                        <asp:Button ID="Button6" runat="server" Text="草稿儲存" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="bt_addrow" runat="server" Text="正式送出" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="Button1" runat="server" Text="回上一頁" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                    </div>
                </td>
            </tr>
        </table>
        <input id="hidmemo8" type="hidden" name="hidmemo8" runat="server">
        <asp:HiddenField ID="Hid_COMIDNO" runat="server" />
        <asp:HiddenField ID="Hid_sender1" runat="server" />
    </form>
</body>
</html>
