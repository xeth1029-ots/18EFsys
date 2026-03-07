<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_04_002_add2.aspx.vb" Inherits="WDAIIP.TC_04_002_add2" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>班級審核作業</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
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
        function create_techtable(tidobj, tnameobj, didobj, dnameobj, Majorobj, tableobj) {
            //create_techtable(
            var TeachID = "0," + tidobj.value;
            var TeachName = "教師姓名," + tnameobj.value;
            //var DegreeID="0,"+didobj.value;
            var DegreeName = "學歷," + dnameobj.value;
            var MajorName = "專業領域," + Majorobj.value;

            var stid = TeachID.split(",");
            var stname = TeachName.split(",");
            //var sdid=DegreeID.split(",");
            var sdname = DegreeName.split(",");
            var Maname = MajorName.split(",");

            var element0, element1, element2, element3, element3b, element4, element5;
            var element4a, element4b, element4c, element4d, element4e;
            var newSpana, newSpanb, newSpand, newSpane; //= document.createElement("span");

            element0 = document.createDocumentFragment();
            element1 = document.createElement("table");
            setStyle(element1, "height:6px;width:100%;border-collapse:collapse;font-family: \"新細明體\", \"細明體\", \"Arial, \"Helvetica\", sans-serif\";font-size: 9pt;line-height: 16px;");
            element1.setAttribute("cellspacing", "0");
            element1.setAttribute("rules", "all");
            element1.setAttribute("DESIGNTIMEDRAGDROP", "347");
            element1.setAttribute("border", "1");
            element2 = document.createElement("tbody");
            //alert(stid.length);
            for (i = 0; i < stid.length; i++) {
                //alert(i);
                if (i == 0) {
                    element3 = document.createElement("tr");
                    setStyle(element3, "height:6px;width:100%;border-collapse:collapse;font-family: \"新細明體\", \"細明體\", \"Arial, \"Helvetica\", sans-serif\";font-size: 9pt;line-height: 16px;");
                    element4a = document.createElement("td");
                    element4a.setAttribute("id", i);
                    element4b = document.createElement("td");
                    element4b.setAttribute("id", i);
                    element4d = document.createElement("td");
                    element4d.setAttribute("id", i);
                    element4e = document.createElement("td");
                    element4e.setAttribute("id", i);

                    newSpana = document.createElement("span");
                    newSpana.appendChild(document.createTextNode("序號"));
                    element4a.appendChild(newSpana);
                    //element4b.onclick = function(){godo(this);};
                    newSpanb = document.createElement("span");
                    newSpanb.appendChild(document.createTextNode(stname[i]));
                    element4b.appendChild(newSpanb);
                    //element4d.onclick = function(){godo(this);};
                    newSpand = document.createElement("span");
                    newSpand.appendChild(document.createTextNode(sdname[i]));
                    element4d.appendChild(newSpand);

                    newSpane = document.createElement("span");
                    newSpane.appendChild(document.createTextNode(Maname[i]));
                    element4e.appendChild(newSpane);

                    element3.appendChild(element4a);
                    element3.appendChild(element4b);
                    element3.appendChild(element4d);
                    element3.appendChild(element4e);

                    element2.appendChild(element3);

                }
                else {
                    element3b = document.createElement("tr");
                    setStyle(element3b, "height:6px;width:100%;border-collapse:collapse;font-family: \"新細明體\", \"細明體\", \"Arial, \"Helvetica\", sans-serif\";font-size: 9pt;line-height: 16px;");
                    element4a = document.createElement("td");
                    element4a.setAttribute("id", i);
                    element4b = document.createElement("td");
                    element4b.setAttribute("id", i);
                    element4d = document.createElement("td");
                    element4d.setAttribute("id", i);
                    element4e = document.createElement("td");
                    element4e.setAttribute("id", i);

                    //element4b.onclick = function(){godo(this);};
                    newSpana = document.createElement("span");
                    newSpana.appendChild(document.createTextNode(i));
                    element4a.appendChild(newSpana);

                    newSpanb = document.createElement("span");
                    newSpanb.appendChild(document.createTextNode(stname[i]));
                    element4b.appendChild(newSpanb);

                    //element4d.onclick = function(){godo(this);};
                    newSpand = document.createElement("span");
                    newSpand.appendChild(document.createTextNode(sdname[i]));
                    element4d.appendChild(newSpand);

                    newSpane = document.createElement("span");
                    newSpane.appendChild(document.createTextNode(Maname[i]));
                    element4e.appendChild(newSpane);

                    element3b.appendChild(element4a);
                    element3b.appendChild(element4b);
                    element3b.appendChild(element4d);
                    element3b.appendChild(element4e);

                    element2.appendChild(element3b);
                }

            }
            element1.appendChild(element2);
            element0.appendChild(element1);

            setDiv(tableobj);
            tableobj.appendChild(element0);
        }

        function godo(element) {
            window.alert(element.getAttribute("id"));
        }

    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
<%--<table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0"><tr><td><asp:Label ID="TitleLab1" runat="server"></asp:Label>
<asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;基本資料設定&gt;&gt;<font color="#990000">班級審核作業</font></asp:Label></td></tr></table>--%>
<%--<font color="#990000">-新增(修改)</font> (<font color="#ff0000">*</font>為必填欄位)--%>
<%--<table class="font" id="Table_T" cellspacing="1" cellpadding="1" width="736" border="0"><tr><td></td></tr></table>--%>
        <table cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr class="whitecol">
                <td class="bluecol" style="width: 16%">年度 </td>
                <td>
                    <asp:DropDownList ID="PlanYear" runat="server" Enabled="False">
                    </asp:DropDownList>
                </td>
                <td class="bluecol" style="width: 16%">訓練機構 </td>
                <td>
                    <asp:Label ID="labORGNAME" runat="server"></asp:Label>
                </td>
            </tr>
            <tr class="whitecol">
                <td id="Td15" align="left" class="bluecol" runat="server">課程名稱 </td>
                <td>
                    <asp:TextBox ID="ClassName" runat="server" Columns="30" onfocus="this.blur()"></asp:TextBox>
                </td>
                <td id="Td16" align="left" class="bluecol" runat="server">&nbsp;&nbsp;&nbsp;&nbsp;訓練職能 </td>
                <td>
                    <asp:DropDownList ID="ClassCate" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr id="TRClassid" runat="server" class="whitecol">
                <td id="Td17" align="left" class="bluecol" runat="server">課程班別 </td>
                <td colspan="3">
                    <asp:DropDownList ID="ClassID" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr class="whitecol">
                <td id="Td1" nowrap align="left" class="bluecol" runat="server">訓練人數 </td>
                <td id="Td2" nowrap align="left" runat="server">
                    <asp:TextBox ID="Tnum" runat="server" onfocus="this.blur()"></asp:TextBox>
                </td>
                <td id="Td3" nowrap align="left" class="bluecol" runat="server">訓練時數 </td>
                <td id="Td4" nowrap align="left" runat="server">
                    <asp:TextBox ID="THours" runat="server" onfocus="this.blur()"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="table_title" align="center" colspan="8">上課時間 </td>
            </tr>
            <tr>
                <td colspan="8">
                    <table class="font" id="DataGrid1Table" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" AutoGenerateColumns="False" CssClass="font" CellPadding="8" Width="100%">
                                    <ItemStyle BackColor="White"></ItemStyle>
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:TemplateColumn HeaderText="星期">
                                            <HeaderStyle Width="100px"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="Weeks1" runat="server"></asp:Label>
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:DropDownList ID="Weeks2" runat="server">
                                                </asp:DropDownList>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="上課時段">
                                            <ItemTemplate>
                                                <asp:Label ID="Times1" runat="server"></asp:Label>
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox ID="Times2" runat="server"></asp:TextBox>
                                            </EditItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn Visible="False" HeaderText="功能">
                                            <HeaderStyle Width="100px"></HeaderStyle>
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
                <td class="bluecol" width="16%">起迄日期 </td>
                <td colspan="6" class="whitecol">
                    <span runat="server">
                        <asp:TextBox ID="start_date" runat="server" onfocus="this.blur()" Width="15%"></asp:TextBox>
                        <img id="IMG1" style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30">～
                        <asp:TextBox ID="end_date" runat="server" onfocus="this.blur()" Width="15%"></asp:TextBox>
                        <img id="IMG2" style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30">
                    </span>
                    <asp:CompareValidator ID="comdate" runat="server" ErrorMessage="迄日不得小於起日或日期格式不正確" Display="None" ControlToValidate="end_date" Type="Date" Operator="GreaterThan" ControlToCompare="start_date"></asp:CompareValidator><asp:RequiredFieldValidator ID="mustsdate" runat="server" ErrorMessage="請輸入起日" Display="None" ControlToValidate="start_date"></asp:RequiredFieldValidator><asp:RequiredFieldValidator ID="mustedate" runat="server" ErrorMessage="請輸入迄日" Display="None" ControlToValidate="end_date"></asp:RequiredFieldValidator><asp:ValidationSummary ID="totalmsg" runat="server" ShowSummary="False" ShowMessageBox="True" DisplayMode="List"></asp:ValidationSummary>
                </td>
            </tr>
            <%--<TR><TD id="Td22" style="WIDTH: 126px; HEIGHT: 30px" align="left" class="bluecol" colSpan="2"runat="server">訓練性質</TD>
                <TD style="HEIGHT: 30px" colSpan="6"><asp:dropdownlist id="ProcID" Enabled="False" Width="192px" Runat="server"></asp:dropdownlist></TD>
                </TR><TR><TD id="Td23" style="WIDTH: 126px; HEIGHT: 18px" align="left" class="bluecol" colSpan="2"runat="server">上課時段</TD>
                <TD   colSpan="6"><asp:dropdownlist id="TPeriod" Width="192px" Runat="server"></asp:dropdownlist></TD></TR>--%>
        </table>
        <table id="Table_1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <%--<td bgcolor="#ffcc33">一、規劃與執行能力 </td>--%>
                <td class="table_title" align="center">規劃與執行能力 </td>
            </tr>
        </table>
        <table id="Table_1D" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table cellspacing="1" cellpadding="1" border="0" width="100%">
                        <tr>
                            <td class="bluecol_need" width="16%">教學方法 </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="cblTMethod" runat="server" CssClass="font" RepeatLayout="Flow">
                                    <asp:ListItem Value="01">講授教學法（運用敘述或講演的方式，傳遞教材知識的一種教學方法，提供相關教材或講義）</asp:ListItem>
                                    <asp:ListItem Value="02">討論教學法（指團體成員齊聚一起，經由說、聽和觀察的過程，彼此溝通意見，由講師帶領達成教學目標）</asp:ListItem>
                                    <asp:ListItem Value="03">演練教學法（由講師的帶領下透過設備或教材，進行練習、表現和實作，親自解說示範的技能或程序的一種教學方法）</asp:ListItem>
                                    <asp:ListItem Value="99">其他教學方法：</asp:ListItem>
                                </asp:CheckBoxList>
                                <asp:TextBox runat="server" ID="TMethodOth" Width="40%" MaxLength="100"></asp:TextBox><br />
                                <asp:Label ID="Label2" runat="server" ForeColor="red">(若選"其他教學方法"，需填寫輸入，上限100個字)</asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="16%">訓練需求調查</td>
                            <td class="whitecol">
                                <table cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td>產業人力需求調查：</td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">
                                            <asp:TextBox ID="tPOWERNEED1" runat="server" Width="60%" Height="80px" TextMode="MultiLine" MaxLength="1000" placeholder="(應論述調查期間、區域範圍、調查對象、產業發展趨勢及該產業之訓練需求)"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td>區域人力需求調查：</td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">
                                            <asp:TextBox ID="tPOWERNEED2" runat="server" Width="60%" Height="80px" TextMode="MultiLine" MaxLength="1000" placeholder="(依產業人力需求調查結果，進行區域性的人力需求調查，應論述調查期間、區域範圍、調查對象及該產業於該區域之訓練需求)"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td>訓練需求概述：</td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">
                                            <asp:TextBox ID="tPOWERNEED3" runat="server" Width="60%" Height="80px" TextMode="MultiLine" MaxLength="200"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:CheckBox ID="cbPOWERNEED4" runat="server" />課程須符合目的事業主管機關相關規定： </td>
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
                            <td class="bluecol_need" width="16%">訓練目標</td>
                            <td>
                                <table cellspacing="1" cellpadding="1" width="100%" border="0">
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
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table class="font" id="tbTeacherDesc_AB" style="width: 100%;" cellspacing="1" cellpadding="1" border="0">
                        <tr>
                            <td colspan="6" class="whitecol">
                                <div>
                                    <table class="font" id="tbDataGrid21" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                        <tr>
                                            <td class="table_title" align="center">授課教師 </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <asp:DataGrid ID="DataGrid21" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" CellPadding="8">
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
                                    <table class="font" id="tbDataGrid22" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
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
<%--<td class="bluecol_need">授課教師 - 遴選辦法說明</td>--%>
<%--<asp:TextBox ID="TeacherDesc_A" runat="server" Width="60%" Height="80px" TextMode="MultiLine"></asp:TextBox>
<input id="btn_TCTYPEA" type="button" value="..." runat="server" class="button_b_Mini" />--%>
<%--<td class="bluecol_need">授課助教 - 遴選辦法說明</td>--%>
<%--<asp:TextBox ID="TeacherDesc_B" runat="server" Width="60%" Height="80px" TextMode="MultiLine"></asp:TextBox>
<input id="btn_TCTYPEB" type="button" value="..." runat="server" class="button_b_Mini" />--%>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <table id="Table_1_D" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td class="bluecol" style="width: 20%">學員學歷 </td>
                <td class="whitecol" >
                    <asp:DropDownList ID="CapDegree" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol"  >學員資格 </td>
                <td class="whitecol">
                    <asp:TextBox ID="CapAll" runat="server" Width="80%" Height="80px" TextMode="MultiLine"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="bluecol">訓練費用<br />編列說明 </td>
                <td class="whitecol">
                    <asp:TextBox ID="CostDesc" runat="server" Width="80%" Height="80px" TextMode="MultiLine"></asp:TextBox>
                </td>
            </tr>
        </table>
        <table class="font" id="Table_2" style="width: 100%" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <%--<td bgcolor="#ffcc33">二、裝備與設施 </td>--%>
                <td class="table_title" align="center">裝備與設施 </td>
            </tr>
        </table>
        <table class="font" id="Table_2_D" style="width: 100%; height: 270px" cellspacing="1" cellpadding="1" border="0">
            <tr class="whitecol">
                <td style="width: 10%; height: 125px" class="bluecol" rowspan="2">學科場地 </td>
                <td id="TD11" style="width: 10%; height: 72px" class="bluecol" runat="server">容納人數 </td>
                <td style="width: 42.59%; height: 72px">
                    <asp:TextBox ID="Tnum2" runat="server" Width="100px" onfocus="this.blur()"></asp:TextBox>
                </td>
            </tr>
            <tr class="whitecol">
                <td class="bluecol">硬體設施說明 </td>
                <td style="width: 42.59%">
                    <asp:TextBox ID="HwDesc2" runat="server" Width="90%" Height="78px" onfocus="this.blur()" TextMode="MultiLine"></asp:TextBox>
                </td>
            </tr>
            <tr class="whitecol">
                <td style="width: 10%" class="bluecol" rowspan="2">術科場地&nbsp; </td>
                <td style="width: 10%; height: 72px" class="bluecol">容納人數 </td>
                <td style="width: 42.59%">
                    <asp:TextBox ID="Tnum3" runat="server" Width="100px" onfocus="this.blur()"></asp:TextBox>
                </td>
            </tr>
            <tr class="whitecol">
                <td class="bluecol">硬體設施說明 </td>
                <td style="width: 42.59%">
                    <asp:TextBox ID="HwDesc3" runat="server" Width="90%" Height="78px" onfocus="this.blur()" TextMode="MultiLine"></asp:TextBox>
                </td>
            </tr>
            <tr class="whitecol">
                <td class="bluecol" colspan="2">其他器材設備 </td>
                <td style="width: 42.59%">
                    <asp:TextBox ID="OtherDesc23" runat="server" Width="90%" Height="78px" onfocus="this.blur()" TextMode="MultiLine"></asp:TextBox>
                </td>
            </tr>
        </table>
        <%--<table class="font" id="Table_3" style="width: 100%; height: 20px" cellspacing="1" cellpadding="1" width="848" border="0">
            <tr>
                <td bgcolor="#ffcc33">三、訓練模式特色與創新性 </td>
            </tr>
        </table>--%>
        <table class="font" id="Table_3_D" style="width: 100%;" cellspacing="1" cellpadding="1" border="0">
            <%--<tr>
                <td width="16%" class="bluecol">訓練方式 </td>
                <td>
                    <asp:TextBox ID="TrainMode" runat="server" Width="90%" TextMode="MultiLine" Rows="6"></asp:TextBox>
                </td>
            </tr>--%>
            <tr>
                <td colspan="2">
                    <table class="font" id="Datagrid3Table" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td class="table_title" align="center">課程大綱 </td>
                        </tr>
                        <tr>
                            <td class="whitecol">
                                <asp:DataGrid ID="Datagrid3" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font">
                                    <ItemStyle BackColor="White"></ItemStyle>
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:TemplateColumn HeaderText="日期">
                                            <HeaderStyle Width="100px"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="STrainDateLabel" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="授課時間">
                                            <HeaderStyle Width="100px"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="PNameLabel" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="時數">
                                            <ItemTemplate>
                                                <asp:Label ID="PHourLabel" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="課程進度／內容">
                                            <ItemTemplate>
                                                <asp:TextBox ID="PContText" runat="server" onfocus="this.blur()" Width="98%" Columns="50" TextMode="MultiLine" Rows="5" Enabled="False" Height="58px"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="學／術科">
                                            <HeaderStyle Width="120px"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:DropDownList ID="drpClassification1" runat="server" Enabled="False" AutoPostBack="True">
                                                    <asp:ListItem Value="1">學科</asp:ListItem>
                                                    <asp:ListItem Value="2">術科</asp:ListItem>
                                                </asp:DropDownList>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="上課地點">
                                            <ItemTemplate>
                                                <asp:DropDownList ID="drpPTID" runat="server" Width="90%" Enabled="False">
                                                </asp:DropDownList>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="遠距教學">
                                            <HeaderStyle Width="100px"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:CheckBox ID="cb_FARLEARNi" runat="server" Enabled="False" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="任課教師">
                                            <HeaderStyle Width="120px"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <input id="Tech1Value" type="hidden" name="Tech1Value" runat="server">
                                                <asp:TextBox ID="Tech1Text" runat="server" onfocus="this.blur()" Columns="10" Enabled="False"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <%--
				<tr>
					<td width="80" class="bluecol" height="88">課程大綱</td>
					<td width="454" height="88"><asp:textbox id="Content" runat="server" Width="624px" Height="78px" TextMode="MultiLine"></asp:textbox></td>
				</tr>
            --%>
        </table>
        <table class="font" id="Table_4" style="width: 100%; height: 20px" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <%--<td bgcolor="#ffcc33">四、訓練績效評估 </td>--%>
                <td class="table_title" align="center" colspan="2">訓練績效評估</td>
            </tr>
        </table>
        <table class="font" id="Table_4_D" cellspacing="1" cellpadding="1" border="0" width="100%">
            <tr class="whitecol">
                <td width="16%" class="bluecol">1.反應評估 </td>
                <td>
                    <asp:TextBox ID="RecDesc" runat="server" Width="80%"></asp:TextBox>(滿意度調查機制) </td>
            </tr>
            <tr class="whitecol">
                <td class="bluecol">2.學習評估 </td>
                <td>
                    <asp:TextBox ID="LearnDesc" runat="server" Width="80%"></asp:TextBox>(考試或報告機制) </td>
            </tr>
            <tr class="whitecol">
                <td class="bluecol">3.行為評估 </td>
                <td>
                    <asp:TextBox ID="ActDesc" runat="server" Width="80%"></asp:TextBox>(課後行動計畫調查機制) </td>
            </tr>
            <tr class="whitecol">
                <td class="bluecol">4.成果評估 </td>
                <td>
                    <asp:TextBox ID="ResultDesc" runat="server" Width="80%"></asp:TextBox>(工作行動調查機制) </td>
            </tr>
            <tr class="whitecol">
                <td class="bluecol">5.其它機制 </td>
                <td>
                    <asp:TextBox ID="OtherDesc" runat="server" Width="80%"></asp:TextBox></td>
            </tr>
        </table>
        <table class="font" id="Table_5" style="width: 100%" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <%--<td bgcolor="#ffcc33">五、促進學習機制 </td>--%>
                <td class="table_title" align="center" colspan="2">促進學習機制</td>
            </tr>
        </table>
        <table class="font" id="Table_5_D" style="width: 100%; height: 173px" cellspacing="1" cellpadding="1" border="0">
            <tr>
                <td width="16%" class="bluecol">是否為iCAP課程</td>
                <td width="84%" class="whitecol">
                    <br />
                    <table width="100%" id="tb_ISiCAPCOUR" runat="server">
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
                <td width="16%" class="bluecol">招訓方式</td>
                <td width="84%" class="whitecol">
                    <asp:TextBox ID="Recruit" runat="server" Width="88%" Height="80px" TextMode="MultiLine"></asp:TextBox></td>
            </tr>
            <tr>
                <td width="16%" class="bluecol">遴選方式</td>
                <td width="84%" class="whitecol">
                    <asp:TextBox ID="Selmethod" runat="server" Width="88%" Height="80px" TextMode="MultiLine"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="bluecol" width="16%">學員激勵辦法 </td>
                <td width="84%" class="whitecol">
                    <asp:TextBox ID="Inspire" runat="server" Width="88%" Height="80px" TextMode="MultiLine"></asp:TextBox></td>
            </tr>
        </table>
        <table class="font" id="Table_6" style="width: 100%; height: 20px" cellspacing="1" cellpadding="1" width="848" border="0">
            <tr>
                <%--<td bgcolor="#ffcc33">六、學分費或訓練費 </td>--%>
                <td class="table_title" align="center" colspan="2">學分費或訓練費</td>
            </tr>
        </table>
        <table class="font" id="Table_6_D" style="width: 100%; height: 172px" cellspacing="1" cellpadding="1" border="0">
            <tr class="whitecol">
                <td style="width: 44%">政府補助
				    <asp:TextBox ID="DefGovCost" runat="server" Width="100px" onfocus="this.blur()"></asp:TextBox>&nbsp;元／每班
				    <asp:TextBox ID="DefGovCost_Tnum" runat="server" Width="100px" onfocus="this.blur()"></asp:TextBox>&nbsp;元／每人 </td>
                <td id="TD_6" rowspan="3">
                    <table class="font" id="table6" style="width: 100%; height: 152px" width="336" border="0">
                        <tr>
                            <td valign="top" class="TC_TD3">審核結果<font color="#ff0000">*</font> </td>
                        </tr>
                        <tr>
                            <td valign="top">
                                <asp:DropDownList ID="VerSeqNo_ch1" runat="server" Width="40%">
                                </asp:DropDownList>
                                <asp:Label ID="Label1" runat="server" Width="200px"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td valign="top" class="TC_TD3">不合格原因 </td>
                        </tr>
                        <tr>
                            <td valign="top">
                                <asp:TextBox ID="VerReason_ch1" runat="server" Width="88%" Height="77px" TextMode="MultiLine"></asp:TextBox>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr class="whitecol">
                <td style="width: 44%">學員自付
				<asp:TextBox ID="DefStdCost" runat="server" Width="100px" onfocus="this.blur()"></asp:TextBox>&nbsp;元／每班
				<asp:TextBox ID="DefStdCost_Tnum" runat="server" Width="100px" onfocus="this.blur()"></asp:TextBox>&nbsp;元／每人 </td>
            </tr>
            <tr class="whitecol">
                <td style="width: 44%">總計
				<asp:TextBox ID="TotalCost" runat="server" Width="100px" onfocus="this.blur()"></asp:TextBox>&nbsp; </td>
            </tr>
        </table>
        <table class="font" id="Table_S" style="width: 100%; height: 28px" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <div align="center" class="whitecol">
                        <asp:Button ID="bt_addrow" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="Button1" runat="server" Text="回上一頁" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                    </div>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="Hid_ComIDNO" runat="server" />
        <asp:HiddenField ID="Hid_RIDValue" runat="server" />
        <asp:HiddenField ID="Hid_SENTERDATE" runat="server" />
        <asp:HiddenField ID="Hid_FENTERDATE" runat="server" />

        <%--<asp:HiddenField ID="RIDValue" runat="server" />--%>
        <%-- <asp:HiddenField ID="HiddenField1" runat="server" />--%>
        <%--<input id="RIDValue" type="hidden" runat="server" />--%>
        <%-- <asp:HiddenField ID="Hid_clsid" runat="server" />
        <asp:HiddenField ID="Hid_CyclType" runat="server" />
        <asp:HiddenField ID="Hid_OCID" runat="server" />--%>
    </form>
</body>
</html>
