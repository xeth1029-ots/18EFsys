<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SelectPlan.aspx.vb" Inherits="WDAIIP.SelectPlan" MasterPageFile="~/LayoutNoFunc.Master" %>

<asp:Content ContentPlaceHolderID="MainCPH" runat="server">

    <div class="col-sm-10 col-md-offset-1">
        <div class="login-bar">
            <h3 class="loginTitleA"><img src="/images/icon-arrow.svg" alt="項目符號" />選擇 年度/計畫</h3>

    <div class="col-sm-12">
            <div class="theme-news">
                <div class="form-border">
                    <div class="form-group">
                        <label class="col-sm-2 control-label label-set"><span class="mark-red">＊</span>年度</label>
                        <div class="col-sm-10 form-inline">
                            <asp:DropDownList ID="YR" runat="server" AutoPostBack="false" CssClass="form-control formbar-bg">
							</asp:DropDownList>
                        </div>
                    </div>
                    <div class="form-group">
                        <label class="control-label col-sm-2 label-set"><span class="mark-red">＊</span>計畫別</label>
                        <div class="col-sm-10 form-inline">
                            <asp:DropDownList ID="PLANID" runat="server" AutoPostBack="false" CssClass="form-control formbar-bg">
							</asp:DropDownList>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        <div class="col-sm-12 yr-stp-set">
            <div class="btn-group btn-group-right btn-group-sm clearfix">
                <asp:Button runat="server" ID="btnSubmit" CssClass="btn btn-info" Text="確認" OnClick="btnSubmit_Click"/>
                <%--<asp:Button ID="bt_back1" type="back1" runat="server" CssClass="btn btn-info" Text="&nbsp;&nbsp;&nbsp;回登入頁&nbsp;&nbsp;&nbsp;" />--%>
            </div>
        </div>

    </div>
    </div>
    <div style="position: absolute; top: 508px; left: 1px;" id="div12" runat="server">
    <%--ForeColor="White"--%>
    <asp:Label ID="Labmsg1" runat="server" Text="Labx" ForeColor="#D5EEFF"></asp:Label>
    </div>
    <div style="position: absolute; left: -20px; top: 22px; height: 17px; width: 36px;" id="divC" runat="server">
    <a id="A1" title="關閉" class="l" href="#" onclick="window.opener=null; window.open('','_self'); window.close();" style="color:#D5EEFF">關閉</a>
    </div>
<script type="text/javascript">
    $(document).ready(function () {
        $("select[id*=YR]")
            .change(function () {
                RefreshPlanID(this.value);
            })
            .trigger("change");
    });
    function RefreshPlanID(selYR) {
        var parms = {
            "OP": "Ajax",
            "YR": selYR
        };
        var url = "<%=Request.ApplicationPath%>SelectPlan.aspx";
        ajaxLoadMore(url, parms, function(resp) {
            if (resp != undefined) {
                $("select[id*=PLANID]").html(resp);
            }
            else {
                blockAlert("Ajax載入年度計畫清單失敗!");
            }
        });
    }
</script>
</asp:Content>

