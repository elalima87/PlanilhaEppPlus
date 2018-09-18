<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="PlanilhaEppPlus._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <script type="text/javascript">
        function carregaRelatorio(relatorio) {
            alert(relatorio);

            //$('#divRelatorio').css('display', 'block');
            $('#iframe').prop('src', "Default.aspx?pTipoOperacao=X&pRelatorio=" + relatorio);


        }



    </script>

    <div class="container">
        <div class="col-md-4 col-xs-12">

            <div class="row">
                <div class="col-4">
                    <div class="list-group" id="list-tab" role="tablist">
                        <a class="list-group-item list-group-item-action active" id="list-home-list" data-toggle="list" href="javascript:carregaRelatorio('Rel1');" role="tab" aria-controls="relatorio1">1° Relatório</a>
                        <a class="list-group-item list-group-item-action disabled" id="list-profile-list" data-toggle="list" href="javascript:carregaRelatorio('Rel2');" role="tab" aria-controls="relatorio2">2° Relatório</a>
                        <%--<a class="list-group-item list-group-item-action" id="list-messages-list" data-toggle="list" href="#list-messages" role="tab" aria-controls="messages">Messages</a>
                        <a class="list-group-item list-group-item-action" id="list-settings-list" data-toggle="list" href="#list-settings" role="tab" aria-controls="settings">Settings</a>--%>
                        <iframe id="iframe" style="display: none;"></iframe>
                    </div>
                </div>



            </div>
        </div>
    </div>
</asp:Content>
