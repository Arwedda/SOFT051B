<%@ Page Language="VB" %>
<script runat="server">
    Sub Page_Load(s As Object, e As EventArgs)
        Session.Abandon()
        Response.Redirect("Login.aspx")
    End Sub
</script>