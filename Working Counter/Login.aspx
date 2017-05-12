<%@ Page Language="VB" %>
<%@ Import Namespace="System.Data.OleDb" %>
<script runat="server">
    Dim CS As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + _
                        "Data Source=" + Server.MapPath("Phrasebook.accdb") + ";"
    Dim CN As New OleDbConnection(CS)
    Dim CMD As OleDbCommand
    Dim Reader As OleDbDataReader
    Dim SQL As String

    Sub Page_Load(s As Object, e As EventArgs)
        Session("Logged") = 0
        txtEmail.Focus()
    End Sub

    Sub btnLog_Clicked(s As Object, e As EventArgs) Handles btnLog.ServerClick
        Dim em As String
        Dim pw As String

        em = txtEmail.Value
        pw = txtPass.Value
        SQL = "SELECT ID, Email, Password FROM Student WHERE Email = '" & em & "' AND Password = '" & pw & "';"
        CMD = New OleDbCommand(SQL, CN)
        CN.Open()
        Reader = CMD.ExecuteReader()
        If Reader.Read() Then
            Session("Logged") = Reader("ID")
            Response.Redirect("Phrases.aspx")
        Else
            parMsg.InnerHtml = "<center>Login details incorrect. Please check that CAPSLOCK isn't on.</center>"
        End If
    End Sub
</script>
<html>
    <head><title>Authorised Users Only</title></head>
    <body bgcolor="#FFF5FF">
        <form id="Form" runat="server">
            <div style="background-color:#E08D8A; border-width:2px; border-style:solid; border-color:#E55451; border-radius:8px;">
                <center><h1>Phrasebook - Login</h1></center>
            </div>
            <br /><br />
            <br />
            <p id="parLog" style="text-align:center"><b>Please login:</b></p>
            <table align="center">
            <tr><td>Email:</td><td><input id="txtEmail" type="text" runat="server"  /></td></tr>
            <tr><td>Password: </td><td><input id="txtPass" type="password" runat="server"  /></td></tr>
            <tr><td></td><td><input id="btnLog" type="submit" value="Login" runat="server" /></td></tr></table>
            <p id="parMsg" runat="server" ></p>
        </form>
    </body>
</html>