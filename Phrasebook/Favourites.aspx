<%@ Page Language="VB" %>
<%@ Import Namespace="System.Data.OleDb" %>
<script runat="server">
    Dim CS As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + _
                        "Data Source=" + Server.MapPath("Phrasebook.accdb") + ";"
    Dim CN As New OleDbConnection(CS)
    Dim CMD As OleDbCommand
    Dim Reader As OleDbDataReader
    Dim SQL As String
    Dim HTML As String
    Dim Count As Single

    Sub Page_Load(s As Object, e As EventArgs)
        If Session("Logged") = 0 Then
            Response.Redirect("Login.aspx")
        End If
        Count = 0
        SQL = "SELECT * FROM Favourite, Phrase WHERE Favourite.StudentID = " & Session("Logged") & " AND Phrase.ID = Favourite.PhraseID;"
        CreateList()
    End Sub
    
    Sub CreateList()
        CMD = New OleDbCommand(SQL, CN)
        CN.Open()
        Reader = CMD.ExecuteReader()
        HTML = "<table align='center'>"
        Do While Reader.Read()
            HTML = HTML & "<tr><td><a href='Details.aspx?PID=" & Reader("ID") & "'>" & Reader("EnglishText") & "</a></td><td>" & Reader("PinYin") & "</td></tr>"
            Count = Count + 1
        Loop
        CN.Close()
        HTML = HTML & "</table>"
        parList.InnerHtml = HTML
        If Count = 0 Then
            parCount.InnerHtml = "<center>You don't have any favourite phrases to display! :-(</center>"
        ElseIf Count = 1 Then
            parCount.InnerHtml = "<center>Displaying your " & Count & " selected favourite phrase.</center>"
        Else
            parCount.InnerHtml = "<center>Displaying your " & Count & " selected favourite phrases.</center>"
        End If
    End Sub
</script>
<HTML>
    <head><title>Favourites</title></head>
    <body bgcolor="#FFF5FF">
        <form id="Form" runat="server">
            <div style="background-color:#E08D8A; border-width:2px; border-style:solid; border-color:#E55451; border-radius:8px;">
                <center><h1>Phrasebook - Favourites</h1>
                <h3><a href="Phrases.aspx">Phrases</a> | <a href="Favourites.aspx">Favourites</a> | <a href="Logout.aspx">Logout</a></h3></center>
            </div>
            <p id="parList" runat="server" ></p>
            <p id="parCount" runat="server"></p>
        </form>
    </body>
</HTML>