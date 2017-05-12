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
    Dim Shown As Counter
    Dim Total As Counter
    Dim FirstCount As Boolean
    
    Sub Page_Load(s As Object, e As EventArgs)
        If Session("Logged") = 0 Then
            Response.Redirect("Login.aspx")
        End If
        
        If Session("Shown") Is Nothing Then
            Session("Shown") = New Counter
            Session("Total") = New Counter
        End If
        Shown = Session("Shown")
        Total = Session("Total")
        
        If FirstCount = False Then
            Total.Reset()
        End If
        
        SQL = "SELECT ID, EnglishText, Pinyin FROM Phrase;"
        CreateList(Total)
        parCount.InnerHtml = "<table align='center'><td>Displaying all " & Total.GetCount & " Phrases."
    End Sub
    
    Sub btnSearch_Clicked(s As Object, e As EventArgs) Handles btnSearch.ServerClick
        Dim Search As String
        Shown.Reset()
        Search = txtSearch.Value
        SQL = "SELECT ID, EnglishText, Pinyin FROM Phrase WHERE EnglishText Like '%" & Search & "%' OR PinYin Like '%" & Search & "%';"
        CreateList(Shown)
        parCount.InnerHtml = "<table align='center'><td>Displaying " & Shown.GetCount & " of " & Total.GetCount & " Phrases."
    End Sub
    
    Sub CreateList(Counting As Counter)
        CMD = New OleDbCommand(SQL, CN)
        CN.Open()
        Reader = CMD.ExecuteReader()
        HTML = "<table align='center'>"
        Do While Reader.Read()
            HTML = HTML & "<tr><td><a href='Details.aspx?PID=" & Reader("ID") & "'>" & Reader("EnglishText") & "</a></td><td>" & Reader("PinYin") & "</td></tr>"
            Counting.Up()
        Loop
        CN.Close()
        HTML = HTML & "</table>"
        parList.InnerHtml = HTML
    End Sub
</script>
<HTML>
    <head><title>Phrasebook</title></head>
    <body bgcolor="#FFF5FF">
        <form id="Form" runat="server">
            <div style="background-color:#E08D8A; border-width:2px; border-style:solid; border-color:#E55451; border-radius:8px;">
                <center><h1>Phrasebook - Phrases</h1>
                <h3><a href="Phrases.aspx">Phrases</a> | <a href="Favourites.aspx">Favourites</a> | <a href="Logout.aspx">Logout</a></h3></center>
            </div>
            <center><input id="txtSearch" type="text" runat="server"  />
            <input id="btnSearch" type="submit" value="Search" runat="server" /></center><br />
            <p id="parList" runat="server"></p>
            <p id="parCount" runat="server"></p>
        </form>
    </body>
</HTML>