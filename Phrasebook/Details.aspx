<%@ Page Language="VB" %>
<%@ Import Namespace="System.Data.OleDb" %>
<script runat="server">
    Dim CS As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + _
                        "Data Source=" + Server.MapPath("Phrasebook.accdb") + ";"
    Dim CN As New OleDbConnection(CS)
    Dim CMD As OleDbCommand
    Dim Reader As OleDbDataReader
    Dim SQL As String
    Dim PID As Single
    Dim HTML As String
    Dim PinYinSnd As String
    
    Sub Page_Load(s As Object, e As EventArgs)
        If Session("Logged") = 0 Then
            Response.Redirect("Login.aspx")
        End If
        PID = Request.QueryString("PID")
        SQL = "SELECT * FROM Phrase WHERE ID = " & PID & ";"
        CMD = New OleDbCommand(SQL, CN)
        CN.Open()
        Reader = CMD.ExecuteReader()
        HTML = ""
        Reader.Read()
        HTML = "PinYin: " & Reader("PinYin") & "<br />English: " & Reader("EnglishText")
        imgPhrase.Src = Reader("Picture")
        PinYinSnd = Reader("Sound")
        parPhrase.InnerHtml = HTML
        
        SQL = "SELECT * FROM Favourite, Student WHERE Favourite.StudentID = Student.ID AND Student.ID = " & Session("Logged") & " AND Favourite.PhraseID = " & PID & ";"
        CMD = New OleDbCommand(SQL, CN)
        Reader = CMD.ExecuteReader()
        If Reader.Read() Then
            ibFave.ImageUrl = "Yes.jpg"
        Else
            ibFave.ImageUrl = "No.jpg"
        End If
        CN.Close()
    End Sub
    
    Sub btnPlay_Click(s As Object, e As EventArgs) Handles btnPlay.ServerClick
        parBGSound.InnerHtml = "<bgsound id='sndPlayer' src='" & PinYinSnd & "' runat='server'></bgsound>"
    End Sub
    
    Sub ibFave_Click(s As Object, e As ImageClickEventArgs) Handles ibFave.Click
        parBGSound.InnerHtml = ""
        CN.Open()
        If ibFave.ImageUrl = "Yes.jpg" Then
            SQL = "DELETE FROM Favourite WHERE StudentID = " & Session("Logged") & " AND PhraseID = " & PID & ";"
            AddOrDelete()
            ibFave.ImageUrl = "No.jpg"
        ElseIf ibFave.ImageUrl = "No.jpg" Then
            SQL = "INSERT INTO Favourite (StudentID, PhraseID) VALUES (" & Session("Logged") & ", " & PID & ");"
            AddOrDelete()
            ibFave.ImageUrl = "Yes.jpg"
        End If
        CN.Close()
    End Sub
    
    Sub AddOrDelete()
        CMD = New OleDbCommand(SQL, CN)
        CMD.ExecuteNonQuery()
    End Sub
    
</script>
<HTML>
    <head>
        <title>Phrase Details</title>
    </head>
    <body bgcolor="#FFF5FF">
        <form id="Form" runat="server">
            <div style="background-color:#E08D8A; border-width:2px; border-style:solid; border-color:#E55451; border-radius:8px;">
                <center><h1>Phrasebook - Phrase Details</h1>
                <h3><a href="Phrases.aspx">Phrases</a> | <a href="Favourites.aspx">Favourites</a> | <a href="Logout.aspx">Logout</a></h3></center>
            </div>
            <table align='center'>
                <tr><td><p id="parPhrase" runat="server" ></p>
                    <input id="btnPlay" type="submit" value="Play" runat="server" /><br />
                    Favourite:<asp:ImageButton ID="ibFave" width="25" height="25" runat="server" /></td>
                    <td><img id='imgPhrase' src=".jpg" width="250" height="250" runat='server' /></td>
                </tr>
            </table>
            <p id="parBGSound" runat="server"></p>
        </form>
    </body>
</HTML>