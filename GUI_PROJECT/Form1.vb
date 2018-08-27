Public Class Form1
    Dim sqlString As String
    Dim da As OleDb.OleDbDataAdapter
    Dim coonsstring As String = "provider=Microsoft.ace.OleDb.12.0; data source= ROADS.mdb"

    


    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        Dim cn As New OleDb.OleDbConnection
        cn.ConnectionString = coonsstring
        cn.Open()
        Dim mycmd As New OleDb.OleDbCommand
        mycmd.Connection = cn
        mycmd.CommandText = "Insert into DailySiteDiary(Contract_Name) values ('" & TextBox1.Text & "')"
        mycmd.ExecuteNonQuery()
        cn.Close()
        MsgBox("DONE SAVING")
    End Sub
   

   
   


    
   
 











End Class
