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

    End Sub
   

   
   


    
   
 








    Private Sub TextBox16_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox16.TextChanged

    End Sub

    Private Sub TextBox105_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox105.TextChanged

    End Sub

    Private Sub TextBox104_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox104.TextChanged

    End Sub

    Private Sub TextBox117_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox117.TextChanged

    End Sub

    Private Sub TextBox13_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox13.TextChanged

    End Sub

    Private Sub TextBox119_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox119.TextChanged

    End Sub

    Private Sub TextBox15_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox77_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox82_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox83_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class
