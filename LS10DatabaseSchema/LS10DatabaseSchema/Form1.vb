'Imports for data access classes 
Imports System.Data.OleDb
'Imports for StringBuilder class
'Used in Click event of btnChange
Imports System.Text
Public Class Form1
    Inherits System.Windows.Forms.Form

    'Assumes Microsoft Access database
    Const strProvider As String = "Microsoft.Jet.OLEDB.4.0"
    'Using persistent connection as in Lesson 9
    'Public so can be accessed in second form code
    Public myConn As New OleDbConnection

    'Call Create Connection procedure at startup
    Private Sub Form1_Load _
    (ByVal sender As System.Object,
    ByVal e As System.EventArgs) _
    Handles MyBase.Load
        CreateConnection()
    End Sub

    'Create database connection
    'Essentially same code as in Form_Load event in Lesson 9 
    Private Sub CreateConnection()
        Dim dr As DialogResult
        dr = dlgOpen.ShowDialog()
        If dr = Windows.Forms.DialogResult.OK Then
            Dim strFile As String = dlgOpen.FileName

            'the next command is setting up the connection to the db via a string there are many like slq as below
            'EX: Dim connectionString AS String = "Server=my_server;Database=name_of_db;User Id=user_name;Password=my_password"
            'EX:  Dim connectionString AS String = Server=myServerAddress;Database=myDataBase;Trusted_Connection=True;
            myConn.ConnectionString = "Provider=" &
            strProvider & ";Data Source=" & strFile & ";"
            myConn.Open()
            'End application if user chooses Cancel
            'Uses Exit method of Application object
        Else
            MessageBox.Show("No database connection . closing")
            Application.Exit()
        End If
    End Sub

    'Creates OleDbCommand
    'Argument is SQL SELECT statement
    'Essentially same code as in Form_Load event in Lesson 9 
    Private Function CreateCommand _
    (ByVal strComm As String) As OleDbCommand
        Dim myCMD As OleDbCommand = New OleDbCommand
        myCMD.CommandText = strComm  'CommandText needed to retrieve the following fields
        myCMD.Connection = myConn
        Return myCMD
    End Function

    'Creates DataSet
    'First argument is OleDbCommand created by CreateCommand
    'Second argument is name of table
    'Essentially same code as in Form_Load event in Lesson 9 
    Private Function CreateDataSet _
    (ByVal myCmd As OleDbCommand,
    ByVal strTable As String) As DataSet
        Dim myAdapter As OleDbDataAdapter = New OleDbDataAdapter
        myAdapter.SelectCommand = myCmd
        Dim ds As DataSet = New DataSet
        ds.Clear()
        myAdapter.Fill(ds, strTable)
        Return ds
    End Function



    'Displaying the Selected Fields in the DataGridView
    'Beginning code shows second form
    Private Sub btnChange_Click _
    (ByVal sender As System.Object,
    ByVal e As System.EventArgs) _
    Handles btnChange.Click ' get access to form2
        Dim frm As New Form2 ' instance refrence to from2

        ' make sure the user clicked the OK button, selected one 
        ' table, And selected at least one field from the ListBoxes
        If frm.ShowDialog(Me) = Windows.Forms.DialogResult.OK And
        frm.lstTables.SelectedItems.Count > 0 And
        frm.lstFields.SelectedItems.Count > 0 Then
            Dim sb As New StringBuilder
            sb.Append("SELECT ") 'build an SQL SELECT statement from the table
            Dim obj As DataRowView

            'For Each Next loop to iterate through each selected item 
            'in the SelectedItems collection.
            For Each obj In frm.lstFields.SelectedItems
                sb.Append(obj("COLUMN_NAME").ToString())
                sb.Append(", ")
            Next

            'The Remove method of the StringBuilder object in the following 
            'code eliminates that trailing comma And space. The first 
            'argument Is the starting point For removal, the second the 
            'Number of characters to remove
            sb.Remove((sb.Length - 2), 2)


            'add to the StringBuilder object the FROM keyword followed 
            'by the name of the selected table.
            sb.Append(" FROM " & frm.lstTables.Text)


            'contents of the StringBuilder object are assigned to a String 
            Dim strCommand As String = sb.ToString()

            'SQL SELECT statement is now passed as an argument to the 
            'CreateCommand procedure to create an OleDbCommand
            Dim myCMD As OleDbCommand = CreateCommand(strCommand)

            ' OleDBCommand variable and the name of the selected table are 
            ' passed to the CreateDataSet procedure to create a DataSet
            Dim dSet As DataSet = CreateDataSet(myCMD, frm.lstTables.Text)

            dgvData.DataSource = dSet ' DataSource property is set to the DataSet dSet,
            dgvData.DataMember = frm.lstTables.Text 'selected table frm.lstTables.Text
        End If
    End Sub

    'Close database connection when application closes
    Private Sub Form1_Closing(ByVal sender As Object,
    ByVal e As System.ComponentModel.CancelEventArgs) _
    Handles MyBase.Closing
        myConn.Close()
    End Sub
End Class