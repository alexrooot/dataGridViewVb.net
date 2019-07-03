'Imports for data access classes 
Imports System.Data.OleDb
Public Class Form2
    Inherits System.Windows.Forms.Form

    'Call LoadDBTables when second form displays
    Private Sub Form2_Load _
    (ByVal sender As System.Object,
    ByVal e As System.EventArgs) _
    Handles MyBase.Load
        LoadDBTables()
        btnShowFields.Enabled = False


    End Sub

    'Loads lstTables ListBox with names of database tables
    'Will be explained in Chapter 4
    Sub LoadDBTables()
        Dim frm1 As Form1 'reference to the current instance of the main form
        frm1 = CType(Me.Owner, Form1) ' give it a specifc Form to that main form
        Dim conn As OleDbConnection = frm1.myConn ' reference to OleDbConnection
        Dim obj() As Object = New Object() {Nothing, Nothing, Nothing, "TABLE"} 'Obtaining the Database Table Information so clear out any data fileds/query
        Dim dt As DataTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, obj) 'Obtaining the Database Table Information


        'Displaying the Database Table Information

        'The DataTable returned by the GetOleDbSchemaTable function Is then used as 
        'the DataSource property for the ListBox lstTables.
        lstTables.DataSource = dt

        'sets the value that we want to display in the lstTables
        lstTables.DisplayMember =
        dt.Columns("TABLE_NAME").ToString()
        lstTables.ClearSelected()


    End Sub

    'Calls LoadDBFields if database table selected
    'Otherwise warns user no database table selected 
    Private Sub btnShowFields_Click _
    (ByVal sender As Object,
    ByVal e As System.EventArgs) _
    Handles btnShowFields.Click
        If lstTables.SelectedItems.Count > 0 Then
            LoadDBFields()
        Else
            MessageBox.Show("Must select table first")
        End If
    End Sub

    'Loads lstFields ListBox with names of table fields
    'Will be explained in Chapter 4
    Sub LoadDBFields() '
        Dim frm1 As Form1 ' again find what db was selected by indicating form parent
        frm1 = CType(Me.Owner, Form1) ' get that specific form parent
        Dim conn As OleDbConnection = frm1.myConn ' local name of db information
        Dim obj() As Object = New Object() _
         {Nothing, Nothing, lstTables.Text, Nothing} ' clear query information before diplaying

        'GetOleDbSchemaTable is OleDbSchemaGuid.Columns instead of OleDbSchemaGuid.Tables.
        Dim dt As DataTable =
            conn.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, obj)
        lstFields.DataSource = dt
        lstFields.DisplayMember = dt.Columns("COLUMN_NAME").ToString() ' populate the listview
        lstFields.ClearSelected()
    End Sub

    Private Sub lstTables_Click(sender As Object, e As EventArgs) Handles lstTables.Click
        btnShowFields.Enabled = True
    End Sub
End Class