'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2005 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



Public Class lookup_tables
    Inherits System.Windows.Forms.Form
    '////////////////////////////////////////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////
    ' The DataSet that holds the data.
    Private m_DataSet As DataSet

    ' A DataView to allow read-only access the data.
    Private m_DataView As DataView

    ' Load the data.

    Dim connect_string As String
    Private select_string As String
    Private data_adapter As OleDb.OleDbDataAdapter




#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents add_pic As System.Windows.Forms.PictureBox
    Friend WithEvents edit_pic As System.Windows.Forms.PictureBox
    Friend WithEvents delete As System.Windows.Forms.PictureBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents user As System.Windows.Forms.RadioButton
    Friend WithEvents pc As System.Windows.Forms.RadioButton
    Friend WithEvents room As System.Windows.Forms.RadioButton
    Friend WithEvents technician As System.Windows.Forms.RadioButton
    Friend WithEvents delete_link As System.Windows.Forms.LinkLabel
    Friend WithEvents edit_link As System.Windows.Forms.LinkLabel
    Friend WithEvents add_link As System.Windows.Forms.LinkLabel
    Friend WithEvents close_button As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(lookup_tables))
        Me.DataGrid1 = New System.Windows.Forms.DataGrid
        Me.Label1 = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.delete_link = New System.Windows.Forms.LinkLabel
        Me.edit_link = New System.Windows.Forms.LinkLabel
        Me.add_pic = New System.Windows.Forms.PictureBox
        Me.edit_pic = New System.Windows.Forms.PictureBox
        Me.delete = New System.Windows.Forms.PictureBox
        Me.add_link = New System.Windows.Forms.LinkLabel
        Me.Label5 = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.technician = New System.Windows.Forms.RadioButton
        Me.room = New System.Windows.Forms.RadioButton
        Me.pc = New System.Windows.Forms.RadioButton
        Me.user = New System.Windows.Forms.RadioButton
        Me.close_button = New System.Windows.Forms.Button
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'DataGrid1
        '
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(216, 8)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.Size = New System.Drawing.Size(504, 312)
        Me.DataGrid1.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Blue
        Me.Label1.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(192, 23)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Tasks"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Panel1.Controls.Add(Me.delete_link)
        Me.Panel1.Controls.Add(Me.edit_link)
        Me.Panel1.Controls.Add(Me.add_pic)
        Me.Panel1.Controls.Add(Me.edit_pic)
        Me.Panel1.Controls.Add(Me.delete)
        Me.Panel1.Controls.Add(Me.add_link)
        Me.Panel1.Location = New System.Drawing.Point(8, 32)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(192, 88)
        Me.Panel1.TabIndex = 6
        '
        'delete_link
        '
        Me.delete_link.Location = New System.Drawing.Point(32, 56)
        Me.delete_link.Name = "delete_link"
        Me.delete_link.Size = New System.Drawing.Size(40, 16)
        Me.delete_link.TabIndex = 11
        Me.delete_link.TabStop = True
        Me.delete_link.Text = "Delete"
        '
        'edit_link
        '
        Me.edit_link.Location = New System.Drawing.Point(32, 32)
        Me.edit_link.Name = "edit_link"
        Me.edit_link.Size = New System.Drawing.Size(24, 16)
        Me.edit_link.TabIndex = 10
        Me.edit_link.TabStop = True
        Me.edit_link.Text = "Edit"
        '
        'add_pic
        '
        Me.add_pic.Image = CType(resources.GetObject("add_pic.Image"), System.Drawing.Image)
        Me.add_pic.Location = New System.Drawing.Point(8, 8)
        Me.add_pic.Name = "add_pic"
        Me.add_pic.Size = New System.Drawing.Size(16, 16)
        Me.add_pic.TabIndex = 7
        Me.add_pic.TabStop = False
        '
        'edit_pic
        '
        Me.edit_pic.Image = CType(resources.GetObject("edit_pic.Image"), System.Drawing.Image)
        Me.edit_pic.Location = New System.Drawing.Point(8, 32)
        Me.edit_pic.Name = "edit_pic"
        Me.edit_pic.Size = New System.Drawing.Size(16, 16)
        Me.edit_pic.TabIndex = 8
        Me.edit_pic.TabStop = False
        '
        'delete
        '
        Me.delete.Image = CType(resources.GetObject("delete.Image"), System.Drawing.Image)
        Me.delete.Location = New System.Drawing.Point(8, 56)
        Me.delete.Name = "delete"
        Me.delete.Size = New System.Drawing.Size(16, 16)
        Me.delete.TabIndex = 9
        Me.delete.TabStop = False
        '
        'add_link
        '
        Me.add_link.Location = New System.Drawing.Point(32, 8)
        Me.add_link.Name = "add_link"
        Me.add_link.Size = New System.Drawing.Size(24, 16)
        Me.add_link.TabIndex = 9
        Me.add_link.TabStop = True
        Me.add_link.Text = "Add"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.Blue
        Me.Label5.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label5.Location = New System.Drawing.Point(8, 136)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(192, 23)
        Me.Label5.TabIndex = 7
        Me.Label5.Text = "Tables"
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.Panel2.Controls.Add(Me.technician)
        Me.Panel2.Controls.Add(Me.room)
        Me.Panel2.Controls.Add(Me.pc)
        Me.Panel2.Controls.Add(Me.user)
        Me.Panel2.Location = New System.Drawing.Point(8, 160)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(192, 160)
        Me.Panel2.TabIndex = 8
        '
        'technician
        '
        Me.technician.Location = New System.Drawing.Point(16, 136)
        Me.technician.Name = "technician"
        Me.technician.Size = New System.Drawing.Size(80, 16)
        Me.technician.TabIndex = 3
        Me.technician.Text = "Technician"
        '
        'room
        '
        Me.room.Location = New System.Drawing.Point(16, 96)
        Me.room.Name = "room"
        Me.room.Size = New System.Drawing.Size(56, 16)
        Me.room.TabIndex = 2
        Me.room.Text = "Room"
        '
        'pc
        '
        Me.pc.Location = New System.Drawing.Point(16, 16)
        Me.pc.Name = "pc"
        Me.pc.Size = New System.Drawing.Size(56, 16)
        Me.pc.TabIndex = 1
        Me.pc.Text = "PC"
        '
        'user
        '
        Me.user.Location = New System.Drawing.Point(16, 56)
        Me.user.Name = "user"
        Me.user.Size = New System.Drawing.Size(56, 16)
        Me.user.TabIndex = 0
        Me.user.Text = "Users"
        '
        'close_button
        '
        Me.close_button.Location = New System.Drawing.Point(640, 328)
        Me.close_button.Name = "close_button"
        Me.close_button.TabIndex = 9
        Me.close_button.Text = "Close"
        '
        'lookup_tables
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.Cornsilk
        Me.ClientSize = New System.Drawing.Size(744, 366)
        Me.ControlBox = False
        Me.Controls.Add(Me.close_button)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DataGrid1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "lookup_tables"
        Me.Text = "Lookup Tables"
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub lookup_tables_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim main As New switch
        connect_string = main.mystream

        pc.Checked = True



        ' Create the OleDbDataAdapter.
        select_string = "SELECT * FROM PC"

        data_adapter = New OleDb.OleDbDataAdapter(select_string, connect_string)

        ' Map Table to Contacts.
        data_adapter.TableMappings.Add("Tables", "PC")

        ' Fill the DataSet.
        m_DataSet = New DataSet
        data_adapter.Fill(m_DataSet, "PC")
        ' Make a DataView for the data.
        m_DataView = New DataView(m_DataSet.Tables("PC"))
        m_DataView.AllowDelete = False
        m_DataView.AllowEdit = False
        m_DataView.AllowNew = False



        ' Bind the DataGrid control to the DataView.
        DataGrid1.DataSource = m_DataView
    End Sub

    Private Sub delete_link_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs)

        ' create a connection string 
        Dim connString As String = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=D:\Documents and Settings\megatron\My Documents\db1.mdb"
        Dim myConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection
        myConnection.ConnectionString = connString
        ' create a data adapter 
        Dim da As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter("Select * from customers", myConnection)
        ' create a new dataset 
        Dim ds As DataSet = New DataSet("MS_Access_Dataset")
        ' fill dataset 
        da.Fill(ds, "customers")
        Dim current_row As Integer = DataGrid1.CurrentRowIndex
        'MessageBox.Show(current_row, "TEST", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        'Set the SelectCommand property for the DataAdapter
        da.SelectCommand = New OleDb.OleDbCommand
        da.SelectCommand.CommandText = "SELECT * FROM Customer"
        'Create a CommandBuilder object and assign MyDataAdapter
        'to be the value of its DataAdapter property
        Dim MyCommandBuilder As New OleDb.OleDbCommandBuilder
        MyCommandBuilder.DataAdapter = da
        Dim NumberOfRowsUpdated As Integer
        NumberOfRowsUpdated = da.Update(ds, "customers")


    End Sub

    Private Sub customer_delete()
        '/////////////////////////////////////////////////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////////
        Dim main As New switch
        connect_string = main.mystream

        ' The DataSet that holds the data.
        Dim m_DataSet As DataSet

        ' A DataView to allow read-only access the data.
        Dim m_DataView As DataView

        ' Load the data.

        Dim select_string As String

        Dim data_adapter As OleDb.OleDbDataAdapter

        ' Compose the database file name.
        ' Modify this if the database is somewhere else.
        ' Compose the connect string. 


        ' Create the OleDbDataAdapter.
        select_string = "SELECT * FROM customers"

        data_adapter = New OleDb.OleDbDataAdapter(select_string, connect_string)

        ' Map Table to Contacts.
        data_adapter.TableMappings.Add("Tables", "customers")

        ' Fill the DataSet.
        m_DataSet = New DataSet
        data_adapter.Fill(m_DataSet, "customers")
        ' Make a DataView for the data.
        m_DataView = New DataView(m_DataSet.Tables("customers"))
        m_DataView.AllowDelete = True
        m_DataView.AllowEdit = False
        m_DataView.AllowNew = False



        ' Bind the DataGrid control to the DataView.
        DataGrid1.DataSource = m_DataView

        '/////////////////////////////////////////////////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////////

    End Sub


    Private Sub user_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles user.CheckedChanged
        '/////////////////////////////////////////////////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////////
        Dim main As New switch
        connect_string = main.mystream
        ' The DataSet that holds the data.
        Dim m_DataSet As DataSet

        ' A DataView to allow read-only access the data.
        Dim m_DataView As DataView

        ' Load the data.

        Dim select_string As String

        Dim data_adapter As OleDb.OleDbDataAdapter

        ' Compose the database file name.
        ' Modify this if the database is somewhere else.
        ' Compose the connect string. 




        ' Create the OleDbDataAdapter.
        select_string = "SELECT * FROM customers"

        data_adapter = New OleDb.OleDbDataAdapter(select_string, connect_string)

        ' Map Table to Contacts.
        data_adapter.TableMappings.Add("Tables", "customers")

        ' Fill the DataSet.
        m_DataSet = New DataSet
        data_adapter.Fill(m_DataSet, "customers")
        ' Make a DataView for the data.
        m_DataView = New DataView(m_DataSet.Tables("customers"))
        m_DataView.AllowDelete = False
        m_DataView.AllowEdit = False
        m_DataView.AllowNew = False



        ' Bind the DataGrid control to the DataView.
        DataGrid1.DataSource = m_DataView
        '/////////////////////////////////////////////////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////////

    End Sub

    Private Sub pc_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pc.CheckedChanged
        '/////////////////////////////////////////////////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////////
        Dim main As New switch
        connect_string = main.mystream

        ' The DataSet that holds the data.
        Dim m_DataSet As DataSet

        ' A DataView to allow read-only access the data.
        Dim m_DataView As DataView

        ' Load the data.

        Dim select_string As String

        Dim data_adapter As OleDb.OleDbDataAdapter

        ' Compose the database file name.
        ' Modify this if the database is somewhere else.
        ' Compose the connect string. 




        ' Create the OleDbDataAdapter.
        select_string = "SELECT * FROM PC"

        data_adapter = New OleDb.OleDbDataAdapter(select_string, connect_string)

        ' Map Table to Contacts.
        data_adapter.TableMappings.Add("Tables", "PC")

        ' Fill the DataSet.
        m_DataSet = New DataSet
        data_adapter.Fill(m_DataSet, "PC")
        ' Make a DataView for the data.
        m_DataView = New DataView(m_DataSet.Tables("PC"))
        m_DataView.AllowDelete = False
        m_DataView.AllowEdit = False
        m_DataView.AllowNew = False



        ' Bind the DataGrid control to the DataView.
        DataGrid1.DataSource = m_DataView
        '/////////////////////////////////////////////////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////////

    End Sub

    Private Sub room_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles room.CheckedChanged
        '/////////////////////////////////////////////////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////////
        Dim main As New switch
        connect_string = main.mystream

        ' The DataSet that holds the data.
        Dim m_DataSet As DataSet

        ' A DataView to allow read-only access the data.
        Dim m_DataView As DataView

        ' Load the data.

        Dim select_string As String

        Dim data_adapter As OleDb.OleDbDataAdapter

        ' Compose the database file name.
        ' Modify this if the database is somewhere else.
        ' Compose the connect string. 




        ' Create the OleDbDataAdapter.
        select_string = "SELECT * FROM room"

        data_adapter = New OleDb.OleDbDataAdapter(select_string, connect_string)

        ' Map Table to Contacts.
        data_adapter.TableMappings.Add("Tables", "room")

        ' Fill the DataSet.
        m_DataSet = New DataSet
        data_adapter.Fill(m_DataSet, "room")
        ' Make a DataView for the data.
        m_DataView = New DataView(m_DataSet.Tables("room"))
        m_DataView.AllowDelete = False
        m_DataView.AllowEdit = False
        m_DataView.AllowNew = False



        ' Bind the DataGrid control to the DataView.
        DataGrid1.DataSource = m_DataView
        '/////////////////////////////////////////////////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////////

    End Sub

    Private Sub technician_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles technician.CheckedChanged
        '/////////////////////////////////////////////////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////////
        Dim main As New switch
        connect_string = main.mystream
        ' The DataSet that holds the data.
        Dim m_DataSet As DataSet

        ' A DataView to allow read-only access the data.
        Dim m_DataView As DataView

        ' Load the data.

        Dim select_string As String

        Dim data_adapter As OleDb.OleDbDataAdapter

        ' Compose the database file name.
        ' Modify this if the database is somewhere else.
        ' Compose the connect string.



        ' Create the OleDbDataAdapter.
        select_string = "SELECT * FROM technician"

        data_adapter = New OleDb.OleDbDataAdapter(select_string, connect_string)

        ' Map Table to Contacts.
        data_adapter.TableMappings.Add("Tables", "technician")

        ' Fill the DataSet.
        m_DataSet = New DataSet
        data_adapter.Fill(m_DataSet, "technician")
        ' Make a DataView for the data.
        m_DataView = New DataView(m_DataSet.Tables("technician"))
        m_DataView.AllowDelete = False
        m_DataView.AllowEdit = False
        m_DataView.AllowNew = False



        ' Bind the DataGrid control to the DataView.
        DataGrid1.DataSource = m_DataView
        '/////////////////////////////////////////////////////////////////////////////////////
        '////////////////////////////////////////////////////////////////////////////////////

    End Sub



    Private Sub delete_link_LinkClicked_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles delete_link.LinkClicked
        If user.Checked = True Then
            delete_task("delete_user", "@user_id", "customers")
        ElseIf pc.Checked = True Then
            delete_task("delete_pc", "@pc_id", "pc")
        ElseIf room.Checked = True Then
            delete_task("delete_room", "@room_id", "room")
        ElseIf technician.Checked = True Then
            delete_task("delete_tech", "@tech_id", "technician")
        End If
    End Sub

    Private Sub delete_task(ByVal sql_query As String, ByVal column_val As String, ByVal table As String)

        'Set the SQL string
        Dim objData As New DALBase
        Dim main As switch
        Dim connect_string As String
        connect_string = main.mystream

        objData.SQL = sql_query
        'Initialize the Command object
        objData.InitializeCommand()
        'Add the Parameters to the Parameters collection
        objData.AddParameter(column_val, _
            OleDb.OleDbType.Integer, 16, DataGrid1.Item(DataGrid1.CurrentRowIndex, 0))

        objData.Connection.ConnectionString = connect_string
        'Open the database connection
        Try
            objData.OpenConnection()
            'Execute the query
            Dim intRowsAffected As Integer = objData.Command.ExecuteNonQuery()
            'Close the database connection
            objData.CloseConnection()
            If intRowsAffected = 0 Then
                Throw New Exception("Delete Task Failed")
            End If
        Catch ex As Exception
            MessageBox.Show("This record has delete protection because it has relations with one or more tables.", "Relation Integrity Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try


        restate(table)

    End Sub


    Private Sub add_link_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles add_link.LinkClicked
        Dim table As String
        If user.Checked = True Then
            Dim form As New user_add
            form.ShowDialog(Me)
            table = "customers"
        ElseIf pc.Checked = True Then
            Dim form As New pc_add
            form.ShowDialog(Me)
            table = "PC"
        ElseIf room.Checked = True Then
            Dim form As New room_add
            form.ShowDialog(Me)
            table = "room"
        ElseIf technician.Checked = True Then
            Dim form As New tech_add
            form.ShowDialog(Me)
            table = "technician"
        End If
        restate(table)

    End Sub



    Private Sub edit_link_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles edit_link.LinkClicked
        Dim table As String
        If user.Checked = True Then
            Dim form As New user_edit1
            form.name1.Text = DataGrid1.Item(DataGrid1.CurrentRowIndex, 1)
            form.dept_text.Text = DataGrid1.Item(DataGrid1.CurrentRowIndex, 2)
            form.ShowDialog(Me)
            table = "customers"
        ElseIf pc.Checked = True Then
            Dim form As New pc_edit
            form.name1.Text = DataGrid1.Item(DataGrid1.CurrentRowIndex, 1)
            form.id1.Text = DataGrid1.Item(DataGrid1.CurrentRowIndex, 0)
            form.tag_text.Text = DataGrid1.Item(DataGrid1.CurrentRowIndex, 2)
            form.mac_address.Text = DataGrid1.Item(DataGrid1.CurrentRowIndex, 3)
            form.name1.Enabled = False
            form.ShowDialog(Me)
            table = "PC"
        ElseIf room.Checked = True Then
            MessageBox.Show("Edit is not available for this table", "STOP", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            table = "room"
        ElseIf technician.Checked = True Then
            MessageBox.Show("Edit is not available for this table", "STOP", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            table = "technician"
        End If

        restate(table)

    End Sub
    Private Sub restate(ByVal table As String)
        ' Create the OleDbDataAdapter.
        Dim main As New switch
        connect_string = main.mystream

        select_string = "SELECT * FROM " & table & ""


        data_adapter = New OleDb.OleDbDataAdapter(select_string, connect_string)

        ' Map Table to Contacts.
        data_adapter.TableMappings.Add("Tables", table)

        ' Fill the DataSet.
        m_DataSet = New DataSet
        data_adapter.Fill(m_DataSet, table)
        ' Make a DataView for the data.
        m_DataView = New DataView(m_DataSet.Tables(table))
        m_DataView.AllowDelete = False
        m_DataView.AllowEdit = False
        m_DataView.AllowNew = False



        ' Bind the DataGrid control to the DataView.
        DataGrid1.DataSource = m_DataView
    End Sub

    Private Sub formload()
        pc.Checked = True
        Dim main As New switch
        connect_string = main.mystream

        ' Create the OleDbDataAdapter.
        select_string = "SELECT * FROM PC"

        data_adapter = New OleDb.OleDbDataAdapter(select_string, connect_string)

        ' Map Table to Contacts.
        data_adapter.TableMappings.Add("Tables", "PC")

        ' Fill the DataSet.
        m_DataSet = New DataSet
        data_adapter.Fill(m_DataSet, "PC")
        ' Make a DataView for the data.
        m_DataView = New DataView(m_DataSet.Tables("PC"))
        m_DataView.AllowDelete = False
        m_DataView.AllowEdit = False
        m_DataView.AllowNew = False



        ' Bind the DataGrid control to the DataView.
        DataGrid1.DataSource = m_DataView
    End Sub

    Private Sub close_button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles close_button.Click
        Me.Close()

    End Sub
End Class

