'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2005 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



Public Class loner_pc
    Inherits System.Windows.Forms.Form
    Dim loaner As Integer
    Dim name1 As String
    Dim tabs As Integer
    Dim check1 As Boolean = False
    Dim check2 As Boolean = False
    Dim check3 As Boolean = False
    Dim check4 As Boolean = False
    Dim check5 As Boolean = False
    Dim main As New switch
    Dim connect_string As String







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
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents in1 As System.Windows.Forms.Button
    Friend WithEvents out1 As System.Windows.Forms.Button
    Friend WithEvents loneruna As System.Windows.Forms.ListBox
    Friend WithEvents lonera As System.Windows.Forms.ListBox
    Friend WithEvents callnumber1 As System.Windows.Forms.TextBox
    Friend WithEvents callbox As System.Windows.Forms.ComboBox
    Friend WithEvents in2 As System.Windows.Forms.Button
    Friend WithEvents out2 As System.Windows.Forms.Button
    Friend WithEvents loneruna2 As System.Windows.Forms.ListBox
    Friend WithEvents lonera2 As System.Windows.Forms.ListBox
    Friend WithEvents ok As System.Windows.Forms.Button
    Friend WithEvents cancel As System.Windows.Forms.Button
    Friend WithEvents userbox1 As System.Windows.Forms.ComboBox
    Friend WithEvents roombox1 As System.Windows.Forms.ComboBox
    Friend WithEvents techbox1 As System.Windows.Forms.ComboBox
    Friend WithEvents date_issue As System.Windows.Forms.TextBox
    Friend WithEvents date_recieved As System.Windows.Forms.TextBox
    Friend WithEvents userbox2 As System.Windows.Forms.ComboBox
    Friend WithEvents roombox2 As System.Windows.Forms.ComboBox
    Friend WithEvents techbox2 As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.date_issue = New System.Windows.Forms.TextBox
        Me.techbox1 = New System.Windows.Forms.ComboBox
        Me.roombox1 = New System.Windows.Forms.ComboBox
        Me.userbox1 = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.in1 = New System.Windows.Forms.Button
        Me.out1 = New System.Windows.Forms.Button
        Me.loneruna = New System.Windows.Forms.ListBox
        Me.lonera = New System.Windows.Forms.ListBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.callnumber1 = New System.Windows.Forms.TextBox
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.techbox2 = New System.Windows.Forms.ComboBox
        Me.roombox2 = New System.Windows.Forms.ComboBox
        Me.userbox2 = New System.Windows.Forms.ComboBox
        Me.date_recieved = New System.Windows.Forms.TextBox
        Me.callbox = New System.Windows.Forms.ComboBox
        Me.in2 = New System.Windows.Forms.Button
        Me.out2 = New System.Windows.Forms.Button
        Me.loneruna2 = New System.Windows.Forms.ListBox
        Me.lonera2 = New System.Windows.Forms.ListBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.ok = New System.Windows.Forms.Button
        Me.cancel = New System.Windows.Forms.Button
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.HotTrack = True
        Me.TabControl1.Location = New System.Drawing.Point(8, 16)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(728, 328)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.Color.Beige
        Me.TabPage1.Controls.Add(Me.date_issue)
        Me.TabPage1.Controls.Add(Me.techbox1)
        Me.TabPage1.Controls.Add(Me.roombox1)
        Me.TabPage1.Controls.Add(Me.userbox1)
        Me.TabPage1.Controls.Add(Me.Label5)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.in1)
        Me.TabPage1.Controls.Add(Me.out1)
        Me.TabPage1.Controls.Add(Me.loneruna)
        Me.TabPage1.Controls.Add(Me.lonera)
        Me.TabPage1.Controls.Add(Me.Label4)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.Label7)
        Me.TabPage1.Controls.Add(Me.callnumber1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 25)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(720, 299)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Check Out"
        '
        'date_issue
        '
        Me.date_issue.Enabled = False
        Me.date_issue.Location = New System.Drawing.Point(128, 40)
        Me.date_issue.Name = "date_issue"
        Me.date_issue.Size = New System.Drawing.Size(192, 20)
        Me.date_issue.TabIndex = 34
        Me.date_issue.Text = ""
        '
        'techbox1
        '
        Me.techbox1.Location = New System.Drawing.Point(520, 64)
        Me.techbox1.Name = "techbox1"
        Me.techbox1.Size = New System.Drawing.Size(184, 21)
        Me.techbox1.TabIndex = 33
        '
        'roombox1
        '
        Me.roombox1.Location = New System.Drawing.Point(520, 40)
        Me.roombox1.Name = "roombox1"
        Me.roombox1.Size = New System.Drawing.Size(184, 21)
        Me.roombox1.TabIndex = 32
        '
        'userbox1
        '
        Me.userbox1.Location = New System.Drawing.Point(520, 16)
        Me.userbox1.Name = "userbox1"
        Me.userbox1.Size = New System.Drawing.Size(184, 21)
        Me.userbox1.TabIndex = 31
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(400, 64)
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 30
        Me.Label5.Text = "Technician"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(400, 40)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 29
        Me.Label1.Text = "Room"
        '
        'in1
        '
        Me.in1.Location = New System.Drawing.Point(320, 200)
        Me.in1.Name = "in1"
        Me.in1.TabIndex = 28
        Me.in1.Text = "<<"
        '
        'out1
        '
        Me.out1.Location = New System.Drawing.Point(320, 160)
        Me.out1.Name = "out1"
        Me.out1.TabIndex = 27
        Me.out1.Text = ">>"
        '
        'loneruna
        '
        Me.loneruna.Location = New System.Drawing.Point(400, 104)
        Me.loneruna.Name = "loneruna"
        Me.loneruna.Size = New System.Drawing.Size(320, 173)
        Me.loneruna.TabIndex = 26
        '
        'lonera
        '
        Me.lonera.Location = New System.Drawing.Point(0, 104)
        Me.lonera.Name = "lonera"
        Me.lonera.Size = New System.Drawing.Size(312, 173)
        Me.lonera.TabIndex = 25
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(400, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "User"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Date Issued"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(16, 16)
        Me.Label7.Name = "Label7"
        Me.Label7.TabIndex = 16
        Me.Label7.Text = "Call Number"
        '
        'callnumber1
        '
        Me.callnumber1.Location = New System.Drawing.Point(128, 16)
        Me.callnumber1.Name = "callnumber1"
        Me.callnumber1.Size = New System.Drawing.Size(192, 20)
        Me.callnumber1.TabIndex = 3
        Me.callnumber1.Text = ""
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.Color.Beige
        Me.TabPage2.Controls.Add(Me.techbox2)
        Me.TabPage2.Controls.Add(Me.roombox2)
        Me.TabPage2.Controls.Add(Me.userbox2)
        Me.TabPage2.Controls.Add(Me.date_recieved)
        Me.TabPage2.Controls.Add(Me.callbox)
        Me.TabPage2.Controls.Add(Me.in2)
        Me.TabPage2.Controls.Add(Me.out2)
        Me.TabPage2.Controls.Add(Me.loneruna2)
        Me.TabPage2.Controls.Add(Me.lonera2)
        Me.TabPage2.Controls.Add(Me.Label12)
        Me.TabPage2.Controls.Add(Me.Label11)
        Me.TabPage2.Controls.Add(Me.Label10)
        Me.TabPage2.Controls.Add(Me.Label8)
        Me.TabPage2.Controls.Add(Me.Label6)
        Me.TabPage2.Location = New System.Drawing.Point(4, 25)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(720, 299)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Check In"
        '
        'techbox2
        '
        Me.techbox2.Location = New System.Drawing.Point(520, 64)
        Me.techbox2.Name = "techbox2"
        Me.techbox2.Size = New System.Drawing.Size(184, 21)
        Me.techbox2.TabIndex = 41
        '
        'roombox2
        '
        Me.roombox2.Location = New System.Drawing.Point(520, 40)
        Me.roombox2.Name = "roombox2"
        Me.roombox2.Size = New System.Drawing.Size(184, 21)
        Me.roombox2.TabIndex = 40
        '
        'userbox2
        '
        Me.userbox2.Location = New System.Drawing.Point(520, 16)
        Me.userbox2.Name = "userbox2"
        Me.userbox2.Size = New System.Drawing.Size(184, 21)
        Me.userbox2.TabIndex = 39
        '
        'date_recieved
        '
        Me.date_recieved.Enabled = False
        Me.date_recieved.Location = New System.Drawing.Point(128, 40)
        Me.date_recieved.Name = "date_recieved"
        Me.date_recieved.Size = New System.Drawing.Size(192, 20)
        Me.date_recieved.TabIndex = 35
        Me.date_recieved.Text = ""
        '
        'callbox
        '
        Me.callbox.Location = New System.Drawing.Point(128, 16)
        Me.callbox.Name = "callbox"
        Me.callbox.Size = New System.Drawing.Size(192, 21)
        Me.callbox.TabIndex = 29
        '
        'in2
        '
        Me.in2.Location = New System.Drawing.Point(320, 200)
        Me.in2.Name = "in2"
        Me.in2.TabIndex = 28
        Me.in2.Text = "<<"
        '
        'out2
        '
        Me.out2.Location = New System.Drawing.Point(320, 160)
        Me.out2.Name = "out2"
        Me.out2.TabIndex = 27
        Me.out2.Text = ">>"
        '
        'loneruna2
        '
        Me.loneruna2.Location = New System.Drawing.Point(400, 104)
        Me.loneruna2.Name = "loneruna2"
        Me.loneruna2.Size = New System.Drawing.Size(320, 173)
        Me.loneruna2.TabIndex = 26
        '
        'lonera2
        '
        Me.lonera2.Location = New System.Drawing.Point(0, 104)
        Me.lonera2.Name = "lonera2"
        Me.lonera2.Size = New System.Drawing.Size(312, 173)
        Me.lonera2.TabIndex = 25
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(408, 64)
        Me.Label12.Name = "Label12"
        Me.Label12.TabIndex = 24
        Me.Label12.Text = "Technician"
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(408, 40)
        Me.Label11.Name = "Label11"
        Me.Label11.TabIndex = 23
        Me.Label11.Text = "Room"
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(408, 16)
        Me.Label10.Name = "Label10"
        Me.Label10.TabIndex = 22
        Me.Label10.Text = "User"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(16, 40)
        Me.Label8.Name = "Label8"
        Me.Label8.TabIndex = 20
        Me.Label8.Text = "Date Recieved"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(16, 16)
        Me.Label6.Name = "Label6"
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "Call Number"
        '
        'ok
        '
        Me.ok.Location = New System.Drawing.Point(576, 360)
        Me.ok.Name = "ok"
        Me.ok.TabIndex = 1
        Me.ok.Text = "OK"
        '
        'cancel
        '
        Me.cancel.Location = New System.Drawing.Point(664, 360)
        Me.cancel.Name = "cancel"
        Me.cancel.TabIndex = 2
        Me.cancel.Text = "Cancel"
        '
        'loner_pc
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.Cornsilk
        Me.ClientSize = New System.Drawing.Size(752, 389)
        Me.ControlBox = False
        Me.Controls.Add(Me.cancel)
        Me.Controls.Add(Me.ok)
        Me.Controls.Add(Me.TabControl1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "loner_pc"
        Me.Text = "Loaner PC"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub loner_pc_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        connect_string = main.mystream

        tabs = 0
        loaner = 0
        Dim ThisMoment As Date
        ThisMoment = Now
        date_issue.Text = ThisMoment.Date

        date_recieved.Text = ThisMoment.Date
        '///////////////init///////////////////////////////
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



        '/////////////////////user combo box/////////////////////////////
        ' Create the OleDbDataAdapter.
        select_string = "SELECT * FROM customers"

        data_adapter = New OleDb.OleDbDataAdapter(select_string, connect_string)

        ' Map Table to Contacts.
        data_adapter.TableMappings.Add("Tables", "customers")

        ' Fill the DataSet.
        m_DataSet = New DataSet
        data_adapter.Fill(m_DataSet, "customers")

        userbox1.DataSource = m_DataSet.Tables("customers")
        userbox1.DisplayMember = "user"
        userbox2.DataSource = m_DataSet.Tables("customers")
        userbox2.DisplayMember = "user"
        '///////////////////////////room combo box/////////////////////////////////////////////////////////////////
        select_string = "SELECT * FROM room"

        data_adapter = New OleDb.OleDbDataAdapter(select_string, connect_string)

        ' Map Table to Contacts.
        data_adapter.TableMappings.Add("Tables", "room")

        ' Fill the DataSet.
        m_DataSet = New DataSet
        data_adapter.Fill(m_DataSet, "room")

        roombox1.DataSource = m_DataSet.Tables("room")
        roombox1.DisplayMember = "room"
        roombox2.DataSource = m_DataSet.Tables("room")
        roombox2.DisplayMember = "room"
        '/////////////////////////////technician combo box
        select_string = "SELECT * FROM technician"

        data_adapter = New OleDb.OleDbDataAdapter(select_string, connect_string)

        ' Map Table to Contacts.
        data_adapter.TableMappings.Add("Tables", "technician")

        ' Fill the DataSet.
        m_DataSet = New DataSet
        data_adapter.Fill(m_DataSet, "technician")

        techbox1.DataSource = m_DataSet.Tables("technician")
        techbox1.DisplayMember = "technician"
        techbox2.DataSource = m_DataSet.Tables("technician")
        techbox2.DisplayMember = "technician"
        '///////////////////////////listbox1/////////////////////////////////////
        select_string = "SELECT * FROM PC WHERE Available=TRUE"

        data_adapter = New OleDb.OleDbDataAdapter(select_string, connect_string)

        ' Map Table to Contacts.
        data_adapter.TableMappings.Add("Tables", "PC")

        ' Fill the DataSet.
        m_DataSet = New DataSet
        data_adapter.Fill(m_DataSet, "PC")

        lonera.DataSource = m_DataSet.Tables("PC")
        lonera.DisplayMember = "pc_name"
        lonera2.DataSource = m_DataSet.Tables("PC")
        lonera2.DisplayMember = "pc_name"

        '///////////////////////////////////////////////////////////////////////////////////////


        '////////////////////////////////call number//////////////////////////////////////////
        select_string = "SELECT * FROM check_in"

        data_adapter = New OleDb.OleDbDataAdapter(select_string, connect_string)

        ' Map Table to Contacts.
        data_adapter.TableMappings.Add("Tables", "check_in")

        ' Fill the DataSet.
        m_DataSet = New DataSet
        data_adapter.Fill(m_DataSet, "check_in")
        '/////////////////////////!_!_!_!_!_!_!_!_!_!_!_!_!_!_!_!_!_!_!_!
        callbox.DataSource = m_DataSet.Tables("check_in")
        callbox.DisplayMember = "call_number"
        '///////////////////////////////////////////////////////////////////////////////////
        '///////////////////////////////////////////////////////////////////////////////////
        '////////////////////////////////////unavailable list///////////////////////////////
        select_string = "SELECT * FROM PC WHERE Available=FALSE"

        data_adapter = New OleDb.OleDbDataAdapter(select_string, connect_string)

        ' Map Table to Contacts.
        data_adapter.TableMappings.Add("Tables", "PC")

        ' Fill the DataSet.
        m_DataSet = New DataSet
        data_adapter.Fill(m_DataSet, "PC")

        'loneruna.DataSource = m_DataSet.Tables("PC")
        'loneruna.DisplayMember = "pc_name"
        'loneruna2.DataSource = m_DataSet.Tables("PC")
        'loneruna2.DisplayMember = "pc_name"

    End Sub

    Private Sub out1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles out1.Click

        loaner = loaner + 1
        If loaner <= 1 Then
            out_process()

        Else
            MessageBox.Show("Only one loner pc per user, please!", "Stop_out", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            loaner = loaner - 1
        End If

    End Sub
    Private Sub update_database(ByVal test As String)

        '////no open connection
        'Set the SQL string
        Dim objData As New DALBase
        Dim connect_string As String
        connect_string = main.mystream
        objData.SQL = "talor"
        'Initialize the Command object
        objData.InitializeCommand()
        'Add the Parameters to the Parameters collection

        objData.AddParameter("@av", _
    OleDb.OleDbType.Boolean, 16, 0)
        objData.AddParameter("@unav", _
                    OleDb.OleDbType.Boolean, 16, 1)
        objData.AddParameter("@pc_name", _
                            OleDb.OleDbType.VarChar, 16, test)
        objData.Connection.ConnectionString = connect_string
        Try
            objData.OpenConnection()
            'Execute the query
            Dim intRowsAffected As Integer = objData.Command.ExecuteNonQuery
            'Close the database connection
            objData.CloseConnection()
            If intRowsAffected = 0 Then
                Throw New Exception("EDIT Task Failed")
            End If
        Catch ex As Exception
            MessageBox.Show("This record has delete protection because it has relations with one or more tables.", "Relation Integrity Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try

        update_database6(test)
        reset_list(test)


    End Sub

    Private Sub reset_list(ByVal temp_test)

        connect_string = main.mystream

        '///////////////init///////////////////////////////
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
        '///////////////////////////listbox/////////////////////////////////////
        select_string = "SELECT * FROM PC WHERE Available=TRUE"

        data_adapter = New OleDb.OleDbDataAdapter(select_string, connect_string)

        ' Map Table to Contacts.
        data_adapter.TableMappings.Add("Tables", "PC")

        ' Fill the DataSet.
        m_DataSet = New DataSet
        data_adapter.Fill(m_DataSet, "PC")

        lonera.DataSource = m_DataSet.Tables("PC")
        lonera.DisplayMember = "pc_name"
        lonera2.DataSource = m_DataSet.Tables("PC")
        lonera2.DisplayMember = "pc_name"

        '////////////////////////////////////unavailable list///////////////////////////////
        select_string = "SELECT * FROM pc_un"

        data_adapter = New OleDb.OleDbDataAdapter(select_string, connect_string)

        ' Map Table to Contacts.
        data_adapter.TableMappings.Add("Tables", "pc_un")

        ' Fill the DataSet.
        m_DataSet = New DataSet
        data_adapter.Fill(m_DataSet, "pc_un")

        loneruna.DataSource = m_DataSet.Tables("pc_un")
        loneruna.DisplayMember = "pc_name"
        loneruna2.DataSource = m_DataSet.Tables("pc_un")
        loneruna2.DisplayMember = "pc_name"

    End Sub

    Private Sub in1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles in1.Click

        connect_string = main.mystream
        Dim test As Integer
        test = loneruna.SelectedIndex


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
        select_string = "SELECT * FROM pc_un"

        data_adapter = New OleDb.OleDbDataAdapter(select_string, connect_string)

        ' Map Table to Contacts.
        data_adapter.TableMappings.Add("Tables", "PC")

        ' Fill the DataSet.

        m_DataSet = New DataSet
        data_adapter.Fill(m_DataSet, "PC")

        If loaner = 1 Then
            name1 = m_DataSet.Tables("PC").Rows(0).Item("pc_name")
            loaner = loaner - 1
            update_database2(name1)

        Else
            loaner = 2
            MessageBox.Show("You cannot add anymore PCs back into the queue", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        If loaner >= 0 And loaner < 2 Then
            update_database7()

        Else : loaner = 2
            loaner = 0
        End If

        reset_list(test)



    End Sub
    Private Sub update_database2(ByVal test As String)


        'Set the SQL string
        Dim objData As New DALBase
        Dim connect_string As String
        connect_string = main.mystream
        objData.SQL = "talor"
        'Initialize the Command object
        objData.InitializeCommand()
        'Add the Parameters to the Parameters collection

        objData.AddParameter("@av", _
    OleDb.OleDbType.Boolean, 16, 1)
        objData.AddParameter("@unav", _
                    OleDb.OleDbType.Boolean, 16, 0)
        objData.AddParameter("@pc_name", _
                            OleDb.OleDbType.VarChar, 16, test)
        objData.Connection.ConnectionString = connect_string

        Try
            objData.OpenConnection()
            'Execute the query
            Dim intRowsAffected As Integer = objData.Command.ExecuteNonQuery
            'Close the database connection
            objData.CloseConnection()
            If intRowsAffected = 0 Then
                Throw New Exception("EDIT Task Failed")
            End If
        Catch ex As Exception
        End Try
        update_database9()




    End Sub
    Private Sub update_database3()
        'Set the SQL string
        Dim objData As New DALBase
        Dim connect_string As String
        connect_string = main.mystream
        objData.SQL = "delete_checkin"
        'Initialize the Command object
        objData.InitializeCommand()
        'Add the Parameters to the Parameters collection
        objData.Connection.ConnectionString = connect_string

        objData.AddParameter("@call_number", _
        OleDb.OleDbType.VarChar, 16, callbox.Text)


        Try
            objData.OpenConnection()
            'Execute the query
            Dim intRowsAffected As Integer = objData.Command.ExecuteNonQuery
            'Close the database connection
            objData.CloseConnection()
            If intRowsAffected = 0 Then
                Throw New Exception("EDIT Task Failed")
            End If
        Catch ex As Exception
        End Try
        update_database5()
        Dim test As String
        reset_list(test)
    End Sub


    Private Sub out_process()
        Dim test As Integer
        test = lonera.SelectedIndex

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
        select_string = "SELECT * FROM PC WHERE Available=TRUE"

        data_adapter = New OleDb.OleDbDataAdapter(select_string, connect_string)

        ' Map Table to Contacts.
        data_adapter.TableMappings.Add("Tables", "PC")

        ' Fill the DataSet.

        m_DataSet = New DataSet
        data_adapter.Fill(m_DataSet, "PC")
        Dim hello As Integer = m_DataSet.Tables("PC").Rows.Count

        If hello <> 0 Then
            name1 = m_DataSet.Tables("PC").Rows(test).Item("pc_name")


        Else
            MessageBox.Show("Their are no more loaner pcs to loan out. Sorry!", "Stop1", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        update_database(name1)
    End Sub

    Private Sub cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cancel.Click
        Me.Close()

    End Sub

    Private Sub ok_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ok.Click
        If tabs = 0 And check1 = True And check3 = True And check4 = True And check5 = True And loaner = 1 Then
            'Set the SQL string
            Dim objData As New DALBase
            Dim connect_string As String
            connect_string = main.mystream

            objData.SQL = "add_check"
            'Initialize the Command object
            objData.InitializeCommand()
            'Add the Parameters to the Parameters collection
            objData.AddParameter("@call_number", _
                OleDb.OleDbType.VarChar, 16, callnumber1.Text)
            objData.AddParameter("@date_issued", _
                OleDb.OleDbType.VarChar, 16, date_issue.Text)
            objData.AddParameter("@user", _
                        OleDb.OleDbType.VarChar, 16, userbox1.Text)
            objData.AddParameter("@pc_name", _
                                OleDb.OleDbType.VarChar, 16, name1)
            objData.AddParameter("@room", _
                                OleDb.OleDbType.VarChar, 16, roombox1.Text)
            objData.AddParameter("@technician", _
                                OleDb.OleDbType.VarChar, 16, techbox1.Text)
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
            update_database7()
            updatedatabase4()

        ElseIf check1 = False And tabs = 0 Then
            MessageBox.Show("Please enter a call number", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        ElseIf check3 = False And tabs = 0 Then
            MessageBox.Show("Please choose an user name", "STOP", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        ElseIf check4 = False And tabs = 0 Then
            MessageBox.Show("Please choose the room number", "STOP", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        ElseIf check5 = False And tabs = 0 Then
            MessageBox.Show("Please choose the room number", "STOP", MessageBoxButtons.OK, MessageBoxIcon.Stop)
        Else
            update_database3()
            Me.Close()
        End If
    End Sub


    Private Sub callbox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles callbox.SelectedIndexChanged
        '///////////////init///////////////////////////////
        ' The DataSet that holds the data.
        connect_string = main.mystream

        Dim m_DataSet As DataSet
        Dim table_name As Integer = callbox.Text
        ' A DataView to allow read-only access the data.
        Dim m_DataView As DataView

        ' Load the data.

        Dim select_string As String

        Dim data_adapter As OleDb.OleDbDataAdapter

        ' Compose the database file name.
        ' Modify this if the database is somewhere else.
        ' Compose the connect string. 

        select_string = "SELECT * FROM check_in WHERE call_number=" & table_name & ""

        data_adapter = New OleDb.OleDbDataAdapter(select_string, connect_string)

        ' Map Table to Contacts.
        data_adapter.TableMappings.Add("Tables", "check_in")

        ' Fill the DataSet.
        m_DataSet = New DataSet
        data_adapter.Fill(m_DataSet, "check_in")

        loneruna2.DataSource = m_DataSet.Tables("check_in")
        loneruna2.DisplayMember = "pc_name"
        userbox2.DataSource = m_DataSet.Tables("check_in")
        userbox2.DisplayMember = "user"
        roombox2.DataSource = m_DataSet.Tables("check_in")
        roombox2.DisplayMember = "room"
        techbox2.DataSource = m_DataSet.Tables("check_in")
        techbox2.DisplayMember = "technician"

    End Sub


    Private Sub out2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles out2.Click
        MessageBox.Show("You cannot add pc's at this time, please click on the ""Check Out"" tab", "Stop", MessageBoxButtons.OK, MessageBoxIcon.Stop)
    End Sub

    Private Sub in2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles in2.Click
        If loneruna2.SelectedIndex < 0 Then
            MessageBox.Show("There is only one loaner pc per user", "Warning2", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            update_database8()
        End If


    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        If tabs = 0 Then
            tabs = 1
        Else
            tabs = 0
        End If
    End Sub


    Private Sub updatedatabase4()

        'Set the SQL string
        Dim objData As New DALBase
        Dim connect_string As String
        connect_string = main.mystream

        objData.SQL = "add_history"
        'Initialize the Command object
        objData.InitializeCommand()
        'Add the Parameters to the Parameters collection
        objData.AddParameter("@call_number", _
            OleDb.OleDbType.VarChar, 16, callnumber1.Text)
        objData.AddParameter("@user_id", _
            OleDb.OleDbType.VarChar, 16, userbox1.Text)
        objData.AddParameter("@pc_id", _
                    OleDb.OleDbType.VarChar, 16, name1)
        objData.AddParameter("@room_id", _
                            OleDb.OleDbType.VarChar, 16, roombox1.Text)
        objData.AddParameter("@tech_id", _
                    OleDb.OleDbType.VarChar, 16, techbox1.Text)
        objData.AddParameter("@date_issued", _
                    OleDb.OleDbType.Date, 16, date_issue.Text)
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
        End Try
        Me.Close()

    End Sub
    Private Sub update_database5()
        'Set the SQL string
        Dim objData As New DALBase
        Dim connect_string As String
        connect_string = main.mystream

        objData.SQL = "edit_history"
        'Initialize the Command object
        objData.InitializeCommand()
        'Add the Parameters to the Parameters collection


        objData.AddParameter("@date_recieved", _
        OleDb.OleDbType.VarChar, 16, date_recieved.Text)
        objData.AddParameter("@call_number", _
                OleDb.OleDbType.VarChar, 16, callbox.Text)
        objData.Connection.ConnectionString = connect_string

        Try
            objData.OpenConnection()
            'Execute the query
            Dim intRowsAffected As Integer = objData.Command.ExecuteNonQuery
            'Close the database connection
            objData.CloseConnection()
            If intRowsAffected = 0 Then
                Throw New Exception("EDIT Task Failed")
            End If
        Catch ex As Exception
        End Try
        Dim test As String
        reset_list(test)
    End Sub
    Private Sub update_database6(ByVal test As String)
        'Set the SQL string
        Dim objData As New DALBase
        Dim connect_string As String
        connect_string = main.mystream
        objData.SQL = "add_pc_un"
        'Initialize the Command object
        objData.InitializeCommand()
        'Add the Parameters to the Parameters collection

        objData.AddParameter("@pc_name", _
                    OleDb.OleDbType.VarChar, 16, test)
        objData.Connection.ConnectionString = connect_string
        Try
            objData.OpenConnection()
            'Execute the query
            Dim intRowsAffected As Integer = objData.Command.ExecuteNonQuery
            'Close the database connection
            objData.CloseConnection()
            If intRowsAffected = 0 Then
                Throw New Exception("EDIT Task Failed")
            End If
        Catch ex As Exception
            MessageBox.Show("This record has delete protection because it has relations with one or more tables.", "Relation Integrity Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub


    Private Sub update_database7()
        'Set the SQL string
        Dim objData As New DALBase
        Dim connect_string As String
        connect_string = main.mystream

        objData.SQL = "delete_pc_un"
        'Initialize the Command object
        objData.InitializeCommand()
        'Add the Parameters to the Parameters collection
        objData.Connection.ConnectionString = connect_string

        Try
            objData.OpenConnection()
            'Execute the query
            Dim intRowsAffected As Integer = objData.Command.ExecuteNonQuery
            'Close the database connection
            objData.CloseConnection()
            If intRowsAffected = 0 Then
                Throw New Exception("EDIT Task Failed")
            End If
        Catch ex As Exception
            MessageBox.Show("This record has delete protection because it has relations with one or more tables.", "Relation Integrity Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub update_database8()

        Dim test As Integer
        Dim connect_string As String
        connect_string = main.mystream
        test = loneruna2.SelectedIndex

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


        select_string = "SELECT * FROM check_in WHERE call_number=" & callbox.Text & ""
        data_adapter = New OleDb.OleDbDataAdapter(select_string, connect_string)

        ' Map Table to Contacts.
        data_adapter.TableMappings.Add("Tables", "check_in")

        ' Fill the DataSet.

        m_DataSet = New DataSet
        data_adapter.Fill(m_DataSet, "check_in")
        Dim hello As Integer = m_DataSet.Tables("check_in").Rows.Count


        If hello <> 0 Then
            name1 = m_DataSet.Tables("check_in").Rows(test).Item("pc_name")
        Else
            MessageBox.Show("Their are no more loaner pcs to loan out. Sorry!", "Stop1", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
        update_database2(name1)
    End Sub
    Private Sub update_database9()
        '///////////////init///////////////////////////////
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


        '///////////////////////////listbox1/////////////////////////////////////
        select_string = "SELECT * FROM PC WHERE Available=TRUE"

        data_adapter = New OleDb.OleDbDataAdapter(select_string, connect_string)

        ' Map Table to Contacts.
        data_adapter.TableMappings.Add("Tables", "PC")

        ' Fill the DataSet.
        m_DataSet = New DataSet
        data_adapter.Fill(m_DataSet, "PC")

        lonera.DataSource = m_DataSet.Tables("PC")
        lonera.DisplayMember = "pc_name"
        lonera2.DataSource = m_DataSet.Tables("PC")
        lonera2.DisplayMember = "pc_name"

        '////////////////////////////////////////////////////////////////
        '////////////////////////////////////unavailable list///////////////////////////////
        select_string = "SELECT * FROM pc_un"

        data_adapter = New OleDb.OleDbDataAdapter(select_string, connect_string)

        ' Map Table to Contacts.
        data_adapter.TableMappings.Add("Tables", "pc_un")

        ' Fill the DataSet.
        m_DataSet = New DataSet
        data_adapter.Fill(m_DataSet, "pc_un")


        loneruna2.DataSource = m_DataSet.Tables("pc_un")
        loneruna2.DisplayMember = "pc_name"

    End Sub

    Private Sub callnumber1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles callnumber1.TextChanged
        check1 = True

    End Sub

    Private Sub userbox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles userbox1.SelectedIndexChanged
        check3 = True
    End Sub

    Private Sub roombox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles roombox1.SelectedIndexChanged
        check4 = True
    End Sub

    Private Sub techbox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles techbox1.SelectedIndexChanged
        check5 = True
    End Sub
End Class
