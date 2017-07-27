'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2005 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////




Public Class history_lookup
    Inherits System.Windows.Forms.Form
    '////////////////////////////////////////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////
    ' The DataSet that holds the data.
    Private m_DataSet As DataSet

    ' A DataView to allow read-only access the data.
    Private m_DataView As DataView

    ' Load the data.

    Private select_string As String
    Dim enable_flag As Boolean

    Dim connect_string As String
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
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents date1 As System.Windows.Forms.ComboBox
    Friend WithEvents field2 As System.Windows.Forms.ComboBox
    Friend WithEvents field1 As System.Windows.Forms.ComboBox
    Friend WithEvents date2 As System.Windows.Forms.ComboBox
    Friend WithEvents label_date As System.Windows.Forms.Label
    Friend WithEvents search As System.Windows.Forms.Button
    Friend WithEvents close_button As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.DataGrid1 = New System.Windows.Forms.DataGrid
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.search = New System.Windows.Forms.Button
        Me.date2 = New System.Windows.Forms.ComboBox
        Me.label_date = New System.Windows.Forms.Label
        Me.date1 = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.field2 = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.field1 = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.close_button = New System.Windows.Forms.Button
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'DataGrid1
        '
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(16, 88)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.Size = New System.Drawing.Size(704, 232)
        Me.DataGrid1.TabIndex = 1
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.NavajoWhite
        Me.Panel1.Controls.Add(Me.search)
        Me.Panel1.Controls.Add(Me.date2)
        Me.Panel1.Controls.Add(Me.label_date)
        Me.Panel1.Controls.Add(Me.date1)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.field2)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.field1)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(16, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(696, 80)
        Me.Panel1.TabIndex = 2
        '
        'search
        '
        Me.search.BackColor = System.Drawing.Color.Silver
        Me.search.Location = New System.Drawing.Point(8, 48)
        Me.search.Name = "search"
        Me.search.Size = New System.Drawing.Size(75, 23)
        Me.search.TabIndex = 8
        Me.search.Text = "Search"
        Me.search.UseVisualStyleBackColor = False
        '
        'date2
        '
        Me.date2.FormatString = "d"
        Me.date2.Location = New System.Drawing.Point(496, 40)
        Me.date2.Name = "date2"
        Me.date2.Size = New System.Drawing.Size(121, 21)
        Me.date2.TabIndex = 7
        '
        'label_date
        '
        Me.label_date.Location = New System.Drawing.Point(656, 24)
        Me.label_date.Name = "label_date"
        Me.label_date.Size = New System.Drawing.Size(16, 16)
        Me.label_date.TabIndex = 6
        Me.label_date.Text = "to"
        '
        'date1
        '
        Me.date1.FormatString = "d"
        Me.date1.Location = New System.Drawing.Point(496, 8)
        Me.date1.Name = "date1"
        Me.date1.Size = New System.Drawing.Size(121, 21)
        Me.date1.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(472, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(8, 16)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "="
        '
        'field2
        '
        Me.field2.Items.AddRange(New Object() {"All", "Call Number", "Date Issued", "Date Recieved", "User", "PC Name", "Room", "Technician"})
        Me.field2.Location = New System.Drawing.Point(336, 8)
        Me.field2.Name = "field2"
        Me.field2.Size = New System.Drawing.Size(121, 21)
        Me.field2.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(192, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(136, 16)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "From     History     Where"
        '
        'field1
        '
        Me.field1.Items.AddRange(New Object() {"All", "Call Number", "Date Issued", "Date Recieved", "User", "PC Name", "Room", "Technician"})
        Me.field1.Location = New System.Drawing.Point(56, 8)
        Me.field1.Name = "field1"
        Me.field1.Size = New System.Drawing.Size(121, 21)
        Me.field1.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Select"
        '
        'close_button
        '
        Me.close_button.Location = New System.Drawing.Point(640, 328)
        Me.close_button.Name = "close_button"
        Me.close_button.Size = New System.Drawing.Size(75, 23)
        Me.close_button.TabIndex = 3
        Me.close_button.Text = "Close"
        '
        'history_lookup
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.Cornsilk
        Me.ClientSize = New System.Drawing.Size(736, 358)
        Me.ControlBox = False
        Me.Controls.Add(Me.close_button)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.DataGrid1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "history_lookup"
        Me.Text = "Lookup Tables"
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub history_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim main As switch
        connect_string = main.mystream
        date2.Visible = False
        label_date.Visible = False
        search.Enabled = False

    End Sub


    Private Sub field2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles field2.SelectedIndexChanged
        Select Case field2.Text
            Case "All"

                If (enable_flag = True) Then
                    search.Enabled = True
                ElseIf (enable_flag = False) Then
                    enable_flag = True
                End If
            Case "Call Number"
                If (enable_flag = True) Then
                    search.Enabled = False
                ElseIf (enable_flag = False) Then
                    enable_flag = True
                End If
            Case "Date Issued"
                If (enable_flag = True) Then
                    search.Enabled = False
                ElseIf (enable_flag = False) Then
                    enable_flag = True
                End If
            Case "Date Recieved"
                If (enable_flag = True) Then
                    search.Enabled = False
                ElseIf (enable_flag = False) Then
                    enable_flag = True
                End If
            Case "User"
                If (enable_flag = True) Then
                    search.Enabled = False
                ElseIf (enable_flag = False) Then
                    enable_flag = True
                End If
            Case "PC Name"
                If (enable_flag = True) Then
                    search.Enabled = False
                ElseIf (enable_flag = False) Then
                    enable_flag = True
                End If
            Case "Room"
                If (enable_flag = True) Then
                    search.Enabled = False
                ElseIf (enable_flag = False) Then
                    enable_flag = True
                End If
            Case "Technician"
                If (enable_flag = True) Then
                    search.Enabled = False
                ElseIf (enable_flag = False) Then
                    enable_flag = True
                End If
        End Select

        Dim asf_field2 As String
        If field2.Text = "Date Recieved" Or field2.Text = "Date Issued" Then
            date1.Visible = True
            date2.Visible = True
            label_date.Visible = True
        ElseIf field2.Text = "All" Then
            date1.Visible = False
            date2.Visible = False
        Else
            date1.Visible = True
            date2.Visible = False
            label_date.Visible = False
        End If

        '/////////////////////user combo box/////////////////////////////
        ' Create the OleDbDataAdapter.

        Select Case field2.Text
            Case "All"
                asf_field2 = "*"
            Case "Call Number"
                asf_field2 = "call_number"
            Case "Date Issued"
                asf_field2 = "date_issued"
            Case "Date Recieved"
                asf_field2 = "date_recieved"
            Case "User"
                asf_field2 = "user_id"
            Case "PC Name"
                asf_field2 = "pc_id"
            Case "Room"
                asf_field2 = "room_id"
            Case "Technician"
                asf_field2 = "tech_id"
            Case Else
        End Select

        select_string = "SELECT * FROM history"

        data_adapter = New OleDb.OleDbDataAdapter(select_string, connect_string)

        ' Map Table to Contacts.
        data_adapter.TableMappings.Add("Tables", "history")

        ' Fill the DataSet.
        m_DataSet = New DataSet
        data_adapter.Fill(m_DataSet, "history")

        date1.DataSource = m_DataSet.Tables("history")
        date1.DisplayMember = asf_field2
        'date1.FormatString.Remove()


        date2.DataSource = m_DataSet.Tables("history")
        date2.DisplayMember = asf_field2

        If (date1.Visible = True) Or (date2.Visible = True) Then
            search.Enabled = False
        Else
            search.Enabled = True

        End If
        '///////////////////////////room combo box/////////////////////////////////////////////////////////////////


    End Sub

    Private Sub search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles search.Click
        ' Create the OleDbDataAdapter.
        Dim new_date1, new_date2 As String
        Dim ast1 As String
        Dim temp = False
        Select Case field1.Text
            Case "All"
                ast1 = "*"
            Case "Call Number"
                ast1 = "call_number"
            Case "Date Issued"
                ast1 = "date_issued"
            Case "Date Recieved"
                ast1 = "date_recieved"
            Case "User"
                ast1 = "user_id"
            Case "PC Name"
                ast1 = "pc_id"
            Case "Room"
                ast1 = "room_id"
            Case "Technician"
                ast1 = "tech_id"
        End Select

        Dim ast2 As String

        Select Case field2.Text
            Case "All"
                ast2 = "*"
            Case "Call Number"
                ast2 = "call_number"
            Case "Date Issued"
                ast2 = "date_issued"
            Case "Date Recieved"
                ast2 = "date_recieved"
            Case "User"
                ast2 = "user_id"
            Case "PC Name"
                ast2 = "pc_id"
            Case "Room"
                ast2 = "room_id"
            Case "Technician"
                ast2 = "tech_id"
        End Select



        If ast2 = "*" Then
            select_string = "SELECT " & ast1 & "  FROM history "
        ElseIf field2.Text = "Date Recieved" Or field2.Text = "Date Issued" Then
            new_date1 = date1.Text
            new_date2 = date2.Text
            new_date1 = new_date1.Remove(10, 12)
            new_date2 = new_date2.Remove(10, 12)
            select_string = "SELECT " & ast1 & "  FROM history where " & ast2 & " BETWEEN #" & new_date1 & "# And #" & new_date2 & "#"
        Else

            select_string = "SELECT " & ast1 & "  FROM history where " & ast2 & " = """ & date1.Text & """"
        End If
        '/////////'
        data_adapter = New OleDb.OleDbDataAdapter(select_string, connect_string)

        ' Map Table to Contacts.
        data_adapter.TableMappings.Add("Tables", "history")

        ' Fill the DataSet.
        m_DataSet = New DataSet
        data_adapter.Fill(m_DataSet, "history")
        ' Make a DataView for the data.
        m_DataView = New DataView(m_DataSet.Tables("history"))
        m_DataView.AllowDelete = False
        m_DataView.AllowEdit = False
        m_DataView.AllowNew = False



        ' Bind the DataGrid control to the DataView.
        DataGrid1.DataSource = m_DataView
    End Sub

    Private Sub close_button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles close_button.Click
        Me.Close()
    End Sub

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub field1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles field1.SelectedIndexChanged
        Select Case field1.Text
            Case "All"
                If (enable_flag = True) Then
                    search.Enabled = True
                ElseIf (enable_flag = False) Then
                    enable_flag = True
                End If
            Case "Call Number"
                enable_flag = False
                search.Enabled = False
                If (enable_flag = True) Then
                    search.Enabled = True
                ElseIf (enable_flag = False) Then
                    enable_flag = True
                End If
            Case "Date Issued"
                enable_flag = False
                search.Enabled = False
                If (enable_flag = True) Then
                    search.Enabled = True
                ElseIf (enable_flag = False) Then
                    enable_flag = True
                End If
            Case "Date Recieved"
                enable_flag = False
                search.Enabled = False
                If (enable_flag = True) Then
                    search.Enabled = True
                ElseIf (enable_flag = False) Then
                    enable_flag = True
                End If
            Case "User"
                enable_flag = False
                search.Enabled = False
                If (enable_flag = True) Then
                    search.Enabled = True
                ElseIf (enable_flag = False) Then
                    enable_flag = True
                End If
            Case "PC Name"
                enable_flag = False
                search.Enabled = False
                If (enable_flag = True) Then
                    search.Enabled = True
                ElseIf (enable_flag = False) Then
                    enable_flag = True
                End If
            Case "Room"
                enable_flag = False
                search.Enabled = False
                If (enable_flag = True) Then
                    search.Enabled = True
                ElseIf (enable_flag = False) Then
                    enable_flag = True
                End If
            Case "Technician"
                enable_flag = False
                search.Enabled = False
                If (enable_flag = True) Then
                    search.Enabled = True
                ElseIf (enable_flag = False) Then
                    enable_flag = True
                End If
        End Select
    End Sub

    Private Sub date1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles date1.SelectedIndexChanged
        If date2.Visible = True Then
            search.Enabled = False
        Else
            search.Enabled = True
        End If

    End Sub

    Private Sub date2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles date2.SelectedIndexChanged
        search.Enabled = True

    End Sub
End Class

