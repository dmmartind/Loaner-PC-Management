'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2005 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Public Class lookup_tables_check
    Inherits System.Windows.Forms.Form
    '////////////////////////////////////////////////////////////////////////////////////
    '///////////////////////////////////////////////////////////////////////////////////
    ' The DataSet that holds the data.
    Private m_DataSet As DataSet

    ' A DataView to allow read-only access the data.
    Private m_DataView As DataView

    ' Load the data.

    Private select_string As String
    Private data_adapter As OleDb.OleDbDataAdapter
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
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents close_button As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.DataGrid1 = New System.Windows.Forms.DataGrid
        Me.close_button = New System.Windows.Forms.Button
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGrid1
        '
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGrid1.Location = New System.Drawing.Point(8, 8)
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.Size = New System.Drawing.Size(568, 312)
        Me.DataGrid1.TabIndex = 1
        '
        'close_button
        '
        Me.close_button.Location = New System.Drawing.Point(496, 328)
        Me.close_button.Name = "close_button"
        Me.close_button.TabIndex = 2
        Me.close_button.Text = "Close"
        '
        'lookup_tables_check
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.Cornsilk
        Me.ClientSize = New System.Drawing.Size(582, 356)
        Me.ControlBox = False
        Me.Controls.Add(Me.close_button)
        Me.Controls.Add(Me.DataGrid1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "lookup_tables_check"
        Me.Text = "Check Out Lookup"
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub lookup_tables_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        connect_string = main.mystream
        ' Create the OleDbDataAdapter.
        select_string = "SELECT * FROM check_in"

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

