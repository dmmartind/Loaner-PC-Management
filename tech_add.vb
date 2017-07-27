'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2005 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Public Class tech_add
    Inherits System.Windows.Forms.Form

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
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents ok_button As System.Windows.Forms.Button
    Friend WithEvents tech_name As System.Windows.Forms.TextBox
    Friend WithEvents cancel_button As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label2 = New System.Windows.Forms.Label
        Me.tech_name = New System.Windows.Forms.TextBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.ok_button = New System.Windows.Forms.Button
        Me.cancel_button = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 16)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Technician Name"
        '
        'tech_name
        '
        Me.tech_name.Location = New System.Drawing.Point(16, 24)
        Me.tech_name.Name = "tech_name"
        Me.tech_name.Size = New System.Drawing.Size(232, 20)
        Me.tech_name.TabIndex = 8
        Me.tech_name.Text = ""
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.Control
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Location = New System.Drawing.Point(0, 56)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(272, 4)
        Me.Panel1.TabIndex = 9
        '
        'ok_button
        '
        Me.ok_button.BackColor = System.Drawing.SystemColors.Control
        Me.ok_button.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.ok_button.Location = New System.Drawing.Point(16, 80)
        Me.ok_button.Name = "ok_button"
        Me.ok_button.TabIndex = 10
        Me.ok_button.Text = "OK"
        '
        'cancel_button
        '
        Me.cancel_button.BackColor = System.Drawing.SystemColors.Control
        Me.cancel_button.Location = New System.Drawing.Point(176, 80)
        Me.cancel_button.Name = "cancel_button"
        Me.cancel_button.TabIndex = 11
        Me.cancel_button.Text = "Cancel"
        '
        'tech_add
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.Cornsilk
        Me.ClientSize = New System.Drawing.Size(272, 118)
        Me.ControlBox = False
        Me.Controls.Add(Me.cancel_button)
        Me.Controls.Add(Me.ok_button)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.tech_name)
        Me.Controls.Add(Me.Label2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "tech_add"
        Me.Text = "Technician (ADD)"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub ok_button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ok_button.Click
        'Set the SQL string
        Dim objData As New DALBase
        Dim main As switch
        Dim connect_string As String
        connect_string = main.mystream
        objData.SQL = "add_tech"
        'Initialize the Command object
        objData.InitializeCommand()
        'Add the Parameters to the Parameters collection
        objData.AddParameter("@tech_id", _
            OleDb.OleDbType.VarChar, 16, tech_name.Text)
        objData.Connection.ConnectionString = connect_string
        'Open the database connection
        objData.OpenConnection()
        'Execute the query
        Dim intRowsAffected As Integer = objData.Command.ExecuteNonQuery()
        'Close the database connection
        objData.CloseConnection()
        If intRowsAffected = 0 Then
            Throw New Exception("Add Task Failed")
        End If
        'Clear the input fields

        Me.Close()


    End Sub

    Private Sub tech_name_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tech_name.TextChanged
        ok_button.Enabled = True
    End Sub

    Private Sub tech_add_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ok_button.Enabled = False
    End Sub

    Private Sub cancel_button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cancel_button.Click
        Me.Close()

    End Sub
End Class
