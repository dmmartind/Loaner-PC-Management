'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Copyright 2005 David Martin. All Rights Reserved.
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Public Class room_add
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
    Friend WithEvents cancel_button As System.Windows.Forms.Button
    Friend WithEvents room_text As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.room_text = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.ok_button = New System.Windows.Forms.Button
        Me.cancel_button = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.SuspendLayout()
        '
        'room_text
        '
        Me.room_text.Location = New System.Drawing.Point(24, 32)
        Me.room_text.Name = "room_text"
        Me.room_text.Size = New System.Drawing.Size(232, 20)
        Me.room_text.TabIndex = 11
        Me.room_text.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(40, 16)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Room"
        '
        'ok_button
        '
        Me.ok_button.BackColor = System.Drawing.SystemColors.Control
        Me.ok_button.Location = New System.Drawing.Point(24, 96)
        Me.ok_button.Name = "ok_button"
        Me.ok_button.TabIndex = 14
        Me.ok_button.Text = "OK"
        '
        'cancel_button
        '
        Me.cancel_button.BackColor = System.Drawing.SystemColors.Control
        Me.cancel_button.Location = New System.Drawing.Point(176, 96)
        Me.cancel_button.Name = "cancel_button"
        Me.cancel_button.TabIndex = 15
        Me.cancel_button.Text = "Cancel"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.Control
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Location = New System.Drawing.Point(0, 72)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(272, 4)
        Me.Panel1.TabIndex = 16
        '
        'room_add
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.Cornsilk
        Me.ClientSize = New System.Drawing.Size(272, 126)
        Me.ControlBox = False
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.cancel_button)
        Me.Controls.Add(Me.ok_button)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.room_text)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "room_add"
        Me.Text = "Room (Add)"
        Me.ResumeLayout(False)

    End Sub

#End Region



    Private Sub ok_button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ok_button.Click
        'Set the SQL string
        Dim objData As New DALBase
        Dim main As switch
        Dim connect_string As String
        connect_string = main.mystream
        objData.SQL = "add_room"
        'Initialize the Command object
        objData.InitializeCommand()
        'Add the Parameters to the Parameters collection
        objData.AddParameter("@room", _
            OleDb.OleDbType.VarChar, 16, room_text.Text)
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


    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles room_text.TextChanged
        ok_button.Enabled = True


    End Sub

    Private Sub room_add_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ok_button.Enabled = False
    End Sub

    Private Sub cancel_button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cancel_button.Click
        Me.Close()


    End Sub
End Class
