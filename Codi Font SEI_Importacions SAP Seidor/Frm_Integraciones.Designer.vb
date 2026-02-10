<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Frm_Integraciones
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm_Integraciones))
        Me.lblMsg = New System.Windows.Forms.Label()
        Me.lblConectar = New System.Windows.Forms.Label()
        Me.bIntegracions = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblMsg
        '
        Me.lblMsg.AutoSize = True
        Me.lblMsg.Location = New System.Drawing.Point(33, 181)
        Me.lblMsg.Name = "lblMsg"
        Me.lblMsg.Size = New System.Drawing.Size(0, 13)
        Me.lblMsg.TabIndex = 2
        '
        'lblConectar
        '
        Me.lblConectar.AutoSize = True
        Me.lblConectar.Location = New System.Drawing.Point(33, 133)
        Me.lblConectar.Name = "lblConectar"
        Me.lblConectar.Size = New System.Drawing.Size(0, 13)
        Me.lblConectar.TabIndex = 4
        '
        'bIntegracions
        '
        Me.bIntegracions.Location = New System.Drawing.Point(70, 32)
        Me.bIntegracions.Name = "bIntegracions"
        Me.bIntegracions.Size = New System.Drawing.Size(199, 23)
        Me.bIntegracions.TabIndex = 6
        Me.bIntegracions.Text = "Importar"
        Me.bIntegracions.UseVisualStyleBackColor = True
        '
        'Frm_Integraciones
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(343, 210)
        Me.Controls.Add(Me.bIntegracions)
        Me.Controls.Add(Me.lblConectar)
        Me.Controls.Add(Me.lblMsg)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Frm_Integraciones"
        Me.Text = "Importaciones v.1.0.0"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblMsg As System.Windows.Forms.Label
    Friend WithEvents lblConectar As System.Windows.Forms.Label
    Friend WithEvents bIntegracions As System.Windows.Forms.Button
End Class
