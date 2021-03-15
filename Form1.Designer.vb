<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.offunit = New System.Windows.Forms.ListBox()
        Me.offCO = New System.Windows.Forms.ListBox()
        Me.Offter = New System.Windows.Forms.ListBox()
        Me.DefUnit = New System.Windows.Forms.ListBox()
        Me.DefCO = New System.Windows.Forms.ListBox()
        Me.Defter = New System.Windows.Forms.ListBox()
        Me.Subweapon1 = New System.Windows.Forms.CheckBox()
        Me.Subweapon2 = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Offpower = New System.Windows.Forms.CheckBox()
        Me.Defpower = New System.Windows.Forms.CheckBox()
        Me.Luck1 = New System.Windows.Forms.TextBox()
        Me.Luck2 = New System.Windows.Forms.TextBox()
        Me.Offtow = New System.Windows.Forms.TextBox()
        Me.DefTow = New System.Windows.Forms.TextBox()
        Me.OffHP = New System.Windows.Forms.TextBox()
        Me.DefHP = New System.Windows.Forms.TextBox()
        Me.Offlev = New System.Windows.Forms.TextBox()
        Me.Deflev = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.OffResult = New System.Windows.Forms.Label()
        Me.DefResult = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'offunit
        '
        Me.offunit.FormattingEnabled = True
        Me.offunit.ItemHeight = 12
        Me.offunit.Items.AddRange(New Object() {"Infantry", "Mech", "Bike", "Recon", "Flare", "Antiair", "Tank", "MediumTank", "WarTank", "Artillery", "Antitank", "Rockets", "Missiles", "Rig", "Fighter", "Bomber", "Duster", "BattleCopter", "TransportCopter", "Seaplane", "Battleship", "Carrier", "Submarine", "Crusier", "Lander", "Gunboat"})
        Me.offunit.Location = New System.Drawing.Point(18, 24)
        Me.offunit.Name = "offunit"
        Me.offunit.Size = New System.Drawing.Size(98, 472)
        Me.offunit.TabIndex = 0
        '
        'offCO
        '
        Me.offCO.FormattingEnabled = True
        Me.offCO.ItemHeight = 12
        Me.offCO.Items.AddRange(New Object() {"No CO (T2)", "Brenner (T2)", "Caulder (T2)", "Forsythe (T2)", "Gage (T2)", "Grat (T2)", "Greyfield (T2)", "Hawk (T2)", "Isabella (T2)", "Lee (T2)", "Lin (T2)", "Mira (T2)", "Tabitha (T2)", "Tasha (T2)", "Tinker (T2)", "Waylon (T2)", "Will (T2)", "Zhaotiantong (T2)"})
        Me.offCO.Location = New System.Drawing.Point(131, 24)
        Me.offCO.Name = "offCO"
        Me.offCO.Size = New System.Drawing.Size(98, 472)
        Me.offCO.TabIndex = 1
        '
        'Offter
        '
        Me.Offter.FormattingEnabled = True
        Me.Offter.ItemHeight = 12
        Me.Offter.Items.AddRange(New Object() {"Plain", "River", "Sea", "Beach", "Road", "Bridge", "Wood", "Mountain", "Wasteland", "Ruins", "Rough", "Mistonsea", "Reef", "Silo", "Headquarters", "City", "Commandtower", "Radar", "Factory", "Airport", "Seaport", "TempAirport", "TempSeaport", "MistonPlain", "MistonRiver", "MistonBeach"})
        Me.Offter.Location = New System.Drawing.Point(244, 24)
        Me.Offter.Name = "Offter"
        Me.Offter.Size = New System.Drawing.Size(98, 472)
        Me.Offter.TabIndex = 2
        '
        'DefUnit
        '
        Me.DefUnit.FormattingEnabled = True
        Me.DefUnit.ItemHeight = 12
        Me.DefUnit.Items.AddRange(New Object() {"Infantry", "Mech", "Bike", "Recon", "Flare", "Antiair", "Tank", "MediumTank", "WarTank", "Artillery", "Antitank", "Rockets", "Missiles", "Rig", "Fighter", "Bomber", "Duster", "BattleCopter", "TransportCopter", "Seaplane", "Battleship", "Carrier", "Submarine", "Crusier", "Lander", "Gunboat", "Meteor"})
        Me.DefUnit.Location = New System.Drawing.Point(527, 24)
        Me.DefUnit.Name = "DefUnit"
        Me.DefUnit.Size = New System.Drawing.Size(98, 472)
        Me.DefUnit.TabIndex = 3
        '
        'DefCO
        '
        Me.DefCO.FormattingEnabled = True
        Me.DefCO.ItemHeight = 12
        Me.DefCO.Items.AddRange(New Object() {"No CO (T2)", "Brenner (T2)", "Caulder (T2)", "Forsythe (T2)", "Gage (T2)", "Grat (T2)", "Greyfield (T2)", "Hawk (T2)", "Isabella (T2)", "Lee (T2)", "Lin (T2)", "Mira (T2)", "Tabitha (T2)", "Tasha (T2)", "Tinker (T2)", "Waylon (T2)", "Will (T2)", "Zhaotiantong (T2)"})
        Me.DefCO.Location = New System.Drawing.Point(640, 24)
        Me.DefCO.Name = "DefCO"
        Me.DefCO.Size = New System.Drawing.Size(98, 472)
        Me.DefCO.TabIndex = 4
        '
        'Defter
        '
        Me.Defter.FormattingEnabled = True
        Me.Defter.ItemHeight = 12
        Me.Defter.Items.AddRange(New Object() {"Plain", "River", "Sea", "Beach", "Road", "Bridge", "Wood", "Mountain", "Wasteland", "Ruins", "Rough", "Mistonsea", "Reef", "Silo", "Headquarters", "City", "Commandtower", "Radar", "Factory", "Airport", "Seaport", "TempAirport", "TempSeaport", "MistonPlain", "MistonRiver", "MistonBeach"})
        Me.Defter.Location = New System.Drawing.Point(753, 24)
        Me.Defter.Name = "Defter"
        Me.Defter.Size = New System.Drawing.Size(98, 472)
        Me.Defter.TabIndex = 5
        '
        'Subweapon1
        '
        Me.Subweapon1.AutoSize = True
        Me.Subweapon1.Location = New System.Drawing.Point(20, 509)
        Me.Subweapon1.Name = "Subweapon1"
        Me.Subweapon1.Size = New System.Drawing.Size(84, 16)
        Me.Subweapon1.TabIndex = 6
        Me.Subweapon1.Text = "2nd Weapon"
        Me.Subweapon1.UseVisualStyleBackColor = True
        '
        'Subweapon2
        '
        Me.Subweapon2.AutoSize = True
        Me.Subweapon2.Location = New System.Drawing.Point(752, 509)
        Me.Subweapon2.Name = "Subweapon2"
        Me.Subweapon2.Size = New System.Drawing.Size(84, 16)
        Me.Subweapon2.TabIndex = 7
        Me.Subweapon2.Text = "2nd Weapon"
        Me.Subweapon2.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(130, 510)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(617, 12)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Note: ""2nd weapon"" is used to check damage of 2nd weapons which are not used unti" &
    "l ammos are used up. "
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(410, 219)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(29, 12)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Luck"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(406, 259)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(35, 12)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Tower"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(403, 300)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(47, 12)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "HP(100)"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(394, 328)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(65, 24)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "Unit Level" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "0,1,2,3"
        '
        'Offpower
        '
        Me.Offpower.AutoSize = True
        Me.Offpower.Location = New System.Drawing.Point(352, 387)
        Me.Offpower.Name = "Offpower"
        Me.Offpower.Size = New System.Drawing.Size(54, 16)
        Me.Offpower.TabIndex = 13
        Me.Offpower.Text = "Power"
        Me.Offpower.UseVisualStyleBackColor = True
        '
        'Defpower
        '
        Me.Defpower.AutoSize = True
        Me.Defpower.Location = New System.Drawing.Point(442, 384)
        Me.Defpower.Name = "Defpower"
        Me.Defpower.Size = New System.Drawing.Size(54, 16)
        Me.Defpower.TabIndex = 14
        Me.Defpower.Text = "Power"
        Me.Defpower.UseVisualStyleBackColor = True
        '
        'Luck1
        '
        Me.Luck1.Location = New System.Drawing.Point(352, 215)
        Me.Luck1.Name = "Luck1"
        Me.Luck1.Size = New System.Drawing.Size(44, 21)
        Me.Luck1.TabIndex = 15
        '
        'Luck2
        '
        Me.Luck2.Location = New System.Drawing.Point(450, 215)
        Me.Luck2.Name = "Luck2"
        Me.Luck2.Size = New System.Drawing.Size(45, 21)
        Me.Luck2.TabIndex = 16
        '
        'Offtow
        '
        Me.Offtow.Location = New System.Drawing.Point(354, 259)
        Me.Offtow.Name = "Offtow"
        Me.Offtow.Size = New System.Drawing.Size(42, 21)
        Me.Offtow.TabIndex = 17
        '
        'DefTow
        '
        Me.DefTow.Location = New System.Drawing.Point(450, 260)
        Me.DefTow.Name = "DefTow"
        Me.DefTow.Size = New System.Drawing.Size(45, 21)
        Me.DefTow.TabIndex = 18
        '
        'OffHP
        '
        Me.OffHP.Location = New System.Drawing.Point(354, 296)
        Me.OffHP.Name = "OffHP"
        Me.OffHP.Size = New System.Drawing.Size(42, 21)
        Me.OffHP.TabIndex = 19
        '
        'DefHP
        '
        Me.DefHP.Location = New System.Drawing.Point(452, 296)
        Me.DefHP.Name = "DefHP"
        Me.DefHP.Size = New System.Drawing.Size(43, 21)
        Me.DefHP.TabIndex = 20
        '
        'Offlev
        '
        Me.Offlev.Location = New System.Drawing.Point(355, 332)
        Me.Offlev.Name = "Offlev"
        Me.Offlev.Size = New System.Drawing.Size(40, 21)
        Me.Offlev.TabIndex = 21
        '
        'Deflev
        '
        Me.Deflev.Location = New System.Drawing.Point(456, 334)
        Me.Deflev.Name = "Deflev"
        Me.Deflev.Size = New System.Drawing.Size(39, 21)
        Me.Deflev.TabIndex = 22
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(385, 37)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(94, 25)
        Me.Button1.TabIndex = 23
        Me.Button1.Text = "Calculate"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'OffResult
        '
        Me.OffResult.AutoSize = True
        Me.OffResult.Location = New System.Drawing.Point(385, 77)
        Me.OffResult.Name = "OffResult"
        Me.OffResult.Size = New System.Drawing.Size(0, 12)
        Me.OffResult.TabIndex = 24
        '
        'DefResult
        '
        Me.DefResult.AutoSize = True
        Me.DefResult.Location = New System.Drawing.Point(381, 116)
        Me.DefResult.Name = "DefResult"
        Me.DefResult.Size = New System.Drawing.Size(0, 12)
        Me.DefResult.TabIndex = 25
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(875, 534)
        Me.Controls.Add(Me.DefResult)
        Me.Controls.Add(Me.OffResult)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Deflev)
        Me.Controls.Add(Me.Offlev)
        Me.Controls.Add(Me.DefHP)
        Me.Controls.Add(Me.OffHP)
        Me.Controls.Add(Me.DefTow)
        Me.Controls.Add(Me.Offtow)
        Me.Controls.Add(Me.Luck2)
        Me.Controls.Add(Me.Luck1)
        Me.Controls.Add(Me.Defpower)
        Me.Controls.Add(Me.Offpower)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Subweapon2)
        Me.Controls.Add(Me.Subweapon1)
        Me.Controls.Add(Me.Defter)
        Me.Controls.Add(Me.DefCO)
        Me.Controls.Add(Me.DefUnit)
        Me.Controls.Add(Me.Offter)
        Me.Controls.Add(Me.offCO)
        Me.Controls.Add(Me.offunit)
        Me.Name = "Form1"
        Me.Text = "TInywarsDamageCalculator_20210315"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents offunit As System.Windows.Forms.ListBox
    Friend WithEvents offCO As System.Windows.Forms.ListBox
    Friend WithEvents Offter As System.Windows.Forms.ListBox
    Friend WithEvents DefUnit As System.Windows.Forms.ListBox
    Friend WithEvents DefCO As System.Windows.Forms.ListBox
    Friend WithEvents Defter As System.Windows.Forms.ListBox
    Friend WithEvents Subweapon1 As System.Windows.Forms.CheckBox
    Friend WithEvents Subweapon2 As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Offpower As System.Windows.Forms.CheckBox
    Friend WithEvents Defpower As System.Windows.Forms.CheckBox
    Friend WithEvents Luck1 As System.Windows.Forms.TextBox
    Friend WithEvents Luck2 As System.Windows.Forms.TextBox
    Friend WithEvents Offtow As System.Windows.Forms.TextBox
    Friend WithEvents DefTow As System.Windows.Forms.TextBox
    Friend WithEvents OffHP As System.Windows.Forms.TextBox
    Friend WithEvents DefHP As System.Windows.Forms.TextBox
    Friend WithEvents Offlev As System.Windows.Forms.TextBox
    Friend WithEvents Deflev As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents OffResult As System.Windows.Forms.Label
    Friend WithEvents DefResult As System.Windows.Forms.Label

End Class
