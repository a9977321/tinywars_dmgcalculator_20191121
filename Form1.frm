VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   12090
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Luck2 
      Height          =   270
      Left            =   6720
      TabIndex        =   25
      Top             =   2760
      Width           =   735
   End
   Begin VB.CheckBox Subweapon2 
      Caption         =   "2nd Weapon"
      Height          =   255
      Left            =   10680
      TabIndex        =   23
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CheckBox Subweapon1 
      Caption         =   "2nd Weapon"
      Height          =   375
      Left            =   240
      TabIndex        =   22
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox Deflev 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   1
      EndProperty
      Height          =   270
      Left            =   6960
      TabIndex        =   18
      Top             =   4200
      Width           =   735
   End
   Begin VB.TextBox Offlev 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   1
      EndProperty
      Height          =   270
      Left            =   4440
      TabIndex        =   17
      Top             =   4200
      Width           =   735
   End
   Begin VB.CheckBox Defpower 
      Caption         =   "Power"
      Height          =   375
      Left            =   6720
      TabIndex        =   14
      Top             =   4680
      Width           =   975
   End
   Begin VB.CheckBox Offpower 
      Caption         =   "Power"
      Height          =   375
      Left            =   4680
      TabIndex        =   13
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox Luck1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   1
      EndProperty
      Height          =   270
      Left            =   4680
      TabIndex        =   12
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox DefHP 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   1
      EndProperty
      Height          =   270
      Left            =   6720
      TabIndex        =   11
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox OffHP 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   1
      EndProperty
      Height          =   270
      Left            =   4680
      TabIndex        =   10
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox DefTow 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   1
      EndProperty
      Height          =   270
      Left            =   6720
      TabIndex        =   9
      Top             =   3240
      Width           =   735
   End
   Begin VB.TextBox OffTow 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   1
      EndProperty
      Height          =   270
      Left            =   4680
      TabIndex        =   8
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   5400
      TabIndex        =   7
      Top             =   960
      Width           =   1455
   End
   Begin VB.ListBox DefTer 
      Height          =   4740
      ItemData        =   "Form1.frx":0000
      Left            =   10680
      List            =   "Form1.frx":0002
      TabIndex        =   5
      Top             =   360
      Width           =   1335
   End
   Begin VB.ListBox DefCO 
      Height          =   4740
      ItemData        =   "Form1.frx":0004
      Left            =   9120
      List            =   "Form1.frx":0006
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.ListBox defunit 
      Height          =   4740
      ItemData        =   "Form1.frx":0008
      Left            =   7800
      List            =   "Form1.frx":000A
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.ListBox OffTer 
      Height          =   4740
      ItemData        =   "Form1.frx":000C
      Left            =   3000
      List            =   "Form1.frx":000E
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.ListBox OffCO 
      Height          =   4740
      ItemData        =   "Form1.frx":0010
      Left            =   1560
      List            =   "Form1.frx":0012
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.ListBox offunit 
      Height          =   4740
      ItemData        =   "Form1.frx":0014
      Left            =   240
      List            =   "Form1.frx":0016
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Note:""2nd Weapon"" is used to check damage of 2nd weapons which are not used until ammos are used up"
      Height          =   255
      Left            =   1560
      TabIndex        =   24
      Top             =   5280
      Width           =   9135
   End
   Begin VB.Label DefResult 
      Height          =   375
      Left            =   5040
      TabIndex        =   21
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label OffResult 
      Height          =   255
      Left            =   5160
      TabIndex        =   20
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Unit level,0,1,2,3"
      Height          =   420
      Left            =   5160
      TabIndex        =   19
      Top             =   4200
      Width           =   1770
   End
   Begin VB.Label Label3 
      Caption         =   "HP(100)"
      Height          =   255
      Left            =   5760
      TabIndex        =   16
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Tower"
      Height          =   255
      Left            =   5760
      TabIndex        =   15
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Luck"
      Height          =   255
      Left            =   5760
      TabIndex        =   6
      Top             =   2760
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Dim basedmg1 As Double, offunit1 As Double, defunit1 As Double, subweapon01 As Double, subweapon02 As Double, offcomod1 As Double, defcomod1 As Double
Dim offco1 As Double, defco1 As Double, tow1 As Double, tow2 As Double, lev2 As Double, offpow1 As Double, defpow1 As Double
Dim HP1 As Double, HP2 As Double, offter1 As Double, defter1 As Double, offlevmod1 As Double, deflevmod1 As Double, offlevmod2 As Double, deflevmod2 As Double
Dim lev1 As Double, luc1 As Double, luc2 As Double, offresult1 As Double, defresult1 As Double, offcomod2 As Double, defcomod2 As Double, basedmg2 As Double
Dim cflag As Boolean
'Obtaining Data
offunit1 = offunit.ListIndex
defunit1 = defunit.ListIndex
subweapon01 = Subweapon1.value
subweapon02 = Subweapon2.value
offco1 = OffCO.ListIndex
defco1 = DefCO.ListIndex
offpow1 = Offpower.value
defpow1 = Defpower.value
lev1 = Val(Offlev.Text)
lev2 = Val(Deflev.Text)
tow1 = Val(OffTow.Text)
tow2 = Val(DefTow.Text)
HP1 = Val(OffHP.Text)
HP2 = Val(DefHP.Text)
offter1 = OffTer.ListIndex
defter1 = DefTer.ListIndex
luc1 = Val(Luck1.Text)
luc2 = Val(Luck2.Text)
offcomod1 = 100
defcomod1 = 100
offcomod2 = 100
defcomod2 = 100
cflag = False
'Error Page


'Formal
basedmg1 = basedmgcal(offunit1, defunit1, subweapon01)
offcomod1 = offcomod1 + offcomodcal(offco1, offunit1, offpow1, lev1)
offlevmod1 = levdmgmod(lev1)
deflevmod1 = levdefmod(lev2)
defter1 = tercal(defter1)
offter1 = tercal(offter1)
defcomod1 = defcomod1 + defcomodcal(defco1, defunit1, defpow1, lev2)
offresult1 = Fix((basedmg1 * (offcomod1 + offlevmod1 + 5 * tow1) + luc1 * 100) * ((HP1 + 9) \ 10) / (defcomod1 + defter1 * ((HP2 + 9) \ 10) / 10 + 5 * tow2 + deflevmod1) / 10)
HP2 = HP2 - offresult1

If HP2 > 0 Then
cflag = True
End If
Select Case offunit1
Case 9, 11, 12, 20
cflag = False
End Select
Select Case defunit1
Case 9, 11, 12, 20
cflag = False
End Select

If cflag Then
basedmg2 = basedmgcal(defunit1, offunit1, subweapon02)
offcomod2 = offcomod2 + offcomodcal(defco1, defunit1, defpow1, lev2)
offlevmod2 = levdmgmod(lev2)
deflevmod2 = levdmgmod(lev1)
defcomod2 = defcomod2 + defcomodcal(offco1, offunit1, offpow1, lev1)
defresult1 = Fix((basedmg2 * (offcomod2 + offlevmod2 + 5 * tow2) + luc2 * 100) * ((HP2 + 9) \ 10) / (defcomod2 + offter1 * ((HP1 + 9) \ 10) / 10 + 5 * tow1 + deflevmod2) / 10)
Else
defresult1 = 0
End If

'Results
OffResult.Caption = "Attack" & Space(2) & Str(offresult1)
DefResult.Caption = "Counterattack" & Space(1) & Str(defresult1)
End Sub
'Level modifier
Private Function levdmgmod(ByVal lev As Double) As Double
Select Case lev
Case 0
levdmgmod = 0
Case 1
levdmgmod = 5
Case 2
levdmgmod = 10
Case 3
levdmgmod = 20
End Select
End Function
Private Function levdefmod(ByVal lev As Double) As Double
If lev = 3 Then
levdefmod = 20
Else
levdefmod = 0
End If
End Function

Private Function offcomodcal(ByRef offco1 As Double, ByRef offunit1 As Double, ByRef pow As Double, ByRef lev As Double) As Double
'CO damage modifier calculator
Dim a As Double
a = 0
Select Case offco1
Case 0
a = 0
Case 1
Call dmgmod0112(a)
Case 2
If pow = 0 Then
Call dmgmod0102(a)
Else
Call dmgmod0108(a)
End If
Case 3
Call dmgmod0104(a)
Case 4
If pow = 0 Then
Call dmgmod0210(offunit1, a)
Else
Call dmgmod02a1(offunit1, a)
End If
Case 5
If pow = 0 Then
a = 0
Else
Call dmgmod0104(a)
End If
Case 6
If pow = 0 Then
a = 0
Else
Call dmgmod0310(offunit1, a)
End If
Case 7
Call dmgmod0404(lev, a)
Case 8
Call dmgmod0112(a)
Case 9
Call dmgmod0102(a)
Case 10
Call dmgmod0106(a)
Case 11
Call dmgmod0102(a)
Call dmgmod0504(offunit1, a)
Case 12
Call dmgmod0112(a)
Case 13
Call dmgmod0102(a)
Call dmgmod0604(offunit1, a)
Case 14
Call dmgmod0102(a)
Call dmgmod0704(offunit1, a)
Case 15
Call dmgmod0104(a)
Case 16
a = 0
Case 17
Call dmgmod0102(a)
Call dmgmod0814(offunit1, a)
Case 18
If pow = 0 Then
Call dmgmod0102(a)
Else
Call dmgmod0104(a)
End If
Case 19
If pow = 0 Then
Call dmgmod0102(a)
Else
Call dmgmod0106(a)
End If
Case 20
If pow = 0 Then
Call dmgmod0102(a)
Call dmgmod0914(offunit1, a)
Else
Call dmgmod0102(a)
Call dmgmod0918(offunit1, a)
End If
Case 21
Call dmgmod0102(a)
Call dmgmod1004(offunit1, a)
Case 22
Call dmgmod0104(a)
Case 23
Call dmgmod0102(a)
Call dmgmod1102(offunit1, a)
Case 24
Call dmgmod0102(a)
Call dmgmod0608(offunit1, a)
Case 25
If pow = 0 Then
Call dmgmod0102(a)
Else
Call dmgmod0108(a)
End If
Case 26
a = 0
Case 27
Call dmgmod0120(a)
End Select
offcomodcal = a
End Function



'All units +10% offense
Private Sub dmgmod0102(ByRef dmg As Double)
dmg = dmg + 10
End Sub
'All units +20% offense
Private Sub dmgmod0104(ByRef dmg As Double)
dmg = dmg + 20
End Sub
'All units +30% offense
Private Sub dmgmod0106(ByRef dmg As Double)
dmg = dmg + 30
End Sub
'All units+40% offense
Private Sub dmgmod0108(ByRef dmg As Double)
dmg = dmg + 40
End Sub
'All units +60% offense
Private Sub dmgmod0112(ByRef dmg As Double)
dmg = dmg + 60
End Sub
'All units +100% offense
Private Sub dmgmod0120(ByRef dmg As Double)
dmg = dmg + 100
End Sub
'All footsoldiers +50% offense
Private Sub dmgmod0210(ByVal a As Double, ByRef dmg As Double)
Select Case a
Case 0, 1, 2
dmg = dmg + 50
End Select
End Sub
'All footsoldiers +1000% offense
Private Sub dmgmod02a1(ByVal a As Double, ByRef dmg As Double)
Select Case a
Case 0, 1, 2
dmg = dmg + 1000
End Select
End Sub
'All Indirects +50% offense
Private Sub dmgmod0310(ByVal a As Double, ByRef dmg As Double)
Select Case a
Case 9, 10, 11, 12, 20
dmg = dmg + 50
End Select
End Sub
'All veteran units +20% offense
Private Sub dmgmod0404(ByVal a As Double, ByRef dmg As Double)
Select Case a
Case 3
dmg = dmg + 20
End Select
End Sub
'All naval & indirect units +20% offense
Private Sub dmgmod0504(ByVal a As Double, ByRef dmg As Double)
Select Case a
Case 9, 10, 11, 12, 20, 21, 22, 23, 24, 25
dmg = dmg + 20
End Select
End Sub
'All air units +20% offense
Private Sub dmgmod0604(ByVal a As Double, ByRef dmg As Double)
Select Case a
Case 14, 15, 16, 17, 18, 19
dmg = dmg + 20
End Select
End Sub
'All air units +40% offense
Private Sub dmgmod0608(ByVal a As Double, ByRef dmg As Double)
Select Case a
Case 14, 15, 16, 17, 18, 19
dmg = dmg + 40
End Select
End Sub
'All land direct units +20% offense
Private Sub dmgmod0704(ByVal a As Double, ByRef dmg As Double)
Select Case a
Case 0, 1, 2, 3, 4, 5, 6, 7, 8
dmg = dmg + 20
End Select
End Sub
'All bombers/fighters +70% offense
Private Sub dmgmod0814(ByVal a As Double, ByRef dmg As Double)
Select Case a
Case 14, 15
dmg = dmg + 70
End Select
End Sub
'All mediumtanks/wartanks +70% offense
Private Sub dmgmod0914(ByVal a As Double, ByRef dmg As Double)
Select Case a
Case 7, 8
dmg = dmg + 70
End Select
End Sub
'All mediumtanks/wartanks +90% offense
Private Sub dmgmod0918(ByVal a As Double, ByRef dmg As Double)
Select Case a
Case 7, 8
dmg = dmg + 90
End Select
End Sub
'All land units +20% offense
Private Sub dmgmod1004(ByVal a As Double, ByRef dmg As Double)
Select Case a
Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13
dmg = dmg + 20
End Select
End Sub
'All copters,seaplane & naval units +10% offense
Private Sub dmgmod1102(ByVal a As Double, ByRef dmg As Double)
Select Case a
Case 17, 18, 19, 20, 21, 22, 23, 24, 25
dmg = dmg + 10
End Select
End Sub

Private Function defcomodcal(ByVal co As Double, ByVal unit As Double, ByVal power As Double, ByVal lev As Double) As Double
'CO defense modifier calculator
Dim a As Double
a = 0
Select Case co
Case 0
a = 0
Case 1
Call defmod0112(a)
Case 2
If power = 0 Then
Call defmod0102(a)
Else
Call defmod0104(a)
End If
Case 3
Call defmod0104(a)
Case 4
If power = 0 Then
Call defmod0204(unit, a)
Else
Call defmod0260(unit, a)
End If
Case 5
If power = 0 Then
a = 0
Else
Call defmod0104(a)
End If
Case 6
If power = 0 Then
a = 0
Else
Call defmod0106(a)
End If
Case 7
Call defmod0304(lev, a)
Case 8
Call defmod0112(a)
Case 9
Call defmod0106(a)
Case 10
Call defmod0106(a)
Case 11
Call defmod0102(a)
Call defmod0402(unit, a)
Case 12
Call defmod0112(a)
Case 13
If power = 0 Then
Call defmod0102(a)
Call defmod0506(unit, a)
Else
Call defmod0102(a)
Call defmod0560(unit, a)
End If
Case 14
Call defmod0102(a)
Case 15
Call defmod0102(a)
Case 16
a = 0
Case 17
If power = 0 Then
Call defmod0102(a)
Call defmod0614(unit, a)
Else
Call defmod0102(a)
Call defmod06a1(unit, a)
End If
Case 18
Call defmod0102(a)
Case 19
If power = 0 Then
Call defmod0102(a)
Else
Call defmod01a1(a)
End If
Case 20
If power = 0 Then
Call defmod0102(a)
Call defmod0714(unit, a)
Else
Call defmod0102(a)
Call defmod0718(unit, a)
End If
Case 21
Call defmod0102(a)
Call defmod0804(unit, a)
Case 22
Call defmod0104(a)
Case 23
Call defmod0102(a)
Call defmod0908(unit, a)
Case 24
Call defmod0102(a)
Call defmod0504(unit, a)
Case 25
If power = 0 Then
Call defmod0102(a)
Else
Call defmod01a2(a)
End If
Case 26
a = 0
Case 27
Call defmod0120(a)
End Select
defcomodcal = a
End Function
'All units -10% defense
Private Sub defmod01a1(ByRef def As Double)
def = def - 10
End Sub
'All units -30% defense
Private Sub defmod01a2(ByRef def As Double)
def = def - 30
End Sub
'All units +10% defense
Private Sub defmod0102(ByRef def As Double)
def = def + 10
End Sub
'All units +20% defense
Private Sub defmod0104(ByRef def As Double)
def = def + 20
End Sub
'All units+30% defense
Private Sub defmod0106(ByRef def As Double)
def = def + 30
End Sub
'All units +60% defense
Private Sub defmod0112(ByRef def As Double)
def = def + 60
End Sub
'All units +100% defense
Private Sub defmod0120(ByRef def As Double)
def = def + 100
End Sub
'All footsoldiers +20% defense
Private Sub defmod0204(ByVal a As Double, ByRef def As Double)
Select Case a
Case 0, 1, 2
def = def + 20
End Select
End Sub
'All footsoldiers +300% defense
Private Sub defmod0260(ByVal a As Double, ByRef def As Double)
Select Case a
Case 0, 1, 2
def = def + 300
End Select
End Sub
'All veteran units +20% defense
Private Sub defmod0304(ByVal a As Double, ByRef def As Double)
Select Case a
Case 3
def = def + 20
End Select
End Sub
'All naval & indirect units +10% defense
Private Sub defmod0402(ByVal a As Double, ByRef def As Double)
Select Case a
Case 9, 10, 11, 12, 20, 21, 22, 23, 24, 25
def = def + 10
End Select
End Sub
'All air units +20% defense
Private Sub defmod0504(ByVal a As Double, ByRef def As Double)
Select Case a
Case 14, 15, 16, 17, 18, 19
def = def + 20
End Select
End Sub
'All air units +30% defense
Private Sub defmod0506(ByVal a As Double, ByRef def As Double)
Select Case a
Case 14, 15, 16, 17, 18, 19
def = def + 30
End Select
End Sub
'All air units +300% defense
Private Sub defmod0560(ByVal a As Double, ByRef def As Double)
Select Case a
Case 14, 15, 16, 17, 18, 19
def = def + 300
End Select
End Sub
'All fighters/bombers +70% defense
Private Sub defmod0614(ByVal a As Double, ByRef def As Double)
Select Case a
Case 14, 15
def = def + 70
End Select
End Sub
'All fighters/bombers +10000% defense
Private Sub defmod06a1(ByVal a As Double, ByRef def As Double)
Select Case a
Case 14, 15
def = def + 10000
End Select
End Sub
'All mediumtanks/wartanks +70% defense
Private Sub defmod0714(ByVal a As Double, ByRef def As Double)
Select Case a
Case 7, 8
def = def + 70
End Select
End Sub
'All mediumtanks/wartanks +90% defense
Private Sub defmod0718(ByVal a As Double, ByRef def As Double)
Select Case a
Case 7, 8
def = def + 90
End Select
End Sub
'All land units +20% defense
Private Sub defmod0804(ByVal a As Double, ByRef def As Double)
Select Case a
Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13
def = def + 20
End Select
End Sub
'All copters,seaplane & naval units +40% defense
Private Sub defmod0908(ByVal a As Double, ByRef def As Double)
Select Case a
Case 17, 18, 19, 20, 21, 22, 23, 24, 25
def = def + 40
End Select
End Sub
Private Function basedmgcal(ByRef offunit1 As Double, ByRef defunit1 As Double, ByRef c As Double) As Double
'Base Damage Chart, c is used for cases with 2 weapons available
Dim value As Double
value = 100 * offunit1 + defunit1
Select Case value
Case 0
basedmgcal = 55
Case 1
basedmgcal = 45
Case 2
basedmgcal = 45
Case 3
basedmgcal = 12
Case 4
basedmgcal = 10
Case 5
basedmgcal = 3
Case 6
basedmgcal = 5
Case 7
basedmgcal = 5
Case 8
basedmgcal = 1
Case 9
basedmgcal = 10
Case 10
basedmgcal = 30
Case 11
basedmgcal = 20
Case 12
basedmgcal = 20
Case 13
basedmgcal = 14
Case 17
basedmgcal = 8
Case 18
basedmgcal = 30
Case 26
basedmgcal = 1


Case 100
basedmgcal = 65
Case 101
basedmgcal = 55
Case 102
basedmgcal = 55
Case 103
If c = 0 Then
basedmgcal = 85
Else
basedmgcal = 18
End If
Case 104
If c = 0 Then
basedmgcal = 80
Else
basedmgcal = 15
End If
Case 105
If c = 0 Then
basedmgcal = 55
Else
basedmgcal = 5
End If
Case 106
If c = 0 Then
basedmgcal = 55
Else
basedmgcal = 8
End If
Case 107
If c = 0 Then
basedmgcal = 25
Else
basedmgcal = 5
End If
Case 108
If c = 0 Then
basedmgcal = 15
Else
basedmgcal = 1
End If
Case 109
If c = 0 Then
basedmgcal = 70
Else
basedmgcal = 15
End If
Case 110
If c = 0 Then
basedmgcal = 55
Else
basedmgcal = 35
End If
Case 111
If c = 0 Then
basedmgcal = 85
Else
basedmgcal = 35
End If
Case 112
If c = 0 Then
basedmgcal = 85
Else
basedmgcal = 35
End If
Case 113
If c = 0 Then
basedmgcal = 75
Else
basedmgcal = 20
End If
Case 117
basedmgcal = 12
Case 118
basedmgcal = 35
Case 126
If c = 0 Then
basedmgcal = 15
Else
basedmgcal = 1
End If


Case 200
basedmgcal = 65
Case 201
basedmgcal = 55
Case 202
basedmgcal = 55
Case 203
basedmgcal = 18
Case 204
basedmgcal = 15
Case 205
basedmgcal = 5
Case 206
basedmgcal = 8
Case 207
basedmgcal = 5
Case 208
basedmgcal = 1
Case 209
basedmgcal = 15
Case 210
basedmgcal = 35
Case 211
basedmgcal = 35
Case 212
basedmgcal = 35
Case 213
basedmgcal = 20
Case 217
basedmgcal = 12
Case 218
basedmgcal = 35
Case 226
basedmgcal = 1


Case 300
basedmgcal = 75
Case 301
basedmgcal = 65
Case 302
basedmgcal = 65
Case 303
basedmgcal = 35
Case 304
basedmgcal = 30
Case 305
basedmgcal = 8
Case 306
basedmgcal = 8
Case 307
basedmgcal = 5
Case 308
basedmgcal = 1
Case 309
basedmgcal = 45
Case 310
basedmgcal = 25
Case 311
basedmgcal = 55
Case 312
basedmgcal = 55
Case 313
basedmgcal = 45
Case 317
basedmgcal = 18
Case 318
basedmgcal = 35
Case 326
basedmgcal = 1


Case 400
basedmgcal = 80
Case 401
basedmgcal = 70
Case 402
basedmgcal = 70
Case 403
basedmgcal = 60
Case 404
basedmgcal = 50
Case 405
basedmgcal = 45
Case 406
basedmgcal = 10
Case 407
basedmgcal = 5
Case 408
basedmgcal = 1
Case 409
basedmgcal = 45
Case 410
basedmgcal = 25
Case 411
basedmgcal = 55
Case 412
basedmgcal = 55
Case 413
basedmgcal = 45
Case 417
basedmgcal = 18
Case 418
basedmgcal = 35
Case 426
basedmgcal = 5


Case 500
basedmgcal = 105
Case 501
basedmgcal = 105
Case 502
basedmgcal = 105
Case 503
basedmgcal = 60
Case 504
basedmgcal = 50
Case 505
basedmgcal = 45
Case 506
basedmgcal = 15
Case 507
basedmgcal = 10
Case 508
basedmgcal = 5
Case 509
basedmgcal = 50
Case 510
basedmgcal = 25
Case 511
basedmgcal = 55
Case 512
basedmgcal = 55
Case 513
basedmgcal = 50
Case 514
basedmgcal = 70
Case 515
basedmgcal = 70
Case 516
basedmgcal = 75
Case 517
basedmgcal = 105
Case 518
basedmgcal = 120
Case 519
basedmgcal = 75
Case 526
basedmgcal = 10


Case 600
basedmgcal = 75
Case 601
basedmgcal = 70
Case 602
basedmgcal = 70
Case 603
If c = 0 Then
basedmgcal = 85
Else
basedmgcal = 40
End If
Case 604
If c = 0 Then
basedmgcal = 80
Else
basedmgcal = 35
End If
Case 605
If c = 0 Then
basedmgcal = 75
Else
basedmgcal = 8
End If
Case 606
If c = 0 Then
basedmgcal = 55
Else
basedmgcal = 8
End If
Case 607
If c = 0 Then
basedmgcal = 35
Else
basedmgcal = 5
End If
Case 608
If c = 0 Then
basedmgcal = 20
Else
basedmgcal = 1
End If
Case 609
If c = 0 Then
basedmgcal = 70
Else
basedmgcal = 45
End If
Case 610
If c = 0 Then
basedmgcal = 30
Else
basedmgcal = 1
End If
Case 611
If c = 0 Then
basedmgcal = 85
Else
basedmgcal = 55
End If
Case 612
If c = 0 Then
basedmgcal = 85
Else
basedmgcal = 55
End If
Case 613
If c = 0 Then
basedmgcal = 75
Else
basedmgcal = 45
End If
Case 617
basedmgcal = 18
Case 618
basedmgcal = 35
Case 620
basedmgcal = 8
Case 621
basedmgcal = 8
Case 622
basedmgcal = 9
Case 623
basedmgcal = 9
Case 624
basedmgcal = 18
Case 625
basedmgcal = 55
Case 626
If c = 0 Then
basedmgcal = 20
Else
basedmgcal = 1
End If


Case 700
basedmgcal = 90
Case 701
basedmgcal = 80
Case 702
basedmgcal = 80
Case 703
If c = 0 Then
basedmgcal = 95
Else
basedmgcal = 40
End If
Case 704
If c = 0 Then
basedmgcal = 90
Else
basedmgcal = 35
End If
Case 705
If c = 0 Then
basedmgcal = 90
Else
basedmgcal = 8
End If
Case 706
If c = 0 Then
basedmgcal = 70
Else
basedmgcal = 8
End If
Case 707
If c = 0 Then
basedmgcal = 55
Else
basedmgcal = 5
End If
Case 708
If c = 0 Then
basedmgcal = 35
Else
basedmgcal = 1
End If
Case 709
If c = 0 Then
basedmgcal = 85
Else
basedmgcal = 45
End If
Case 710
If c = 0 Then
basedmgcal = 35
Else
basedmgcal = 1
End If
Case 711
If c = 0 Then
basedmgcal = 90
Else
basedmgcal = 60
End If
Case 712
If c = 0 Then
basedmgcal = 90
Else
basedmgcal = 60
End If
Case 173
If c = 0 Then
basedmgcal = 90
Else
basedmgcal = 45
End If
Case 717
basedmgcal = 24
Case 718
basedmgcal = 40
Case 720
basedmgcal = 10
Case 721
basedmgcal = 10
Case 722
basedmgcal = 12
Case 723
basedmgcal = 12
Case 724
basedmgcal = 22
Case 725
basedmgcal = 55
Case 726
If c = 0 Then
basedmgcal = 35
Else
basedmgcal = 1
End If



Case 800
basedmgcal = 105
Case 801
basedmgcal = 95
Case 802
basedmgcal = 95
Case 803
If c = 0 Then
basedmgcal = 105
Else
basedmgcal = 45
End If
Case 804
If c = 0 Then
basedmgcal = 105
Else
basedmgcal = 40
End If
Case 805
If c = 0 Then
basedmgcal = 105
Else
basedmgcal = 10
End If
Case 806
If c = 0 Then
basedmgcal = 85
Else
basedmgcal = 10
End If
Case 807
If c = 0 Then
basedmgcal = 75
Else
basedmgcal = 10
End If
Case 808
If c = 0 Then
basedmgcal = 55
Else
basedmgcal = 1
End If
Case 809
If c = 0 Then
basedmgcal = 105
Else
basedmgcal = 45
End If
Case 810
If c = 0 Then
basedmgcal = 40
Else
basedmgcal = 1
End If
Case 811
If c = 0 Then
basedmgcal = 105
Else
basedmgcal = 65
End If
Case 812
If c = 0 Then
basedmgcal = 105
Else
basedmgcal = 65
End If
Case 813
If c = 0 Then
basedmgcal = 105
Else
basedmgcal = 45
End If
Case 817
basedmgcal = 35
Case 818
basedmgcal = 45
Case 820
basedmgcal = 12
Case 821
basedmgcal = 12
Case 822
basedmgcal = 14
Case 823
basedmgcal = 14
Case 824
basedmgcal = 28
Case 825
basedmgcal = 65
Case 826
If c = 0 Then
basedmgcal = 55
Else
basedmgcal = 1
End If


Case 900
basedmgcal = 90
Case 901
basedmgcal = 85
Case 902
basedmgcal = 85
Case 903
basedmgcal = 80
Case 904
basedmgcal = 75
Case 905
basedmgcal = 65
Case 906
basedmgcal = 60
Case 907
basedmgcal = 45
Case 908
basedmgcal = 35
Case 909
basedmgcal = 75
Case 910
basedmgcal = 55
Case 911
basedmgcal = 80
Case 912
basedmgcal = 80
Case 913
basedmgcal = 70
Case 920
basedmgcal = 45
Case 921
basedmgcal = 45
Case 922
basedmgcal = 55
Case 923
basedmgcal = 55
Case 924
basedmgcal = 65
Case 925
basedmgcal = 100
Case 926
basedmgcal = 45


Case 1000
basedmgcal = 75
Case 1001
basedmgcal = 65
Case 1002
basedmgcal = 65
Case 1003
basedmgcal = 75
Case 1004
basedmgcal = 75
Case 1005
basedmgcal = 75
Case 1006
basedmgcal = 75
Case 1007
basedmgcal = 65
Case 1008
basedmgcal = 55
Case 1009
basedmgcal = 65
Case 1010
basedmgcal = 55
Case 1011
basedmgcal = 70
Case 1012
basedmgcal = 70
Case 1013
basedmgcal = 65
Case 1017
basedmgcal = 45
Case 1018
basedmgcal = 55
Case 1026
basedmgcal = 55


Case 1100
basedmgcal = 95
Case 1101
basedmgcal = 90
Case 1102
basedmgcal = 90
Case 1103
basedmgcal = 90
Case 1104
basedmgcal = 85
Case 1105
basedmgcal = 75
Case 1106
basedmgcal = 70
Case 1107
basedmgcal = 55
Case 1108
basedmgcal = 45
Case 1109
basedmgcal = 80
Case 1110
basedmgcal = 65
Case 1111
basedmgcal = 85
Case 1112
basedmgcal = 85
Case 1113
basedmgcal = 80
Case 1120
basedmgcal = 55
Case 1121
basedmgcal = 55
Case 1122
basedmgcal = 65
Case 1123
basedmgcal = 65
Case 1124
basedmgcal = 75
Case 1125
basedmgcal = 105
Case 1125
basedmgcal = 55


Case 1214
basedmgcal = 100
Case 1215
basedmgcal = 100
Case 1216
basedmgcal = 100
Case 1217
basedmgcal = 120
Case 1218
basedmgcal = 120
Case 1219
basedmgcal = 100


Case 1414
basedmgcal = 55
Case 1415
basedmgcal = 65
Case 1416
basedmgcal = 80
Case 1417
basedmgcal = 120
Case 1418
basedmgcal = 120
Case 1419
basedmgcal = 65


Case 1500
basedmgcal = 115
Case 1501
basedmgcal = 110
Case 1502
basedmgcal = 110
Case 1503
basedmgcal = 105
Case 1504
basedmgcal = 105
Case 1505
basedmgcal = 85
Case 1506
basedmgcal = 105
Case 1507
basedmgcal = 95
Case 1508
basedmgcal = 75
Case 1509
basedmgcal = 105
Case 1510
basedmgcal = 80
Case 1511
basedmgcal = 105
Case 1512
basedmgcal = 95
Case 1513
basedmgcal = 105
Case 1520
basedmgcal = 85
Case 1521
basedmgcal = 85
Case 1522
basedmgcal = 95
Case 1523
basedmgcal = 50
Case 1524
basedmgcal = 95
Case 1525
basedmgcal = 120
Case 1526
basedmgcal = 90


Case 1600
basedmgcal = 55
Case 1601
basedmgcal = 45
Case 1602
basedmgcal = 45
Case 1603
basedmgcal = 18
Case 1604
basedmgcal = 15
Case 1605
basedmgcal = 5
Case 1606
basedmgcal = 8
Case 1607
basedmgcal = 5
Case 1608
basedmgcal = 1
Case 1609
basedmgcal = 15
Case 1610
basedmgcal = 5
Case 1611
basedmgcal = 20
Case 1612
basedmgcal = 20
Case 1613
basedmgcal = 15
Case 1614
basedmgcal = 40
Case 1615
basedmgcal = 45
Case 1616
basedmgcal = 55
Case 1617
basedmgcal = 75
Case 1618
basedmgcal = 90
Case 1619
basedmgcal = 45
Case 1626
basedmgcal = 1


Case 1700
basedmgcal = 75
Case 1701
basedmgcal = 60
Case 1702
basedmgcal = 60
Case 1703
If c = 0 Then
basedmgcal = 75
Else
basedmgcal = 30
End If
Case 1704
If c = 0 Then
basedmgcal = 75
Else
basedmgcal = 30
End If
Case 1705
If c = 0 Then
basedmgcal = 10
Else
basedmgcal = 1
End If
Case 1706
If c = 0 Then
basedmgcal = 70
Else
basedmgcal = 8
End If
Case 1707
If c = 0 Then
basedmgcal = 45
Else
basedmgcal = 8
End If
Case 1708
If c = 0 Then
basedmgcal = 35
Else
basedmgcal = 1
End If
Case 1709
If c = 0 Then
basedmgcal = 65
Else
basedmgcal = 25
End If
Case 1710
If c = 0 Then
basedmgcal = 20
Else
basedmgcal = 1
End If
Case 1711
If c = 0 Then
basedmgcal = 75
Else
basedmgcal = 35
End If
Case 1712
If c = 0 Then
basedmgcal = 55
Else
basedmgcal = 25
End If
Case 1713
If c = 0 Then
basedmgcal = 70
Else
basedmgcal = 20
End If
Case 1717
basedmgcal = 65
Case 1718
basedmgcal = 85
Case 1720
basedmgcal = 25
Case 1721
basedmgcal = 25
Case 1722
basedmgcal = 25
Case 1723
basedmgcal = 5
Case 1724
basedmgcal = 25
Case 1725
basedmgcal = 85
Case 1726
basedmgcal = 20


Case 1900
basedmgcal = 90
Case 1901
basedmgcal = 85
Case 1902
basedmgcal = 85
Case 1903
basedmgcal = 80
Case 1904
basedmgcal = 80
Case 1905
basedmgcal = 45
Case 1906
basedmgcal = 75
Case 1907
basedmgcal = 65
Case 1908
basedmgcal = 55
Case 1909
basedmgcal = 70
Case 1910
basedmgcal = 50
Case 1911
basedmgcal = 80
Case 1912
basedmgcal = 70
Case 1913
basedmgcal = 75
Case 1914
basedmgcal = 45
Case 1915
basedmgcal = 55
Case 1916
basedmgcal = 65
Case 1917
basedmgcal = 85
Case 1918
basedmgcal = 95
Case 1919
basedmgcal = 55
Case 1920
basedmgcal = 45
Case 1921
basedmgcal = 65
Case 1922
basedmgcal = 55
Case 1923
basedmgcal = 40
Case 1924
basedmgcal = 85
Case 1925
basedmgcal = 105
Case 1926
basedmgcal = 55


Case 2000
basedmgcal = 75
Case 2001
basedmgcal = 70
Case 2002
basedmgcal = 70
Case 2003
basedmgcal = 70
Case 2004
basedmgcal = 70
Case 2005
basedmgcal = 65
Case 2006
basedmgcal = 65
Case 2007
basedmgcal = 50
Case 2008
basedmgcal = 40
Case 2009
basedmgcal = 70
Case 2010
basedmgcal = 55
Case 2011
basedmgcal = 75
Case 2012
basedmgcal = 75
Case 2013
basedmgcal = 65
Case 2020
basedmgcal = 45
Case 2021
basedmgcal = 50
Case 2022
basedmgcal = 65
Case 2023
basedmgcal = 65
Case 2024
basedmgcal = 75
Case 2025
basedmgcal = 95
Case 2026
basedmgcal = 55


Case 2114
basedmgcal = 35
Case 2115
basedmgcal = 35
Case 2116
basedmgcal = 40
Case 2117
basedmgcal = 45
Case 2118
basedmgcal = 55
Case 2119
basedmgcal = 40


Case 2220
basedmgcal = 80
Case 2221
basedmgcal = 110
Case 2222
basedmgcal = 55
Case 2223
basedmgcal = 20
Case 2224
basedmgcal = 85
Case 2225
basedmgcal = 120


Case 2314
basedmgcal = 105
Case 2315
basedmgcal = 105
Case 2316
basedmgcal = 105
Case 2317
basedmgcal = 120
Case 2318
basedmgcal = 120
Case 2319
basedmgcal = 105
Case 2320
basedmgcal = 38
Case 2321
basedmgcal = 38
Case 2322
basedmgcal = 95
Case 2323
basedmgcal = 28
Case 2324
basedmgcal = 40
Case 2325
basedmgcal = 85


Case 2520
basedmgcal = 40
Case 2521
basedmgcal = 40
Case 2522
basedmgcal = 40
Case 2523
basedmgcal = 40
Case 2524
basedmgcal = 55
Case 2525
basedmgcal = 75


Case Else
basedmgcal = -1

End Select
End Function

'Terrain Defense
Private Function tercal(ByVal a As Double) As Double
Dim b As Double
b = 0
Select Case a
Case 0, 9, 11, 21, 22, 24, 25
b = 10
Case 1, 2, 3, 4, 5
b = 0
Case 6, 16, 17, 18, 19, 20
b = 30
Case 7, 14
b = 40
Case 8, 10, 12, 13, 15, 23
b = 20
End Select
tercal = b
End Function


Private Sub Form_Load()
'Initialization
'Initializing the Unit List
offunit.AddItem ("Infantry")
offunit.AddItem ("Mech")
offunit.AddItem ("Bike")
offunit.AddItem ("Recon")
offunit.AddItem ("Flare")
offunit.AddItem ("Antiair")
offunit.AddItem ("Tank")
offunit.AddItem ("MediumTank")
offunit.AddItem ("WarTank")
offunit.AddItem ("Artillery")
offunit.AddItem ("Antitank")
offunit.AddItem ("Rockets")
offunit.AddItem ("Missiles")
offunit.AddItem ("Rig")
offunit.AddItem ("Fighter")
offunit.AddItem ("Bomber")
offunit.AddItem ("Duster")
offunit.AddItem ("BattleCopter")
offunit.AddItem ("TransportCopter")
offunit.AddItem ("Seaplane")
offunit.AddItem ("Battleship")
offunit.AddItem ("Carrier")
offunit.AddItem ("Submarine")
offunit.AddItem ("Crusier")
offunit.AddItem ("Lander")
offunit.AddItem ("Gunboat")


defunit.AddItem ("Infantry")
defunit.AddItem ("Mech")
defunit.AddItem ("Bike")
defunit.AddItem ("Recon")
defunit.AddItem ("Flare")
defunit.AddItem ("Antiair")
defunit.AddItem ("Tank")
defunit.AddItem ("MediumTank")
defunit.AddItem ("WarTank")
defunit.AddItem ("Artillery")
defunit.AddItem ("Antitank")
defunit.AddItem ("Rockets")
defunit.AddItem ("Missiles")
defunit.AddItem ("Rig")
defunit.AddItem ("Fighter")
defunit.AddItem ("Bomber")
defunit.AddItem ("Duster")
defunit.AddItem ("BattleCopter")
defunit.AddItem ("TransportCopter")
defunit.AddItem ("Seaplane")
defunit.AddItem ("Battleship")
defunit.AddItem ("Carrier")
defunit.AddItem ("Submarine")
defunit.AddItem ("Crusier")
defunit.AddItem ("Lander")
defunit.AddItem ("Gunboat")
defunit.AddItem ("Meteor")

'Initializing the CO list
OffCO.AddItem ("No CO")
OffCO.AddItem ("Caulder t0")
OffCO.AddItem ("Grat t0")
OffCO.AddItem ("Isabella")
OffCO.AddItem ("Dracula")
OffCO.AddItem ("Grat t1")
OffCO.AddItem ("Mauser")
OffCO.AddItem ("Alexender")
OffCO.AddItem ("Tabitha t1")
OffCO.AddItem ("Brenner")
OffCO.AddItem ("Caulder t2")
OffCO.AddItem ("Gage")
OffCO.AddItem ("Tabitha t2")
OffCO.AddItem ("Waylon")
OffCO.AddItem ("Will")
OffCO.AddItem ("Hawk")
OffCO.AddItem ("Mira")
OffCO.AddItem ("Albatross")
OffCO.AddItem ("Lee")
OffCO.AddItem ("Tinker t2")
OffCO.AddItem ("Mammoth")
OffCO.AddItem ("Lin")
OffCO.AddItem ("Forsythe")
OffCO.AddItem ("Greyfield")
OffCO.AddItem ("Tasha")
OffCO.AddItem ("Tinker t3")
OffCO.AddItem ("zhaotiantong")
OffCO.AddItem ("Tabitha t0")

DefCO.AddItem ("No CO")
DefCO.AddItem ("Caulder t0")
DefCO.AddItem ("Grat t0")
DefCO.AddItem ("Isabella")
DefCO.AddItem ("Dracula")
DefCO.AddItem ("Grat t1")
DefCO.AddItem ("Mauser")
DefCO.AddItem ("Alexender")
DefCO.AddItem ("Tabitha t1")
DefCO.AddItem ("Brenner")
DefCO.AddItem ("Caulder t2")
DefCO.AddItem ("Gage")
DefCO.AddItem ("Tabitha t2")
DefCO.AddItem ("Waylon")
DefCO.AddItem ("Will")
DefCO.AddItem ("Hawk")
DefCO.AddItem ("Mira")
DefCO.AddItem ("Albatross")
DefCO.AddItem ("Lee")
DefCO.AddItem ("Tinker t2")
DefCO.AddItem ("Mammoth")
DefCO.AddItem ("Lin")
DefCO.AddItem ("Forsythe")
DefCO.AddItem ("Greyfield")
DefCO.AddItem ("Tasha")
DefCO.AddItem ("Tinker t3")
DefCO.AddItem ("zhaotiantong")
DefCO.AddItem ("Tabitha t0")

'Initializing Terrain Lists
OffTer.AddItem ("Plain")
OffTer.AddItem ("River")
OffTer.AddItem ("Sea")
OffTer.AddItem ("Beach")
OffTer.AddItem ("Road")
OffTer.AddItem ("Bridge")
OffTer.AddItem ("Wood")
OffTer.AddItem ("Mountain")
OffTer.AddItem ("Wasteland")
OffTer.AddItem ("Ruins")
OffTer.AddItem ("Rough")
OffTer.AddItem ("Mistonsea")
OffTer.AddItem ("Reef")
OffTer.AddItem ("Silo")
OffTer.AddItem ("Headquarters")
OffTer.AddItem ("City")
OffTer.AddItem ("Commandtower")
OffTer.AddItem ("Radar")
OffTer.AddItem ("Factory")
OffTer.AddItem ("Airport")
OffTer.AddItem ("Seaport")
OffTer.AddItem ("TempAirport")
OffTer.AddItem ("TempSeaport")
OffTer.AddItem ("MistonPlain")
OffTer.AddItem ("MistonRiver")
OffTer.AddItem ("MistonBeach")

DefTer.AddItem ("Plain")
DefTer.AddItem ("River")
DefTer.AddItem ("Sea")
DefTer.AddItem ("Beach")
DefTer.AddItem ("Road")
DefTer.AddItem ("Bridge")
DefTer.AddItem ("Wood")
DefTer.AddItem ("Mountain")
DefTer.AddItem ("Wasteland")
DefTer.AddItem ("Ruins")
DefTer.AddItem ("Rough")
DefTer.AddItem ("Mistonsea")
DefTer.AddItem ("Reef")
DefTer.AddItem ("Silo")
DefTer.AddItem ("Headquarters")
DefTer.AddItem ("City")
DefTer.AddItem ("Commandtower")
DefTer.AddItem ("Radar")
DefTer.AddItem ("Factory")
DefTer.AddItem ("Airport")
DefTer.AddItem ("Seaport")
DefTer.AddItem ("TempAirport")
DefTer.AddItem ("TempSeaport")
DefTer.AddItem ("MistonPlain")
DefTer.AddItem ("MistonRiver")
DefTer.AddItem ("MistonBeach")

'Data Initialization
Luck1.Text = "0"
Luck2.Text = "0"
OffTow.Text = "0"
DefTow.Text = "0"
OffHP.Text = "100"
DefHP.Text = "100"
Offlev = "0"
Deflev = "0"

End Sub

