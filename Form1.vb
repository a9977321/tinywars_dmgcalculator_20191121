Public Class Form1

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim basedmg1 As Double, offunit1 As Double, defunit1 As Double, subweapon01 As Boolean, subweapon02 As Boolean, offcomod1 As Double, defcomod1 As Double
        Dim offco1 As Double, defco1 As Double, tow1 As Double, tow2 As Double, lev2 As Double, offpow1 As Boolean, defpow1 As Boolean
        Dim HP1 As Double, HP2 As Double, offter1 As Double, defter1 As Double, offlevmod1 As Double, deflevmod1 As Double, offlevmod2 As Double, deflevmod2 As Double
        Dim lev1 As Double, luc1 As Double, luc2 As Double, offresult1 As Double, defresult1 As Double, offcomod2 As Double, defcomod2 As Double, basedmg2 As Double
        Dim cflag As Boolean
        'Obtaining Data
        offunit1 = offunit.SelectedIndex
        defunit1 = DefUnit.SelectedIndex
        subweapon01 = Subweapon1.Checked
        subweapon02 = Subweapon2.Checked
        offco1 = offCO.SelectedIndex
        defco1 = DefCO.SelectedIndex
        offpow1 = Offpower.Checked
        defpow1 = Defpower.Checked
        lev1 = Val(Offlev.Text)
        lev2 = Val(Deflev.Text)
        tow1 = Val(Offtow.Text)
        tow2 = Val(DefTow.Text)
        HP1 = Val(OffHP.Text)
        HP2 = Val(DefHP.Text)
        offter1 = Offter.SelectedIndex
        defter1 = Defter.SelectedIndex
        luc1 = Val(Luck1.Text)
        luc2 = Val(Luck2.Text)
        offcomod1 = 100
        defcomod1 = 100
        offcomod2 = 100
        defcomod2 = 100
        cflag = False
        'Error Page
        'Booleans are used as numbers in some cases, but the function remains. 

        'Data calculation
        basedmg1 = basedmgcal(offunit1, defunit1, subweapon01)
        offcomod1 = offcomod1 + offcomodcal(offco1, offunit1, offpow1)
        offlevmod1 = levdmgmod(lev1)
        deflevmod1 = levdefmod(lev2)
        defter1 = tercal(defter1)
        offter1 = tercal(offter1)

        'Air units and meteors do not enjoy terrain defense
        Select Case offunit1
            Case 14, 15, 16, 17, 18, 19
                offter1 = 0
        End Select
        Select Case defunit1
            Case 14, 15, 16, 17, 18, 19, 26
                defter1 = 0
        End Select

        defcomod1 = defcomod1 + defcomodcal(defco1, defunit1, defpow1)
        offresult1 = Fix((basedmg1 * (offcomod1 + offlevmod1 + 5 * tow1) + luc1 * 100) * ((HP1 + 9) \ 10) / (defcomod1 + defter1 * ((HP2 + 9) \ 10) / 10 + 5 * tow2 + deflevmod1) / 10)
        HP2 = HP2 - offresult1

        'Destroyed units cannot counterattack
        If HP2 > 0 Then
            cflag = True
        End If
        'Indirects cannot be counterattacked
        Select Case offunit1
            Case 9, 11, 12, 20
                cflag = False
        End Select
        'Indirects cannot counterattack
        Select Case defunit1
            Case 9, 11, 12, 20
                cflag = False
        End Select

        If cflag Then
            basedmg2 = basedmgcal(defunit1, offunit1, subweapon02)
            offcomod2 = offcomod2 + offcomodcal(defco1, defunit1, defpow1)
            offlevmod2 = levdmgmod(lev2)
            deflevmod2 = levdmgmod(lev1)
            defcomod2 = defcomod2 + defcomodcal(offco1, offunit1, offpow1)
            defresult1 = Fix((basedmg2 * (offcomod2 + offlevmod2 + 5 * tow2) + luc2 * 100) * ((HP2 + 9) \ 10) / (defcomod2 + offter1 * ((HP1 + 9) \ 10) / 10 + 5 * tow1 + deflevmod2) / 10)
        Else
            defresult1 = 0
        End If

        'Results output
        OffResult.Text = "Attack" & Space(2) & Str(offresult1)
        DefResult.Text = "Counterattack" & Space(1) & Str(defresult1)
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
            Case Else
                levdmgmod = 0
        End Select
    End Function
    Private Function levdefmod(ByVal lev As Double) As Double
        If lev = 3 Then
            levdefmod = 20
        Else
            levdefmod = 0
        End If
    End Function

    Private Function offcomodcal(ByRef co As Double, ByRef unit As Double, ByRef pow As Double) As Double
        'CO damage modifier calculator
        Dim a As Double
        a = 0
        Select Case co
            Case 0
                'No CO T2
                a = 0
            Case 1
                'Brenner T2
                Call dmgmod0102(a)
            Case 2
                'Caulder T2
                Call dmgmod0105(a)
            Case 3
                'Forsythe T2
                Call dmgmod0104(a)
            Case 4
                'Gage T2
                Call dmgmod0102(a)
                Call dmgmod0504(unit, a)
            Case 5
                'Grat T2
                If pow = 0 Then
                    Call dmgmod0102(a)
                Else
                    Call dmgmod0104(a)
                End If
            Case 6
                'Greyfield T2
                Call dmgmod0102(a)
                Call dmgmod1102(unit, a)
            Case 7
                'Hawk T2
                Call dmgmod0104(a)
            Case 8
                'Isabella T2
                Call dmgmod0104(a)
            Case 9
                'Lee T2
                If pow = 0 Then
                    Call dmgmod0102(a)
                Else
                    Call dmgmod0104(a)
                End If
            Case 10
                'Lin T2
                Call dmgmod0102(a)
                Call dmgmod1004(unit, a)
            Case 11
                'Mira T2
                a = 0
            Case 12
                'Tabitha T2
                Call dmgmod0112(a)
            Case 13
                'Tasha T2
                Call dmgmod0102(a)
                Call dmgmod0608(unit, a)
            Case 14
                'Tinker T2
                If pow = 0 Then
                    Call dmgmod0102(a)
                Else
                    Call dmgmod0108(a)
                End If
            Case 15
                'Waylon T2
                Call dmgmod0102(a)
                Call dmgmod0604(unit, a)
            Case 16
                'Will T2
                Call dmgmod0102(a)
                Call dmgmod0704(unit, a)
            Case 17
                'zhaotiantong T2
                Call dmgmod0102(a)
        End Select
        offcomodcal = a
    End Function



    'All units +10% offense    Common
    Private Sub dmgmod0102(ByRef dmg As Double)
        dmg = dmg + 10
    End Sub
    'All units +20% offense    Forsythe T2, Grat T2, Hawk T2, Isabella T2, Lee T2
    Private Sub dmgmod0104(ByRef dmg As Double)
        dmg = dmg + 20
    End Sub
    'All units +25% offense    Caulder T2
    Private Sub dmgmod0105(ByRef dmg As Double)
        dmg = dmg + 25
    End Sub

    'All units +40% offense    Tinker T2
    Private Sub dmgmod0108(ByRef dmg As Double)
        dmg = dmg + 40
    End Sub

    'All units +60% offense    Tabitha T2
    Private Sub dmgmod0112(ByRef dmg As Double)
        dmg = dmg + 60
    End Sub

    'All naval & indirect units +20% offense    Gage T2
    Private Sub dmgmod0504(ByVal a As Double, ByRef dmg As Double)
        Select Case a
            Case 9, 10, 11, 12, 20, 21, 22, 23, 24, 25
                dmg = dmg + 20
        End Select
    End Sub
    'All air units +20% offense    Waylon T2
    Private Sub dmgmod0604(ByVal a As Double, ByRef dmg As Double)
        Select Case a
            Case 14, 15, 16, 17, 18, 19
                dmg = dmg + 20
        End Select
    End Sub
    'All air units +40% offense    Tasha T2
    Private Sub dmgmod0608(ByVal a As Double, ByRef dmg As Double)
        Select Case a
            Case 14, 15, 16, 17, 18, 19
                dmg = dmg + 40
        End Select
    End Sub
    'All land direct units +20% offense    Will T2
    Private Sub dmgmod0704(ByVal a As Double, ByRef dmg As Double)
        Select Case a
            Case 0, 1, 2, 3, 4, 5, 6, 7, 8
                dmg = dmg + 20
        End Select
    End Sub

    'All land units +20% offense    Lin T2
    Private Sub dmgmod1004(ByVal a As Double, ByRef dmg As Double)
        Select Case a
            Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13
                dmg = dmg + 20
        End Select
    End Sub
    'All copters,seaplane & naval units +10% offense    Greyfield T2
    Private Sub dmgmod1102(ByVal a As Double, ByRef dmg As Double)
        Select Case a
            Case 17, 18, 19, 20, 21, 22, 23, 24, 25
                dmg = dmg + 10
        End Select
    End Sub

    Private Function defcomodcal(ByVal co As Double, ByVal unit As Double, ByVal pow As Double) As Double
        'CO defense modifier calculator
        Dim a As Double
        a = 0
        Select Case co
            Case 0
                'No CO
                a = 0
            Case 1
                'Brenner T2
                Call defmod0106(a)
            Case 2
                'Caulder T2
                Call defmod0105(a)
            Case 3
                'Forsythe T2
                Call defmod0104(a)
            Case 4
                'Gage T2
                Call defmod0102(a)
                Call defmod0402(unit, a)
            Case 5
                'Grat T2
                If pow = 0 Then
                    Call defmod0102(a)
                Else
                    Call defmod0104(a)
                End If
            Case 6
                'Greyfield T2
                Call defmod0102(a)
                Call defmod0908(unit, a)
            Case 7
                'Hawk T2
                Call defmod0102(a)
            Case 8
                'Isabella T2
                Call defmod0104(a)
            Case 9
                'Lee T2
                Call defmod0102(a)
            Case 10
                'Lin T2
                Call defmod0102(a)
                Call defmod0804(unit, a)
            Case 11
                'Mira T2
                a = 0
            Case 12
                'Tabitha T2
                Call defmod0112(a)
            Case 13
                'Tasha T2
                Call defmod0102(a)
                Call defmod0504(unit, a)
            Case 14
                'Tinker T2
                If pow = 0 Then
                    Call defmod0102(a)
                Else
                    Call defmod01a1(a)
                End If
            Case 15
                'Waylon T2
                If pow = 0 Then
                    Call defmod0102(a)
                    Call defmod0506(unit, a)
                Else
                    Call defmod0102(a)
                    Call defmod0560(unit, a)
                End If
            Case 16
                'Will T2
                Call defmod0102(a)
            Case 17
                'zhaotiantong T2
                Call defmod0102(a)
        End Select
        defcomodcal = a
    End Function
    'All units -10% defense    Tinker T2
    Private Sub defmod01a1(ByRef def As Double)
        def = def - 10
    End Sub
    'All units +10% defense    Common
    Private Sub defmod0102(ByRef def As Double)
        def = def + 10
    End Sub
    'All units +20% defense    Forsythe T2, Grat T2, Isabella T2
    Private Sub defmod0104(ByRef def As Double)
        def = def + 20
    End Sub
    'All units +25% defense    Caulder T2
    Private Sub defmod0105(ByRef def As Double)
        def = def + 25
    End Sub
    'All units +30% defense    Brenner T2
    Private Sub defmod0106(ByRef def As Double)
        def = def + 30
    End Sub
    'All units +60% defense    Tabitha T2
    Private Sub defmod0112(ByRef def As Double)
        def = def + 60
    End Sub
    'All naval & indirect units +10% defense    Gage T2
    Private Sub defmod0402(ByVal a As Double, ByRef def As Double)
        Select Case a
            Case 9, 10, 11, 12, 20, 21, 22, 23, 24, 25
                def = def + 10
        End Select
    End Sub
    'All air units +20% defense    Tasha T2
    Private Sub defmod0504(ByVal a As Double, ByRef def As Double)
        Select Case a
            Case 14, 15, 16, 17, 18, 19
                def = def + 20
        End Select
    End Sub
    'All air units +30% defense    Waylon T2
    Private Sub defmod0506(ByVal a As Double, ByRef def As Double)
        Select Case a
            Case 14, 15, 16, 17, 18, 19
                def = def + 30
        End Select
    End Sub
    'All air units +300% defense    Waylon T2
    Private Sub defmod0560(ByVal a As Double, ByRef def As Double)
        Select Case a
            Case 14, 15, 16, 17, 18, 19
                def = def + 300
        End Select
    End Sub
    'All land units +20% defense    Lin T2
    Private Sub defmod0804(ByVal a As Double, ByRef def As Double)
        Select Case a
            Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13
                def = def + 20
        End Select
    End Sub
    'All copters,seaplane & naval units +40% defense    Greyfield T2
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




    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        'Data Initialization
        Luck1.Text = "0"
        Luck2.Text = "0"
        Offtow.Text = "0"
        DefTow.Text = "0"
        OffHP.Text = "100"
        DefHP.Text = "100"
        Offlev.Text = "0"
        Deflev.Text = "0"
    End Sub
End Class
