Attribute VB_Name = "modGameLogic"
Private Board(1 To BOARD_ROWS, 1 To BOARD_COLS, 1 To TIME_END) As String    '(i,j,1) = current, (i,j,2-TIME_END) = animation
Private BG(1 To BOARD_ROWS, 1 To BOARD_COLS, 1 To TIME_END) As String
Private StructureStatus(1 To BOARD_COLS, 1 To TIME_END) As Integer   'Column, Time  = 0-4
Private StructureStatusRepairs(1 To 2, 1 To BOARD_COLS) As Integer
Private InProcess As Boolean
Private Wave As Integer
Private TurnCount As Integer
Private GameScore As Long
Private ShotCount As Integer

Private EFighters(1 To EF_MAX_COUNT, 1 To 4, 1 To TIME_END) As Integer
Private EFFire(1 To EF_MAX_COUNT, 1 To 3, 1 To TIME_END) As Integer 'Aircraft #, Row/Column/Direction, Time

Private EBombers(1 To EB_MAX_COUNT, 1 To 5, 1 To TIME_END) As Integer        '(Aircraft #, Row/Column/Speed(Rnd(1, 2, or 3))/Direction(1 or 2 [L or R])/Damage(0 = No, 1 = Yes), Time)
Private EBFirePrime(1 To EB_MAX_COUNT) As Integer             'Fire at 3, counts on turns
Private EBFire(1 To EB_MAX_COUNT, 1 To 2, 1 To TIME_END) As Integer          '(Aircraft #, Row/Column)

Private RAFStatus(1 To RAF_MAX_COUNT, 1 To 7, 1 To TIME_END) As Integer  '(Aircraft #, Row/Column/Speed(1 - 4)/Direction/# Onboard Rockets/OnGround(1) vs InAir(2)/Health, Time)
Private RAFRepairs(1 To 5, 1 To 2) As Integer           '(Aircraft # (5 = total change), repairs, rocketchange)
Private Rockets As Integer                                    'Number of Rockets In Storage
Private RAFFire(1 To RAF_MAX_COUNT, 1 To 4, 1 To TIME_END) As Integer  '(Aircraft #, Row/Column/Type(1 = Norm, 2 = Rockets)/Direction, Time)
Private RAFPos As Integer
Private RAFCull As Boolean
Private RocketsChange As Integer

Private NumEF As Integer
Private NumEFRemain As Integer
Private NumEB As Integer
Private NumEBRemain As Integer
Private NumRAFRemain As Integer

Private AirfieldLeft As Integer
Private RepairServ1 As Integer
Private RepairServ2 As Integer
Private NumRepairServRemain As Integer
Private ComC As Integer
Private Bunker1 As Integer
Private Bunker2 As Integer
Private NumBunkerRemain As Integer
Private CityStruct(1 To 12) As Integer

Private Turrets(1 To BOARD_COLS, 1 To TIME_END) As Turret

Private TurretAimNom As TurretFireDir
Private SkipTurretNom As Boolean
Private NumTurRemain As Integer
Private TurretPos As Integer
Private TurretCount As Integer
Private PlacingTurretsInit As Boolean
Private PlacingTurrets As Boolean

Private PurchaseCost As Integer
Private RepairActive As Boolean
Private AmmoActive As Boolean
Private RAFShopActive As Boolean
Private RAFPreLaunch As Boolean
Private RAFComm As Boolean

#If Win64 Then
   Declare PtrSafe Function GetTickCount64 Lib "kernel32" () As LongLong
#Else
   Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#End If

Sub StartGame_Click()
    If InProcess = True Then
        Restart = 0
        Restart = MsgBox("Previous game has not yet finished. Do you wish to restart?", vbYesNo, "Abort game?")
        If Restart = vbYes Then InProcess = False
    End If
    If InProcess = False Then
    
        Worksheets("London Siege").StartGame.Caption = "Restart"
        
        Cells(12, 12) = " Loading..."
        Application.Wait (Now + TimeValue("0:00:01"))
    
        Call CleanHouse

        InProcess = True
        
        'Place airfield
        Randomize
        AirfieldLeft = Int((BOARD_COLS - RAF_MAX_COUNT + 1) * Rnd() + 1)
        For i = 0 To RAF_MAX_COUNT - 1
            Board(STRUCT_ROW, AirfieldLeft + i, 1) = "     "
            Range(Cells(GRID_BOT_BRDR - STRUCT_ROW, GRID_LEFT_BRDR + AirfieldLeft + i).Address).Font.Underline = xlUnderlineStyleSingle
        Next i
        
        'Board(1, AirfieldLeft, 1) = RAF_ICON
        'RAFStatus(1, 1, 1) = 1
        'RAFStatus(1, 2, 1) = AirfieldLeft
        'RAFStatus(1, 5, 1) = 4
        'RAFStatus(1, 6, 1) = 1
        'RAFStatus(1, 7, 1) = 3
        'NumRAFRemain = 1
        
        'Place repair shops
        RepairServ1p = False
        Do While RepairServ1p = False
            Randomize
            RepairServ1 = Int(BOARD_COLS * Rnd() + 1)
            If Board(STRUCT_ROW, RepairServ1, 1) = "" Then
                Board(STRUCT_ROW, RepairServ1, 1) = "#"
                RepairServ1p = True
            End If
        Loop
        RepairServ2p = False
        Do While RepairServ2p = False
            Randomize
            RepairServ2 = Int(BOARD_COLS * Rnd() + 1)
            If Board(STRUCT_ROW, RepairServ2, 1) = "" Then
                Board(STRUCT_ROW, RepairServ2, 1) = "#"
                RepairServ2p = True
            End If
        Loop
        NumRepairServRemain = 2
        'Place command center
        ComCp = False
        Do While ComCp = False
            Randomize
            ComC = Int(BOARD_COLS * Rnd() + 1)
            If Board(STRUCT_ROW, ComC, 1) = "" Then
                Board(STRUCT_ROW, ComC, 1) = "*"
                ComCp = True
            End If
        Loop
        'Place weapons bunkers
        Bunker1p = False
        Do While Bunker1p = False
            Randomize
            Bunker1 = Int(BOARD_COLS * Rnd() + 1)
            If Board(STRUCT_ROW, Bunker1, 1) = "" Then
                Board(STRUCT_ROW, Bunker1, 1) = "$"
                Bunker1p = True
            End If
        Loop
        Bunker2p = False
        Do While Bunker2p = False
            Randomize
            Bunker2 = Int(BOARD_COLS * Rnd() + 1)
            If Board(STRUCT_ROW, Bunker2, 1) = "" Then
                Board(STRUCT_ROW, Bunker2, 1) = "$"
                Bunker2p = True
            End If
        Loop
        NumBunkerRemain = 2
        
        R = 1
        For j = 1 To BOARD_COLS
            If Board(STRUCT_ROW, j, 1) = "" Then
                Board(STRUCT_ROW, j, 1) = "l=l"
                CityStruct(R) = j
                R = R + 1
            End If
        Next j
        
        Application.ScreenUpdating = False
        
        For i = 1 To BOARD_ROWS
            For j = 1 To BOARD_COLS
                Cells(GRID_BOT_BRDR - i, GRID_LEFT_BRDR + j) = Board(i, j, 1)
            Next j
        Next i
        
        
        Cells(14, 24) = "Command          *"
        Cells(14, 25) = 1
        Cells(15, 24) = "Repair Services #"
        Cells(15, 25) = "2 of 2"
        Cells(16, 24) = "Ammo Bunkers $"
        Cells(16, 25) = "2 of 2"
        Cells(17, 24) = "Airfield           ___"
        Cells(17, 25) = 1
        Cells(18, 24) = "City Structures"
        Cells(18, 25) = "12 of 12"
        Cells(23, 24) = GameScore
        
        Worksheets("London Siege").CheckTur1.Visible = True
        Worksheets("London Siege").CheckTur2.Visible = True
        Worksheets("London Siege").CheckTur3.Visible = True
        Worksheets("London Siege").CheckTur4.Visible = True
        Worksheets("London Siege").CheckTur5.Visible = True
        Worksheets("London Siege").CheckTur6.Visible = True
        Worksheets("London Siege").CheckTur7.Visible = True
        Worksheets("London Siege").CheckTur8.Visible = True
        Worksheets("London Siege").CheckTur9.Visible = True
        Worksheets("London Siege").CheckTur10.Visible = True
        Worksheets("London Siege").CheckTur11.Visible = True
        Worksheets("London Siege").CheckTur12.Visible = True
        Worksheets("London Siege").CheckTur13.Visible = True
        Worksheets("London Siege").CheckTur14.Visible = True
        Worksheets("London Siege").CheckTur15.Visible = True
        Worksheets("London Siege").CheckTur16.Visible = True
        Worksheets("London Siege").CheckTur17.Visible = True
        Worksheets("London Siege").CheckTur18.Visible = True
        Worksheets("London Siege").CheckTur19.Visible = True
        Worksheets("London Siege").CheckTur20.Visible = True
        Worksheets("London Siege").CheckTur21.Visible = True
        
        Cells(13, 24) = "Turrets"
        Cells(13, 25) = TurretCount
        
        Cells(6, 12) = "Place 3 Turrets"
        PlacingTurretsInit = True
        
        Worksheets("London Siege").NextWave.Caption = "Night 1"
        Worksheets("London Siege").NextWave.Enabled = False
        Worksheets("London Siege").NextTurn.Enabled = False
        Worksheets("London Siege").QuitGame.Enabled = True
        Worksheets("London Siege").Protect UserInterfaceOnly:=True
        Application.ScreenUpdating = True
        
    End If
End Sub
        
Sub StartFinish()

        Application.ScreenUpdating = False
        PlacingTurretsInit = False
        
        Cells(6, 12) = ""
        Worksheets("London Siege").CheckTur1.Visible = False
        Worksheets("London Siege").CheckTur2.Visible = False
        Worksheets("London Siege").CheckTur3.Visible = False
        Worksheets("London Siege").CheckTur4.Visible = False
        Worksheets("London Siege").CheckTur5.Visible = False
        Worksheets("London Siege").CheckTur6.Visible = False
        Worksheets("London Siege").CheckTur7.Visible = False
        Worksheets("London Siege").CheckTur8.Visible = False
        Worksheets("London Siege").CheckTur9.Visible = False
        Worksheets("London Siege").CheckTur10.Visible = False
        Worksheets("London Siege").CheckTur11.Visible = False
        Worksheets("London Siege").CheckTur12.Visible = False
        Worksheets("London Siege").CheckTur13.Visible = False
        Worksheets("London Siege").CheckTur14.Visible = False
        Worksheets("London Siege").CheckTur15.Visible = False
        Worksheets("London Siege").CheckTur16.Visible = False
        Worksheets("London Siege").CheckTur17.Visible = False
        Worksheets("London Siege").CheckTur18.Visible = False
        Worksheets("London Siege").CheckTur19.Visible = False
        Worksheets("London Siege").CheckTur20.Visible = False
        Worksheets("London Siege").CheckTur21.Visible = False

        For i = 1 To BOARD_ROWS
            For j = 1 To BOARD_COLS
                Cells(GRID_BOT_BRDR - i, GRID_LEFT_BRDR + j) = Board(i, j, 1)
            Next j
        Next i
        
        Worksheets("London Siege").NextWave.Caption = "Night 1"
        Worksheets("London Siege").NextWave.Enabled = True
        Worksheets("London Siege").NextTurn.Enabled = False
        Worksheets("London Siege").QuitGame.Enabled = True
        
        Wave = 1
        
        Application.ScreenUpdating = True
    
End Sub

Function AddTurret(j As Integer)
    Turrets(j, 1).Icon = TURRET_ICON_SINGLE
    TurretCount = TurretCount + 1
    NumTurRemain = NumTurRemain + 1
    Cells(13, 25) = NumTurRemain
    Cells(21, 1 + j) = Turrets(j, 1).Icon
    Board(TURRET_ROW, j, 1) = Turrets(j, 1).Icon
    If PlacingTurretsInit = True And TurretCount = 3 Then
        StartFinish
    End If
End Function

Sub CheckTur1()
    Worksheets("London Siege").CheckTur1.Visible = False
    addl = AddTurret(1)
End Sub

Sub CheckTur2()
    Worksheets("London Siege").CheckTur2.Visible = False
    addl = AddTurret(2)
End Sub

Sub CheckTur3()
    Worksheets("London Siege").CheckTur3.Visible = False
    addl = AddTurret(3)
End Sub

Sub CheckTur4()
    Worksheets("London Siege").CheckTur4.Visible = False
    addl = AddTurret(4)
End Sub

Sub CheckTur5()
    Worksheets("London Siege").CheckTur5.Visible = False
    addl = AddTurret(5)
End Sub

Sub CheckTur6()
    Worksheets("London Siege").CheckTur6.Visible = False
    addl = AddTurret(6)
End Sub

Sub CheckTur7()
    Worksheets("London Siege").CheckTur7.Visible = False
    addl = AddTurret(7)
End Sub

Sub CheckTur8()
    Worksheets("London Siege").CheckTur8.Visible = False
    addl = AddTurret(8)
End Sub

Sub CheckTur9()
    Worksheets("London Siege").CheckTur9.Visible = False
    addl = AddTurret(9)
End Sub

Sub CheckTur10()
    Worksheets("London Siege").CheckTur10.Visible = False
    addl = AddTurret(10)
End Sub

Sub CheckTur11()
    Worksheets("London Siege").CheckTur11.Visible = False
    addl = AddTurret(11)
End Sub

Sub CheckTur12()
    Worksheets("London Siege").CheckTur12.Visible = False
    addl = AddTurret(12)
End Sub

Sub CheckTur13()
    Worksheets("London Siege").CheckTur13.Visible = False
    addl = AddTurret(13)
End Sub

Sub CheckTur14()
    Worksheets("London Siege").CheckTur14.Visible = False
    addl = AddTurret(14)
End Sub

Sub CheckTur15()
    Worksheets("London Siege").CheckTur15.Visible = False
    addl = AddTurret(15)
End Sub

Sub CheckTur16()
    Worksheets("London Siege").CheckTur16.Visible = False
    addl = AddTurret(16)
End Sub

Sub CheckTur17()
    Worksheets("London Siege").CheckTur17.Visible = False
    addl = AddTurret(17)
End Sub

Sub CheckTur18()
    Worksheets("London Siege").CheckTur18.Visible = False
    addl = AddTurret(18)
End Sub

Sub CheckTur19()
    Worksheets("London Siege").CheckTur19.Visible = False
    addl = AddTurret(19)
End Sub

Sub CheckTur20()
    Worksheets("London Siege").CheckTur20.Visible = False
    addl = AddTurret(20)
End Sub

Sub CheckTur21()
    Worksheets("London Siege").CheckTur21.Visible = False
    addl = AddTurret(21)
End Sub

Sub NextTurn_Click()

    Application.ScreenUpdating = False
    Worksheets("London Siege").Protect UserInterfaceOnly:=True
    Worksheets("London Siege").StartGame.Enabled = False
    Worksheets("London Siege").NextWave.Enabled = False
    Worksheets("London Siege").NextTurn.Enabled = False
    Worksheets("London Siege").QuitGame.Enabled = False
    
    Cells(23, 26) = ""                   'Remove bonus counts from last turn
    
    FSpeDir = 1
    Do While FSpeDir <= NumEF
        FSpe = 0
        FDir = 0
        If EFighters(FSpeDir, 3, 1) = 1 Then
            Randomize
            FSpe = Int(3 * Rnd + 1)
            Randomize
            FDir = Int(8 * Rnd + 1)
        ElseIf EFighters(FSpeDir, 3, 1) = 2 Then
            Randomize
            FSpe = Int(4 * Rnd + 1)
            Randomize
            FDir = (Int(5 * Rnd + 1) - 3) + EFighters(FSpeDir, 4, 1)
            If FDir > 8 Then
                FDir = FDir - 8
            ElseIf FDir < 1 Then
                FDir = FDir + 8
            End If
        ElseIf EFighters(FSpeDir, 3, 1) = 3 Then
            Randomize
            FSpe = Int(4 * Rnd + 1)
            Randomize
            FDir = (Int(3 * Rnd + 1) - 2) + EFighters(FSpeDir, 4, 1)
            If FDir > 8 Then
                FDir = FDir - 8
            ElseIf FDir < 1 Then
                FDir = FDir + 8
            End If
        ElseIf EFighters(FSpeDir, 3, 1) = 4 Then
            Randomize
            FSpe = Int(4 * Rnd + 1)
            If FSpe < 3 Then FSpe = 3
            Randomize
            FDir = (Int(3 * Rnd + 1) - 2) + EFighters(FSpeDir, 4, 1)
            If FDir > 8 Then
                FDir = FDir - 8
            ElseIf FDir < 1 Then
                FDir = FDir + 8
            End If
        End If
'If RAF are nearby, skew fighter movement in their direction
        If NumRAFRemain > 0 Then
            For B = 1 To 4
                If RAFStatus(B, 1, 1) <> 0 Then
                    RAFdirect = 0
                    ydif = EFighters(FSpeDir, 1, 1) - RAFStatus(B, 1, 1)
                    xdif = EFighters(FSpeDir, 2, 1) - RAFStatus(B, 2, 1)
                    If (ydif <= 4 And ydif >= 2 And xdif <= -2 And xdif >= -4) Or (ydif = 1 And xdif = -1) Then
                        RAFdirect = 1
                    ElseIf (ydif <= 4 And ydif >= 2 And xdif <= 4 And xdif >= 2) Or (ydif = 1 And xdif = 1) Then
                        RAFdirect = 3
                    ElseIf (ydif <= -2 And ydif >= -4 And xdif <= 4 And xdif >= 2) Or (ydif = -1 And xdif = 1) Then
                        RAFdirect = 5
                    ElseIf (ydif <= -2 And ydif >= -4 And xdif <= -2 And xdif >= -4) Or (ydif = -1 And xdif = -1) Then
                        RAFdirect = 7
                    ElseIf ydif <= 4 And ydif > 0 And xdif <= 1 And xdif >= -1 Then
                        RAFdirect = 2
                    ElseIf ydif <= 1 And ydif >= -1 And xdif <= -4 And xdif > 0 Then
                        RAFdirect = 4
                    ElseIf ydif < 0 And ydif >= -4 And xdif <= 1 And xdif >= -1 Then
                        RAFdirect = 6
                    ElseIf ydif <= 1 And ydif >= -1 And xdif < 0 And xdif >= -4 Then
                        RAFdirect = 8
                    End If
                    If EFighters(FSpeDir, 3, 1) = 1 And RAFdirect <> 0 Then
                        FDir = RAFdirect
                    ElseIf RAFdirect <> 0 Then
                        DirCheck = 0
                        If EFighters(FSpeDir, 4, 1) = 2 Then
                            DirCheck = 2
                        ElseIf EFighters(FSpeDir, 4, 1) >= 3 Then
                            DirCheck = 1
                        End If
                        UpB = EFighters(FSpeDir, 4, 1) + DirCheck
                        Adju = 8
                        If UpB > 8 Then Adju = -8
                        If (RAFdirect <= UpB And RAFdirect >= UpB - 2 * DirCheck) Or (RAFdirect <= UpB - Adju And RAFdirect >= UpB - Adju - 2 * DirCheck) Then
                            FDir = RAFdirect
                        End If
                    End If
                End If
            Next B
        End If
        
        EFighters(FSpeDir, 3, 1) = FSpe
        EFighters(FSpeDir, 4, 1) = FDir
        
        If EFighters(FSpeDir, 1, 1) <> 0 Then
            If FDir <= 3 Then
                EFFire(FSpeDir, 1, 1) = EFighters(FSpeDir, 1, 1)
                EFFire(FSpeDir, 2, 1) = EFighters(FSpeDir, 2, 1)
                EFFire(FSpeDir, 3, 1) = FDir
            ElseIf FDir = 4 Then
                For eraf = 1 To 4
                    If RAFStatus(eraf, 2, 1) <= EFighters(FSpeDir, 2, 1) And RAFStatus(eraf, 1, 1) >= EFighters(FSpeDir, 1, 1) - 3 And _
                    RAFStatus(eraf, 1, 1) <= EFighters(FSpeDir, 1, 1) + 3 And RAFStatus(eraf, 1, 1) <> 0 Then
                        EFFire(FSpeDir, 1, 1) = EFighters(FSpeDir, 1, 1)
                        EFFire(FSpeDir, 2, 1) = EFighters(FSpeDir, 2, 1)
                        EFFire(FSpeDir, 3, 1) = FDir
                        eraf = 4
                    End If
                Next eraf
            ElseIf FDir = 5 Then
                For eraf = 1 To 4
                    If RAFStatus(eraf, 1, 1) >= EFighters(FSpeDir, 1, 1) And RAFStatus(eraf, 2, 1) <= EFighters(FSpeDir, 2, 1) And RAFStatus(eraf, 1, 1) <> 0 Then
                        EFFire(FSpeDir, 1, 1) = EFighters(FSpeDir, 1, 1)
                        EFFire(FSpeDir, 2, 1) = EFighters(FSpeDir, 2, 1)
                        EFFire(FSpeDir, 3, 1) = FDir
                        eraf = 4
                    End If
                Next eraf
            ElseIf FDir = 6 Then
                For eraf = 1 To 4
                    If RAFStatus(eraf, 1, 1) >= EFighters(FSpeDir, 1, 1) And RAFStatus(eraf, 2, 1) >= EFighters(FSpeDir, 2, 1) - 3 And _
                    RAFStatus(eraf, 2, 1) <= EFighters(FSpeDir, 2, 1) + 3 And RAFStatus(eraf, 1, 1) <> 0 Then
                        EFFire(FSpeDir, 1, 1) = EFighters(FSpeDir, 1, 1)
                        EFFire(FSpeDir, 2, 1) = EFighters(FSpeDir, 2, 1)
                        EFFire(FSpeDir, 3, 1) = FDir
                        eraf = 4
                    End If
                Next eraf
            ElseIf FDir = 7 Then
                For eraf = 1 To 4
                    If RAFStatus(eraf, 1, 1) >= EFighters(FSpeDir, 1, 1) And RAFStatus(eraf, 2, 1) >= EFighters(FSpeDir, 2, 1) And RAFStatus(eraf, 1, 1) <> 0 Then
                        EFFire(FSpeDir, 1, 1) = EFighters(FSpeDir, 1, 1)
                        EFFire(FSpeDir, 2, 1) = EFighters(FSpeDir, 2, 1)
                        EFFire(FSpeDir, 3, 1) = FDir
                        eraf = 4
                    End If
                Next eraf
            ElseIf FDir = 8 Then
                For eraf = 1 To 4
                    If RAFStatus(eraf, 2, 1) >= EFighters(FSpeDir, 2, 1) And RAFStatus(eraf, 1, 1) >= EFighters(FSpeDir, 1, 1) - 3 And _
                    RAFStatus(eraf, 1, 1) <= EFighters(FSpeDir, 1, 1) + 3 And RAFStatus(eraf, 1, 1) <> 0 Then
                        EFFire(FSpeDir, 1, 1) = EFighters(FSpeDir, 1, 1)
                        EFFire(FSpeDir, 2, 1) = EFighters(FSpeDir, 2, 1)
                        EFFire(FSpeDir, 3, 1) = FDir
                        eraf = 4
                    End If
                Next eraf
            End If
            FSpex = 0
            FSpey = 0
        
            'FDir determines speed in x and y directions of active enemy fighters
            If FDir = 1 Then
                FSpex = FSpe
                FSpey = (0 - FSpe)
            ElseIf FDir = 2 Then
                FSpey = (0 - FSpe)
            ElseIf FDir = 3 Then
                FSpex = (0 - FSpe)
                FSpey = (0 - FSpe)
            ElseIf FDir = 4 Then
                FSpex = (0 - FSpe)
            ElseIf FDir = 5 Then
                FSpex = (0 - FSpe)
                FSpey = FSpe
            ElseIf FDir = 6 Then
                FSpey = FSpe
            ElseIf FDir = 7 Then
                FSpex = FSpe
                FSpey = FSpe
            ElseIf FDir = 8 Then
                FSpex = FSpe
            End If
            
            'Fill movement matrix for enemy fighters
            'Note: 0.4 is accelerative factor
            '      to timeshift ahead movement of fighters with speed = 3
            If FSpey < 0 Then
                FSpey1 = (-1) * Int(Abs(FSpey) / 4 + 0.4)
                FSpey2 = (-1) * Int(Abs(FSpey) / 2 + 0.4)
                FSpey3 = (-1) * Int(3 * Abs(FSpey) / 4 + 0.4)
            Else
                FSpey1 = Int(FSpey / 4 + 0.4)
                FSpey2 = Int(FSpey / 2 + 0.4)
                FSpey3 = Int(3 * FSpey / 4 + 0.4)
            End If
            If FSpex < 0 Then
                FSpex1 = (-1) * Int(Abs(FSpex) / 4 + 0.4)
                FSpex2 = (-1) * Int(Abs(FSpex) / 2 + 0.4)
                FSpex3 = (-1) * Int(3 * Abs(FSpex) / 4 + 0.4)
            Else
                FSpex1 = Int(FSpex / 4 + 0.4)
                FSpex2 = Int(FSpex / 2 + 0.4)
                FSpex3 = Int(3 * FSpex / 4 + 0.4)
            End If
            
            If EFFire(FSpeDir, 3, 1) <> 0 Then
                FireValid = True
                Randomize
                FireStart = Int(2 * Rnd + 2)
                If FireStart = 3 Then
                        EFFire(FSpeDir, 1, 2) = EFFire(FSpeDir, 1, 1)
                        EFFire(FSpeDir, 2, 2) = EFFire(FSpeDir, 2, 1)
                        EFFire(FSpeDir, 3, 2) = EFFire(FSpeDir, 3, 1)
                End If
                For incr = FireStart To TIME_END
                    If FireValid = True Then
                        If EFFire(FSpeDir, 3, 1) <= 3 Then
                            EFFire(FSpeDir, 1, incr) = EFFire(FSpeDir, 1, incr - 1) - 1
                        ElseIf EFFire(FSpeDir, 3, 1) >= 5 And EFFire(FSpeDir, 3, 1) <= 7 Then
                            EFFire(FSpeDir, 1, incr) = EFFire(FSpeDir, 1, incr - 1) + 1
                        Else
                            EFFire(FSpeDir, 1, incr) = EFFire(FSpeDir, 1, incr - 1)
                        End If
                        If EFFire(FSpeDir, 3, 1) >= 3 And EFFire(FSpeDir, 3, 1) <= 5 Then
                            EFFire(FSpeDir, 2, incr) = EFFire(FSpeDir, 2, incr - 1) - 1
                        ElseIf EFFire(FSpeDir, 3, 1) >= 7 Or EFFire(FSpeDir, 3, 1) = 1 Then
                            EFFire(FSpeDir, 2, incr) = EFFire(FSpeDir, 2, incr - 1) + 1
                        Else
                            EFFire(FSpeDir, 2, incr) = EFFire(FSpeDir, 2, incr - 1)
                        End If
                        EFFire(FSpeDir, 3, incr) = EFFire(FSpeDir, 3, incr - 1)
                        If EFFire(FSpeDir, 1, incr) <= 0 Or EFFire(FSpeDir, 1, incr) >= 22 Or _
                        EFFire(FSpeDir, 2, incr) <= 0 Or EFFire(FSpeDir, 2, incr) >= 22 Then
                            FireValid = False
                            EFFire(FSpeDir, 1, incr) = 0
                            EFFire(FSpeDir, 2, incr) = 0
                            EFFire(FSpeDir, 3, incr) = 0
                        End If
                    End If
                Next incr
            End If
            Randomize
            move1 = Int(4 * Rnd + 4)
            move2 = Int(4 * Rnd + 9)
            move3 = Int(4 * Rnd + 14)
            move4 = Int(4 * Rnd + 19)
            For incr = 2 To TIME_END
                For i = 1 To 4
                    EFighters(FSpeDir, i, incr) = EFighters(FSpeDir, i, incr - 1)
                Next i
                If incr <> move1 And incr <> move2 And incr <> move3 And incr <> move4 Then
                    'Procedure Bypass
                ElseIf incr = move1 Then
                    EFighters(FSpeDir, 1, incr) = EFighters(FSpeDir, 1, 1) + FSpey1
                    EFighters(FSpeDir, 2, incr) = EFighters(FSpeDir, 2, 1) + FSpex1
                ElseIf incr = move2 Then
                    EFighters(FSpeDir, 1, incr) = EFighters(FSpeDir, 1, 1) + FSpey2
                    EFighters(FSpeDir, 2, incr) = EFighters(FSpeDir, 2, 1) + FSpex2
                ElseIf incr = move3 Then
                    EFighters(FSpeDir, 1, incr) = EFighters(FSpeDir, 1, 1) + FSpey3
                    EFighters(FSpeDir, 2, incr) = EFighters(FSpeDir, 2, 1) + FSpex3
                ElseIf incr = move4 Then
                    EFighters(FSpeDir, 1, incr) = EFighters(FSpeDir, 1, 1) + FSpey
                    EFighters(FSpeDir, 2, incr) = EFighters(FSpeDir, 2, 1) + FSpex
                End If
                If EFighters(FSpeDir, 1, incr) > BOARD_ROWS Then
                    EFighters(FSpeDir, 1, incr) = BOARD_ROWS
                    EFighters(FSpeDir, 3, incr) = 3
                    EFighters(FSpeDir, 4, incr) = 2
                ElseIf EFighters(FSpeDir, 1, incr) < 4 Then
                    EFighters(FSpeDir, 1, incr) = 4
                    EFighters(FSpeDir, 3, incr) = 4
                    EFighters(FSpeDir, 4, incr) = 6
                End If
                If (EFighters(FSpeDir, 2, incr) > BOARD_COLS) Or (EFighters(FSpeDir, 2, incr) = 0) Then
                    EFighters(FSpeDir, 2, incr) = Abs(EFighters(FSpeDir, 2, incr) - BOARD_COLS)
                ElseIf EFighters(FSpeDir, 2, incr) < 0 Then
                    EFighters(FSpeDir, 2, incr) = EFighters(FSpeDir, 2, incr) + BOARD_COLS
                End If
            Next incr
        End If
        FSpeDir = FSpeDir + 1
    Loop
    
    BSpeDir = 1
    Do While BSpeDir <= NumEB
        BSpe = 0
        If EBombers(BSpeDir, 1, 1) <> 0 Then
            Randomize
            BSpe = Int(3 * Rnd + 1)
            If EBFirePrime(BSpeDir) = 4 Then
                    EBFire(BSpeDir, 1, 1) = EBombers(BSpeDir, 1, 1)
                    EBFire(BSpeDir, 2, 1) = EBombers(BSpeDir, 2, 1)
                    EBFirePrime(BSpeDir) = 0
            Else
                Randomize
                FireAccel = Int(10 * Rnd + 1)              'Adds chance to increase frequency of bomb drop
                If FireAccel <= 4 Then
                    EBFirePrime(BSpeDir) = EBFirePrime(BSpeDir) + 1
                ElseIf FireAccel <= 7 Then
                    EBFirePrime(BSpeDir) = EBFirePrime(BSpeDir) + 2
                ElseIf FireAccel <= 9 Then
                    EBFirePrime(BSpeDir) = EBFirePrime(BSpeDir) + 3
                ElseIf FireAccel = 10 Then
                    EBFirePrime(BSpeDir) = EBFirePrime(BSpeDir) + 4
                End If
                If EBFirePrime(BSpeDir) > 4 Then
                    EBFirePrime(BSpeDir) = 4
                End If
            End If
            
            If EBombers(BSpeDir, 4, 1) = 1 Then
                BSpe1 = (-1) * Int(Abs(BSpe) / 4 + 0.4)
                BSpe2 = (-1) * Int(Abs(BSpe) / 2 + 0.4)
                BSpe3 = (-1) * Int(3 * Abs(BSpe) / 4 + 0.4)
                BSpe = (-1) * BSpe
            Else
                BSpe1 = Int(BSpe / 4 + 0.4)
                BSpe2 = Int(BSpe / 2 + 0.4)
                BSpe3 = Int(3 * BSpe / 4 + 0.4)
            End If
            
            If EBFire(BSpeDir, 1, 1) <> 0 Then
                FireValid = True
                Randomize
                FireStart = Int(2 * Rnd + 2)
                If FireStart = 3 Then
                        EBFire(BSpeDir, 1, 2) = EBFire(BSpeDir, 1, 1)
                        EBFire(BSpeDir, 2, 2) = EBFire(BSpeDir, 2, 1)
                End If
                For incr = FireStart To TIME_END
                    If FireValid = True Then
                        EBFire(BSpeDir, 1, incr) = EBFire(BSpeDir, 1, incr - 1) - 1
                        EBFire(BSpeDir, 2, incr) = EBFire(BSpeDir, 2, incr - 1)
                        If EBFire(BSpeDir, 1, incr) = 0 Then
                            FireValid = False
                            EBFire(BSpeDir, 1, incr) = 0
                            EBFire(BSpeDir, 2, incr) = 0
                        End If
                    End If
                Next incr
            End If
            Randomize
            move1 = Int(4 * Rnd + 4)
            move2 = Int(4 * Rnd + 9)
            move3 = Int(4 * Rnd + 14)
            move4 = Int(4 * Rnd + 19)
            For incr = 2 To TIME_END
                For i = 1 To 5
                    EBombers(BSpeDir, i, incr) = EBombers(BSpeDir, i, incr - 1)
                Next i
                If incr <> move1 And incr <> move2 And incr <> move3 And incr <> move4 Then
                    'Procedure Bypass
                ElseIf incr = move1 Then
                    EBombers(BSpeDir, 2, incr) = EBombers(BSpeDir, 2, 1) + BSpe1
                ElseIf incr = move2 Then
                    EBombers(BSpeDir, 2, incr) = EBombers(BSpeDir, 2, 1) + BSpe2
                ElseIf incr = move3 Then
                    EBombers(BSpeDir, 2, incr) = EBombers(BSpeDir, 2, 1) + BSpe3
                ElseIf incr = move4 Then
                    EBombers(BSpeDir, 2, incr) = EBombers(BSpeDir, 2, 1) + BSpe
                End If
                If (EBombers(BSpeDir, 2, incr) > BOARD_COLS) Or (EBombers(BSpeDir, 2, incr) = 0) Then
                    EBombers(BSpeDir, 2, incr) = Abs(EBombers(BSpeDir, 2, incr) - BOARD_COLS)
                ElseIf EBombers(BSpeDir, 2, incr) < 0 Then
                    EBombers(BSpeDir, 2, incr) = EBombers(BSpeDir, 2, incr) + BOARD_COLS
                End If
            Next incr
        End If
        BSpeDir = BSpeDir + 1
    Loop
    
    TurretPos = 1
    
    NextTurret
    
End Sub

Sub NextTurret()

    Application.ScreenUpdating = False
    Worksheets("London Siege").FireTurret.Visible = False
    Worksheets("London Siege").SpinTurretAim.Visible = False
    Worksheets("London Siege").SkipTurret.Visible = False
    
    For it = 1 To 23
        For jt = 1 To 23
            Cells(it, 25 + jt) = ""
        Next jt
    Next it

    TurretFind = False
    
    Do While TurretFind = False And TurretPos <= BOARD_COLS                                                                                'Fill TurretFire with user input
        If Turrets(TurretPos, 1).Icon <> "" And Turrets(TurretPos, 1).Health < 4 Then
            With Range(Cells(21, 1 + TurretPos).Address())
                .Borders(xlEdgeTop).ThemeColor = 1
                .Borders(xlEdgeTop).TintAndShade = 0
                .Borders(xlEdgeRight).ThemeColor = 1
                .Borders(xlEdgeRight).TintAndShade = 0
                .Borders(xlEdgeBottom).ThemeColor = 1
                .Borders(xlEdgeBottom).TintAndShade = 0
                .Borders(xlEdgeLeft).ThemeColor = 1
                .Borders(xlEdgeLeft).TintAndShade = 0
            End With
            TurretAimNom = Dir90
            TurretAimAnimate
            Worksheets("London Siege").FireTurret.Visible = True
            Worksheets("London Siege").SpinTurretAim.Visible = True
            Worksheets("London Siege").SkipTurret.Visible = True
            Cells(10, 35) = Turrets(TurretPos, 1).Icon
            Range(Cells(10, 35).Address()).Font.Size = 60
            Cells(17, 37) = "Column " & TurretPos
            Cells(18, 37) = "Status: " & (100 - Turrets(TurretPos, 1).Health * 25) & "%"
            TurretFind = True
        End If
        If TurretFind = False Then TurretPos = TurretPos + 1
    Loop
    
    If TurretPos = 22 Then
        RAFPos = 1
        NextTurnRAF
    End If
    
    Application.ScreenUpdating = True

End Sub

Sub FireTurret()

    Turrets(TurretPos, 1).Fire.Row = TURRET_ROW                          'Starting row
    Turrets(TurretPos, 1).Fire.Column = TurretPos             'Starting column
    Turrets(TurretPos, 1).Fire.Direction = TurretAimNom      'Turret Aim
    Turrets(TurretPos, 1).Fire.Type = 1                          'Turret Fire Type, will change
    
    Application.ScreenUpdating = False
    With Range(Cells(21, 1 + TurretPos).Address())
        .Borders(xlEdgeTop).ColorIndex = 0
        .Borders(xlEdgeTop).TintAndShade = 0
        .Borders(xlEdgeRight).ColorIndex = 0
        .Borders(xlEdgeRight).TintAndShade = 0
        .Borders(xlEdgeBottom).ColorIndex = 0
        .Borders(xlEdgeBottom).TintAndShade = 0
        .Borders(xlEdgeLeft).ColorIndex = 0
        .Borders(xlEdgeLeft).TintAndShade = 0
    End With
    
    ShotCount = ShotCount + 1
    TurretPos = TurretPos + 1
    NextTurret

End Sub

Sub SkipTurret()

    Application.ScreenUpdating = False
    With Range(Cells(21, 1 + TurretPos).Address())
        .Borders(xlEdgeTop).ColorIndex = 0
        .Borders(xlEdgeTop).TintAndShade = 0
        .Borders(xlEdgeRight).ColorIndex = 0
        .Borders(xlEdgeRight).TintAndShade = 0
        .Borders(xlEdgeBottom).ColorIndex = 0
        .Borders(xlEdgeBottom).TintAndShade = 0
        .Borders(xlEdgeLeft).ColorIndex = 0
        .Borders(xlEdgeLeft).TintAndShade = 0
    End With

    TurretPos = TurretPos + 1
    NextTurret
    
End Sub

Sub SpinTurretAimLeft()
    If TurretAimNom < Dir153 Then
        TurretAimNom = TurretAimNom + 1
        TurretAimAnimate
    End If
End Sub

Sub SpinTurretAimRight()
    If TurretAimNom > Dir27 Then
        TurretAimNom = TurretAimNom - 1
        TurretAimAnimate
    End If
End Sub

Sub TurretAimAnimate()
    
    Application.ScreenUpdating = False
    
    If TurretAimNom = Dir27 Then
        Cells(16, 37) = "27°"
    ElseIf TurretAimNom = Dir45 Then
        Cells(16, 37) = "45°"
    ElseIf TurretAimNom = Dir63 Then
        Cells(16, 37) = "63°"
    ElseIf TurretAimNom = Dir72 Then
        Cells(16, 37) = "72°"
    ElseIf TurretAimNom = Dir76 Then
        Cells(16, 37) = "76°"
    ElseIf TurretAimNom = Dir90 Then
        Cells(16, 37) = "90°"
    ElseIf TurretAimNom = Dir104 Then
        Cells(16, 37) = "104°"
    ElseIf TurretAimNom = Dir108 Then
        Cells(16, 37) = "108°"
    ElseIf TurretAimNom = Dir117 Then
        Cells(16, 37) = "117°"
    ElseIf TurretAimNom = Dir135 Then
        Cells(16, 37) = "135°"
    ElseIf TurretAimNom = Dir153 Then
        Cells(16, 37) = "153°"
    End If
    
    For it = 1 To 9
        For jt = 1 To 23
            Cells(it, 25 + jt) = ""
        Next jt
    Next it
    
    If TurretAimNom = Dir27 Then
        For it = 1 To 11
            If it = 1 Then
                Cells(4, 49 - it) = "o"
            ElseIf it <= 3 Then
                Cells(5, 49 - it) = "o"
            ElseIf it <= 5 Then
                Cells(6, 49 - it) = "o"
            ElseIf it <= 7 Then
                Cells(7, 49 - it) = "o"
            ElseIf it <= 9 Then
                Cells(8, 49 - it) = "o"
            Else
                Cells(9, 49 - it) = "o"
            End If
        Next it
    ElseIf TurretAimNom = Dir45 Then
        For it = 1 To 9
            Cells(it, 47 - it) = "o"
        Next it
    ElseIf TurretAimNom = Dir63 Then
        For it = 1 To 9
            If it = 1 Then
                Cells(it, 42) = "o"
            ElseIf it <= 3 Then
                Cells(it, 41) = "o"
            ElseIf it <= 5 Then
                Cells(it, 40) = "o"
            ElseIf it <= 7 Then
                Cells(it, 39) = "o"
            Else
                Cells(it, 38) = "o"
            End If
        Next it
    ElseIf TurretAimNom = Dir72 Then
        For it = 1 To 9
            If it <= 3 Then
                Cells(it, 40) = "o"
            ElseIf it <= 6 Then
                Cells(it, 39) = "o"
            Else
                Cells(it, 38) = "o"
            End If
        Next it
    ElseIf TurretAimNom = Dir76 Then
        For it = 1 To 9
            If it <= 3 Then
                Cells(it, 39) = "o"
            ElseIf it <= 7 Then
                Cells(it, 38) = "o"
            Else
                Cells(it, 37) = "o"
            End If
        Next it
    ElseIf TurretAimNom = Dir90 Then
        For it = 1 To 9
            Cells(it, 37) = "o"
        Next it
    ElseIf TurretAimNom = Dir104 Then
        For it = 1 To 9
            If it <= 3 Then
                Cells(it, 35) = "o"
            ElseIf it <= 7 Then
                Cells(it, 36) = "o"
            Else
                Cells(it, 37) = "o"
            End If
        Next it
    ElseIf TurretAimNom = Dir108 Then
        For it = 1 To 9
            If it <= 3 Then
                Cells(it, 34) = "o"
            ElseIf it <= 6 Then
                Cells(it, 35) = "o"
            Else
                Cells(it, 36) = "o"
            End If
        Next it
    ElseIf TurretAimNom = Dir117 Then
        For it = 1 To 9
            If it = 1 Then
                Cells(it, 32) = "o"
            ElseIf it <= 3 Then
                Cells(it, 33) = "o"
            ElseIf it <= 5 Then
                Cells(it, 34) = "o"
            ElseIf it <= 7 Then
                Cells(it, 35) = "o"
            Else
                Cells(it, 36) = "o"
            End If
        Next it
    ElseIf TurretAimNom = Dir135 Then
        For it = 1 To 9
            Cells(it, 27 + it) = "o"
        Next it
    ElseIf TurretAimNom = Dir153 Then
        For it = 1 To 11
            If it = 1 Then
                Cells(4, 25 + it) = "o"
            ElseIf it <= 3 Then
                Cells(5, 25 + it) = "o"
            ElseIf it <= 5 Then
                Cells(6, 25 + it) = "o"
            ElseIf it <= 7 Then
                Cells(7, 25 + it) = "o"
            ElseIf it <= 9 Then
                Cells(8, 25 + it) = "o"
            Else
                Cells(9, 25 + it) = "o"
            End If
        Next it
    End If
    
    Application.ScreenUpdating = True

End Sub

Sub NextTurnRAF()
    If StructureStatus(ComC, 1) >= 10 Then    'Unclear what this code block does, StructureStatus(ComC, 1) should not exceed 4
        RAFSpeDir = 1
        For RAFSpeDir = 1 To 4
            RAFSpe = 0
            RAFDir = 0
            If RAFStatus(RAFSpeDir, 6, 1) = 1 Then
                Randomize
                LaunchRnd = Int(7 * Rnd + 1)
                If LaunchRnd < 6 Then
                    RAFStatus(RAFSpeDir, 3, 1) = 2
                    RAFSpe = 2
                    RAFStatus(RAFSpeDir, 4, 1) = 6
                    RAFDir = 6
                End If
            ElseIf RAFStatus(RAFSpeDir, 6, 1) = 2 Then
                Randomize
                RAFSpe = Int(3 * Rnd + 2)
                Randomize
                If RAFStatus(RAFSpeDir, 3, 1) <= 2 Then
                    RAFDir = (Int(5 * Rnd + 1) - 3) + RAFStatus(RAFSpeDir, 4, 1)
                ElseIf RAFStatus(RAFSpeDir, 3, 1) <= 4 Then
                    RAFDir = (Int(3 * Rnd + 1) - 2) + RAFStatus(RAFSpeDir, 4, 1)
                End If
                If RAFDir > 8 Then
                    RAFDir = RAFDir - 8
                ElseIf RAFDir < 1 Then
                    RAFDir = RAFDir + 8
                End If
                RAFStatus(RAFSpeDir, 3, 1) = RAFSpe
                RAFStatus(RAFSpeDir, 4, 1) = RAFDir
                If RAFDir = 1 Then
                    For ef = 1 To 20
                        If (EBombers(ef, 1, 1) <= RAFStatus(RAFSpeDir, 1, 1) And EBombers(ef, 2, 1) >= RAFStatus(RAFSpeDir, 2, 1) And EBombers(ef, 1, 1) <> 0) Or _
                        (EFighters(ef, 1, 1) <= RAFStatus(RAFSpeDir, 1, 1) And EFighters(ef, 2, 1) >= RAFStatus(RAFSpeDir, 2, 1) And EFighters(ef, 1, 1) <> 0) Then
                            RAFFire(RAFSpeDir, 1, 1) = RAFStatus(RAFSpeDir, 1, 1)
                            RAFFire(RAFSpeDir, 2, 1) = RAFStatus(RAFSpeDir, 2, 1)
                            RAFFire(RAFSpeDir, 3, 1) = 1
                            RAFFire(RAFSpeDir, 4, 1) = RAFDir
                            ef = 20
                        End If
                    Next ef
                ElseIf RAFDir = 2 Then
                    For ef = 1 To 20
                        If (EBombers(ef, 1, 1) <= RAFStatus(RAFSpeDir, 1, 1) And EBombers(ef, 2, 1) >= RAFStatus(RAFSpeDir, 2, 1) - 3 And _
                        EBombers(ef, 2, 1) <= RAFStatus(RAFSpeDir, 2, 1) + 3 And EBombers(ef, 1, 1) <> 0) Or _
                        (EFighters(ef, 1, 1) <= RAFStatus(RAFSpeDir, 1, 1) And EFighters(ef, 2, 1) >= RAFStatus(RAFSpeDir, 2, 1) - 3 And _
                        EFighters(ef, 2, 1) <= RAFStatus(RAFSpeDir, 2, 1) + 3 And EFighters(ef, 1, 1) <> 0) Then
                            RAFFire(RAFSpeDir, 1, 1) = RAFStatus(RAFSpeDir, 1, 1)
                            RAFFire(RAFSpeDir, 2, 1) = RAFStatus(RAFSpeDir, 2, 1)
                            RAFFire(RAFSpeDir, 3, 1) = 1
                            RAFFire(RAFSpeDir, 4, 1) = RAFDir
                            ef = 20
                        End If
                    Next ef
                ElseIf RAFDir = 3 Then
                    For ef = 1 To 20
                        If (EBombers(ef, 1, 1) <= RAFStatus(RAFSpeDir, 1, 1) And EBombers(ef, 2, 1) <= RAFStatus(RAFSpeDir, 2, 1) And EBombers(ef, 1, 1) <> 0) Or _
                        (EFighters(ef, 1, 1) <= RAFStatus(RAFSpeDir, 1, 1) And EFighters(ef, 2, 1) <= RAFStatus(RAFSpeDir, 2, 1) And EFighters(ef, 1, 1) <> 0) Then
                            RAFFire(RAFSpeDir, 1, 1) = RAFStatus(RAFSpeDir, 1, 1)
                            RAFFire(RAFSpeDir, 2, 1) = RAFStatus(RAFSpeDir, 2, 1)
                            RAFFire(RAFSpeDir, 3, 1) = 1
                            RAFFire(RAFSpeDir, 4, 1) = RAFDir
                            ef = 20
                        End If
                    Next ef
                ElseIf RAFDir = 4 Then
                    For ef = 1 To 20
                        If (EBombers(ef, 2, 1) <= RAFStatus(RAFSpeDir, 2, 1) And EBombers(ef, 1, 1) >= RAFStatus(RAFSpeDir, 1, 1) - 3 And _
                        EBombers(ef, 1, 1) <= RAFStatus(RAFSpeDir, 1, 1) + 3 And EBombers(ef, 1, 1) <> 0) Or _
                        (EFighters(ef, 2, 1) <= RAFStatus(RAFSpeDir, 2, 1) And EFighters(ef, 1, 1) >= RAFStatus(RAFSpeDir, 1, 1) - 3 And _
                        EFighters(ef, 1, 1) <= RAFStatus(RAFSpeDir, 1, 1) + 3 And EFighters(ef, 1, 1) <> 0) Then
                            RAFFire(RAFSpeDir, 1, 1) = RAFStatus(RAFSpeDir, 1, 1)
                            RAFFire(RAFSpeDir, 2, 1) = RAFStatus(RAFSpeDir, 2, 1)
                            RAFFire(RAFSpeDir, 3, 1) = 1
                            RAFFire(RAFSpeDir, 4, 1) = RAFDir
                            ef = 20
                        End If
                    Next ef
                ElseIf RAFDir = 5 Then
                    For ef = 1 To 20
                        If (EBombers(ef, 1, 1) >= RAFStatus(RAFSpeDir, 1, 1) And EBombers(ef, 2, 1) <= RAFStatus(RAFSpeDir, 2, 1) And EBombers(ef, 1, 1) <> 0) Or _
                        (EFighters(ef, 1, 1) >= RAFStatus(RAFSpeDir, 1, 1) And EFighters(ef, 2, 1) <= RAFStatus(RAFSpeDir, 2, 1) And EFighters(ef, 1, 1) <> 0) Then
                            RAFFire(RAFSpeDir, 1, 1) = RAFStatus(RAFSpeDir, 1, 1)
                            RAFFire(RAFSpeDir, 2, 1) = RAFStatus(RAFSpeDir, 2, 1)
                            RAFFire(RAFSpeDir, 3, 1) = 1
                            RAFFire(RAFSpeDir, 4, 1) = RAFDir
                            ef = 20
                        End If
                    Next ef
                ElseIf RAFDir = 6 Then
                    For ef = 1 To 20
                        If (EBombers(ef, 1, 1) >= RAFStatus(RAFSpeDir, 1, 1) And EBombers(ef, 2, 1) >= RAFStatus(RAFSpeDir, 2, 1) - 3 And _
                        EBombers(ef, 2, 1) <= RAFStatus(RAFSpeDir, 2, 1) + 3 And EBombers(ef, 1, 1) <> 0) Or _
                        (EFighters(ef, 1, 1) >= RAFStatus(RAFSpeDir, 1, 1) And EFighters(ef, 2, 1) >= RAFStatus(RAFSpeDir, 2, 1) - 3 And _
                        EFighters(ef, 2, 1) <= RAFStatus(RAFSpeDir, 2, 1) + 3 And EFighters(ef, 1, 1) <> 0) Then
                            RAFFire(RAFSpeDir, 1, 1) = RAFStatus(RAFSpeDir, 1, 1)
                            RAFFire(RAFSpeDir, 2, 1) = RAFStatus(RAFSpeDir, 2, 1)
                            RAFFire(RAFSpeDir, 3, 1) = 1
                            RAFFire(RAFSpeDir, 4, 1) = RAFDir
                            ef = 20
                        End If
                    Next ef
                ElseIf RAFDir = 7 Then
                    For ef = 1 To 20
                        If (EBombers(ef, 1, 1) >= RAFStatus(RAFSpeDir, 1, 1) And EBombers(ef, 2, 1) >= RAFStatus(RAFSpeDir, 2, 1) And EBombers(ef, 1, 1) <> 0) Or _
                        (EFighters(ef, 1, 1) >= RAFStatus(RAFSpeDir, 1, 1) And EFighters(ef, 2, 1) >= RAFStatus(RAFSpeDir, 2, 1) And EFighters(ef, 1, 1) <> 0) Then
                            RAFFire(RAFSpeDir, 1, 1) = RAFStatus(RAFSpeDir, 1, 1)
                            RAFFire(RAFSpeDir, 2, 1) = RAFStatus(RAFSpeDir, 2, 1)
                            RAFFire(RAFSpeDir, 3, 1) = 1
                            RAFFire(RAFSpeDir, 4, 1) = RAFDir
                            ef = 20
                        End If
                    Next ef
                ElseIf RAFDir = 8 Then
                    For ef = 1 To 20
                        If (EBombers(ef, 2, 1) >= RAFStatus(RAFSpeDir, 2, 1) And EBombers(ef, 1, 1) >= RAFStatus(RAFSpeDir, 1, 1) - 3 And _
                        EBombers(ef, 1, 1) <= RAFStatus(RAFSpeDir, 1, 1) + 3 And EBombers(ef, 1, 1) <> 0) Or _
                        (EFighters(ef, 2, 1) >= RAFStatus(RAFSpeDir, 2, 1) And EFighters(ef, 1, 1) >= RAFStatus(RAFSpeDir, 1, 1) - 3 And _
                        EFighters(ef, 1, 1) <= RAFStatus(RAFSpeDir, 1, 1) + 3 And EFighters(ef, 1, 1) <> 0) Then
                            RAFFire(RAFSpeDir, 1, 1) = RAFStatus(RAFSpeDir, 1, 1)
                            RAFFire(RAFSpeDir, 2, 1) = RAFStatus(RAFSpeDir, 2, 1)
                            RAFFire(RAFSpeDir, 3, 1) = 1
                            RAFFire(RAFSpeDir, 4, 1) = RAFDir
                            ef = 20
                        End If
                    Next ef
                End If
            End If
            If RAFStatus(RAFSpeDir, 1, 1) <> 0 Then
                RAFSpex = 0
                RAFSpey = 0
        
                'RAFDir determines speed in x and y directions of active RAF
                If RAFDir = 1 Then
                    RAFSpex = RAFSpe
                    RAFSpey = (0 - RAFSpe)
                ElseIf RAFDir = 2 Then
                    RAFSpey = (0 - RAFSpe)
                ElseIf RAFDir = 3 Then
                    RAFSpex = (0 - RAFSpe)
                    RAFSpey = (0 - RAFSpe)
                ElseIf RAFDir = 4 Then
                    RAFSpex = (0 - RAFSpe)
                ElseIf RAFDir = 5 Then
                    RAFSpex = (0 - RAFSpe)
                    RAFSpey = RAFSpe
                ElseIf RAFDir = 6 Then
                    RAFSpey = RAFSpe
                ElseIf RAFDir = 7 Then
                    RAFSpex = RAFSpe
                    RAFSpey = RAFSpe
                ElseIf RAFDir = 8 Then
                    RAFSpex = RAFSpe
                End If
            
                'Fill movement matrix for enemy fighters
                'Note: 0.4 is accelerative factor
                '      to timeshift ahead movement of fighters with speed = 3
                If RAFSpey < 0 Then
                    RAFSpey1 = (-1) * Int(Abs(RAFSpey) / 4 + 0.4)
                    RAFSpey2 = (-1) * Int(Abs(RAFSpey) / 2 + 0.4)
                    RAFSpey3 = (-1) * Int(3 * Abs(RAFSpey) / 4 + 0.4)
                Else
                    RAFSpey1 = Int(RAFSpey / 4 + 0.4)
                    RAFSpey2 = Int(RAFSpey / 2 + 0.4)
                    RAFSpey3 = Int(3 * RAFSpey / 4 + 0.4)
                End If
                If RAFSpex < 0 Then
                    RAFSpex1 = (-1) * Int(Abs(RAFSpex) / 4 + 0.4)
                    RAFSpex2 = (-1) * Int(Abs(RAFSpex) / 2 + 0.4)
                    RAFSpex3 = (-1) * Int(3 * Abs(RAFSpex) / 4 + 0.4)
                Else
                    RAFSpex1 = Int(RAFSpex / 4 + 0.4)
                    RAFSpex2 = Int(RAFSpex / 2 + 0.4)
                    RAFSpex3 = Int(3 * RAFSpex / 4 + 0.4)
                End If
             
                If RAFFire(RAFSpeDir, 3, 1) <> 0 Then
                    FireValid = True
                    ShotCount = ShotCount + 1
                    Randomize
                    FireStart = Int(2 * Rnd + 2)
                    If FireStart = 3 Then
                        RAFFire(RAFSpeDir, 1, 2) = RAFFire(RAFSpeDir, 1, 1)
                        RAFFire(RAFSpeDir, 2, 2) = RAFFire(RAFSpeDir, 2, 1)
                        RAFFire(RAFSpeDir, 3, 2) = RAFFire(RAFSpeDir, 3, 1)
                        RAFFire(RAFSpeDir, 4, 2) = RAFFire(RAFSpeDir, 4, 1)
                    End If
                    For incr = FireStart To TIME_END
                        If FireValid = True Then
                            RocketFire = RAFFire(RAFSpeDir, 3, 1) - 1      'If rockets are being fired, RocketFire will double traveling speed of shot
                            If RAFFire(RAFSpeDir, 4, 1) <= 3 Then
                                RAFFire(RAFSpeDir, 1, incr) = RAFFire(RAFSpeDir, 1, incr - 1) - 1 - RocketFire
                            ElseIf RAFFire(RAFSpeDir, 4, 1) >= 5 And RAFFire(RAFSpeDir, 4, 1) <= 7 Then
                                RAFFire(RAFSpeDir, 1, incr) = RAFFire(RAFSpeDir, 1, incr - 1) + 1 + RocketFire
                            Else
                                RAFFire(RAFSpeDir, 1, incr) = RAFFire(RAFSpeDir, 1, incr - 1)
                            End If
                            If RAFFire(RAFSpeDir, 4, 1) >= 3 And RAFFire(RAFSpeDir, 4, 1) <= 5 Then
                                RAFFire(RAFSpeDir, 2, incr) = RAFFire(RAFSpeDir, 2, incr - 1) - 1 - RocketFire
                            ElseIf RAFFire(RAFSpeDir, 4, 1) >= 7 Or RAFFire(RAFSpeDir, 4, 1) = 1 Then
                                RAFFire(RAFSpeDir, 2, incr) = RAFFire(RAFSpeDir, 2, incr - 1) + 1 + RocketFire
                            Else
                                RAFFire(RAFSpeDir, 2, incr) = RAFFire(RAFSpeDir, 2, incr - 1)
                            End If
                            RAFFire(RAFSpeDir, 3, incr) = RAFFire(RAFSpeDir, 3, incr - 1)
                            RAFFire(RAFSpeDir, 4, incr) = RAFFire(RAFSpeDir, 4, incr - 1)
                            If RAFFire(RAFSpeDir, 1, incr) <= 0 Or RAFFire(RAFSpeDir, 1, incr) >= 22 Or _
                            RAFFire(RAFSpeDir, 2, incr) <= 0 Or RAFFire(RAFSpeDir, 2, incr) >= 22 Then
                                FireValid = False
                                RAFFire(RAFSpeDir, 1, incr) = 0
                                RAFFire(RAFSpeDir, 2, incr) = 0
                                RAFFire(RAFSpeDir, 3, incr) = 0
                                RAFFire(RAFSpeDir, 4, incr) = 0
                            End If
                        End If
                    Next incr
                End If
                Randomize
                move1 = Int(4 * Rnd + 4)
                move2 = Int(4 * Rnd + 9)
                move3 = Int(4 * Rnd + 14)
                move4 = Int(4 * Rnd + 19)
                For incr = 2 To TIME_END
                    For i = 1 To 7
                        RAFStatus(RAFSpeDir, i, incr) = RAFStatus(RAFSpeDir, i, incr - 1)
                    Next i
                    If incr <> move1 And incr <> move2 And incr <> move3 And incr <> move4 Then
                        'Procedure Bypass
                    ElseIf incr = move1 Then
                        RAFStatus(RAFSpeDir, 1, incr) = RAFStatus(RAFSpeDir, 1, 1) + RAFSpey1
                        RAFStatus(RAFSpeDir, 2, incr) = RAFStatus(RAFSpeDir, 2, 1) + RAFSpex1
                    ElseIf incr = move2 Then
                        RAFStatus(RAFSpeDir, 1, incr) = RAFStatus(RAFSpeDir, 1, 1) + RAFSpey2
                        RAFStatus(RAFSpeDir, 2, incr) = RAFStatus(RAFSpeDir, 2, 1) + RAFSpex2
                    ElseIf incr = move3 Then
                        RAFStatus(RAFSpeDir, 1, incr) = RAFStatus(RAFSpeDir, 1, 1) + RAFSpey3
                        RAFStatus(RAFSpeDir, 2, incr) = RAFStatus(RAFSpeDir, 2, 1) + RAFSpex3
                    ElseIf incr = move4 Then
                        RAFStatus(RAFSpeDir, 1, incr) = RAFStatus(RAFSpeDir, 1, 1) + RAFSpey
                        RAFStatus(RAFSpeDir, 2, incr) = RAFStatus(RAFSpeDir, 2, 1) + RAFSpex
                    End If
                    If RAFStatus(RAFSpeDir, 1, incr) > BOARD_ROWS Then
                        RAFStatus(RAFSpeDir, 1, incr) = BOARD_ROWS
                        RAFStatus(RAFSpeDir, 3, incr) = 3
                        RAFStatus(RAFSpeDir, 4, incr) = 2
                    ElseIf RAFStatus(RAFSpeDir, 1, incr) = 2 And RAFStatus(RAFSpeDir, 1, incr - 1) = 1 Then
                        Board(RAFStatus(RAFSpeDir, 1, incr - 1), RAFStatus(RAFSpeDir, 2, incr - 1), incr) = "     "
                        RAFStatus(RAFSpeDir, 6, incr) = 2
                    ElseIf RAFStatus(RAFSpeDir, 1, incr) < 2 And RAFStatus(RAFSpeDir, 6, 1) <> 1 Then
                        RAFStatus(RAFSpeDir, 1, incr) = 2
                        RAFStatus(RAFSpeDir, 3, incr) = 4
                        RAFStatus(RAFSpeDir, 4, incr) = 6
                    End If
                    If (RAFStatus(RAFSpeDir, 2, incr) > BOARD_COLS) Or (RAFStatus(RAFSpeDir, 2, incr) = 0) Then
                        RAFStatus(RAFSpeDir, 2, incr) = Abs(RAFStatus(RAFSpeDir, 2, incr) - BOARD_COLS)
                    ElseIf RAFStatus(RAFSpeDir, 2, incr) < 0 Then
                        RAFStatus(RAFSpeDir, 2, incr) = RAFStatus(RAFSpeDir, 2, incr) + BOARD_COLS
                    End If
                Next incr
            End If
        Next RAFSpeDir
        NextTurnPostFire
    ElseIf NumRAFRemain > 0 Then
        RAFPos = 1
        
        For i = 1 To 5
            For j = i To 2
                RAFRepairs(i, j) = 0
            Next j
        Next i
        
        NextTurnRAFInitialize
    Else
        NextTurnPostFire
    End If
End Sub

Sub NextTurnRAFInitialize()
    Application.ScreenUpdating = False
    
    For it = 1 To 23
        For jt = 1 To 23
            Cells(it, 25 + jt) = ""
        Next jt
    Next it
    
    fighterfound = False
    Do While fighterfound = False And RAFPos <= 4
        If RAFStatus(RAFPos, 1, 1) <> 0 Then
            fighterfound = True
            Cells(10, 35) = RAF_ICON
            Cells(2, 28) = "Current"
            Cells(2, 36) = "Next"
            Cells(3, 32) = "Direction"
            Cells(4, 32) = "Speed"
            Cells(5, 32) = "Altitude"
            Cells(2, 42) = "Health"
            Range(Cells(3, 42).Address).Font.Size = 5
        Else
            RAFPos = RAFPos + 1
        End If
    Loop
    
    If RAFPos <= 4 Then
        If RAFStatus(RAFPos, 6, 1) = 2 Then
            If RAFStatus(RAFPos, 4, 1) <= 3 Then
                Cells(3, 28) = "Down"
            ElseIf RAFStatus(RAFPos, 4, 1) >= 5 Or RAFStatus(RAFPos, 4, 1) <= 7 Then
                Cells(3, 28) = "Up"
            End If
            If RAFStatus(RAFPos, 4, 1) <> 2 And RAFStatus(RAFPos, 4, 1) <> 6 And Cells(3, 28) <> "" Then
                Cells(3, 28) = Cells(3, 28) & "-"
            End If
            If RAFStatus(RAFPos, 4, 1) >= 3 And RAFStatus(RAFPos, 4, 1) <= 5 Then
                Cells(3, 28) = Cells(3, 28) & "Left"
            ElseIf RAFStatus(RAFPos, 4, 1) >= 7 Or RAFStatus(RAFPos, 4, 1) = 1 Then
                Cells(3, 28) = Cells(3, 28) & "Right"
            End If
            Cells(4, 28) = RAFStatus(RAFPos, 3, 1) * 50 & " m/s"
            Cells(5, 28) = RAFStatus(RAFPos, 1, 1) & "00 m"
            Cells(3, 36) = Cells(3, 28)
            Cells(4, 36) = Cells(4, 28)
            Cells(5, 36) = Cells(5, 28)
            RAFRepairs(RAFPos, 1) = RAFStatus(RAFPos, 3, 1)
            RAFRepairs(RAFPos, 2) = RAFStatus(RAFPos, 4, 1)
            Cells(3, 42) = Int((RAFStatus(RAFPos, 7, 1) / 3) * 100) & "%"
        
            With Range(Cells(GRID_BOT_BRDR - RAFStatus(RAFPos, 1, 1), 1 + RAFStatus(RAFPos, 2, 1)).Address())
                .Borders(xlEdgeTop).ThemeColor = 1
                .Borders(xlEdgeTop).TintAndShade = 0
                .Borders(xlEdgeRight).ThemeColor = 1
                .Borders(xlEdgeRight).TintAndShade = 0
                .Borders(xlEdgeBottom).ThemeColor = 1
                .Borders(xlEdgeBottom).TintAndShade = 0
                .Borders(xlEdgeLeft).ThemeColor = 1
                .Borders(xlEdgeLeft).TintAndShade = 0
            End With
        
            Worksheets("London Siege").RAFCheckFire.Value = False
            Worksheets("London Siege").RAFCheckFire.Visible = True
            Worksheets("London Siege").RAFCheckRockets.Value = False
            Worksheets("London Siege").RAFCheckRockets.Visible = False
            Worksheets("London Siege").RAFSpeSpin.Visible = True
            Worksheets("London Siege").RAFDirSpin.Visible = True
            Worksheets("London Siege").RAFNext.Visible = True
            RAFComm = True
            
            RAFAnimate
            
        ElseIf RAFStatus(RAFPos, 6, 1) = 1 Then
            Cells(3, 28) = "NA"
            Cells(4, 28) = "NA"
            Cells(5, 28) = "NA"
            Cells(3, 36) = Cells(3, 28)
            Cells(4, 36) = Cells(4, 28)
            Cells(5, 36) = Cells(5, 28)
            RAFRepairs(RAFPos, 1) = RAFStatus(RAFPos, 3, 1)
            RAFRepairs(RAFPos, 2) = RAFStatus(RAFPos, 4, 1)
            Cells(3, 42) = Int((RAFStatus(RAFPos, 7, 1) / 3) * 100) & "%"
        
            With Range(Cells(GRID_BOT_BRDR - RAFStatus(RAFPos, 1, 1), 1 + RAFStatus(RAFPos, 2, 1)).Address())
                .Borders(xlEdgeTop).ThemeColor = 1
                .Borders(xlEdgeTop).TintAndShade = 0
                .Borders(xlEdgeRight).ThemeColor = 1
                .Borders(xlEdgeRight).TintAndShade = 0
                .Borders(xlEdgeBottom).ThemeColor = 1
                .Borders(xlEdgeBottom).TintAndShade = 0
                .Borders(xlEdgeLeft).ThemeColor = 1
                .Borders(xlEdgeLeft).TintAndShade = 0
            End With
        
            Worksheets("London Siege").RAFCheckFire.Value = False
            Worksheets("London Siege").RAFCheckFire.Visible = False
            Worksheets("London Siege").RAFCheckRockets.Value = False
            Worksheets("London Siege").RAFCheckRockets.Visible = False
            Worksheets("London Siege").RAFSpeSpin.Visible = False
            Worksheets("London Siege").RAFDirSpin.Visible = False
            Worksheets("London Siege").RAFCommLaunch.Value = False
            Worksheets("London Siege").RAFCommLaunch.Visible = True
            Worksheets("London Siege").RAFNext.Visible = True
            RAFComm = True
        End If
    Else
        RAFComm = False
        Range(Cells(3, 42).Address).Font.Size = 11
        NextTurnPostFire
    End If
    Application.ScreenUpdating = True
End Sub

Sub RAFCommLaunch()
    If RAFComm = True Then
        If Worksheets("London Siege").RAFCommLaunch.Value = True Then
            Cells(3, 36) = "Up"
            Cells(4, 36) = "100 m/s"
            Cells(5, 36) = "300 m"
            RAFRepairs(RAFPos, 1) = 2
            RAFRepairs(RAFPos, 2) = 6
            Cells(8, 37) = RAF_ICON
            Cells(9, 37) = RAF_ICON
        ElseIf Worksheets("London Siege").RAFCommLaunch.Value = False Then
            Cells(3, 36) = "NA"
            Cells(4, 36) = "NA"
            Cells(5, 36) = "NA"
            RAFRepairs(RAFPos, 1) = 0
            RAFRepairs(RAFPos, 2) = 0
            Cells(8, 37) = ""
            Cells(9, 37) = ""
        End If
    End If
End Sub

Sub RAFCheckFire()
    If RAFComm = True Then
        If Worksheets("London Siege").RAFCheckFire.Value = True And RAFStatus(RAFPos, 5, 1) > 0 Then
            Worksheets("London Siege").RAFCheckRockets.Visible = True
            Cells(5, 46) = "(" & RAFStatus(RAFPos, 5, 1) & " Left)"
        ElseIf Worksheets("London Siege").RAFCheckFire.Value = False Then
            Worksheets("London Siege").RAFCheckRockets.Value = False
            Worksheets("London Siege").RAFCheckRockets.Visible = False
            Cells(5, 46) = ""
        End If
        RAFAnimate
    End If
End Sub

Sub RAFCheckRockets()
    If RAFComm = True Then
        RAFAnimate
    End If
End Sub

Sub RAFDirSpinLeft()
    If RAFComm = True Then
        RAFRepairs(RAFPos, 2) = RAFRepairs(RAFPos, 2) - 1
        DirCheck = 0
        If RAFStatus(RAFPos, 3, 1) > 2 Then
            DirCheck = 1
        ElseIf RAFStatus(RAFPos, 3, 1) <= 2 Then
            DirCheck = 2
        End If
        UpB = RAFStatus(RAFPos, 4, 1) + DirCheck
        If UpB > 8 Then
            Adju = -8
        Else
            Adju = 8
        End If

        If (RAFRepairs(RAFPos, 2) <= UpB And RAFRepairs(RAFPos, 2) >= UpB - 2 * DirCheck) Or _
        (RAFRepairs(RAFPos, 2) <= UpB + Adju And RAFRepairs(RAFPos, 2) >= UpB + Adju - 2 * DirCheck) Then
            'No straddle
        Else
            RAFRepairs(RAFPos, 2) = RAFRepairs(RAFPos, 2) + (2 * DirCheck + 1)
        End If
        
        If RAFRepairs(RAFPos, 2) > 8 Then
            RAFRepairs(RAFPos, 2) = RAFRepairs(RAFPos, 2) - 8
        ElseIf RAFRepairs(RAFPos, 2) < 1 Then
            RAFRepairs(RAFPos, 2) = RAFRepairs(RAFPos, 2) + 8
        End If
        
        RAFAnimate
        
    End If
End Sub

Sub RAFDirSpinRight()
    If RAFComm = True Then
        RAFRepairs(RAFPos, 2) = RAFRepairs(RAFPos, 2) + 1
        DirCheck = 0
        If RAFStatus(RAFPos, 3, 1) > 2 Then
            DirCheck = 1
        ElseIf RAFStatus(RAFPos, 3, 1) <= 2 Then
            DirCheck = 2
        End If
        UpB = RAFStatus(RAFPos, 4, 1) + DirCheck
        If UpB > 8 Then
            Adju = -8
        Else
            Adju = 8
        End If

        If (RAFRepairs(RAFPos, 2) <= UpB And RAFRepairs(RAFPos, 2) >= UpB - 2 * DirCheck) Or _
        (RAFRepairs(RAFPos, 2) <= UpB + Adju And RAFRepairs(RAFPos, 2) >= UpB + Adju - 2 * DirCheck) Then
            'No straddle
        Else
            RAFRepairs(RAFPos, 2) = RAFRepairs(RAFPos, 2) - (2 * DirCheck + 1)
        End If
        
        If RAFRepairs(RAFPos, 2) > 8 Then
            RAFRepairs(RAFPos, 2) = RAFRepairs(RAFPos, 2) - 8
        ElseIf RAFRepairs(RAFPos, 2) < 1 Then
            RAFRepairs(RAFPos, 2) = RAFRepairs(RAFPos, 2) + 8
        End If
        
        RAFAnimate
        
    End If
End Sub

Sub RAFSpeSpinLeft()
    If RAFComm = True And RAFRepairs(RAFPos, 1) > 2 Then
        RAFRepairs(RAFPos, 1) = RAFRepairs(RAFPos, 1) - 1
        RAFAnimate
    End If
End Sub

Sub RAFSpeSpinRight()
    If RAFComm = True And RAFRepairs(RAFPos, 1) < 4 Then
        RAFRepairs(RAFPos, 1) = RAFRepairs(RAFPos, 1) + 1
        RAFAnimate
    End If
End Sub

Sub RAFAnimate()
    If RAFComm = True Then
        Application.ScreenUpdating = False
        ias = 12
        jas = 37
        signi = 0
        signj = 0
        Cells(3, 36) = ""
        Cells(4, 36) = ""
        Cells(5, 36) = ""
        If RAFRepairs(RAFPos, 2) <= 3 Then
            Cells(3, 36) = "Down"
            If RAFStatus(RAFPos, 1, 1) - RAFRepairs(RAFPos, 1) < 2 Then
                Cells(5, 36) = "Low"
            Else
                Cells(5, 36) = RAFStatus(RAFPos, 1, 1) - RAFRepairs(RAFPos, 1) & "00 m"
            End If
            ias = 15
            signi = -1
        ElseIf RAFRepairs(RAFPos, 2) >= 5 And RAFRepairs(RAFPos, 2) <= 7 Then
            Cells(3, 36) = "Up"
            If RAFStatus(RAFPos, 1, 1) + RAFRepairs(RAFPos, 1) > BOARD_ROWS Then
                Cells(5, 36) = "High"
            Else
                Cells(5, 36) = RAFStatus(RAFPos, 1, 1) + RAFRepairs(RAFPos, 1) & "00 m"
            End If
            ias = 9
            signi = 1
        End If
        If RAFRepairs(RAFPos, 2) <> 2 And RAFRepairs(RAFPos, 2) <> 6 And Cells(3, 36) <> "" Then
            Cells(3, 36) = Cells(3, 36) & "-"
        End If
        If RAFRepairs(RAFPos, 2) >= 3 And RAFRepairs(RAFPos, 2) <= 5 Then
            Cells(3, 36) = Cells(3, 36) & "Left"
            jas = 34
            signj = -1
        ElseIf RAFRepairs(RAFPos, 2) >= 7 Or RAFRepairs(RAFPos, 2) = 1 Then
            Cells(3, 36) = Cells(3, 36) & "Right"
            jas = 40
            signj = 1
        End If
        Cells(4, 36) = RAFRepairs(RAFPos, 1) * 50 & " m/s"
        If RAFRepairs(RAFPos, 2) = 4 Or RAFRepairs(RAFPos, 2) = 8 Then
            Cells(5, 36) = RAFStatus(RAFPos, 1, 1) & "00 m"
        End If
        
        For i = 6 To 18
            For j = 31 To 43
                If i >= 10 And i <= 14 And j >= 35 And j <= 39 Then
                    'Don't erase center
                Else
                    Cells(i, j) = ""
                End If
            Next j
        Next i
        
        icons = ""
        
        If Worksheets("London Siege").RAFCheckRockets.Value = True Then
            icons = "+"
        ElseIf Worksheets("London Siege").RAFCheckFire.Value = True Then
            icons = "o"
        Else
            icons = RAF_ICON
        End If
        
        For i = 0 To RAFRepairs(RAFPos, 1) - 1
            Cells(ias - i * signi, jas + i * signj) = icons
        Next i
        
        Application.ScreenUpdating = True
    End If
End Sub

Sub RAFNext()
    If RAFComm = True Then
        RAFSpeDir = RAFPos             'Adapt code for use with Comm active
        RAFSpe = RAFRepairs(RAFPos, 1)
        RAFDir = RAFRepairs(RAFPos, 2)
        RAFStatus(RAFSpeDir, 3, 1) = RAFSpe
        RAFStatus(RAFPos, 4, 1) = RAFDir
        
        If Worksheets("London Siege").RAFCheckFire.Value = True Then
            RAFFire(RAFSpeDir, 1, 1) = RAFStatus(RAFSpeDir, 1, 1)
            RAFFire(RAFSpeDir, 2, 1) = RAFStatus(RAFSpeDir, 2, 1)
            If Worksheets("London Siege").RAFCheckRockets.Value = True Then
                RAFFire(RAFSpeDir, 3, 1) = 2
                RAFStatus(RAFPos, 5, 1) = RAFStatus(RAFPos, 5, 1) - 1
            Else
                RAFFire(RAFSpeDir, 3, 1) = 1
            End If
            RAFFire(RAFSpeDir, 4, 1) = RAFDir
        End If
        
        If Worksheets("London Siege").RAFCommLaunch.Value = True Then
            RAFStatus(RAFPos, 3, 1) = RAFRepairs(RAFPos, 1)
            RAFStatus(RAFPos, 4, 1) = RAFRepairs(RAFPos, 2)
        End If
        
        RAFSpex = 0
        RAFSpey = 0
        
        'RAFDir determines speed in x and y directions of active RAF
        If RAFDir = 1 Then
            RAFSpex = RAFSpe
            RAFSpey = (0 - RAFSpe)
        ElseIf RAFDir = 2 Then
            RAFSpey = (0 - RAFSpe)
        ElseIf RAFDir = 3 Then
            RAFSpex = (0 - RAFSpe)
            RAFSpey = (0 - RAFSpe)
        ElseIf RAFDir = 4 Then
            RAFSpex = (0 - RAFSpe)
        ElseIf RAFDir = 5 Then
            RAFSpex = (0 - RAFSpe)
            RAFSpey = RAFSpe
        ElseIf RAFDir = 6 Then
            RAFSpey = RAFSpe
        ElseIf RAFDir = 7 Then
            RAFSpex = RAFSpe
            RAFSpey = RAFSpe
        ElseIf RAFDir = 8 Then
            RAFSpex = RAFSpe
        End If
            
        'Fill movement matrix for enemy fighters
        'Note: 0.4 is accelerative factor
        '      to timeshift ahead movement of fighters with speed = 3
        If RAFSpey < 0 Then
            RAFSpey1 = (-1) * Int(Abs(RAFSpey) / 4 + 0.4)
            RAFSpey2 = (-1) * Int(Abs(RAFSpey) / 2 + 0.4)
            RAFSpey3 = (-1) * Int(3 * Abs(RAFSpey) / 4 + 0.4)
        Else
            RAFSpey1 = Int(RAFSpey / 4 + 0.4)
            RAFSpey2 = Int(RAFSpey / 2 + 0.4)
            RAFSpey3 = Int(3 * RAFSpey / 4 + 0.4)
        End If
        If RAFSpex < 0 Then
            RAFSpex1 = (-1) * Int(Abs(RAFSpex) / 4 + 0.4)
            RAFSpex2 = (-1) * Int(Abs(RAFSpex) / 2 + 0.4)
            RAFSpex3 = (-1) * Int(3 * Abs(RAFSpex) / 4 + 0.4)
        Else
            RAFSpex1 = Int(RAFSpex / 4 + 0.4)
            RAFSpex2 = Int(RAFSpex / 2 + 0.4)
            RAFSpex3 = Int(3 * RAFSpex / 4 + 0.4)
        End If
             
        If RAFFire(RAFSpeDir, 3, 1) <> 0 Then
            FireValid = True
            ShotCount = ShotCount + 1
            Randomize
            FireStart = Int(2 * Rnd + 2)
            If FireStart = 3 Then
                RAFFire(RAFSpeDir, 1, 2) = RAFFire(RAFSpeDir, 1, 1)
                RAFFire(RAFSpeDir, 2, 2) = RAFFire(RAFSpeDir, 2, 1)
                RAFFire(RAFSpeDir, 3, 2) = RAFFire(RAFSpeDir, 3, 1)
                RAFFire(RAFSpeDir, 4, 2) = RAFFire(RAFSpeDir, 4, 1)
            End If
            For incr = FireStart To TIME_END
                If FireValid = True Then
                    RocketFire = RAFFire(RAFSpeDir, 3, 1) - 1      'If rockets are being fired, RocketFire will double traveling speed of shot
                    If RAFFire(RAFSpeDir, 4, 1) <= 3 Then
                        RAFFire(RAFSpeDir, 1, incr) = RAFFire(RAFSpeDir, 1, incr - 1) - 1 - RocketFire
                    ElseIf RAFFire(RAFSpeDir, 4, 1) >= 5 And RAFFire(RAFSpeDir, 4, 1) <= 7 Then
                        RAFFire(RAFSpeDir, 1, incr) = RAFFire(RAFSpeDir, 1, incr - 1) + 1 + RocketFire
                    Else
                        RAFFire(RAFSpeDir, 1, incr) = RAFFire(RAFSpeDir, 1, incr - 1)
                    End If
                    If RAFFire(RAFSpeDir, 4, 1) >= 3 And RAFFire(RAFSpeDir, 4, 1) <= 5 Then
                        RAFFire(RAFSpeDir, 2, incr) = RAFFire(RAFSpeDir, 2, incr - 1) - 1 - RocketFire
                    ElseIf RAFFire(RAFSpeDir, 4, 1) >= 7 Or RAFFire(RAFSpeDir, 4, 1) = 1 Then
                        RAFFire(RAFSpeDir, 2, incr) = RAFFire(RAFSpeDir, 2, incr - 1) + 1 + RocketFire
                    Else
                        RAFFire(RAFSpeDir, 2, incr) = RAFFire(RAFSpeDir, 2, incr - 1)
                    End If
                    RAFFire(RAFSpeDir, 3, incr) = RAFFire(RAFSpeDir, 3, incr - 1)
                    RAFFire(RAFSpeDir, 4, incr) = RAFFire(RAFSpeDir, 4, incr - 1)
                    If RAFFire(RAFSpeDir, 1, incr) <= 0 Or RAFFire(RAFSpeDir, 1, incr) >= 22 Or _
                    RAFFire(RAFSpeDir, 2, incr) <= 0 Or RAFFire(RAFSpeDir, 2, incr) >= 22 Then
                        FireValid = False
                        RAFFire(RAFSpeDir, 1, incr) = 0
                        RAFFire(RAFSpeDir, 2, incr) = 0
                        RAFFire(RAFSpeDir, 3, incr) = 0
                        RAFFire(RAFSpeDir, 4, incr) = 0
                    End If
                End If
            Next incr
        End If
        Randomize
        move1 = Int(4 * Rnd + 4)
        move2 = Int(4 * Rnd + 9)
        move3 = Int(4 * Rnd + 14)
        move4 = Int(4 * Rnd + 19)
        For incr = 2 To TIME_END
            For i = 1 To 7
                RAFStatus(RAFSpeDir, i, incr) = RAFStatus(RAFSpeDir, i, incr - 1)
            Next i
            If incr <> move1 And incr <> move2 And incr <> move3 And incr <> move4 Then
                'Procedure Bypass
            ElseIf incr = move1 Then
                RAFStatus(RAFSpeDir, 1, incr) = RAFStatus(RAFSpeDir, 1, 1) + RAFSpey1
                RAFStatus(RAFSpeDir, 2, incr) = RAFStatus(RAFSpeDir, 2, 1) + RAFSpex1
            ElseIf incr = move2 Then
                RAFStatus(RAFSpeDir, 1, incr) = RAFStatus(RAFSpeDir, 1, 1) + RAFSpey2
                RAFStatus(RAFSpeDir, 2, incr) = RAFStatus(RAFSpeDir, 2, 1) + RAFSpex2
            ElseIf incr = move3 Then
                RAFStatus(RAFSpeDir, 1, incr) = RAFStatus(RAFSpeDir, 1, 1) + RAFSpey3
                RAFStatus(RAFSpeDir, 2, incr) = RAFStatus(RAFSpeDir, 2, 1) + RAFSpex3
            ElseIf incr = move4 Then
                RAFStatus(RAFSpeDir, 1, incr) = RAFStatus(RAFSpeDir, 1, 1) + RAFSpey
                RAFStatus(RAFSpeDir, 2, incr) = RAFStatus(RAFSpeDir, 2, 1) + RAFSpex
            End If
            If RAFStatus(RAFSpeDir, 1, incr) > BOARD_ROWS Then
                RAFStatus(RAFSpeDir, 1, incr) = BOARD_ROWS
                RAFStatus(RAFSpeDir, 3, incr) = 3
                RAFStatus(RAFSpeDir, 4, incr) = 2
            ElseIf RAFStatus(RAFSpeDir, 1, incr) = 2 And RAFStatus(RAFSpeDir, 1, incr - 1) = 1 Then
                Board(RAFStatus(RAFSpeDir, 1, incr - 1), RAFStatus(RAFSpeDir, 2, incr - 1), incr) = "     "
                RAFStatus(RAFSpeDir, 6, incr) = 2
            ElseIf RAFStatus(RAFSpeDir, 1, incr) < 2 And RAFStatus(RAFSpeDir, 6, 1) <> 1 Then
                RAFStatus(RAFSpeDir, 1, incr) = 2
                RAFStatus(RAFSpeDir, 3, incr) = 4
                RAFStatus(RAFSpeDir, 4, incr) = 6
            End If
            If (RAFStatus(RAFSpeDir, 2, incr) > BOARD_COLS) Or (RAFStatus(RAFSpeDir, 2, incr) = 0) Then
                RAFStatus(RAFSpeDir, 2, incr) = Abs(RAFStatus(RAFSpeDir, 2, incr) - BOARD_COLS)
            ElseIf RAFStatus(RAFSpeDir, 2, incr) < 0 Then
                RAFStatus(RAFSpeDir, 2, incr) = RAFStatus(RAFSpeDir, 2, incr) + BOARD_COLS
            End If
        Next incr
        
        RAFComm = False
        Worksheets("London Siege").RAFCheckFire.Value = False
        Worksheets("London Siege").RAFCheckFire.Visible = False
        Worksheets("London Siege").RAFCheckRockets.Value = False
        Worksheets("London Siege").RAFCheckRockets.Visible = False
        Worksheets("London Siege").RAFSpeSpin.Visible = False
        Worksheets("London Siege").RAFDirSpin.Visible = False
        Worksheets("London Siege").RAFCommLaunch.Value = False
        Worksheets("London Siege").RAFCommLaunch.Visible = False
        Worksheets("London Siege").RAFNext.Visible = False
        
        With Range(Cells(GRID_BOT_BRDR - RAFStatus(RAFPos, 1, 1), 1 + RAFStatus(RAFPos, 2, 1)).Address())
            .Borders(xlEdgeTop).ColorIndex = 0
            .Borders(xlEdgeTop).TintAndShade = 0
            .Borders(xlEdgeRight).ColorIndex = 0
            .Borders(xlEdgeRight).TintAndShade = 0
            .Borders(xlEdgeBottom).ColorIndex = 0
            .Borders(xlEdgeBottom).TintAndShade = 0
            .Borders(xlEdgeLeft).ColorIndex = 0
            .Borders(xlEdgeLeft).TintAndShade = 0
        End With
        
        RAFPos = RAFPos + 1
        
        NextTurnRAFInitialize
    End If

End Sub

Sub NextTurnPostFire()

    Worksheets("London Siege").StartGame.Enabled = False
    Worksheets("London Siege").QuitGame.Enabled = False
    
    For turi = 1 To BOARD_COLS
        FireDelay = 0
        For incr = 2 To TIME_END
            If Turrets(turi, incr - 1).Fire.Type <> 0 Then
                Turrets(turi, incr).Fire.Direction = Turrets(turi, incr - 1).Fire.Direction
                Turrets(turi, incr).Fire.Type = Turrets(turi, incr - 1).Fire.Type
                If incr = 2 Then
                    Randomize
                    FireDelay = Int(2 * Rnd)
                End If
                If FireDelay = 1 And incr = 2 Then
                    Turrets(turi, incr).Fire.Row = Turrets(turi, incr - 1).Fire.Row
                    Turrets(turi, incr).Fire.Column = Turrets(turi, incr - 1).Fire.Column
                Else
                    If Turrets(turi, incr).Fire.Direction >= Dir45 And Turrets(turi, incr).Fire.Direction <= Dir135 Then
                        Turrets(turi, incr).Fire.Row = Turrets(turi, incr - 1).Fire.Row + 1
                    ElseIf (incr + FireDelay) Mod 2 = 0 Then
                        Turrets(turi, incr).Fire.Row = Turrets(turi, incr - 1).Fire.Row + 1
                    Else
                        Turrets(turi, incr).Fire.Row = Turrets(turi, incr - 1).Fire.Row
                    End If
                    If Turrets(turi, incr).Fire.Direction <= Dir45 Then
                        Turrets(turi, incr).Fire.Column = Turrets(turi, incr - 1).Fire.Column + 1
                    ElseIf Turrets(turi, incr).Fire.Direction = Dir63 And (incr + FireDelay) Mod 2 = 0 Then
                        Turrets(turi, incr).Fire.Column = Turrets(turi, incr - 1).Fire.Column + 1
                    ElseIf Turrets(turi, incr).Fire.Direction = Dir72 And (incr - FireDelay - 2) Mod 3 = 0 Then
                        Turrets(turi, incr).Fire.Column = Turrets(turi, incr - 1).Fire.Column + 1
                    ElseIf Turrets(turi, incr).Fire.Direction = Dir76 And (incr - FireDelay - 2) Mod 4 = 2 Then
                        Turrets(turi, incr).Fire.Column = Turrets(turi, incr - 1).Fire.Column + 1
                    ElseIf Turrets(turi, incr).Fire.Direction = Dir104 And (incr - FireDelay - 2) Mod 4 = 2 Then
                        Turrets(turi, incr).Fire.Column = Turrets(turi, incr - 1).Fire.Column - 1
                    ElseIf Turrets(turi, incr).Fire.Direction = Dir108 And (incr - FireDelay - 2) Mod 3 = 0 Then
                        Turrets(turi, incr).Fire.Column = Turrets(turi, incr - 1).Fire.Column - 1
                    ElseIf Turrets(turi, incr).Fire.Direction = Dir117 And (incr + FireDelay) Mod 2 = 0 Then
                        Turrets(turi, incr).Fire.Column = Turrets(turi, incr - 1).Fire.Column - 1
                    ElseIf Turrets(turi, incr).Fire.Direction >= Dir135 Then
                        Turrets(turi, incr).Fire.Column = Turrets(turi, incr - 1).Fire.Column - 1
                    Else
                        Turrets(turi, incr).Fire.Column = Turrets(turi, incr - 1).Fire.Column
                    End If
                    If Turrets(turi, incr).Fire.Row > BOARD_ROWS Or Turrets(turi, incr).Fire.Column > BOARD_COLS _
                    Or Turrets(turi, incr).Fire.Column < 1 Then
                        Dim clearTurretFire As TurretFire
                        Turrets(turi, incr).Fire = clearTurretFire
                    End If
                End If
            End If
        Next incr
    Next turi
    
    Dim targetRow As Integer
    Dim targetCol As Integer

    For incr = 2 To TIME_END
'Transfer structures through time
        For j = 1 To BOARD_COLS
            StructureStatus(j, incr) = StructureStatus(j, incr - 1)
            If Board(STRUCT_ROW, j, incr) = "" Then Board(STRUCT_ROW, j, incr) = Board(STRUCT_ROW, j, incr - 1)
            Turrets(j, incr).Icon = Turrets(j, incr - 1).Icon
            Turrets(j, incr).Health = Turrets(j, incr - 1).Health
            If Turrets(j, incr).Icon <> "" Then Board(TURRET_ROW, j, incr) = Turrets(j, incr).Icon
        Next j
'EFighters
        For B = 1 To 20
            If EFighters(B, 1, incr) <> 0 Then     'Skip destroyed or inactive fighters
                If Board(EFighters(B, 1, incr), EFighters(B, 2, incr), incr) <> "." Then
                    Board(EFighters(B, 1, incr), EFighters(B, 2, incr), incr) = EF_ICON
                End If
            End If
'EFFire
            If EFFire(B, 3, incr) <> 0 Then
                If EFFire(B, 1, incr) = STRUCT_ROW And StructureStatus(EFFire(B, 2, incr), incr) < 4 Then
                    StructureStatus(EFFire(B, 2, incr), incr) = StructureStatus(EFFire(B, 2, incr), incr) + 1
                    For yu = 1 To 3
                        For zu = incr To TIME_END
                            EFFire(B, yu, zu) = 0
                        Next zu
                    Next yu
                ElseIf EFFire(B, 1, incr) = TURRET_ROW And Turrets(EFFire(B, 2, incr), 1).Icon <> "" And Turrets(EFFire(B, 2, incr), incr).Health < 4 Then
                    Turrets(EFFire(B, 2, incr), incr).Health = Turrets(EFFire(B, 2, incr), incr).Health + 1
                    For yu = 1 To 3
                        For zu = incr To TIME_END
                            EFFire(B, yu, zu) = 0
                        Next zu
                    Next yu
                ElseIf Board(EFFire(B, 1, incr), EFFire(B, 2, incr), incr) = "" Then
                    If EFFire(B, 3, incr) = 1 Or EFFire(B, 3, incr) = 5 Then
                        Board(EFFire(B, 1, incr), EFFire(B, 2, incr), incr) = "`"
                    ElseIf EFFire(B, 3, incr) = 2 Or EFFire(B, 3, incr) = 6 Then
                        Board(EFFire(B, 1, incr), EFFire(B, 2, incr), incr) = "''"
                    ElseIf EFFire(B, 3, incr) = 3 Or EFFire(B, 3, incr) = 7 Then
                        Board(EFFire(B, 1, incr), EFFire(B, 2, incr), incr) = ","
                    ElseIf EFFire(B, 3, incr) = 4 Or EFFire(B, 3, incr) = 8 Then
                        Board(EFFire(B, 1, incr), EFFire(B, 2, incr), incr) = "'-"
                    End If
                End If
            End If
        Next B
'EBombers
        For B = 1 To 20
            'First place bombers that don't move in turn
            If EBombers(B, 1, incr) <> 0 And EBombers(B, 2, incr) = EBombers(B, 2, incr - 1) Then
                If EBombers(B, 4, incr) = 1 Then
                    If EBombers(B, 5, incr) = 0 Then Board(EBombers(B, 1, incr), EBombers(B, 2, incr), incr) = EB_ICON_LEFT
                    If EBombers(B, 5, incr) = 1 Then Board(EBombers(B, 1, incr), EBombers(B, 2, incr), incr) = EB_ICON_LEFT_DMG
                ElseIf EBombers(B, 4, incr) = 2 Then
                    If EBombers(B, 5, incr) = 0 Then Board(EBombers(B, 1, incr), EBombers(B, 2, incr), incr) = EB_ICON_RIGHT
                    If EBombers(B, 5, incr) = 1 Then Board(EBombers(B, 1, incr), EBombers(B, 2, incr), incr) = EB_ICON_RIGHT_DMG
                End If
            End If
        Next B

        For B = 1 To 20
            Dim BomberUp As Boolean
            Dim BomberEven As Boolean
            Dim BomberDown As Boolean
            If EBombers(B, 1, incr) <> 0 And EBombers(B, 2, incr) <> EBombers(B, 2, incr - 1) Then     'Skip destroyed or inactive bombers
                'Determine bomber's possible moves (will not go below row 13)
                'Bomber only changes altitude on turn with forward movement
                BomberUp = False
                BomberEven = False
                BomberDown = False
                If Board(EBombers(B, 1, incr), EBombers(B, 2, incr), incr) <> EF_ICON And _
                Board(EBombers(B, 1, incr), EBombers(B, 2, incr), incr) <> EB_ICON_LEFT And _
                Board(EBombers(B, 1, incr), EBombers(B, 2, incr), incr) <> EB_ICON_RIGHT And _
                Board(EBombers(B, 1, incr), EBombers(B, 2, incr), incr) <> EB_ICON_LEFT_DMG And _
                Board(EBombers(B, 1, incr), EBombers(B, 2, incr), incr) <> EB_ICON_RIGHT_DMG Then
                    BomberEven = True
                End If
                If EBombers(B, 1, incr) < BOARD_ROWS Then
                    If Board(EBombers(B, 1, incr) + 1, EBombers(B, 2, incr), incr) <> EF_ICON And _
                    Board(EBombers(B, 1, incr) + 1, EBombers(B, 2, incr), incr) <> EB_ICON_LEFT And _
                    Board(EBombers(B, 1, incr) + 1, EBombers(B, 2, incr), incr) <> EB_ICON_RIGHT And _
                    Board(EBombers(B, 1, incr) + 1, EBombers(B, 2, incr), incr) <> EB_ICON_LEFT_DMG And _
                    Board(EBombers(B, 1, incr) + 1, EBombers(B, 2, incr), incr) <> EB_ICON_RIGHT_DMG Then
                        BomberUp = True
                    End If
                End If
                If EBombers(B, 1, incr) > EB_MIN_ROW Then
                    If Board(EBombers(B, 1, incr) - 1, EBombers(B, 2, incr), incr) <> EF_ICON And _
                    Board(EBombers(B, 1, incr) - 1, EBombers(B, 2, incr), incr) <> EB_ICON_LEFT And _
                    Board(EBombers(B, 1, incr) - 1, EBombers(B, 2, incr), incr) <> EB_ICON_RIGHT And _
                    Board(EBombers(B, 1, incr) - 1, EBombers(B, 2, incr), incr) <> EB_ICON_LEFT_DMG And _
                    Board(EBombers(B, 1, incr) - 1, EBombers(B, 2, incr), incr) <> EB_ICON_RIGHT_DMG Then
                        BomberDown = True
                    End If
                End If
                If BomberUp = True Or BomberEven = True Or BomberDown = True Then
                    SpaceFound = False
                    Do While SpaceFound = False
                        Randomize
                        SpaceFoundInt = Int(3 * Rnd + 1)
                        If (SpaceFoundInt = 1 And BomberUp = True) Or (SpaceFoundInt = 2 And BomberEven = True) Or _
                        (SpaceFoundInt = 3 And BomberDown = True) Then
                            InitialAlt = EBombers(B, 1, incr)
                            'Populate remaining bomber path
                            For kvert = incr To TIME_END
                                EBombers(B, 1, kvert) = InitialAlt - (SpaceFoundInt - 2)
                            Next kvert
                            SpaceFound = True
                        End If
                    Loop
                Else
                    'If no evasion is possible, bomber will take any move inbounds
                    SpaceFound = False
                    Do While SpaceFound = False
                        Randomize
                        SpaceFoundInt = Int(3 * Rnd + 1)
                        If (EBombers(B, 1, incr) - (SpaceFoundInt - 2) > 12) And (EBombers(B, 1, incr) - (SpaceFoundInt - 2) < 22) Then
                            InitialAlt = EBombers(B, 1, incr)
                            For kvert = incr To TIME_END
                                EBombers(B, 1, kvert) = InitialAlt - (SpaceFoundInt - 2)
                            Next kvert
                            SpaceFound = True
                        End If
                    Loop
                End If
                If EBombers(B, 4, incr) = 1 Then
                    If EBombers(B, 5, incr) = 0 Then Board(EBombers(B, 1, incr), EBombers(B, 2, incr), incr) = EB_ICON_LEFT
                    If EBombers(B, 5, incr) = 1 Then Board(EBombers(B, 1, incr), EBombers(B, 2, incr), incr) = EB_ICON_LEFT_DMG
                ElseIf EBombers(B, 4, incr) = 2 Then
                    If EBombers(B, 5, incr) = 0 Then Board(EBombers(B, 1, incr), EBombers(B, 2, incr), incr) = EB_ICON_RIGHT
                    If EBombers(B, 5, incr) = 1 Then Board(EBombers(B, 1, incr), EBombers(B, 2, incr), incr) = EB_ICON_RIGHT_DMG
                End If
            End If
'EBFire
            If EBFire(B, 1, incr) <> 0 Then
                targetRow = EBFire(B, 1, incr)
                targetCol = EBFire(B, 2, incr)
                If targetRow = STRUCT_ROW Or _
                (targetRow = TURRET_ROW And Turrets(targetCol, incr).Icon <> "" And Turrets(targetCol, incr).Health < 4) Then
            
                    Call ColorizerBomb(targetRow, targetCol, incr)
            
                    If targetRow = STRUCT_ROW Then
                        If StructureStatus(targetCol, incr) >= 3 Then
                            StructureStatus(targetCol, incr) = 4
                        Else
                            StructureStatus(targetCol, incr) = StructureStatus(targetCol, incr) + 2
                        End If
                    ElseIf targetRow = TURRET_ROW Then
                        If Turrets(targetCol, incr).Health >= 3 Then
                            Turrets(targetCol, incr).Health = 4
                        Else
                            Turrets(targetCol, incr).Health = Turrets(targetCol, incr).Health + 2
                        End If
                    End If
                    If targetCol > 1 Then
                        'Damage done to structures to the left
                        If StructureStatus(targetCol - 1, incr) < 4 Then
                            StructureStatus(targetCol - 1, incr) = StructureStatus(targetCol - 1, incr) + 1
                        End If
                        If Turrets(targetCol - 1, incr).Health < 4 And Turrets(targetCol - 1, 1).Icon <> "" Then
                            Turrets(targetCol - 1, incr).Health = Turrets(targetCol - 1, incr).Health + 1
                        End If
                    End If
                    If targetCol < BOARD_COLS Then
                        'Damage done to structures to the right
                        If StructureStatus(targetCol + 1, incr) < 4 Then
                            StructureStatus(targetCol + 1, incr) = StructureStatus(targetCol + 1, incr) + 1
                        End If
                        If Turrets(targetCol + 1, incr).Health < 4 And Turrets(targetCol + 1, 1).Icon <> "" Then
                            Turrets(targetCol + 1, incr).Health = Turrets(targetCol + 1, incr).Health + 1
                        End If
                    End If
                    For yu = 1 To 2
                        For zu = incr To TIME_END
                            EBFire(B, yu, zu) = 0
                        Next zu
                    Next yu
                ElseIf targetRow > TURRET_ROW Or (targetRow = TURRET_ROW And (Turrets(targetCol, incr).Icon = "" Or Turrets(targetCol, incr).Health >= 4)) Then
                    Select Case Board(targetRow, targetCol, incr)
                        Case "", EF_ICON_FIRE_VERT, EF_ICON_FIRE_HORIZ, EF_ICON_FIRE_DIAG_POS, EF_ICON_FIRE_DIAG_NEG
                            Board(targetRow, targetCol, incr) = "!"
                    End Select
                    'Add tail where applicable
                    For i = 1 To 2
                        If targetRow < EBFire(B, 1, 1) - i Then
                            If Board(targetRow + i, targetCol, incr) = "" Then
                                Board(targetRow + i, targetCol, incr) = "|"
                            End If
                        End If
                    Next i
                End If
            End If
        Next B
'RAF
        For i = 1 To 4
            If RAFStatus(i, 1, incr) <> 0 Then     'Skip destroyed or inactive fighters
                If (Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = EF_ICON Or Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = EB_ICON_RIGHT Or _
                Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = EB_ICON_LEFT Or Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = EB_ICON_RIGHT_DMG Or _
                Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = EB_ICON_LEFT_DMG) And (RAFStatus(i, 1, incr) <> RAFStatus(i, 1, incr - 1) Or _
                RAFStatus(i, 2, incr) <> RAFStatus(i, 2, incr - 1)) Then                                                                          'RAF avoids collisions, but they can get run into
                    Randomize
                    dirchoice = Int(2 * Rnd + 1)
                    dircur = RAFStatus(i, 4, incr)
                    ishift = 0
                    jshift = 0
                    If dircur = 1 Then
                        If dirchoice = 1 Then
                            ishift = 1
                        ElseIf dirchoice = 2 Then
                            jshift = -1
                        End If
                    ElseIf dircur = 2 Then
                        If dirchoice = 1 Then
                            jshift = 1
                        ElseIf dirchoice = 2 Then
                            jshift = -1
                        End If
                    ElseIf dircur = 3 Then
                        If dirchoice = 1 Then
                            jshift = 1
                        ElseIf dirchoice = 2 Then
                            ishift = 1
                        End If
                    ElseIf dircur = 4 Then
                        If dirchoice = 1 Then
                            ishift = -1
                        ElseIf dirchoice = 2 Then
                            ishift = 1
                        End If
                    ElseIf dircur = 5 Then
                        If dirchoice = 1 Then
                            ishift = -1
                        ElseIf dirchoice = 2 Then
                            jshift = 1
                        End If
                    ElseIf dircur = 6 Then
                        If dirchoice = 1 Then
                            jshift = -1
                        ElseIf dirchoice = 2 Then
                            jshift = 1
                        End If
                    ElseIf dircur = 7 Then
                        If dirchoice = 1 Then
                            jshift = -1
                        ElseIf dirchoice = 2 Then
                            ishift = -1
                        End If
                    ElseIf dircur = 8 Then
                        If dirchoice = 1 Then
                            ishift = 1
                        ElseIf dirchoice = 2 Then
                            ishift = -1
                        End If
                    End If
                    If ishift <> 0 Then
                        For k = incr To TIME_END
                            RAFStatus(i, 1, k) = RAFStatus(i, 1, k) + ishift
                        Next k
                    ElseIf jshift <> 0 Then
                        For k = incr To TIME_END
                            RAFStatus(i, 2, k) = RAFStatus(i, 2, k) + jshift
                        Next k
                    End If
                End If
                If Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = "" Or Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = "." Or _
                Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = "|" Or Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = "+" Then
                    Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = RAF_ICON
                ElseIf Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = EF_ICON Then
                    fighterfound = False
                    B = 1
                    Do While fighterfound = False And B <= 20
                        If EFighters(B, 1, incr) = RAFStatus(i, 1, incr) And EFighters(B, 2, incr) = RAFStatus(i, 2, incr) Then
                            Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = ""
                            BG(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = "sx1"
                            If incr <= 24 Then BG(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr + 1) = "sx2"
                            If incr <= 23 Then BG(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr + 2) = "sx3"
                            If incr <= 22 And BG(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr + 3) = "" Then BG(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr + 3) = "blu"
                            For ja = 1 To 7
                                For ka = incr To TIME_END
                                    If ja <= 4 Then EFighters(B, ja, ka) = 0
                                    RAFStatus(i, ja, ka) = 0
                                Next ka
                            Next ja
                            fighterfound = True
                        End If
                        B = B + 1
                    Loop
                ElseIf Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = EB_ICON_RIGHT Or Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = EB_ICON_LEFT Then
                    bomberfound = False
                    B = 1
                    Do While bomberfound = False And B <= 20
                        If EBombers(B, 1, incr) = RAFStatus(i, 1, incr) And EBombers(B, 2, incr) = RAFStatus(i, 2, incr) And EBombers(B, 5, incr) = 0 Then
                            Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = ""
                            BG(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = "sx1"
                            If incr <= 24 Then BG(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr + 1) = "sx2"
                            If incr <= 23 Then BG(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr + 2) = "sx3"
                            If incr <= 22 And BG(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr + 3) = "" Then BG(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr + 3) = "blu"
                            For ja = 1 To 7
                                For ka = incr To TIME_END
                                    If ja <= 5 Then EBombers(B, ja, ka) = 0
                                    RAFStatus(i, ja, ka) = 0
                                Next ka
                            Next ja
                            bomberfound = True
                        End If
                        B = B + 1
                    Loop
                ElseIf Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = EB_ICON_RIGHT_DMG Or Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = EB_ICON_LEFT_DMG Then
                    bomberfound = False
                    B = 1
                    Do While bomberfound = False And B <= 20
                        If EBombers(B, 1, incr) = RAFStatus(i, 1, incr) And EFighters(B, 2, incr) = RAFStatus(i, 2, incr) And EBombers(B, 5, incr) = 1 Then
                            Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = ""
                            BG(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = "sx1"
                            If incr <= 24 Then BG(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr + 1) = "sx2"
                            If incr <= 23 Then BG(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr + 2) = "sx3"
                            If incr <= 22 And BG(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr + 3) = "" Then BG(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr + 3) = "blu"
                            For ja = 1 To 7
                                For ka = incr To TIME_END
                                    If ja <= 5 Then EBombers(B, ja, ka) = 0
                                    RAFStatus(i, ja, ka) = 0
                                Next ka
                            Next ja
                            bomberfound = True
                        End If
                        B = B + 1
                    Loop
                ElseIf Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = "," Or Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = "''" Or _
                Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = "`" Or Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = "'-" Then
                    firefound = False
                    B = 1
                    Do While firefound = False And B <= 20
                        If EFFire(B, 1, incr) = RAFStatus(i, 1, incr) And EFFire(B, 2, incr) = RAFStatus(i, 2, incr) Then
                            If RAFStatus(i, 7, incr) > 1 Then
                                For ka = incr To TIME_END
                                    RAFStatus(i, 7, ka) = RAFStatus(i, 7, ka) - 1
                                Next ka
                                Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = RAF_ICON
                                BG(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = "sx3"
                                If incr <= 24 And BG(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr + 1) = "" Then BG(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr + 1) = "blu"
                                For ja = 1 To 3
                                    For ka = incr To TIME_END
                                        EFFire(B, ja, ka) = 0
                                    Next ka
                                Next ja
                            Else
                                Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = ""
                                BG(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = "sx1"
                                If incr <= 24 Then BG(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr + 1) = "sx2"
                                If incr <= 23 Then BG(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr + 2) = "sx3"
                                If incr <= 22 And BG(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr + 3) = "" Then BG(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr + 3) = "blu"
                                For ja = 1 To 7
                                    For ka = incr To TIME_END
                                        If ja <= 3 Then EFFire(B, ja, ka) = 0
                                        RAFStatus(i, ja, ka) = 0
                                    Next ka
                                Next ja
                            End If
                            firefound = True
                        End If
                        B = B + 1
                    Loop
                ElseIf Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = "!" Then
                    firefound = False
                    B = 1
                    Do While firefound = False And B <= 20
                        If EBFire(B, 1, incr) = RAFStatus(i, 1, incr) And EBFire(B, 2, incr) = RAFStatus(i, 2, incr) Then
                            Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr) = ""
                            
                            Call ColorizerBomb(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), incr)

                            For ja = 1 To 7
                                For ka = incr To TIME_END
                                    If ja <= 2 Then EBFire(B, ja, ka) = 0
                                    RAFStatus(i, ja, ka) = 0
                                Next ka
                            Next ja
                            firefound = True
                        End If
                        B = B + 1
                    Loop
                End If
            End If
'RAFFire (Rockets - intermediate space)
            If RAFFire(i, 3, incr) = 2 Then
                inti = (RAFFire(i, 1, incr) - RAFFire(i, 1, incr - 1)) / 2 + RAFFire(i, 1, incr - 1)
                intj = (RAFFire(i, 2, incr) - RAFFire(i, 2, incr - 1)) / 2 + RAFFire(i, 2, incr - 1)
                If inti <= 2 Then
                    For yu = 1 To 4
                        For zu = incr To TIME_END
                            RAFFire(i, yu, zu) = 0
                        Next zu
                    Next yu
                ElseIf Board(inti, intj, incr) = EF_ICON Then
                    fighterfound = False
                    B = 1
                    Do While fighterfound = False And B <= 20
                        If EFighters(B, 1, incr) = inti And EFighters(B, 2, incr) = intj Then
                            Board(inti, intj, incr) = ""
                            BG(inti, intj, incr) = "sx1"
                            If incr <= 24 Then BG(inti, intj, incr + 1) = "sx2"
                            If incr <= 23 Then BG(inti, intj, incr + 2) = "sx3"
                            If incr <= 22 And BG(inti, intj, incr + 3) = "" Then BG(inti, intj, incr + 3) = "blu"
                            For ja = 1 To 4
                                For ka = incr To TIME_END
                                    EFighters(B, ja, ka) = 0
                                    RAFFire(i, ja, ka) = 0
                                Next ka
                            Next ja
                            fighterfound = True
                        End If
                        B = B + 1
                    Loop
                ElseIf Board(inti, intj, incr) = EB_ICON_RIGHT Or Board(inti, intj, incr) = EB_ICON_LEFT Or _
                Board(inti, intj, incr) = EB_ICON_RIGHT_DMG Or Board(inti, intj, incr) = EB_ICON_LEFT_DMG Then
                    bomberfound = False
                    B = 1
                    Do While bomberfound = False And B <= 20
                        If EBombers(B, 1, incr) = inti And EBombers(B, 2, incr) = intj Then
                            Board(inti, intj, incr) = ""
                            
                            Call ColorizerBomb(inti, intj, incr)
                            
                            For ja = 1 To 5
                                For ka = incr To TIME_END
                                    EBombers(B, ja, ka) = 0
                                    If ja <= 4 Then RAFFire(i, ja, ka) = 0
                                Next ka
                            Next ja
                            bomberfound = True
                        End If
                        B = B + 1
                    Loop
                ElseIf Board(inti, intj, incr) = "," Or Board(inti, intj, incr) = "''" Or Board(inti, intj, incr) = "`" Or Board(inti, intj, incr) = "'-" Then
                    firefound = False
                    B = 1
                    Do While firefound = False And B <= 20
                        If EFFire(B, 1, incr) = inti And EFFire(B, 2, incr) = intj Then
                            Board(inti, intj, incr) = ""
                            BG(inti, intj, incr) = "sx1"
                            If incr <= 24 Then BG(inti, intj, incr + 1) = "sx2"
                            If incr <= 23 Then BG(inti, intj, incr + 2) = "sx3"
                            If incr <= 22 And BG(inti, intj, incr + 3) = "" Then BG(inti, intj, incr + 3) = "blu"
                            For ja = 1 To 4
                                For ka = incr To TIME_END
                                    If ja <= 3 Then EFFire(B, ja, ka) = 0
                                    RAFFire(i, ja, ka) = 0
                                Next ka
                            Next ja
                            firefound = True
                        End If
                        B = B + 1
                    Loop
                ElseIf Board(inti, intj, incr) = "!" Then
                    firefound = False
                    B = 1
                    Do While firefound = False And B <= 20
                        If EBFire(B, 1, incr) = inti And EBFire(B, 2, incr) = intj Then
                            Board(inti, intj, incr) = ""
                            
                            Call ColorizerBomb(inti, intj, incr)

                            For ja = 1 To 4
                                For ka = incr To TIME_END
                                    If ja <= 2 Then EBFire(B, ja, ka) = 0
                                    RAFFire(i, ja, ka) = 0
                                Next ka
                            Next ja
                            firefound = True
                        End If
                        B = B + 1
                    Loop
                End If
            End If
'RAFFire
            If RAFFire(i, 3, incr) <> 0 Then
                If RAFFire(i, 1, incr) <= 2 Then
                    For yu = 1 To 4
                        For zu = incr To TIME_END
                            RAFFire(i, yu, zu) = 0
                        Next zu
                    Next yu
                ElseIf Board(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr) = "" And RAFFire(i, 3, incr) = 1 Then
                    Board(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr) = "."
                ElseIf Board(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr) = "" And RAFFire(i, 3, incr) = 2 Then
                    Board(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr) = "+"
                ElseIf Board(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr) = EF_ICON Then
                    fighterfound = False
                    B = 1
                    Do While fighterfound = False And B <= 20
                        If EFighters(B, 1, incr) = RAFFire(i, 1, incr) And _
                        EFighters(B, 2, incr) = RAFFire(i, 2, incr) Then
                            Board(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr) = ""
                            BG(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr) = "sx1"
                            If incr <= 24 Then BG(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr + 1) = "sx2"
                            If incr <= 23 Then BG(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr + 2) = "sx3"
                            If incr <= 22 And BG(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr + 3) = "" Then BG(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr + 3) = "blu"
                            For ja = 1 To 4
                                For ka = incr To TIME_END
                                    EFighters(B, ja, ka) = 0
                                    RAFFire(i, ja, ka) = 0
                                Next ka
                            Next ja
                            fighterfound = True
                        End If
                        B = B + 1
                    Loop
                ElseIf (Board(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr) = EB_ICON_RIGHT Or Board(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr) = EB_ICON_LEFT) And RAFFire(i, 3, incr) = 1 Then
                    bomberfound = False
                    B = 1
                    Do While bomberfound = False And B <= 20
                        If EBombers(B, 1, incr) = RAFFire(i, 1, incr) And _
                        EBombers(B, 2, incr) = RAFFire(i, 2, incr) Then
                            If EBombers(B, 4, incr) = 1 Then Board(EBombers(B, 1, incr), EBombers(B, 2, incr), incr) = EB_ICON_LEFT_DMG
                            If EBombers(B, 4, incr) = 2 Then Board(EBombers(B, 1, incr), EBombers(B, 2, incr), incr) = EB_ICON_RIGHT_DMG
                            BG(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr) = "sx3"
                            If incr <= 24 And BG(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr + 1) = "" Then BG(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr + 1) = "blu"
                            For ja = 1 To 4
                                For ka = incr To TIME_END
                                    If ja = 1 Then
                                        EBombers(B, 5, ka) = 1   'Limited to executing just once per bomber per ka, else ja = 1 is arbitrary
                                    End If
                                    RAFFire(i, ja, ka) = 0
                                Next ka
                            Next ja
                            bomberfound = True
                        End If
                        B = B + 1
                    Loop
                ElseIf Board(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr) = EB_ICON_RIGHT_DMG Or Board(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr) = EB_ICON_LEFT_DMG Or _
                ((Board(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr) = EB_ICON_RIGHT Or Board(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr) = EB_ICON_LEFT) And RAFFire(i, 3, incr) = 2) Then
                    bomberfound = False
                    B = 1
                    Do While bomberfound = False And B <= 20
                        If EBombers(B, 1, incr) = RAFFire(i, 1, incr) And EBombers(B, 2, incr) = RAFFire(i, 2, incr) Then
                            Board(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr) = ""
                            
                            Call ColorizerBomb(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr)
                            
                            For ja = 1 To 5
                                For ka = incr To TIME_END
                                    EBombers(B, ja, ka) = 0
                                    If ja <= 4 Then RAFFire(i, ja, ka) = 0
                                Next ka
                            Next ja
                            bomberfound = True
                        End If
                        B = B + 1
                    Loop
                ElseIf Board(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr) = "," Or Board(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr) = "''" Or _
                Board(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr) = "`" Or Board(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr) = "'-" Then
                    firefound = False
                    B = 1
                    Do While firefound = False And B <= 20
                        If EFFire(B, 1, incr) = RAFFire(i, 1, incr) And EFFire(B, 2, incr) = RAFFire(i, 2, incr) Then
                            Board(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr) = ""
                            If RAFFire(i, 3, incr) = 1 Then
                                BG(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr) = "sx3"
                                If incr <= 24 And BG(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr + 3) = "" Then BG(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr + 3) = "blu"
                            ElseIf RAFFire(i, 3, incr) = 2 Then
                                BG(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr) = "sx1"
                                If incr <= 24 Then BG(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr + 1) = "sx2"
                                If incr <= 23 Then BG(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr + 2) = "sx3"
                                If incr <= 22 And BG(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr + 3) = "" Then BG(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr + 3) = "blu"
                            End If
                            For ja = 1 To 4
                                For ka = incr To TIME_END
                                    If ja <= 3 Then EFFire(B, ja, ka) = 0
                                    RAFFire(i, ja, ka) = 0
                                Next ka
                            Next ja
                            firefound = True
                        End If
                        B = B + 1
                    Loop
                ElseIf Board(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr) = "!" Then
                    firefound = False
                    B = 1
                    Do While firefound = False And B <= 20
                        If EBFire(B, 1, incr) = RAFFire(i, 1, incr) And _
                        EBFire(B, 2, incr) = RAFFire(i, 2, incr) Then
                            Board(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr) = ""
                            
                            Call ColorizerBomb(RAFFire(i, 1, incr), RAFFire(i, 2, incr), incr)

                            For ja = 1 To 4
                                For ka = incr To TIME_END
                                    If ja <= 2 Then EBFire(B, ja, ka) = 0
                                    RAFFire(i, ja, ka) = 0
                                Next ka
                            Next ja
                            firefound = True
                        End If
                        B = B + 1
                    Loop
                End If
            End If
        Next i
'TurretFire
        Dim cleanTurretFire As TurretFire
        
        For i = 1 To BOARD_ROWS
            If Turrets(i, incr).Fire.Type = 1 Then           'Normal Turret Fire
                targetRow = Turrets(i, incr).Fire.Row
                targetCol = Turrets(i, incr).Fire.Column
                If Board(targetRow, targetCol, incr) = "" Then
                    Board(targetRow, targetCol, incr) = "."
                ElseIf Board(targetRow, targetCol, incr) = EF_ICON Then
                    fighterfound = False
                    B = 1
                    Do While fighterfound = False And B <= 20
                        If EFighters(B, 1, incr) = targetRow And _
                        EFighters(B, 2, incr) = targetCol Then
                            Board(targetRow, targetCol, incr) = ""
                            BG(targetRow, targetCol, incr) = "sx1"
                            If incr <= 24 Then BG(targetRow, targetCol, incr + 1) = "sx2"
                            If incr <= 23 Then BG(targetRow, targetCol, incr + 2) = "sx3"
                            If incr <= 22 And BG(targetRow, targetCol, incr + 3) = "" Then BG(targetRow, targetCol, incr + 3) = "blu"
                            For ka = incr To TIME_END
                                For ja = 1 To 4
                                    EFighters(B, ja, ka) = 0
                                Next ja
                                Turrets(i, ka).Fire = cleanTurretFire
                            Next ka
                            fighterfound = True
                        End If
                        B = B + 1
                    Loop
                ElseIf Board(targetRow, targetCol, incr) = EB_ICON_RIGHT Or _
                Board(targetRow, targetCol, incr) = EB_ICON_LEFT Then
                    bomberfound = False
                    B = 1
                    Do While bomberfound = False And B <= 20
                        If EBombers(B, 1, incr) = targetRow And _
                        EBombers(B, 2, incr) = targetCol Then
                            If EBombers(B, 4, incr) = 1 Then Board(EBombers(B, 1, incr), EBombers(B, 2, incr), incr) = EB_ICON_LEFT_DMG
                            If EBombers(B, 4, incr) = 2 Then Board(EBombers(B, 1, incr), EBombers(B, 2, incr), incr) = EB_ICON_RIGHT_DMG
                            BG(targetRow, targetCol, incr) = "sx3"
                            If incr <= 24 And BG(targetRow, targetCol, incr + 1) = "" Then BG(targetRow, targetCol, incr + 1) = "blu"
                            For ka = incr To TIME_END
                                EBombers(B, 5, ka) = 1
                                Turrets(i, ka).Fire = cleanTurretFire
                            Next ka
                            bomberfound = True
                        End If
                        B = B + 1
                    Loop
                ElseIf Board(targetRow, targetCol, incr) = EB_ICON_RIGHT_DMG Or _
                Board(targetRow, targetCol, incr) = EB_ICON_LEFT_DMG Then
                    bomberfound = False
                    B = 1
                    Do While bomberfound = False And B <= 20
                        If EBombers(B, 1, incr) = targetRow And _
                        EBombers(B, 2, incr) = targetCol Then
                            Board(targetRow, targetCol, incr) = ""
                            
                            Call ColorizerBomb(targetRow, targetCol, incr)
                            
                            For ka = incr To TIME_END
                                For ja = 1 To 5
                                    EBombers(B, ja, ka) = 0
                                Next ja
                                Turrets(i, ka).Fire = cleanTurretFire
                            Next ka
                            bomberfound = True
                        End If
                        B = B + 1
                    Loop
                ElseIf Board(targetRow, targetCol, incr) = "," Or Board(targetRow, targetCol, incr) = "''" Or _
                Board(targetRow, targetCol, incr) = "`" Or Board(targetRow, targetCol, incr) = "'-" Then
                    firefound = False
                    B = 1
                    Do While firefound = False And B <= 20
                        If EFFire(B, 1, incr) = targetRow And _
                        EFFire(B, 2, incr) = targetCol Then
                            Board(targetRow, targetCol, incr) = ""
                            BG(targetRow, targetCol, incr) = "sx3"
                            If incr <= TIME_END - 1 And BG(Turrets(i, incr).Fire.Row, targetCol, incr + 3) = "" Then BG(Turrets(i, incr).Fire.Row, targetCol, incr + 3) = "blu"
                            For ka = incr To TIME_END
                                For ja = 1 To 3
                                    EFFire(B, ja, ka) = 0
                                Next ja
                                Turrets(i, ka).Fire = cleanTurretFire
                            Next ka
                            firefound = True
                        End If
                        B = B + 1
                    Loop
                ElseIf Board(Turrets(i, incr).Fire.Row, targetCol, incr) = "!" Then
                    firefound = False
                    B = 1
                    Do While firefound = False And B <= 20
                        If EBFire(B, 1, incr) = Turrets(i, incr).Fire.Row And _
                        EBFire(B, 2, incr) = targetCol Then
                            Board(Turrets(i, incr).Fire.Row, targetCol, incr) = ""
                            
                            Call ColorizerBomb(Turrets(i, incr).Fire.Row, targetCol, incr)

                            For ka = incr To TIME_END
                                For ja = 1 To 2
                                    EBFire(B, ja, ka) = 0
                                Next ja
                                Turrets(i, ka).Fire = cleanTurretFire
                            Next ka
                            firefound = True
                        End If
                        B = B + 1
                    Loop
                End If
            End If
        Next i
'Structure Adjustments
        For j = 1 To BOARD_COLS
        
            Call ReconcileStructureDamage(TURRET_ROW, CInt(j), CInt(incr))
            Call ReconcileStructureDamage(STRUCT_ROW, CInt(j), CInt(incr))
            
        Next j
    Next incr
    
    'Avoid redundant formatting cues
    For i = 1 To BOARD_ROWS
        For j = 1 To BOARD_COLS
            For incr = TIME_END To 2 Step -1
                If BG(i, j, incr) = BG(i, j, incr - 1) Then
                    BG(i, j, incr) = ""
                End If
            Next incr
        Next j
    Next i
    
    'CommenceMovement = MsgBox("Commencing movement", vbOKOnly)
    NumEFRemain1 = NumEFRemain
    NumEBRemain1 = NumEBRemain
    NumTurRemain1 = NumTurRemain
    NumRAFRemain1 = NumRAFRemain
    Dim turnBonus As Long                                   'turnBonus used here to calculate points every turn from shooting down aircraft/losing turrets
    Dim turnBonusPrev As Long
    Dim turnBonusText As String

    Application.ScreenUpdating = True

    For incr = 2 To TIME_END
        Application.ScreenUpdating = False
    
        Call AnimateBoard(incr)
        
        Application.ScreenUpdating = True
        
        DoEvents
        
        'Evaluate status of aircraft, turrets, and buildings
        NumEFRemain = 0
        For i = 1 To 20
            If EFighters(i, 1, incr) <> 0 Then
                NumEFRemain = NumEFRemain + 1
            End If
        Next i
        If Cells(10, 24) <> "" Then Cells(10, 25) = NumEFRemain
        
        NumEBRemain = 0
        For i = 1 To 20
            If EBombers(i, 1, incr) <> 0 Then
                NumEBRemain = NumEBRemain + 1
            End If
        Next i
        If Cells(11, 24) <> "" Then Cells(11, 25) = NumEBRemain
        
        NumTurRemain = 0
        For j = 1 To BOARD_COLS
            If Turrets(j, incr).Health < 4 And Turrets(j, 1).Icon <> "" Then
                NumTurRemain = NumTurRemain + 1
            End If
        Next j
        Cells(13, 25) = NumTurRemain
        
        NumRAFRemain = 0
        For i = 1 To 4
            If RAFStatus(i, 1, incr) > 0 Then
                NumRAFRemain = NumRAFRemain + 1
            End If
        Next i
        Cells(20, 25) = NumRAFRemain
        
        NumEACLess = NumEFRemain1 + NumEBRemain1 - NumEFRemain - NumEBRemain                   'Total number of enemy aircraft destroyed this turn
        
        turnBonusPrev = turnBonus
        
        turnBonus = (NumEFRemain1 - NumEFRemain) * 5 * (2 ^ NumEACLess) * Wave _
        + (NumEBRemain1 - NumEBRemain) * 10 * (2 ^ NumEACLess) * Wave _
        - (NumTurRemain1 - NumTurRemain) * 500 _
        - (NumRAFRemain1 - NumRAFRemain) * 100                                              'Point calculation for aircraft destruction/turret loss
        
        GameScore = GameScore + turnBonus - turnBonusPrev                                                            'Update total turnBonus
        turnBonusText = ""
        
        If NumEFRemain1 - NumEFRemain > 0 Then
            turnBonusText = "+" & (NumEFRemain1 - NumEFRemain) * 5 * (2 ^ NumEACLess) * Wave & " Fighter "
            If NumEFRemain1 - NumEFRemain > 1 Then
                turnBonusText = turnBonusText & "x" & (NumEFRemain1 - NumEFRemain) & " "
            End If
        End If
        If NumEBRemain1 - NumEBRemain > 0 Then
            turnBonusText = turnBonusText & "+" & (NumEBRemain1 - NumEBRemain) * 10 * (2 ^ NumEACLess) * Wave & " Bomber "
            If NumEBRemain1 - NumEBRemain > 1 Then
                turnBonusText = turnBonusText & "x" & (NumEBRemain1 - NumEBRemain) & " "
            End If
        End If
        If NumTurRemain1 - NumTurRemain > 0 Then
            turnBonusText = turnBonusText & "-" & (NumTurRemain1 - NumTurRemain) * 500 & " Turret "
            If NumTurRemain1 - NumTurRemain > 1 Then
                turnBonusText = turnBonusText & "x" & (NumTurRemain1 - NumTurRemain) & " "
            End If
        End If
        If NumRAFRemain1 - NumRAFRemain > 0 Then
            turnBonusText = turnBonusText & "-" & (NumRAFRemain1 - NumRAFRemain) * 100 & " RAF Fighter "
            If NumRAFRemain1 - NumRAFRemain > 1 Then
                turnBonusText = turnBonusText & "x" & (NumRAFRemain1 - NumRAFRemain) & " "
            End If
        End If
        
        Cells(23, 24) = GameScore
        Cells(23, 26) = turnBonusText
        
        Cells(14, 25) = (4 - StructureStatus(ComC, incr)) / 4
        
        If StructureStatus(RepairServ1, incr) >= 4 And StructureStatus(RepairServ1, incr - 1) < 4 Then
            NumRepairServRemain = NumRepairServRemain - 1
        End If
        If StructureStatus(RepairServ2, incr) >= 4 And StructureStatus(RepairServ2, incr - 1) < 4 Then
            NumRepairServRemain = NumRepairServRemain - 1
        End If
        
        Cells(15, 25) = NumRepairServRemain & " of 2"
        
        If StructureStatus(Bunker1, incr) >= 4 And StructureStatus(Bunker1, incr - 1) < 4 Then
            NumBunkerRemain = NumBunkerRemain - 1
        End If
        If StructureStatus(Bunker2, incr) >= 4 And StructureStatus(Bunker2, incr - 1) < 4 Then
            NumBunkerRemain = NumBunkerRemain - 1
        End If
        
        Cells(16, 25) = NumBunkerRemain & " of 2"
        
        AirfieldStatus = 0
        For i = 0 To 3
            AirfieldStatus = AirfieldStatus + StructureStatus(AirfieldLeft + i, incr)
        Next i
        Cells(17, 25) = (16 - AirfieldStatus) / 16
        
        CityStatus = 0
        For i = 1 To 12
            If StructureStatus(CityStruct(i), incr) < 4 Then CityStatus = CityStatus + 1
        Next i
        Cells(18, 25) = CityStatus & " of 12"
        
        
        Application.ScreenUpdating = True
        
        Cells(1, 23) = incr / TIME_END
        
        delayStart = Timer
        
        Do While Timer < delayStart + ANIMATION_RATE
        Loop

        If (NumEFRemain = 0 And NumEBRemain = 0) Or (NumTurRemain = 0 And NumRAFRemain = 0) Then             'End turn immediately if all aircraft destroyed or all turrets lost
            For j = 1 To BOARD_COLS
                StructureStatus(j, TIME_END) = StructureStatus(j, incr)
                Turrets(j, TIME_END) = Turrets(j, incr)
                For i = 1 To 2
                    Board(i, j, TIME_END) = Board(i, j, incr)
                Next i
                If Turrets(j, 1).Icon <> "" Then
                    Turrets(j, 1).Icon = Board(TURRET_ROW, j, incr)
                Else
                    Board(TURRET_ROW, j, TIME_END) = ""
                End If
            Next j
            For i = 1 To 4
                If RAFStatus(i, 1, incr) <> 0 Then
                    If RAFStatus(i, 1, TIME_END) <> 0 Then Board(RAFStatus(i, 1, TIME_END), RAFStatus(i, 2, TIME_END), TIME_END) = ""              'Relocate RAF from where they were going to end up to where they currently are
                    Board(RAFStatus(i, 1, incr), RAFStatus(i, 2, incr), TIME_END) = RAF_ICON
                    For j = 1 To 7
                        RAFStatus(i, j, TIME_END) = RAFStatus(i, j, incr)
                    Next j
                End If
            Next i
            incr = TIME_END
        End If
    Next incr
    
    TurnCount = TurnCount + 1
    
    
    'Reinitialize and clear time-dependent arrays
    For i = 1 To EB_MAX_COUNT
        For B = 1 To 5
            If B <= 4 Then EFighters(i, B, 1) = EFighters(i, B, TIME_END)
            EBombers(i, B, 1) = EBombers(i, B, TIME_END)
            For k = 1 To TIME_END
                If B <= 3 Then EFFire(i, B, k) = 0
                If B <= 2 Then EBFire(i, B, k) = 0
                If k >= 2 Then
                    If B <= 4 Then EFighters(i, B, k) = 0
                    EBombers(i, B, k) = 0
                End If
            Next k
        Next B
    Next i
            
    For i = 1 To RAF_MAX_COUNT
        For B = 1 To 7
            RAFStatus(i, B, 1) = RAFStatus(i, B, TIME_END)
            For k = 1 To TIME_END
                If B <= 4 Then RAFFire(i, B, k) = 0
                If k >= 2 Then RAFStatus(i, B, k) = 0
            Next k
        Next B
    Next i
    
    For j = 1 To BOARD_COLS
        StructureStatus(j, 1) = StructureStatus(j, TIME_END)
        Turrets(j, 1) = Turrets(j, TIME_END)
        For i = 1 To BOARD_ROWS
           Board(i, j, 1) = Board(i, j, TIME_END)
        Next i
        For k = 1 To TIME_END
            Turrets(j, k).Fire = cleanTurretFire
            If k > 1 Then
                Turrets(j, k).Icon = ""
                Turrets(j, k).Health = 0
                StructureStatus(j, k) = 0
                For i = 1 To BOARD_ROWS
                    Board(i, j, k) = ""
                    BG(i, j, k) = ""
                Next i
            End If
        Next k
    Next j
    
    Cells(1, 23) = ""                'Clear cell with animation completion percentage
    
    If NumEFRemain = 0 And NumEBRemain = 0 And (NumTurRemain > 0 Or NumRAFRemain > 0) Then   'Wave Complete. Note: Wave not complete if all turrets are destroyed simultaneously
        For j = 5 To 19
            Cells(6, j) = ""
            With Range(Cells(6, j).Address()).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next j
        Cells(6, 12) = "Night Complete! All enemy aircraft destroyed."
        Dim LevelBonus As Long
        
        TickIncrement (100)
        
        turnBonus = 0                                                                       'Turn Count Bonus
        If TurnCount < 20 Then
            turnBonus = Int(2000 / TurnCount)
        End If
        turnBonus = Int(turnBonus * Wave)
        For j = 5 To 19
            Cells(7, j) = ""
            With Range(Cells(7, j).Address()).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next j
        Cells(7, 12) = TurnCount & " Turns" & "         " & turnBonus & " Points"
        LevelBonus = LevelBonus + turnBonus
        
        TickIncrement (100)
        
        CityStatus = 0                                                                             'City Status Bonus
        For i = 1 To 12
            If StructureStatus(CityStruct(i), 1) < 4 Then CityStatus = CityStatus + 1
        Next i
        Dim CityBonus As Long
        CityBonus = CityStatus * 100 * Wave
        For j = 5 To 19
            Cells(8, j) = ""
            With Range(Cells(8, j).Address()).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next j
        Cells(8, 12) = CityStatus & " Surviving City Structures" & "        " & CityBonus & " Points"
        LevelBonus = LevelBonus + CityBonus
        
        TickIncrement (100)
        
        Dim ShotBonus As Long                                                                       'Shot Bonus
        If ShotCount <= 10 Then
            ShotBonus = Int(350 * (9 + Wave) / (ShotCount + 1))
        ElseIf ShotCount < 40 Then
            ShotBonus = Int(25 * (40 - ShotCount) / 3)
        End If
        ShotBonus = Int(ShotBonus * (1 + Wave))
        For j = 5 To 19
            Cells(9, j) = ""
            With Range(Cells(9, j).Address()).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next j
        Cells(9, 12) = ShotCount & " Shots" & "         " & ShotBonus & " Points"
        LevelBonus = LevelBonus + ShotBonus
        
        TickIncrement (100)
        
        Dim WaveBonus As Long
        WaveBonus = Wave * 1200                                                           'Wave Bonus
        For j = 5 To 19
            Cells(10, j) = ""
            With Range(Cells(10, j).Address()).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next j
        Cells(10, 12) = "Night " & Wave & " Completion         " & WaveBonus & " Points"
        LevelBonus = LevelBonus + WaveBonus
        
        TickIncrement (100)
        
        For j = 5 To 19                                                                'Total Level Bonus
            Cells(11, j) = ""
            With Range(Cells(11, j).Address()).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next j
        Cells(11, 12) = "Total Level Bonus" & "        " & LevelBonus & " Points"
        GameScore = GameScore + LevelBonus
        Cells(23, 24) = GameScore
        
        If Wave < 10 Then
            If (StructureStatus(AirfieldLeft, 1) < 4 Or StructureStatus(AirfieldLeft + 1, 1) < 4 Or _
            StructureStatus(AirfieldLeft + 2, 1) < 4 Or StructureStatus(AirfieldLeft + 3, 1) < 4) Then
                Worksheets("London Siege").RAFShopPortal.Visible = True
                Worksheets("London Siege").RAFShopPortal.Enabled = True
            Else
                Worksheets("London Siege").RAFShopPortal.Visible = False
            End If
            If NumRepairServRemain > 0 Then
                Worksheets("London Siege").RepairShopPortal.Visible = True
                Worksheets("London Siege").RepairShopPortal.Enabled = True
            Else
                Worksheets("London Siege").RepairShopPortal.Visible = False
            End If
            If NumBunkerRemain > 0 Then
                Worksheets("London Siege").AmmoShopPortal.Visible = True
                Worksheets("London Siege").AmmoShopPortal.Enabled = True
            Else
                Worksheets("London Siege").AmmoShopPortal.Visible = False
            End If
            If NumRAFRemain > 0 Then
                landingspots = 0
                For i = 0 To 3
                    If StructureStatus(AirfieldLeft + i, 1) < 4 Then
                        landingspots = landingspots + 1
                    End If
                Next i
                If landingspots < NumRAFRemain And landingspots > 0 And StructureStatus(ComC, 1) < 4 Then
                    RAFCull = True
                    Cells(2, 37) = "Not enough room for all aircraft to land"
                    Cells(3, 37) = "Choose " & landingspots & " for repairs and rearmament"
                    Cells(7, 30) = "Health"
                    Cells(8, 30) = "Rockets"
                    Cells(9, 30) = "Altitude"
                    For i = 1 To 4
                        If RAFStatus(i, 1, 1) <> 0 Then
                            Cells(5, 32 + 2 * i) = RAF_ICON
                            Cells(7, 32 + 2 * i) = Int(100 / RAFStatus(i, 7, 1))
                            Range(Cells(7, 32 + 2 * i).Address).Font.Size = 5
                            Cells(8, 32 + 2 * i) = RAFStatus(i, 5, 1)
                            Cells(9, 32 + 2 * i) = RAFStatus(i, 1, 1) & "00 m"
                            Range(Cells(9, 32 + 2 * i).Address).Font.Size = 7
                        End If
                    Next i
                    If Cells(5, 34) <> "" Then Worksheets("London Siege").BT08.Visible = True
                    If Cells(5, 36) <> "" Then Worksheets("London Siege").BT10.Visible = True
                    If Cells(5, 38) <> "" Then Worksheets("London Siege").BT12.Visible = True
                    If Cells(5, 40) <> "" Then Worksheets("London Siege").BT14.Visible = True
                    Worksheets("London Siege").AmmoShopPortal.Enabled = False
                    Worksheets("London Siege").RAFShopPortal.Enabled = False
                    Worksheets("London Siege").RepairShopPortal.Enabled = False
                    RAFPos = landingspots
                Else
                    For i = 1 To 4
                        If RAFStatus(i, 1, 1) <> 0 Then
                            For heal = 0 To 3            'Place RAF in healthiest spaces first
                                For j = 0 To 3
                                    If Board(1, AirfieldLeft + j, 1) = "     " And StructureStatus(AirfieldLeft + j, 1) = heal Then
                                        Board(RAFStatus(i, 1, 1), RAFStatus(i, 2, 1), 1) = ""
                                        If Cells(GRID_BOT_BRDR - RAFStatus(i, 1, 1), 1 + RAFStatus(i, 2, 1)) = RAF_ICON Then Cells(GRID_BOT_BRDR - RAFStatus(i, 1, 1), 1 + RAFStatus(i, 2, 1)) = ""
                                        RAFStatus(i, 1, 1) = 1
                                        RAFStatus(i, 2, 1) = AirfieldLeft + j
                                        RAFStatus(i, 3, 1) = 0
                                        RAFStatus(i, 4, 1) = 0
                                        RAFStatus(i, 6, 1) = 1
                                        Board(1, AirfieldLeft + j, 1) = RAF_ICON
                                        Cells(22, 1 + AirfieldLeft + j) = RAF_ICON
                                        j = 3
                                        heal = 3
                                    End If
                                Next j
                            Next heal
                        End If
                    Next i
                End If
            End If
        End If
        
        Wave = Wave + 1
        Worksheets("London Siege").NextWave.Caption = "Night " & Wave
        Worksheets("London Siege").NextWave.Enabled = True
        Worksheets("London Siege").NextTurn.Enabled = False
        
    ElseIf NumTurRemain = 0 And NumRAFRemain = 0 Then              'All defenses lost, game over

        For j = 7 To 17
            Cells(6, j) = ""
            With Range(Cells(6, j).Address()).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next j
        Cells(6, 12) = "All defenses down. London is lost."
        
        TickIncrement (100)
        
        For j = 7 To 17
            Cells(7, j) = ""
            With Range(Cells(7, j).Address()).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Next j
        Cells(7, 12) = "Final Score     " & GameScore & " Points"
        
        TickIncrement (100)
        
        Worksheets("London Siege").NextTurn.Enabled = False
        Worksheets("London Siege").StartGame.Caption = "Play Again"
        InProcess = False
        
    Else
        Worksheets("London Siege").NextTurn.Enabled = True              'Re-enabling of Next Turn button amidst level
        
    End If
    
    Application.ScreenUpdating = False
    If (NumEFRemain <> 0 Or NumEBRemain <> 0) And (NumTurRemain <> 0 Or NumRAFRemain <> 0) Then
        Dim structCode As String
        Dim turretCode As String
        For j = 1 To BOARD_COLS            'Remove lingering background and replace structure status color
        
            structCode = GetStructureBGCode(STRUCT_ROW, StructureStatus(j, 1))
            Call RenderColor(Cells(GRID_BOT_BRDR - STRUCT_ROW, GRID_LEFT_BRDR + j), structCode)
            
            turretCode = GetStructureBGCode(TURRET_ROW, Turrets(j, 1).Health)
            Call RenderColor(Cells(GRID_BOT_BRDR - TURRET_ROW, GRID_LEFT_BRDR + j), turretCode)
            
            For i = TURRET_ROW + 1 To BOARD_ROWS
                Call RenderColor(Cells(GRID_BOT_BRDR - i, GRID_LEFT_BRDR + j), "blu")
            Next i
        Next j
    End If
    If NumTurRemain = 0 Then
        Range(Cells(1, 24).Address()).Select
    ElseIf NumEFRemain = 0 And NumEBRemain = 0 Then
        Range(Cells(3, 24).Address()).Select
    Else
        Range(Cells(5, 24).Address()).Select
    End If
    Application.ScreenUpdating = True
    
    Worksheets("London Siege").Unprotect
    Worksheets("London Siege").StartGame.Enabled = True
    Worksheets("London Siege").QuitGame.Enabled = True
    If RAFCull = True Then Worksheets("London Siege").NextWave.Enabled = False

End Sub

Sub ReconcileStructureDamage(i As Integer, j As Integer, incr As Integer)
    
    Dim targetHealth As Integer
    If i = STRUCT_ROW Then targetHealth = StructureStatus(j, incr)
    If i = TURRET_ROW Then targetHealth = Turrets(j, incr).Health
    
    If BG(i, j, incr) = "" Then
        BG(i, j, incr) = GetStructureBGCode(i, targetHealth)
    End If
    
    If targetHealth >= 4 And incr >= 2 Then
        Dim prevTargetHealth As Integer
        If i = STRUCT_ROW Then prevTargetHealth = StructureStatus(j, incr - 1)
        If i = TURRET_ROW Then prevTargetHealth = Turrets(j, incr - 1).Health
        If prevTargetHealth < 4 Then
            'Grounded RAF gets destroyed if airfield section it's parked on is destroyed
            If i = STRUCT_ROW And Board(i, j, incr) = RAF_ICON Then
                fighterfound = False
                B = 1
                Do While fighterfound = False And B <= 4
                    If RAFStatus(B, 1, incr) = i And RAFStatus(B, 2, incr) = j Then
                        For ja = 1 To 7
                            For ka = incr To TIME_END
                                RAFStatus(B, ja, ka) = 0
                            Next ka
                        Next ja
                        fighterfound = True
                    End If
                B = B + 1
                Loop
                Board(i, j, incr) = ""
            End If
            
            Dim newRuins As String
            
            For RuinsAdd = 1 To 4
                Randomize
                RuinsType = Int(7 * Rnd + 1)
                If RuinsType <= 3 Then
                    newRuins = newRuins & "."
                ElseIf RuinsType = 4 Then
                    newRuins = newRuins & ":"
                ElseIf RuinsType <= 6 Then
                    newRuins = newRuins & ","
                Else
                    newRuins = newRuins & ";"
                End If
            Next RuinsAdd
            
            Board(i, j, incr) = newRuins
            
            If i = TURRET_ROW Then Turrets(j, incr).Icon = newRuins
        End If
    End If
    
End Sub

Sub AmmoShopPortal()
    

End Sub

Sub BT01()

End Sub

Sub BT02()

End Sub

Sub BT03()

End Sub

Sub BT04()

End Sub

Sub BT05()

End Sub

Sub BT06()

End Sub

Sub BT07()

End Sub

Sub BT08()
    If RAFCull = True Then
        Application.ScreenUpdating = False
        RAFPos = RAFPos - 1            'Counts number of spots left to fill
        For i = 0 To 3
            vacancy = True
            For j = 1 To 4
                If StructureStatus(AirfieldLeft + i, 1) < 4 And RAFStatus(j, 2, 1) = AirfieldLeft + i Then
                    j = 4
                    vacancy = False
                End If
            Next j
            If vacancy = True Then
                Board(1, AirfieldLeft + i, 1) = RAF_ICON
                Cells(GRID_BOT_BRDR - RAFStatus(1, 1, 1), 1 + RAFStatus(1, 2, 1)) = ""
                RAFStatus(1, 1, 1) = 1
                RAFStatus(1, 2, 1) = AirfieldLeft + i
                RAFStatus(1, 6, 1) = 1
                Cells(22, 1 + AirfieldLeft + i) = RAF_ICON
                Cells(7, 34) = "100%"
                Cells(9, 34) = "NA"
                i = 3
            End If
        Next i
        Worksheets("London Siege").BT08.Visible = False
        If RAFPos = 0 Then
            RAFCull = False
            For i = 1 To 13
                For j = 1 To 23
                    Cells(i, 25 + j) = ""
                    With Range(Cells(i, 25 + j).Address()).Font
                        .Underline = xlUnderlineStyleNone
                        .Size = 11
                    End With
                    With Range(Cells(i, 25 + j).Address()).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorLight1
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                Next j
            Next i
            Worksheets("London Siege").BT10.Visible = False
            Worksheets("London Siege").BT12.Visible = False
            Worksheets("London Siege").BT14.Visible = False
            Worksheets("London Siege").NextWave.Enabled = True
            Worksheets("London Siege").AmmoShopPortal.Enabled = True
            Worksheets("London Siege").RepairShopPortal.Enabled = True
            Worksheets("London Siege").RAFShopPortal.Enabled = True
        End If
        Application.ScreenUpdating = True
    ElseIf RAFPreLaunch = True Then
        Application.ScreenUpdating = False
        Cells(2, 37) = "Select Starting Location"
        For i = 1 To 3
            Cells(5, 34 + 2 * i) = ""
            Cells(7, 34 + 2 * i) = ""
            Cells(8, 34 + 2 * i) = ""
        Next i
        Worksheets("London Siege").BT08.Visible = False
        Worksheets("London Siege").BT10.Visible = False
        Worksheets("London Siege").BT12.Visible = False
        Worksheets("London Siege").BT14.Visible = False
        Worksheets("London Siege").NextWavePreLaunchDone.Visible = False
        Worksheets("London Siege").LaunchRAF.Visible = True
        Worksheets("London Siege").PreLaunchBack.Visible = True
        For i = 1 To 4
            If RAFStatus(i, 2, 1) = AirfieldLeft Then
                RAFPos = i
                i = 4
            End If
        Next i
        Range("A1").Select
        Application.ScreenUpdating = True
    ElseIf RAFShopActive = True Then
        If Worksheets("London Siege").BT08.Value = True Then
            PurchaseCost = PurchaseCost + 2000
        ElseIf Worksheets("London Siege").BT08.Value = False And PurchaseCost > 0 Then
            PurchaseCost = PurchaseCost - 2000
        End If

        RAFShopVisibility
    End If
                                    
End Sub

Sub BT09()

End Sub

Sub BT10()
    If RAFCull = True Then
        Application.ScreenUpdating = False
        RAFPos = RAFPos - 1
        For i = 0 To 3
            vacancy = True
            For j = 1 To 4
                If StructureStatus(AirfieldLeft + i, 1) < 4 And RAFStatus(j, 2, 1) = AirfieldLeft + i Then
                    j = 4
                    vacancy = False
                End If
            Next j
            If vacancy = True Then
                Board(1, AirfieldLeft + i, 1) = RAF_ICON
                Cells(GRID_BOT_BRDR - RAFStatus(2, 1, 1), 1 + RAFStatus(2, 2, 1)) = ""
                RAFStatus(2, 1, 1) = 1
                RAFStatus(2, 2, 1) = AirfieldLeft + i
                RAFStatus(2, 6, 1) = 1
                Cells(22, 1 + AirfieldLeft + i) = RAF_ICON
                Cells(7, 36) = "100%"
                Cells(9, 36) = "NA"
                i = 3
            End If
        Next i
        Worksheets("London Siege").BT10.Visible = False
        If RAFPos = 0 Then
            RAFCull = False
            For i = 1 To 13
                For j = 1 To 23
                    Cells(i, 25 + j) = ""
                    With Range(Cells(i, 25 + j).Address()).Font
                        .Underline = xlUnderlineStyleNone
                        .Size = 11
                    End With
                    With Range(Cells(i, 25 + j).Address()).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorLight1
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                Next j
            Next i
            Worksheets("London Siege").BT08.Visible = False
            Worksheets("London Siege").BT12.Visible = False
            Worksheets("London Siege").BT14.Visible = False
            Worksheets("London Siege").NextWave.Enabled = True
            Worksheets("London Siege").AmmoShopPortal.Enabled = True
            Worksheets("London Siege").RepairShopPortal.Enabled = True
            Worksheets("London Siege").RAFShopPortal.Enabled = True
        End If
        Application.ScreenUpdating = True
    ElseIf RAFPreLaunch = True Then
        Application.ScreenUpdating = False
        Cells(2, 37) = "Select Starting Location"
        For i = 1 To 4
            If i <> 2 Then
                Cells(5, 32 + 2 * i) = ""
                Cells(7, 32 + 2 * i) = ""
                Cells(8, 32 + 2 * i) = ""
            End If
        Next i
        Worksheets("London Siege").BT08.Visible = False
        Worksheets("London Siege").BT10.Visible = False
        Worksheets("London Siege").BT12.Visible = False
        Worksheets("London Siege").BT14.Visible = False
        Worksheets("London Siege").NextWavePreLaunchDone.Visible = False
        Worksheets("London Siege").LaunchRAF.Visible = True
        Worksheets("London Siege").PreLaunchBack.Visible = True
        For i = 1 To 4
            If RAFStatus(i, 2, 1) = AirfieldLeft + 1 Then
                RAFPos = i
                i = 4
            End If
        Next i
        Range("A1").Select
        Application.ScreenUpdating = True
    ElseIf RAFShopActive = True Then
        If Worksheets("London Siege").BT10.Value = True Then
            PurchaseCost = PurchaseCost + 2000
        ElseIf Worksheets("London Siege").BT10.Value = False And PurchaseCost > 0 Then
            PurchaseCost = PurchaseCost - 2000
        End If

        RAFShopVisibility
    End If
End Sub

Sub BT11()

End Sub

Sub BT12()
    If RAFCull = True Then
    Application.ScreenUpdating = False
        RAFPos = RAFPos - 1
        For i = 0 To 3
            vacancy = True
            For j = 1 To 4
                If StructureStatus(AirfieldLeft + i, 1) < 4 And RAFStatus(j, 2, 1) = AirfieldLeft + i Then
                    j = 4
                    vacancy = False
                End If
            Next j
            If vacancy = True Then
                Board(1, AirfieldLeft + i, 1) = RAF_ICON
                Cells(GRID_BOT_BRDR - RAFStatus(3, 1, 1), 1 + RAFStatus(3, 2, 1)) = ""
                RAFStatus(3, 1, 1) = 1
                RAFStatus(3, 2, 1) = AirfieldLeft + i
                RAFStatus(3, 6, 1) = 1
                Cells(22, 1 + AirfieldLeft + i) = RAF_ICON
                Cells(7, 38) = "100%"
                Cells(9, 38) = "NA"
                i = 3
            End If
        Next i
        Worksheets("London Siege").BT12.Visible = False
        If RAFPos = 0 Then
            RAFCull = False
            For i = 1 To 13
                For j = 1 To 23
                    Cells(i, 25 + j) = ""
                    With Range(Cells(i, 25 + j).Address()).Font
                        .Underline = xlUnderlineStyleNone
                        .Size = 11
                    End With
                    With Range(Cells(i, 25 + j).Address()).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorLight1
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                Next j
            Next i
            Worksheets("London Siege").BT08.Visible = False
            Worksheets("London Siege").BT10.Visible = False
            Worksheets("London Siege").BT14.Visible = False
            Worksheets("London Siege").NextWave.Enabled = True
            Worksheets("London Siege").AmmoShopPortal.Enabled = True
            Worksheets("London Siege").RepairShopPortal.Enabled = True
            Worksheets("London Siege").RAFShopPortal.Enabled = True
        End If
        Application.ScreenUpdating = True
    ElseIf RAFPreLaunch = True Then
        Application.ScreenUpdating = False
        Cells(2, 37) = "Select Starting Location"
        For i = 1 To 4
            If i <> 3 Then
                Cells(5, 32 + 2 * i) = ""
                Cells(7, 32 + 2 * i) = ""
                Cells(8, 32 + 2 * i) = ""
            End If
        Next i
        Worksheets("London Siege").BT08.Visible = False
        Worksheets("London Siege").BT10.Visible = False
        Worksheets("London Siege").BT12.Visible = False
        Worksheets("London Siege").BT14.Visible = False
        Worksheets("London Siege").NextWavePreLaunchDone.Visible = False
        Worksheets("London Siege").LaunchRAF.Visible = True
        Worksheets("London Siege").PreLaunchBack.Visible = True
        For i = 1 To 4
            If RAFStatus(i, 2, 1) = AirfieldLeft + 2 Then
                RAFPos = i
                i = 4
            End If
        Next i
        Range("A1").Select
        Application.ScreenUpdating = True
    ElseIf RAFShopActive = True Then
        If Worksheets("London Siege").BT12.Value = True Then
            PurchaseCost = PurchaseCost + 2000
        ElseIf Worksheets("London Siege").BT12.Value = False And PurchaseCost > 0 Then
            PurchaseCost = PurchaseCost - 2000
        End If

        RAFShopVisibility
    End If
End Sub

Sub BT13()

End Sub

Sub BT14()
    If RAFCull = True Then
        Application.ScreenUpdating = False
        RAFPos = RAFPos - 1
        For i = 0 To 3
            vacancy = True
            For j = 1 To 4
                If StructureStatus(AirfieldLeft + i, 1) < 4 And RAFStatus(j, 2, 1) = AirfieldLeft + i Then
                    j = 4
                    vacancy = False
                End If
            Next j
            If vacancy = True Then
                Board(1, AirfieldLeft + i, 1) = RAF_ICON
                Cells(GRID_BOT_BRDR - RAFStatus(3, 1, 1), 1 + RAFStatus(3, 2, 1)) = ""
                RAFStatus(4, 1, 1) = 1
                RAFStatus(4, 2, 1) = AirfieldLeft + i
                RAFStatus(4, 6, 1) = 1
                Cells(22, 1 + AirfieldLeft + i) = RAF_ICON
                Cells(7, 40) = "100%"
                Cells(9, 40) = "NA"
                i = 3
            End If
        Next i
        Worksheets("London Siege").BT14.Visible = False
        If RAFPos = 0 Then
            RAFCull = False
            For i = 1 To 13
                For j = 1 To 23
                    Cells(i, 25 + j) = ""
                    With Range(Cells(i, 25 + j).Address()).Font
                        .Underline = xlUnderlineStyleNone
                        .Size = 11
                    End With
                    With Range(Cells(i, 25 + j).Address()).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorLight1
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                Next j
            Next i
            Worksheets("London Siege").BT08.Visible = False
            Worksheets("London Siege").BT10.Visible = False
            Worksheets("London Siege").BT12.Visible = False
            Worksheets("London Siege").NextWave.Enabled = True
            Worksheets("London Siege").AmmoShopPortal.Enabled = True
            Worksheets("London Siege").RepairShopPortal.Enabled = True
            Worksheets("London Siege").RAFShopPortal.Enabled = True
        End If
        Application.ScreenUpdating = True
    ElseIf RAFPreLaunch = True Then
        Application.ScreenUpdating = False
        Cells(2, 37) = "Select Starting Location"
        For i = 1 To 3
            Cells(5, 32 + 2 * i) = ""
            Cells(7, 32 + 2 * i) = ""
            Cells(8, 32 + 2 * i) = ""
        Next i
        Worksheets("London Siege").BT08.Visible = False
        Worksheets("London Siege").BT10.Visible = False
        Worksheets("London Siege").BT12.Visible = False
        Worksheets("London Siege").BT14.Visible = False
        Worksheets("London Siege").NextWavePreLaunchDone.Visible = False
        Worksheets("London Siege").LaunchRAF.Visible = True
        Worksheets("London Siege").PreLaunchBack.Visible = True
        For i = 1 To 4
            If RAFStatus(i, 2, 1) = AirfieldLeft + 3 Then
                RAFPos = i
                i = 4
            End If
        Next i
        Range("A1").Select
        Application.ScreenUpdating = True
    ElseIf RAFShopActive = True Then
        If Worksheets("London Siege").BT14.Value = True Then
            PurchaseCost = PurchaseCost + 2000
        ElseIf Worksheets("London Siege").BT14.Value = False And PurchaseCost > 0 Then
            PurchaseCost = PurchaseCost - 2000
        End If

        RAFShopVisibility
    End If
End Sub

Sub BT15()

End Sub

Sub BT16()

End Sub

Sub BT17()

End Sub

Sub BT18()

End Sub

Sub BT19()

End Sub

Sub BT20()

End Sub

Sub BT21()

End Sub

Sub RAFShopPortal()
    Application.ScreenUpdating = False
    
    Worksheets("London Siege").RepairsBack.Visible = True
    Worksheets("London Siege").RAFShopReset.Visible = False
    Worksheets("London Siege").AmmoShopPortal.Enabled = False
    Worksheets("London Siege").RAFShopPortal.Enabled = False
    Worksheets("London Siege").RepairShopPortal.Enabled = False
    Worksheets("London Siege").StartGame.Enabled = False
    Worksheets("London Siege").NextWave.Enabled = False
    Worksheets("London Siege").QuitGame.Enabled = False
    
    Worksheets("London Siege").BT08.Value = False
    Worksheets("London Siege").BT10.Value = False
    Worksheets("London Siege").BT12.Value = False
    Worksheets("London Siege").BT14.Value = False

    RAFShopActive = True
    PurchaseCost = 0
    RocketsChange = 0
    For i = 1 To 5
        For j = 1 To 2
            RAFRepairs(i, j) = 0
        Next j
    Next i
    
    Cells(5, 45) = "Rocket"
    Cells(6, 45) = "Storage"
    Cells(7, 45) = Rockets
    Cells(4, 30) = "Health"
    Cells(7, 30) = "Rockets"
    Cells(8, 30) = "Max 6 per Fighter"
    If NumBunkerRemain = 2 Then
        Cells(11, 30) = "Rocket = 100"
    ElseIf NumBunkerRemain = 1 Then
        Cells(11, 30) = "Rocket = 200"
    End If
    If NumRepairServRemain = 2 Then
        Cells(12, 30) = "Repairs = 200"
    ElseIf NumRepairServRemain = 1 Then
        Cells(12, 30) = "Repairs = 400"
    End If
    Cells(13, 30) = "New Fighter = 2000"
    
    RAFShopVisibility
    
    Application.ScreenUpdating = True

End Sub

Sub RAFShopVisibility()
    
    If PurchaseCost = 0 And RAFRepairs(5, 2) = 0 Then
        Cells(2, 37) = ""
        For j = 0 To 3
            If Board(1, AirfieldLeft + j, 1) = RAF_ICON Then
                For i = 1 To 4
                    If RAFStatus(i, 2, 1) = AirfieldLeft + j Then
                        Cells(6, 34 + 2 * j) = RAF_ICON
                        Cells(4, 34 + 2 * j) = Int((RAFStatus(i, 7, 1) / 3) * 100) & "%"
                        Range(Cells(4, 34 + 2 * j).Address).Font.Size = 5
                        Worksheets("London Siege").Shapes("SS" & Format(8 + 2 * j, "00")).Visible = True
                        If RAFStatus(i, 7, 1) < 3 And NumRepairServRemain > 0 Then
                            Worksheets("London Siege").Shapes("ST" & Format(8 + 2 * j, "00")).Visible = True
                        Else
                            Worksheets("London Siege").Shapes("ST" & Format(8 + 2 * j, "00")).Visible = False
                        End If
                        Cells(7, 34 + 2 * j) = RAFStatus(i, 5, 1)
                        Cells(7, 34 + 2 * j).NumberFormat = "General"
                        i = 4
                    End If
                Next i
            End If
        Next j
        If Cells(6, 34) = "" And Board(1, AirfieldLeft, 1) = "     " Then Worksheets("London Siege").BT08.Visible = True
        If Cells(6, 36) = "" And Board(1, AirfieldLeft + 1, 1) = "     " Then Worksheets("London Siege").BT10.Visible = True
        If Cells(6, 38) = "" And Board(1, AirfieldLeft + 2, 1) = "     " Then Worksheets("London Siege").BT12.Visible = True
        If Cells(6, 40) = "" And Board(1, AirfieldLeft + 3, 1) = "     " Then Worksheets("London Siege").BT14.Visible = True
        
        'Ammo Bunkers must survive in order for purchase of rockets
        If NumBunkerRemain > 0 Then
            Worksheets("London Siege").SS19.Visible = True
        End If
        
        Worksheets("London Siege").AirfieldPurchase.Visible = False
        Worksheets("London Siege").RocketArm.Visible = False
        Worksheets("London Siege").RAFShopReset.Visible = False
    'If rocket/repair purchases have been made, only show options to confirm/reset
    ElseIf PurchaseCost > 0 Then
        Cells(2, 37) = "Total Cost: " & PurchaseCost
        Worksheets("London Siege").SS08.Visible = False
        Worksheets("London Siege").SS10.Visible = False
        Worksheets("London Siege").SS12.Visible = False
        Worksheets("London Siege").SS14.Visible = False
        Worksheets("London Siege").AirfieldPurchase.Visible = True
        Worksheets("London Siege").RocketArm.Visible = False
        Worksheets("London Siege").RAFShopReset.Visible = True
    'If rockets have been moved between planes and storage, only show options to confirm
    ElseIf RAFRepairs(5, 2) <> 0 Then
        Cells(2, 37) = ""
        Worksheets("London Siege").BT08.Visible = False
        Worksheets("London Siege").BT10.Visible = False
        Worksheets("London Siege").BT12.Visible = False
        Worksheets("London Siege").BT14.Visible = False
        Worksheets("London Siege").ST08.Visible = False
        Worksheets("London Siege").ST10.Visible = False
        Worksheets("London Siege").ST12.Visible = False
        Worksheets("London Siege").ST14.Visible = False
        Worksheets("London Siege").SS19.Visible = False
        Worksheets("London Siege").AirfieldPurchase.Visible = False
        Worksheets("London Siege").RocketArm.Visible = True
        Worksheets("London Siege").RAFShopReset.Visible = True
    End If
    
End Sub

Sub RocketArm()

    For i = 1 To 4
        RAFStatus(i, 5, 1) = RAFStatus(i, 5, 1) + RAFRepairs(i, 2)
        RAFRepairs(i, 2) = 0
    Next i
    RAFRepairs(5, 2) = 0
    Rockets = Rockets + RocketsChange
    Cells(21, 25) = Rockets
    RocketsChange = 0
    
    RAFShopVisibility
        
End Sub

Sub AirfieldPurchase()
    If PurchaseCost <= GameScore Then
        
        Rockets = Rockets + RocketsChange
        For ix = 1 To 4
            RAFStatus(ix, 7, 1) = RAFStatus(ix, 7, 1) + RAFRepairs(ix, 1)
            RAFRepairs(ix, 1) = 0
        Next ix
        GameScore = GameScore - PurchaseCost
        Cells(23, 24) = GameScore
        RocketsChange = 0
        PurchaseCost = 0
        If Worksheets("London Siege").BT08.Value = True And Cells(6, 34) = "" Then
            For i = 1 To 4
                If RAFStatus(i, 1, 1) = 0 Then
                    RAFStatus(i, 1, 1) = 1
                    RAFStatus(i, 2, 1) = AirfieldLeft
                    RAFStatus(i, 6, 1) = 1
                    RAFStatus(i, 7, 1) = 3
                    Cells(22, 1 + AirfieldLeft) = RAF_ICON
                    Board(1, AirfieldLeft, 1) = RAF_ICON
                    NumRAFRemain = NumRAFRemain + 1
                    Cells(6, 34) = RAF_ICON
                    Cells(4, 34) = "100%"
                    Cells(7, 34) = 0
                    Worksheets("London Siege").BT08.Visible = False
                    i = 4
                End If
            Next i
        End If
        If Worksheets("London Siege").BT10.Value = True And Cells(6, 36) = "" Then
            For i = 1 To 4
                If RAFStatus(i, 1, 1) = 0 Then
                    RAFStatus(i, 1, 1) = 1
                    RAFStatus(i, 2, 1) = AirfieldLeft + 1
                    RAFStatus(i, 6, 1) = 1
                    RAFStatus(i, 7, 1) = 3
                    Cells(22, 1 + AirfieldLeft + 1) = RAF_ICON
                    Board(1, AirfieldLeft + 1, 1) = RAF_ICON
                    NumRAFRemain = NumRAFRemain + 1
                    Cells(6, 36) = RAF_ICON
                    Cells(4, 36) = "100%"
                    Cells(7, 36) = 0
                    Worksheets("London Siege").BT10.Visible = False
                    i = 4
                End If
            Next i
        End If
        If Worksheets("London Siege").BT12.Value = True And Cells(6, 38) = "" Then
            For i = 1 To 4
                If RAFStatus(i, 1, 1) = 0 Then
                    RAFStatus(i, 1, 1) = 1
                    RAFStatus(i, 2, 1) = AirfieldLeft + 2
                    RAFStatus(i, 6, 1) = 1
                    RAFStatus(i, 7, 1) = 3
                    Cells(22, 1 + AirfieldLeft + 2) = RAF_ICON
                    Board(1, AirfieldLeft + 2, 1) = RAF_ICON
                    NumRAFRemain = NumRAFRemain + 1
                    Cells(6, 38) = RAF_ICON
                    Cells(4, 38) = "100%"
                    Cells(7, 38) = 0
                    Worksheets("London Siege").BT12.Visible = False
                    i = 4
                End If
            Next i
        End If
        If Worksheets("London Siege").BT14.Value = True And Cells(6, 40) = "" Then
            For i = 1 To 4
                If RAFStatus(i, 1, 1) = 0 Then
                    RAFStatus(i, 1, 1) = 1
                    RAFStatus(i, 2, 1) = AirfieldLeft + 3
                    RAFStatus(i, 6, 1) = 1
                    RAFStatus(i, 7, 1) = 3
                    Cells(22, 1 + AirfieldLeft + 3) = RAF_ICON
                    Board(1, AirfieldLeft + 3, 1) = RAF_ICON
                    NumRAFRemain = NumRAFRemain + 1
                    Cells(6, 40) = RAF_ICON
                    Cells(4, 40) = "100%"
                    Cells(7, 40) = 0
                    Worksheets("London Siege").BT14.Visible = False
                    i = 4
                End If
            Next i
        End If
        
        If Cells(20, 24) = "" And NumRAFRemain > 0 Then Cells(20, 24) = "Fighters"
        If Cells(20, 24) <> "" Then Cells(20, 25) = NumRAFRemain
        
        If Rockets > 0 And Cells(21, 24) = "" Then Cells(21, 24) = "Rockets"
        If Cells(21, 24) <> "" Then Cells(21, 25) = Rockets
        
        RAFShopVisibility
        
    Else
        notenough = MsgBox("Insufficient funds to make requested repairs/purchases.", vbOKOnly, "Insufficient Funds")
    End If
    
End Sub

Sub RepairShopPortal()
    Application.ScreenUpdating = False
    Worksheets("London Siege").RepairsBack.Visible = True
    Worksheets("London Siege").MakeRepairs.Visible = True
    Worksheets("London Siege").AmmoShopPortal.Enabled = False
    Worksheets("London Siege").RAFShopPortal.Enabled = False
    Worksheets("London Siege").RepairShopPortal.Enabled = False
    Worksheets("London Siege").StartGame.Enabled = False
    Worksheets("London Siege").NextWave.Enabled = False
    Worksheets("London Siege").QuitGame.Enabled = False
    For j = 1 To BOARD_COLS
        If StructureStatus(j, 1) < 4 Then
            Cells(7, 26 + j) = Board(STRUCT_ROW, j, 1)
            If StructureStatus(j, 1) > 0 Then
                Cells(9, 26 + j) = 25 * (4 - StructureStatus(j, 1)) & "%"
                Cells(9, 26 + j).Font.Size = 5
            End If
            If Board(STRUCT_ROW, j, 1) = "     " Or Board(STRUCT_ROW, j, 1) = RAF_ICON Then
                Cells(7, 26 + j).Font.Underline = xlUnderlineStyleSingle
            End If
            If StructureStatus(j, 1) = 1 Then
                Call RenderColor(Cells(7, 26 + j), "ylw")
            ElseIf StructureStatus(j, 1) = 2 Then
                Call RenderColor(Cells(7, 26 + j), "rng")
            ElseIf StructureStatus(j, 1) = 3 Then
                Call RenderColor(Cells(7, 26 + j), "red")
            End If
        End If
        If Turrets(j, 1).Health < 4 Then
            Cells(6, 26 + j) = Board(TURRET_ROW, j, 1)
            If Turrets(j, 1).Health > 0 Then
                Cells(4, 26 + j) = 25 * (4 - Turrets(j, 1).Health) & "%"
                Cells(4, 26 + j).Font.Size = 5
            End If
            If Turrets(j, 1).Health = 1 Then
                Call RenderColor(Cells(6, 26 + j), "ylw")
            ElseIf Turrets(j, 1).Health = 2 Then
                Call RenderColor(Cells(6, 26 + j), "rng")
            ElseIf Turrets(j, 1).Health = 3 Then
                Call RenderColor(Cells(6, 26 + j), "red")
            End If
        End If
    Next j
    If NumRepairServRemain = 2 Then
        Cells(11, 32) = "__ = 50"
        Cells(12, 32) = "l=l = 75"
        Cells(13, 32) = "* = 100"
        Cells(14, 32) = "# = 200"
        Cells(15, 32) = "$ = 200"
    
        For i = 1 To 12
            If Turrets(i, 1).Icon = "A" Then
                Cells(11, 42) = "A = 300"
            ElseIf Turrets(i, 1).Icon = "AA" Then
                Cells(12, 42) = "AA = 400"
            ElseIf Turrets(i, 1).Icon = ":^:" Then
                Cells(13, 42) = ":^: = 500"
            ElseIf Turrets(i, 1).Icon = ":':':" Then
                Cells(14, 42) = ":':': = 600"
            ElseIf Turrets(i, 1).Icon = "W" Then
                Cells(15, 42) = "W = 700"
            ElseIf Turrets(i, 1).Icon = "]^[" Then
                Cells(16, 42) = "]^[ = 800"
            ElseIf Turrets(i, 1).Icon = "lllll" Then
                Cells(17, 42) = "lllll = 900"
            ElseIf Turrets(i, 1).Icon = ")|(" Then
                Cells(18, 42) = ")|( = 1000"
            End If
        Next i
    ElseIf NumRepairServRemain = 1 Then                                      'Prices double if you lose a bunker
        Cells(11, 32) = "__ = 100"
        Cells(12, 32) = "l=l = 150"
        Cells(13, 32) = "* = 200"
        Cells(14, 32) = "# = 400"
        Cells(15, 32) = "$ = 400"
    
        For i = 1 To 12
            If Turrets(i, 1).Icon = "A" Then
                Cells(11, 42) = "A = 600"
            ElseIf Turrets(i, 1).Icon = "AA" Then
                Cells(12, 42) = "AA = 800"
            ElseIf Turrets(i, 1).Icon = ":^:" Then
                Cells(13, 42) = ":^: = 1000"
            ElseIf Turrets(i, 1).Icon = ":':':" Then
                Cells(14, 42) = ":':': = 1200"
            ElseIf Turrets(i, 1).Icon = "W" Then
                Cells(15, 42) = "W = 1400"
            ElseIf Turrets(i, 1).Icon = "]^[" Then
                Cells(16, 42) = "]^[ = 1600"
            ElseIf Turrets(i, 1).Icon = "lllll" Then
                Cells(17, 42) = "lllll = 1800"
            ElseIf Turrets(i, 1).Icon = ")|(" Then
                Cells(18, 42) = ")|( = 2000"
            End If
        Next i
    End If
    
    For i = 1 To 21
        If Cells(6, 26 + i) <> "" And Turrets(i, 1).Health > 0 Then
            Worksheets("London Siege").Shapes("ST" & Format(i, "00")).Visible = True
        Else
            Worksheets("London Siege").Shapes("ST" & Format(i, "00")).Visible = False
        End If
        If Cells(7, 26 + i) <> "" And StructureStatus(i, 1) > 0 Then
            Worksheets("London Siege").Shapes("SS" & Format(i, "00")).Visible = True
        Else
            Worksheets("London Siege").Shapes("SS" & Format(i, "00")).Visible = False
        End If
    Next i
    
    RepairActive = True
    
    PurchaseCost = 0
    
    Application.ScreenUpdating = True
    Worksheets("London Siege").Protect UserInterfaceOnly:=True
End Sub

Sub MakeRepairs()

    Application.ScreenUpdating = False
    
    If PurchaseCost < GameScore Then
        Dim structCode As String
        Dim turretCode As String
        For j = 1 To BOARD_COLS
            StructureStatus(j, 1) = StructureStatus(j, 1) - StructureStatusRepairs(STRUCT_ROW, j)
            Turrets(j, 1).Health = Turrets(j, 1).Health - StructureStatusRepairs(TURRET_ROW, j)
            For i = 1 To 2
                StructureStatusRepairs(i, j) = 0
            Next i
            If StructureStatus(j, 1) = 0 Then Cells(9, 26 + j) = ""
            If Turrets(j, 1).Health = 0 Then Cells(4, 26 + j) = ""
            
            structCode = GetStructureBGCode(STRUCT_ROW, StructureStatus(j, 1))
            Call RenderColor(Cells(GRID_BOT_BRDR - STRUCT_ROW, GRID_LEFT_BRDR + j), structCode)
            
            turretCode = GetStructureBGCode(TURRET_ROW, Turrets(j, 1).Health)
            Call RenderColor(Cells(GRID_BOT_BRDR - TURRET_ROW, GRID_LEFT_BRDR + j), turretCode)
        Next j
        
        For i = 1 To 21
            If Turrets(i, 1).Health = 0 Then
                Worksheets("London Siege").Shapes("ST" & Format(i, "00")).Visible = False
            End If
            If StructureStatus(i, 1) = 0 Then
                Worksheets("London Siege").Shapes("SS" & Format(i, "00")).Visible = False
            End If
        Next i
        
        GameScore = GameScore - PurchaseCost
        PurchaseCost = 0
        Cells(2, 37) = ""
        Cells(23, 24) = GameScore
    Else
        insufficient = MsgBox("Cannot make purchase. Total cost exceeds available funds.", vbOKOnly, "Insufficient Funds")
    End If
    
    Application.ScreenUpdating = True

End Sub

Sub RepairsBack()
    Application.ScreenUpdating = False
    
    RepairActive = False
    RAFShopActive = False
    
    Worksheets("London Siege").RepairsBack.Visible = False
    Worksheets("London Siege").MakeRepairs.Visible = False
    Worksheets("London Siege").RocketArm.Visible = False
    Worksheets("London Siege").AirfieldPurchase.Visible = False
    Worksheets("London Siege").RAFShopReset.Visible = False
    Worksheets("London Siege").AmmoShopPortal.Enabled = True
    Worksheets("London Siege").RAFShopPortal.Enabled = True
    Worksheets("London Siege").RepairShopPortal.Enabled = True
    Worksheets("London Siege").StartGame.Enabled = True
    Worksheets("London Siege").NextWave.Enabled = True
    Worksheets("London Siege").QuitGame.Enabled = True
    
    Worksheets("London Siege").BT08.Visible = False
    Worksheets("London Siege").BT08.Value = False
    Worksheets("London Siege").BT10.Visible = False
    Worksheets("London Siege").BT10.Value = False
    Worksheets("London Siege").BT12.Visible = False
    Worksheets("London Siege").BT12.Value = False
    Worksheets("London Siege").BT14.Visible = False
    Worksheets("London Siege").BT14.Value = False
    
    For i = 1 To 21
        Worksheets("London Siege").Shapes("ST" & Format(i, "00")).Visible = False
        Worksheets("London Siege").Shapes("SS" & Format(i, "00")).Visible = False
    Next i
    
    For i = 1 To 22
        For j = 1 To 23
            If i <= 2 And j <= BOARD_COLS Then
                StructureStatusRepairs(i, j) = 0
            End If
            Cells(i, 25 + j) = ""
            With Range(Cells(i, 25 + j).Address()).Font
                .Underline = xlUnderlineStyleNone
                If i = 4 Or i = 9 Then
                    .Size = 11
                End If
            End With
            
            Call RenderColor(Cells(i, 25 + j), "blk")
            
        Next j
    Next i
    
    RocketsChange = 0
    PurchaseCost = 0
    
    Worksheets("London Siege").Unprotect
    Application.ScreenUpdating = True
End Sub

Function ShopSpin(i As Integer, j As Integer, dir As Integer)
    Dim currentHealth As Integer
    If i = TURRET_ROW Then
        currentHealth = Turrets(j, 1).Health
    ElseIf i = STRUCT_ROW Then
        currentHealth = StructureStatus(j, 1)
    End If

    If RepairActive = True Then
        UpgradeCost = 0
        StrIcon = Cells(8 - i, 26 + j)
        If StrIcon = "     " Or StrIcon = RAF_ICON Then
            UpgradeCost = 50
        ElseIf StrIcon = "l=l" Then
            UpgradeCost = 75
        ElseIf StrIcon = "*" Then
            UpgradeCost = 100
        ElseIf StrIcon = "#" Or StrIcon = "$" Then
            UpgradeCost = 200
        ElseIf StrIcon = "A" Then
            UpgradeCost = 300
        ElseIf StrIcon = "AA" Then
            UpgradeCost = 400
        ElseIf StrIcon = ":^:" Then
            UpgradeCost = 500
        ElseIf StrIcon = ":':':" Then
            UpgradeCost = 600
        ElseIf StrIcon = "W" Then
            UpgradeCost = 700
        ElseIf StrIcon = "]^[" Then
            UpgradeCost = 800
        ElseIf StrIcon = "lllll" Then
            UpgradeCost = 900
        ElseIf StrIcon = ")|(" Then
            UpgradeCost = 1000
        End If
        If NumRepairServRemain = 1 Then
            UpgradeCost = UpgradeCost * 2
        End If
        ChangePoss = False
        If dir = 1 Then
            If StructureStatusRepairs(i, j) > 0 Then
                StructureStatusRepairs(i, j) = StructureStatusRepairs(i, j) - 1
                PurchaseCost = PurchaseCost - UpgradeCost
                ChangePoss = True
            End If
        ElseIf dir = 2 Then
            If currentHealth - StructureStatusRepairs(i, j) > 0 Then
                StructureStatusRepairs(i, j) = StructureStatusRepairs(i, j) + 1
                PurchaseCost = PurchaseCost + UpgradeCost
                ChangePoss = True
            End If
        End If
        If ChangePoss = True Then
            If currentHealth - StructureStatusRepairs(i, j) = 0 Then
                Call RenderColor(Cells(8 - i, 26 + j), "blk")
            ElseIf currentHealth - StructureStatusRepairs(i, j) = 1 Then
                Call RenderColor(Cells(8 - i, 26 + j), "ylw")
            ElseIf currentHealth - StructureStatusRepairs(i, j) = 2 Then
                Call RenderColor(Cells(8 - i, 26 + j), "rng")
            ElseIf currentHealth - StructureStatusRepairs(i, j) = 3 Then
                Call RenderColor(Cells(8 - i, 26 + j), "red")
            End If
            If i = 1 Then
                Cells(9, 26 + j) = 25 * (4 - currentHealth + StructureStatusRepairs(i, j)) & "%"
            ElseIf i = 2 Then
                Cells(4, 26 + j) = 25 * (4 - currentHealth + StructureStatusRepairs(i, j)) & "%"
            End If
            If PurchaseCost > 0 Then
                Cells(2, 37) = "Total Cost: " & PurchaseCost
            Else
                Cells(2, 37) = ""
            End If
        End If
    ElseIf RAFShopActive = True Then
        parkedspot = j
        If j = 8 Or j = 10 Or j = 12 Or j = 14 Then
            parkedspot = AirfieldLeft + (j - 6) / 2 - 1
            For ge = 1 To 4
                If RAFStatus(ge, 2, 1) = parkedspot Then
                    RAFPos = ge
                    ge = 4
                End If
            Next ge
        End If
        If i = 1 And j <> 19 Then
            If dir = 1 And RAFStatus(RAFPos, 5, 1) + RAFRepairs(RAFPos, 2) > 0 Then
                RAFRepairs(5, 2) = RAFRepairs(5, 2) - Abs(RAFRepairs(RAFPos, 2))
                RAFRepairs(RAFPos, 2) = RAFRepairs(RAFPos, 2) - 1
                RAFRepairs(5, 2) = RAFRepairs(5, 2) + Abs(RAFRepairs(RAFPos, 2))
                RocketsChange = RocketsChange + 1
                Cells(7, 26 + j) = RAFStatus(RAFPos, 5, 1) + RAFRepairs(RAFPos, 2)
                Cells(7, 45) = Rockets + RocketsChange
            ElseIf dir = 2 And Rockets + RocketsChange > 0 And RAFStatus(RAFPos, 5, 1) + RAFRepairs(RAFPos, 2) < 6 Then
                RAFRepairs(5, 2) = RAFRepairs(5, 2) - Abs(RAFRepairs(RAFPos, 2))
                RAFRepairs(RAFPos, 2) = RAFRepairs(RAFPos, 2) + 1
                RAFRepairs(5, 2) = RAFRepairs(5, 2) + Abs(RAFRepairs(RAFPos, 2))
                RocketsChange = RocketsChange - 1
                Cells(7, 26 + j) = RAFStatus(RAFPos, 5, 1) + RAFRepairs(RAFPos, 2)
                Cells(7, 45) = Rockets + RocketsChange
            End If
        ElseIf i = 1 And j = 19 Then
            If (dir = 1 And RocketsChange > 0) Or dir = 2 Then
                spinSign = (-1) ^ dir
                RocketsChange = RocketsChange + spinSign
                If NumBunkerRemain = 1 Then
                    PurchaseCost = PurchaseCost + 200 * spinSign
                Else
                    PurchaseCost = PurchaseCost + 100 * spinSign
                End If
                Cells(7, 45) = Rockets + RocketsChange
            End If
        ElseIf i = 2 Then
            If (dir = 1 And RAFRepairs(RAFPos, 1) > 0) _
            Or (dir = 2 And RAFStatus(RAFPos, 7, 1) + RAFRepairs(RAFPos, 1) < 3) Then
                spinSign = (-1) ^ dir
                RAFRepairs(RAFPos, 1) = RAFRepairs(RAFPos, 1) + spinSign
                If NumRepairServRemain = 1 Then
                    PurchaseCost = PurchaseCost + 400 * spinSign
                Else
                    PurchaseCost = PurchaseCost + 200 * spinSign
                End If
                Cells(4, 26 + j) = Int((RAFStatus(RAFPos, 7, 1) + RAFRepairs(RAFPos, 1)) / 3 * 100) & "%"
            End If
        End If
        RAFShopVisibility
    End If
End Function

Sub SS01Down()
    adjust = ShopSpin(1, 1, 1)
End Sub

Sub SS01Up()
    adjust = ShopSpin(1, 1, 2)
End Sub

Sub SS02Down()
    adjust = ShopSpin(1, 2, 1)
End Sub

Sub SS02Up()
    adjust = ShopSpin(1, 2, 2)
End Sub

Sub SS03Down()
    adjust = ShopSpin(1, 3, 1)
End Sub

Sub SS03Up()
    adjust = ShopSpin(1, 3, 2)
End Sub

Sub SS04Down()
    adjust = ShopSpin(1, 4, 1)
End Sub

Sub SS04Up()
    adjust = ShopSpin(1, 4, 2)
End Sub

Sub SS05Down()
    adjust = ShopSpin(1, 5, 1)
End Sub

Sub SS05Up()
    adjust = ShopSpin(1, 5, 2)
End Sub

Sub SS06Down()
    adjust = ShopSpin(1, 6, 1)
End Sub

Sub SS06Up()
    adjust = ShopSpin(1, 6, 2)
End Sub

Sub SS07Down()
    adjust = ShopSpin(1, 7, 1)
End Sub

Sub SS07Up()
    adjust = ShopSpin(1, 7, 2)
End Sub

Sub SS08Down()
    adjust = ShopSpin(1, 8, 1)
End Sub

Sub SS08Up()
    adjust = ShopSpin(1, 8, 2)
End Sub

Sub SS09Down()
    adjust = ShopSpin(1, 9, 1)
End Sub

Sub SS09Up()
    adjust = ShopSpin(1, 9, 2)
End Sub

Sub SS10Down()
    adjust = ShopSpin(1, 10, 1)
End Sub

Sub SS10Up()
    adjust = ShopSpin(1, 10, 2)
End Sub

Sub SS11Down()
    adjust = ShopSpin(1, 11, 1)
End Sub

Sub SS11Up()
    adjust = ShopSpin(1, 11, 2)
End Sub

Sub SS12Down()
    adjust = ShopSpin(1, 12, 1)
End Sub

Sub SS12Up()
    adjust = ShopSpin(1, 12, 2)
End Sub

Sub SS13Down()
    adjust = ShopSpin(1, 13, 1)
End Sub

Sub SS13Up()
    adjust = ShopSpin(1, 13, 2)
End Sub

Sub SS14Down()
    adjust = ShopSpin(1, 14, 1)
End Sub

Sub SS14Up()
    adjust = ShopSpin(1, 14, 2)
End Sub

Sub SS15Down()
    adjust = ShopSpin(1, 15, 1)
End Sub

Sub SS15Up()
    adjust = ShopSpin(1, 15, 2)
End Sub

Sub SS16Down()
    adjust = ShopSpin(1, 16, 1)
End Sub

Sub SS16Up()
    adjust = ShopSpin(1, 16, 2)
End Sub

Sub SS17Down()
    adjust = ShopSpin(1, 17, 1)
End Sub

Sub SS17Up()
    adjust = ShopSpin(1, 17, 2)
End Sub

Sub SS18Down()
    adjust = ShopSpin(1, 18, 1)
End Sub

Sub SS18Up()
    adjust = ShopSpin(1, 18, 2)
End Sub

Sub SS19Down()
    adjust = ShopSpin(1, 19, 1)
End Sub

Sub SS19Up()
    adjust = ShopSpin(1, 19, 2)
End Sub

Sub SS20Down()
    adjust = ShopSpin(1, 20, 1)
End Sub

Sub SS20Up()
    adjust = ShopSpin(1, 20, 2)
End Sub

Sub SS21Down()
    adjust = ShopSpin(1, 21, 1)
End Sub

Sub SS21Up()
    adjust = ShopSpin(1, 21, 2)
End Sub

Sub ST01Down()
    adjust = ShopSpin(2, 1, 1)
End Sub

Sub ST01Up()
    adjust = ShopSpin(2, 1, 2)
End Sub

Sub ST02Down()
    adjust = ShopSpin(2, 2, 1)
End Sub

Sub ST02Up()
    adjust = ShopSpin(2, 2, 2)
End Sub

Sub ST03Down()
    adjust = ShopSpin(2, 3, 1)
End Sub

Sub ST03Up()
    adjust = ShopSpin(2, 3, 2)
End Sub

Sub ST04Down()
    adjust = ShopSpin(2, 4, 1)
End Sub

Sub ST04Up()
    adjust = ShopSpin(2, 4, 2)
End Sub

Sub ST05Down()
    adjust = ShopSpin(2, 5, 1)
End Sub

Sub ST05Up()
    adjust = ShopSpin(2, 5, 2)
End Sub

Sub ST06Down()
    adjust = ShopSpin(2, 6, 1)
End Sub

Sub ST06Up()
    adjust = ShopSpin(2, 6, 2)
End Sub

Sub ST07Down()
    adjust = ShopSpin(2, 7, 1)
End Sub

Sub ST07Up()
    adjust = ShopSpin(2, 7, 2)
End Sub

Sub ST08Down()
    adjust = ShopSpin(2, 8, 1)
End Sub

Sub ST08Up()
    adjust = ShopSpin(2, 8, 2)
End Sub

Sub ST09Down()
    adjust = ShopSpin(2, 9, 1)
End Sub

Sub ST09Up()
    adjust = ShopSpin(2, 9, 2)
End Sub

Sub ST10Down()
    adjust = ShopSpin(2, 10, 1)
End Sub

Sub ST10Up()
    adjust = ShopSpin(2, 10, 2)
End Sub

Sub ST11Down()
    adjust = ShopSpin(2, 11, 1)
End Sub

Sub ST11Up()
    adjust = ShopSpin(2, 11, 2)
End Sub

Sub ST12Down()
    adjust = ShopSpin(2, 12, 1)
End Sub

Sub ST12Up()
    adjust = ShopSpin(2, 12, 2)
End Sub

Sub ST13Down()
    adjust = ShopSpin(2, 13, 1)
End Sub

Sub ST13Up()
    adjust = ShopSpin(2, 13, 2)
End Sub

Sub ST14Down()
    adjust = ShopSpin(2, 14, 1)
End Sub

Sub ST14Up()
    adjust = ShopSpin(2, 14, 2)
End Sub

Sub ST15Down()
    adjust = ShopSpin(2, 15, 1)
End Sub

Sub ST15Up()
    adjust = ShopSpin(2, 15, 2)
End Sub

Sub ST16Down()
    adjust = ShopSpin(2, 16, 1)
End Sub

Sub ST16Up()
    adjust = ShopSpin(2, 16, 2)
End Sub

Sub ST17Down()
    adjust = ShopSpin(2, 17, 1)
End Sub

Sub ST17Up()
    adjust = ShopSpin(2, 17, 2)
End Sub

Sub ST18Down()
    adjust = ShopSpin(2, 18, 1)
End Sub

Sub ST18Up()
    adjust = ShopSpin(2, 18, 2)
End Sub

Sub ST19Down()
    adjust = ShopSpin(2, 19, 1)
End Sub

Sub ST19Up()
    adjust = ShopSpin(2, 19, 2)
End Sub

Sub ST20Down()
    adjust = ShopSpin(2, 20, 1)
End Sub

Sub ST20Up()
    adjust = ShopSpin(2, 20, 2)
End Sub

Sub ST21Down()
    adjust = ShopSpin(2, 21, 1)
End Sub

Sub ST21Up()
    adjust = ShopSpin(2, 21, 2)
End Sub

Sub TickIncrement(Finish As Long)

    #If Win64 Then
        Dim NowTick As LongLong
        Dim EndTick As LongLong
        Dim Finish64 As LongLong
        Finish64 = Finish
    
        DoEvents
    
        EndTick = GetTickCount64 + (Finish64 * 10)

        Do
    
            NowTick = GetTickCount64
        
        Loop Until NowTick >= EndTick
    #Else
        Dim NowTick As Long
        Dim EndTick As Long
    
        DoEvents
    
        EndTick = GetTickCount + (Finish * 10)

        Do
    
            NowTick = GetTickCount
        
        Loop Until NowTick >= EndTick
    #End If
        

End Sub

Sub NextWave_Click()

    Worksheets("London Siege").AmmoShopPortal.Visible = False
    Worksheets("London Siege").RAFShopPortal.Visible = False
    Worksheets("London Siege").RepairShopPortal.Visible = False
    Worksheets("London Siege").NextWave.Enabled = False
    Worksheets("London Siege").StartGame.Enabled = False
    Worksheets("London Siege").QuitGame.Enabled = False

    Application.ScreenUpdating = False

    For i = 2 To 21
        For j = 2 To 22
            If i = 21 And Turrets(j - 1, 1).Icon <> "" Then
            Else
                Cells(i, j) = ""
            End If
        Next j
    Next i
    
    For it = 1 To 23
        For jt = 1 To 23
            Cells(it, 25 + jt) = ""
        Next jt
    Next it
    
    Dim structCode As String
    Dim turretCode As String
    For j = 1 To BOARD_COLS            'Remove lingering background and replace structure status color
    
        structCode = GetStructureBGCode(STRUCT_ROW, StructureStatus(j, 1))
        Call RenderColor(Cells(GRID_BOT_BRDR - STRUCT_ROW, GRID_LEFT_BRDR + j), structCode)
        
        turretCode = GetStructureBGCode(TURRET_ROW, Turrets(j, 1).Health)
        Call RenderColor(Cells(GRID_BOT_BRDR - TURRET_ROW, GRID_LEFT_BRDR + j), turretCode)
        
        For i = TURRET_ROW + 1 To BOARD_ROWS
            Call RenderColor(Cells(GRID_BOT_BRDR - i, GRID_LEFT_BRDR + j), "blu")
        Next i
    Next j

    If NumRAFRemain <> 0 And StructureStatus(ComC, 1) < 4 Then
        prelaunch = MsgBox("Do you wish to deploy RAF fighters prior to arrival of enemy aircraft?", vbYesNo, "RAF Pre-Launch")
        If prelaunch = vbYes Then
            NextWavePreLaunch
        Else
            NextWaveDisplay
        End If
    Else
        NextWaveDisplay
    End If
    
End Sub

Sub NextWavePreLaunch()
    
    Application.ScreenUpdating = False
    Worksheets("London Siege").BT08.Value = False
    Worksheets("London Siege").BT10.Value = False
    Worksheets("London Siege").BT12.Value = False
    Worksheets("London Siege").BT14.Value = False
    Worksheets("London Siege").PreLaunchBack.Visible = False
    Worksheets("London Siege").LaunchRAF.Visible = False
    RAFPreLaunch = True
    If RAFPreLaunch = True Then
        Cells(2, 37) = "Select Fighter"
        Cells(7, 30) = "Health"
        Cells(8, 30) = "Rockets"
        For j = 0 To 3
            If Board(1, AirfieldLeft + j, 1) = RAF_ICON Then
                For i = 1 To 4
                    If RAFStatus(i, 2, 1) = AirfieldLeft + j Then
                        Cells(5, 34 + 2 * j) = RAF_ICON
                        Cells(7, 34 + 2 * j) = Int((RAFStatus(i, 7, 1) / 3) * 100) & "%"
                        Range(Cells(7, 34 + 2 * j).Address).Font.Size = 5
                        Cells(8, 34 + 2 * j) = RAFStatus(i, 5, 1)
                        i = 4
                    End If
                Next i
            End If
        Next j
        If Cells(5, 34) <> "" Then Worksheets("London Siege").BT08.Visible = True
        If Cells(5, 36) <> "" Then Worksheets("London Siege").BT10.Visible = True
        If Cells(5, 38) <> "" Then Worksheets("London Siege").BT12.Visible = True
        If Cells(5, 40) <> "" Then Worksheets("London Siege").BT14.Visible = True
        Worksheets("London Siege").NextWavePreLaunchDone.Visible = True
    End If
    Application.ScreenUpdating = True

End Sub

Sub LaunchRAF()
    If ActiveCell.Row > 1 And ActiveCell.Row < 21 And ActiveCell.Column > 1 And ActiveCell.Column < 23 And RAFPreLaunch = True Then
        If ActiveCell = "" Then
            Application.ScreenUpdating = False
            Board(GRID_BOT_BRDR - ActiveCell.Row, ActiveCell.Column - 1, 1) = RAF_ICON
            ActiveCell = RAF_ICON
            Board(1, RAFStatus(RAFPos, 2, 1), 1) = "     "
            Cells(22, 1 + RAFStatus(RAFPos, 2, 1)) = "     "
            RAFStatus(RAFPos, 1, 1) = GRID_BOT_BRDR - ActiveCell.Row
            RAFStatus(RAFPos, 2, 1) = ActiveCell.Column - 1
            RAFStatus(RAFPos, 3, 1) = Int(3 * Rnd + 2)
            RAFStatus(RAFPos, 4, 1) = Int(8 * Rnd + 1)
            RAFStatus(RAFPos, 6, 1) = 2
            For i = 0 To 5
                Cells(5, 30 + 2 * i) = ""
                Cells(7, 30 + 2 * i) = ""
                Cells(8, 30 + 2 * i) = ""
            Next i
            Application.ScreenUpdating = True
            NextWavePreLaunch
        ElseIf ActiveCell <> "" Then
            nocando = MsgBox("Location already occupied", vbOKOnly, "")
        End If
    Else
        nocando = MsgBox("Invalid selection", vbOKOnly, "")
    End If
            
End Sub

Sub PreLaunchBack()

    For i = 0 To 5
        Cells(5, 30 + 2 * i) = ""
        Cells(7, 30 + 2 * i) = ""
        Cells(8, 30 + 2 * i) = ""
    Next i
    
    NextWavePreLaunch
    
End Sub

Sub NextWavePreLaunchDone()
    Application.ScreenUpdating = False
    Worksheets("London Siege").BT08.Visible = False
    Worksheets("London Siege").BT10.Visible = False
    Worksheets("London Siege").BT12.Visible = False
    Worksheets("London Siege").BT14.Visible = False
    Worksheets("London Siege").NextWavePreLaunchDone.Visible = False
    Worksheets("London Siege").LaunchRAF.Visible = False
    Worksheets("London Siege").PreLaunchBack.Visible = False

    For i = 1 To 22
        For j = 1 To 23
            Cells(i, 25 + j) = ""
            If i = 7 Then
                Range(Cells(i, 25 + j).Address()).Font.Size = 11
            End If
        Next j
    Next i
    
    RAFPreLaunch = False
    
    Application.ScreenUpdating = True
    
    NextWaveDisplay
    
End Sub
    
Sub NextWaveDisplay()

    Worksheets("London Siege").NextWave.Enabled = False
    Worksheets("London Siege").NextTurn.Enabled = True
    
    Cells(12, 12) = "Night " & Wave
    
    Application.ScreenUpdating = True
    
    Application.Wait (Now + TimeValue("0:00:03"))
    Cells(12, 12) = ""
    
    Application.ScreenUpdating = False
    
    If Wave = 1 Then
        NumEF = 3
        NumEB = 0
    ElseIf Wave = 2 Then
        NumEF = 7
        NumEB = 0
    ElseIf Wave = 3 Then
        NumEF = 20
        NumEB = 0
    ElseIf Wave = 4 Then
        NumEF = 5
        NumEB = 2
    ElseIf Wave = 5 Then
        NumEF = 10
        NumEB = 4
    ElseIf Wave = 6 Then
        NumEF = 14
        NumEB = 7
    ElseIf Wave = 7 Then
        NumEF = 0
        NumEB = 12
    ElseIf Wave = 8 Then
        NumEF = 6
        NumEB = 15
    ElseIf Wave = 9 Then
        NumEF = 15
        NumEB = 15
    ElseIf Wave = 10 Then
        NumEF = 20
        NumEB = 20
    End If
    
    If NumEF > 0 Then
        Cells(10, 24) = "Fighters"
        Cells(10, 25) = NumEF
    Else
        Cells(10, 24) = ""
        Cells(10, 25) = ""
    End If
    If NumEB > 0 Then
        Cells(11, 24) = "Bombers"
        Cells(11, 25) = NumEB
    Else
        Cells(11, 24) = ""
        Cells(11, 25) = ""
    End If
    
    NumEFPlaced = 1
    Do While NumEFPlaced <= NumEF
        Randomize
        EFr = 4 + Int(18 * Rnd)
        EFc = Int(BOARD_COLS * Rnd + 1)
        If Board(EFr, EFc, 1) = "" Then
            EFighters(NumEFPlaced, 1, 1) = EFr
            EFighters(NumEFPlaced, 2, 1) = EFc
            EFighters(NumEFPlaced, 3, 1) = Int(4 * Rnd + 1)
            EFighters(NumEFPlaced, 4, 1) = Int(8 * Rnd + 1)
            Board(EFr, EFc, 1) = EF_ICON
            NumEFPlaced = NumEFPlaced + 1
        End If
    Loop
    
    NumEBPlaced = 1
    Do While NumEBPlaced <= NumEB
        Randomize
        EBr = EB_MIN_ROW + Int(9 * Rnd)
        EBc = Int(BOARD_COLS * Rnd + 1)
        If Board(EBr, EBc, 1) = "" Then
            EBombers(NumEBPlaced, 1, 1) = EBr
            EBombers(NumEBPlaced, 2, 1) = EBc
            EBombers(NumEBPlaced, 3, 1) = Int(3 * Rnd + 1)
            EBombers(NumEBPlaced, 4, 1) = Int(2 * Rnd + 1)
            If EBombers(NumEBPlaced, 4, 1) = 1 Then
                Board(EBr, EBc, 1) = EB_ICON_LEFT
            Else
                Board(EBr, EBc, 1) = EB_ICON_RIGHT
            End If
            EBFirePrime(NumEBPlaced) = Int(5 * Rnd)
            NumEBPlaced = NumEBPlaced + 1
        End If
    Loop
    
    NumEFRemain = NumEF
    NumEBRemain = NumEB
    
    For i = 1 To BOARD_ROWS
        For j = 1 To BOARD_COLS
            Cells(GRID_BOT_BRDR - i, GRID_LEFT_BRDR + j) = Board(i, j, 1)
        Next j
    Next i

    TurnCount = 0
    ShotCount = 0
    
    Application.ScreenUpdating = True
    
    Worksheets("London Siege").StartGame.Enabled = True
    Worksheets("London Siege").QuitGame.Enabled = True
    
    Range(Cells(5, 24).Address()).Select
    
End Sub

Sub QuitGame_Click()

    If InProcess = True Then
        ImQuitting = MsgBox("Are you sure you wish to quit? All progress will be lost.", vbYesNo, "Quit Game")
        If ImQuitting = vbYes Then
            InProcess = False
            Cells(12, 12) = "Quitting Game..."
            Application.Wait (Now + TimeValue("0:00:01"))
        End If
    End If

    Call CleanHouse
    
    Application.ScreenUpdating = False
    
    If InProcess = False Then
        Worksheets("London Siege").NextWave.Caption = "Night 1"
        Worksheets("London Siege").NextWave.Enabled = False
        Worksheets("London Siege").NextTurn.Enabled = False
        Worksheets("London Siege").QuitGame.Enabled = False
        Worksheets("London Siege").StartGame.Caption = "Start"
        Worksheets("London Siege").StartGame.Enabled = True
    End If
    
    Application.ScreenUpdating = True
    
End Sub

Sub unprotected()
    Worksheets("London Siege").Unprotect
End Sub

Sub CleanHouse()

    Application.ScreenUpdating = False
    If InProcess = False Then
    
        RepairActive = False
        AmmoActive = False
        RAFShopActive = False
        RAFPreLaunch = False
        RAFComm = False
        RAFCull = False
        PlacingTurretsInit = False
        PlacingTurrets = False
        
        Worksheets("London Siege").AmmoShopPortal.Visible = False
        Worksheets("London Siege").RAFShopPortal.Visible = False
        Worksheets("London Siege").AirfieldPurchase.Visible = False
        Worksheets("London Siege").RocketArm.Visible = False
        Worksheets("London Siege").RAFShopReset.Visible = False
        Worksheets("London Siege").RepairShopPortal.Visible = False
        Worksheets("London Siege").MakeRepairs.Visible = False
        Worksheets("London Siege").RepairsBack.Visible = False
        Worksheets("London Siege").FireTurret.Visible = False
        Worksheets("London Siege").SpinTurretAim.Visible = False
        Worksheets("London Siege").SkipTurret.Visible = False
        Worksheets("London Siege").NextWavePreLaunchDone.Visible = False
        Worksheets("London Siege").PreLaunchBack.Visible = False
        Worksheets("London Siege").LaunchRAF.Visible = False
        Worksheets("London Siege").RAFCheckFire.Value = False
        Worksheets("London Siege").RAFCheckFire.Visible = False
        Worksheets("London Siege").RAFCheckRockets.Value = False
        Worksheets("London Siege").RAFCheckRockets.Visible = False
        Worksheets("London Siege").RAFSpeSpin.Visible = False
        Worksheets("London Siege").RAFDirSpin.Visible = False
        Worksheets("London Siege").RAFNext.Visible = False
        Worksheets("London Siege").RAFCommLaunch.Value = False
        Worksheets("London Siege").RAFCommLaunch.Visible = False
                                                                                                                    'Execute first due to additive implications of changing checkbox values
        For i = 1 To 21
            Worksheets("London Siege").Shapes("CheckTur" & i).Visible = False
            Worksheets("London Siege").OLEObjects("CheckTur" & i).Object.Value = False
            Worksheets("London Siege").Shapes("BT" & Format(i, "00")).Visible = False
            Worksheets("London Siege").OLEObjects("BT" & Format(i, "00")).Object.Value = False
            Worksheets("London Siege").Shapes("ST" & Format(i, "00")).Visible = False
            Worksheets("London Siege").Shapes("SS" & Format(i, "00")).Visible = False
        Next i
        
        Wave = 0
        TurnCount = 0
        GameScore = 0
        ShotCount = 0
        
        Rockets = 0
        RAFPos = 0
        
        RocketsChange = 0
        
        NumEF = 0
        NumEFRemain = 0
        NumEB = 0
        NumEBRemain = 0
        NumRAFRemain = 0
        
        AirfieldLeft = 0
        RepairServ1 = 0
        RepairServ2 = 0
        NumRepairServRemain = 0
        ComC = 0
        Bunker1 = 0
        Bunker2 = 0
        NumBunkerRemain = 0
        
        NumTurRemain = 0
        TurretPos = 0
        TurretCount = 0
        
        PurchaseCost = 0
        
        Dim newTurret As Turret
        
        For j = 1 To BOARD_COLS
            Call RenderColor(Cells(GRID_BOT_BRDR - STRUCT_ROW, GRID_LEFT_BRDR + j), "grs")
            Call RenderColor(Cells(GRID_BOT_BRDR - TURRET_ROW, GRID_LEFT_BRDR + j), "hzn")
            
            For i = 1 To BOARD_ROWS
                
                If j <= 12 Then
                    CityStruct(j) = 0
                End If
                If j <= 20 Then
                    EBFirePrime(j) = 0
                End If
            
                With Cells(GRID_BOT_BRDR - i, GRID_LEFT_BRDR + j).Font
                    If i = 1 Then .Underline = xlUnderlineStyleNone
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = 0
                End With
                With Cells(GRID_BOT_BRDR - i, GRID_LEFT_BRDR + j)
                    .Borders(xlEdgeTop).ColorIndex = 0
                    .Borders(xlEdgeTop).TintAndShade = 0
                    .Borders(xlEdgeRight).ColorIndex = 0
                    .Borders(xlEdgeRight).TintAndShade = 0
                    .Borders(xlEdgeBottom).ColorIndex = 0
                    .Borders(xlEdgeBottom).TintAndShade = 0
                    .Borders(xlEdgeLeft).ColorIndex = 0
                    .Borders(xlEdgeLeft).TintAndShade = 0
                End With

                If i > TURRET_ROW Then
                    Call RenderColor(Cells(GRID_BOT_BRDR - i, GRID_LEFT_BRDR + j), "blu")
                End If
                
                Cells(GRID_BOT_BRDR - i, GRID_LEFT_BRDR + j) = ""
                
                If i <= 14 And i <> 3 And i <> 10 And i <> 13 Then
                    If j <= 2 Then
                    Cells(9 + i, 23 + j) = ""
                    End If
                End If
                
                If i <= 5 Then
                    If j <= 2 Then
                        RAFRepairs(i, j) = 0
                    End If
                End If
                
                For k = 1 To TIME_END
                
                    Board(i, j, k) = ""
                    BG(i, j, k) = ""
                    
                    If i = 1 Then
                        Turrets(j, k) = newTurret
                        StructureStatus(j, k) = 0
                    End If
                    
                    If i <= 20 Then
                        If j <= 4 Then EFighters(i, j, k) = 0
                        If j <= 3 Then EFFire(i, j, k) = 0
                        If j <= 5 Then EBombers(i, j, k) = 0
                        If j <= 2 Then EBFire(i, j, k) = 0
                    End If
                    
                    If i <= 4 Then
                        If j <= 7 Then RAFStatus(i, j, k) = 0
                        If j <= 4 Then RAFFire(i, j, k) = 0
                    End If
                    
                Next k
            Next i
        Next j
        
        For it = 1 To 23
            For jt = 1 To 23
                Cells(it, 25 + jt) = ""
                If it = 4 Or it = 9 Then
                    Range(Cells(it, 25 + jt).Address()).Font.Size = 11
                End If
            Next jt
        Next it
        
        Range(Cells(10, 35).Address()).Font.Size = 60
        
    End If
    
    Application.ScreenUpdating = True

    
End Sub

Sub AnimateBoard(incr)

        For i = 1 To BOARD_ROWS
            For j = 1 To BOARD_COLS
                If Cells(GRID_BOT_BRDR - i, GRID_LEFT_BRDR + j) <> Board(i, j, incr) Then
                    Cells(GRID_BOT_BRDR - i, GRID_LEFT_BRDR + j) = Board(i, j, incr)
                End If
                If i = STRUCT_ROW Then
                    If StructureStatus(j, incr) = 4 And StructureStatus(j, incr - 1) <> 4 Then
                        With Cells(GRID_BOT_BRDR - i, GRID_LEFT_BRDR + j).Font
                            .Underline = xlUnderlineStyleNone
                            .ThemeColor = xlThemeColorLight1
                            .TintAndShade = 0
                        End With
                    End If
                ElseIf i = TURRET_ROW Then
                    If Turrets(j, incr).Health = 4 And Turrets(j, incr - 1).Health <> 4 Then
                        With Cells(GRID_BOT_BRDR - i, GRID_LEFT_BRDR + j).Font
                            .Underline = xlUnderlineStyleNone
                            .ThemeColor = xlThemeColorLight1
                            .TintAndShade = 0
                        End With
                    End If
                End If
                If BG(i, j, incr) <> "" Then
                    Call RenderColor(Cells(GRID_BOT_BRDR - i, GRID_LEFT_BRDR + j), BG(i, j, incr))
                End If
            Next j
        Next i

End Sub

Sub ColorizerBomb(irow, jcol, ktime)
    
    BombLandfall = False
    If irow <= 2 Then
        If irow = 1 Or (irow = 2 And Turrets(jcol, 1).Icon <> "" And Turrets(jcol, ktime).Health < 4) Then
            BombLandfall = True
        End If
    End If
    
    'Initial red explosion
    If irow < BOARD_ROWS Then
        BG(irow + 1, jcol, ktime) = "bum1"
        If jcol > 1 Then BG(irow + 1, jcol - 1, ktime) = "bul1"
        If jcol < BOARD_COLS Then BG(irow + 1, jcol + 1, ktime) = "bur1"
    End If

    BG(irow, jcol, ktime) = "bmm1"
    If jcol > 1 Then
        If irow = 1 Then
            BG(irow, jcol - 1, ktime) = "bml1g"
        Else
            BG(irow, jcol - 1, ktime) = "bml1"
        End If
    End If
    If jcol < BOARD_COLS Then
        If irow = 1 Then
            BG(irow, jcol + 1, ktime) = "bmr1g"
        Else
            BG(irow, jcol + 1, ktime) = "bmr1"
        End If
    End If
    
    If irow > 1 Then
        If BombLandfall = True Or (BombLandfall = False And irow = 2) Then
            BG(irow - 1, jcol, ktime) = "blm1g"
        Else
            BG(irow - 1, jcol, ktime) = "blm1"
        End If
        If jcol > 1 Then
            If irow = 2 Then
                BG(irow - 1, jcol - 1, ktime) = "bll1g"
            Else
                BG(irow - 1, jcol - 1, ktime) = "bll1"
            End If
        End If
        If jcol < BOARD_COLS Then
            If irow = 2 Then
                BG(irow - 1, jcol + 1, ktime) = "blr1g"
            Else
                BG(irow - 1, jcol + 1, ktime) = "blr1"
            End If
        End If
    End If
    
    'Black smoke (+ mushroom cloud if made landfall)
    If ktime <= 24 Then
        If BombLandfall = True Then
            BG(2, jcol, ktime + 1) = "blk"
            BG(irow, jcol, ktime + 1) = "bmm2"
            BG(3, jcol, ktime + 1) = "blk"
            BG(4, jcol, ktime + 1) = "bum2"
            If jcol > 1 Then
                If irow = 1 Then
                    BG(1, jcol - 1, ktime + 1) = "bgl2lf1"
                Else
                    BG(1, jcol - 1, ktime + 1) = "bgl2lf2"
                End If
                BG(2, jcol - 1, ktime + 1) = "bhl2lf"
                BG(3, jcol - 1, ktime + 1) = "bll2"
                BG(4, jcol - 1, ktime + 1) = "bul2"
            End If
            If jcol < BOARD_COLS Then
                If irow = 1 Then
                    BG(1, jcol + 1, ktime + 1) = "bgr2lf1"
                Else
                    BG(1, jcol + 1, ktime + 1) = "bgr2lf2"
                End If
                BG(2, jcol + 1, ktime + 1) = "bhr2lf"
                BG(3, jcol + 1, ktime + 1) = "blr2"
                BG(4, jcol + 1, ktime + 1) = "bur2"
            End If
        Else
            If irow < BOARD_ROWS Then
                BG(irow + 1, jcol, ktime + 1) = "bum2"
                If jcol > 1 Then BG(irow + 1, jcol - 1, ktime + 1) = "bul2"
                If jcol < BOARD_COLS Then BG(irow + 1, jcol + 1, ktime + 1) = "bur2"
            End If
            
            BG(irow, jcol, ktime + 1) = "bmm2"
            If jcol > 1 Then BG(irow, jcol - 1, ktime + 1) = "bml2"
            If jcol < BOARD_COLS Then BG(irow, jcol + 1, ktime + 1) = "bmr2"
            
            If irow = 2 Then
                BG(irow - 1, jcol, ktime + 1) = "blm2g"
                If jcol > 1 Then BG(irow - 1, jcol - 1, ktime + 1) = "bll2g"
                If jcol < BOARD_COLS Then BG(irow - 1, jcol + 1, ktime + 1) = "blr2g"
            ElseIf irow > 2 Then
                BG(irow - 1, jcol, ktime + 1) = "blm2"
                If jcol > 1 Then BG(irow - 1, jcol - 1, ktime + 1) = "bll2"
                If jcol < BOARD_COLS Then BG(irow - 1, jcol + 1, ktime + 1) = "blr2"
            End If
        End If
    End If
    
    
    If ktime <= 23 Then
        If BombLandfall = True Then
            If irow = 1 Then BG(1, jcol, ktime + 2) = "blm3g"
            BG(2, jcol, ktime + 2) = "bum3"
            BG(3, jcol, ktime + 2) = "bpd3lf"
            BG(4, jcol, ktime + 2) = "blm3"
            BG(5, jcol, ktime + 2) = "bum3"
            If jcol > 1 Then
                BG(1, jcol - 1, ktime + 2) = "bgl3lf"
                BG(2, jcol - 1, ktime + 2) = "bhl3lf"
                If BG(3, jcol - 1, ktime + 2) = "" Then BG(3, jcol - 1, ktime + 2) = "blu"
                BG(4, jcol - 1, ktime + 2) = "bll3"
                BG(5, jcol - 1, ktime + 2) = "bul3"
            End If
            If jcol < BOARD_COLS Then
                BG(1, jcol + 1, ktime + 2) = "bgr3lf"
                BG(2, jcol + 1, ktime + 2) = "bhr3lf"
                If BG(3, jcol + 1, ktime + 2) = "" Then BG(3, jcol + 1, ktime + 2) = "blu"
                BG(4, jcol + 1, ktime + 2) = "blr3"
                BG(5, jcol + 1, ktime + 2) = "bur3"
            End If
        Else
            If irow < BOARD_ROWS Then
                BG(irow + 1, jcol, ktime + 2) = "bum3"
                If jcol > 1 Then BG(irow + 1, jcol - 1, ktime + 2) = "bul3"
                If jcol < BOARD_COLS Then BG(irow + 1, jcol + 1, ktime + 2) = "bur3"
            End If
            
            BG(irow, jcol, ktime + 2) = "bmm3"
            If jcol > 1 Then BG(irow, jcol - 1, ktime + 2) = "bml3"
            If jcol < BOARD_COLS Then BG(irow, jcol + 1, ktime + 2) = "bmr3"
            
            If irow = 2 Then
                BG(irow - 1, jcol, ktime + 2) = "blm3g"
                If jcol > 1 Then BG(irow - 1, jcol - 1, ktime + 2) = "bll3g"
                If jcol < BOARD_COLS Then BG(irow - 1, jcol + 1, ktime + 2) = "blr3g"
            ElseIf irow > 2 Then
                BG(irow - 1, jcol, ktime + 2) = "blm3"
                If jcol > 1 Then BG(irow - 1, jcol - 1, ktime + 2) = "bll3"
                If jcol < BOARD_COLS Then BG(irow - 1, jcol + 1, ktime + 2) = "blr3"
            End If
        End If
    End If
    
    If ktime <= 22 Then
        If BombLandfall = True Then
            BG(3, jcol, ktime + 3) = "bpd4lf"
            BG(4, jcol, ktime + 3) = "bpd4lf"
            BG(5, jcol, ktime + 3) = "blm4"
            BG(6, jcol, ktime + 3) = "bum4"
            If jcol > 1 Then
                If BG(4, jcol - 1, ktime + 3) = "" Then BG(4, jcol - 1, ktime + 3) = "blu"
                BG(5, jcol - 1, ktime + 3) = "bll4"
                BG(6, jcol - 1, ktime + 3) = "bul4"
            End If
            If jcol < BOARD_COLS Then
                If BG(4, jcol + 1, ktime + 3) = "" Then BG(4, jcol + 1, ktime + 3) = "blu"
                BG(5, jcol + 1, ktime + 3) = "blr4"
                BG(6, jcol + 1, ktime + 3) = "bur4"
            End If
        Else
            If irow < BOARD_ROWS Then
                BG(irow + 1, jcol, ktime + 3) = "bum4"
                If jcol > 1 Then BG(irow + 1, jcol - 1, ktime + 3) = "bul4"
                If jcol < BOARD_COLS Then BG(irow + 1, jcol + 1, ktime + 3) = "bur4"
            End If
            
            BG(irow, jcol, ktime + 3) = "bmm4"
            If jcol > 1 Then BG(irow, jcol - 1, ktime + 3) = "bml4"
            If jcol < BOARD_COLS Then BG(irow, jcol + 1, ktime + 3) = "bmr4"
            
            If irow = 2 Then
                BG(irow - 1, jcol, ktime + 3) = "blm4g"
                If jcol > 1 Then BG(irow - 1, jcol - 1, ktime + 3) = "bll4g"
                If jcol < BOARD_COLS Then BG(irow - 1, jcol + 1, ktime + 3) = "blr4g"
            ElseIf irow > 2 Then
                BG(irow - 1, jcol, ktime + 3) = "blm4"
                If jcol > 1 Then BG(irow - 1, jcol - 1, ktime + 3) = "bll4"
                If jcol < BOARD_COLS Then BG(irow - 1, jcol + 1, ktime + 3) = "blr4"
            End If
        End If
    End If
    
    'Reset to blue
    If ktime <= 21 Then
        If BombLandfall = True Then
            If BG(3, jcol, ktime + 4) = "" Then BG(3, jcol, ktime + 4) = "blu"
            If BG(4, jcol, ktime + 4) = "" Then BG(4, jcol, ktime + 4) = "blu"
            If BG(5, jcol, ktime + 4) = "" Then BG(5, jcol, ktime + 4) = "blu"
            If BG(6, jcol, ktime + 4) = "" Then BG(6, jcol, ktime + 4) = "blu"
            If jcol > 1 Then
                If BG(5, jcol - 1, ktime + 4) = "" Then BG(5, jcol - 1, ktime + 4) = "blu"
                If BG(6, jcol - 1, ktime + 4) = "" Then BG(6, jcol - 1, ktime + 4) = "blu"
            End If
            If jcol < BOARD_COLS Then
                If BG(5, jcol + 1, ktime + 4) = "" Then BG(5, jcol + 1, ktime + 4) = "blu"
                If BG(6, jcol + 1, ktime + 4) = "" Then BG(6, jcol + 1, ktime + 4) = "blu"
            End If
        Else
            If irow < BOARD_ROWS Then
                If BG(irow + 1, jcol, ktime + 4) = "" Then BG(irow + 1, jcol, ktime + 4) = "blu"
                If jcol > 1 Then
                    If BG(irow + 1, jcol - 1, ktime + 4) = "" Then BG(irow + 1, jcol - 1, ktime + 4) = "blu"
                End If
                If jcol < BOARD_COLS Then
                    If BG(irow + 1, jcol + 1, ktime + 4) = "" Then BG(irow + 1, jcol + 1, ktime + 4) = "blu"
                End If
            End If
            
            If irow > 2 Then
                If BG(irow, jcol, ktime + 4) = "" Then BG(irow, jcol, ktime + 4) = "blu"
                If jcol > 1 Then
                    If BG(irow, jcol - 1, ktime + 4) = "" Then BG(irow, jcol - 1, ktime + 4) = "blu"
                End If
                If jcol < BOARD_COLS Then
                    If BG(irow, jcol + 1, ktime + 4) = "" Then BG(irow, jcol + 1, ktime + 4) = "blu"
                End If
            End If
            
            If irow > 3 Then
                If BG(irow - 1, jcol, ktime + 4) = "" Then BG(irow - 1, jcol, ktime + 4) = "blu"
                If jcol > 1 Then
                    If BG(irow - 1, jcol - 1, ktime + 4) = "" Then BG(irow - 1, jcol - 1, ktime + 4) = "blu"
                End If
                If jcol < BOARD_COLS Then
                    If BG(irow - 1, jcol + 1, ktime + 4) = "" Then BG(irow - 1, jcol + 1, ktime + 4) = "blu"
                End If
            End If
        End If
    End If

End Sub

Function GetStructureBGCode(targetRow As Integer, Health As Integer) As String
    Select Case Health
        Case 1
            GetStructureBGCode = "ylw"
        Case 2
            GetStructureBGCode = "rng"
        Case 3
            GetStructureBGCode = "red"
        Case Else
            Dim output As String
            If targetRow = STRUCT_ROW Then
                output = "grs"
            ElseIf targetRow = TURRET_ROW Then
                output = "hzn"
            End If
            If Health = 4 Then
                output = output & "rns"
            End If
            GetStructureBGCode = output
    End Select
End Function


Sub RenderColor(targetCell As Range, encoded As String)
    Dim code As String
    If encoded = "grsrns" Or encoded = "hznrns" Then
        'Change font color to black for ruins
        With targetCell.Font
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
        End With
        code = Left(encoded, 3)
    ElseIf encoded <> "" Then
        'Reset font to standard color
        With targetCell.Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        code = encoded
    End If
    
    
    Select Case code
        Case "grs"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 90
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .Color = 5009280
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.250984221930601
            End With
        Case "hzn"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 90
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .Color = 5009280
                .TintAndShade = 0
            End With
'Night sky blue
        Case "blu"
            With targetCell.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
                .PatternTintAndShade = 0
            End With
'Structure Status
        Case "ylw"
            With targetCell.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent4
                .TintAndShade = 0.45
                .PatternTintAndShade = 0
            End With
        Case "rng"
            With targetCell.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = 0.2
                .PatternTintAndShade = 0
            End With
        Case "red"
            With targetCell.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 192
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Case "blk"
            With targetCell.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        Case "blkout"
            With targetCell.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With targetCell.Font
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
'Aircraft/fighter fire explosions
        Case "sx1"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 0.5
                .Gradient.RectangleRight = 0.5
                .Gradient.RectangleTop = 0.5
                .Gradient.RectangleBottom = 0.5
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .Color = 255
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.498031556138798
            End With
        Case "sx2"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 0.5
                .Gradient.RectangleRight = 0.5
                .Gradient.RectangleTop = 0.5
                .Gradient.RectangleBottom = 0.5
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = -0.250984221930601
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.498031556138798
            End With
        Case "sx3"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 0.5
                .Gradient.RectangleRight = 0.5
                .Gradient.RectangleTop = 0.5
                .Gradient.RectangleBottom = 0.5
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.250984221930601
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.498031556138798
            End With
'Bomb explosions incr
        Case "bul1"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 1
                .Gradient.RectangleRight = 1
                .Gradient.RectangleTop = 1
                .Gradient.RectangleBottom = 1
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .Color = 255
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.498031556138798
            End With
        Case "bum1"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 90
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .Color = 255
                .TintAndShade = 0
            End With
        Case "bur1"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 0
                .Gradient.RectangleRight = 0
                .Gradient.RectangleTop = 1
                .Gradient.RectangleBottom = 1
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .Color = 255
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.498031556138798
            End With
        Case "bml1"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 0
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .Color = 255
                .TintAndShade = 0
            End With
        Case "bml1g"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 0
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.250984221930601
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .Color = 255
                .TintAndShade = 0
            End With
        Case "bmm1"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 0.5
                .Gradient.RectangleRight = 0.5
                .Gradient.RectangleTop = 0.5
                .Gradient.RectangleBottom = 0.5
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .Color = 255
                .TintAndShade = 0
            End With
        Case "bmr1"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 180
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .Color = 255
                .TintAndShade = 0
            End With
        Case "bmr1g"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 180
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.250984221930601
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .Color = 255
                .TintAndShade = 0
            End With
        Case "bll1"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 1
                .Gradient.RectangleRight = 1
                .Gradient.RectangleTop = 0
                .Gradient.RectangleBottom = 0
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .Color = 255
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.498031556138798
            End With
        Case "bll1g"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 1
                .Gradient.RectangleRight = 1
                .Gradient.RectangleTop = 0
                .Gradient.RectangleBottom = 0
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .Color = 255
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.250984221930601
            End With
        Case "blm1"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 90
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .Color = 255
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
            End With
        Case "blm1g"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 90
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .Color = 255
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.250984221930601
            End With
        Case "blr1"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 0
                .Gradient.RectangleRight = 0
                .Gradient.RectangleTop = 0
                .Gradient.RectangleBottom = 0
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .Color = 255
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.498031556138798
            End With
        Case "blr1g"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 0
                .Gradient.RectangleRight = 0
                .Gradient.RectangleTop = 0
                .Gradient.RectangleBottom = 0
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .Color = 255
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.250984221930601
            End With
' " " incr + 1
        Case "bul2"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 1
                .Gradient.RectangleRight = 1
                .Gradient.RectangleTop = 1
                .Gradient.RectangleBottom = 1
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.498031556138798
            End With
        Case "bum2"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = -90
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
            End With
        Case "bur2"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 0
                .Gradient.RectangleRight = 0
                .Gradient.RectangleTop = 1
                .Gradient.RectangleBottom = 1
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.498031556138798
            End With
        Case "bml2"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 180
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
            End With
        Case "bmm2"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 0.5
                .Gradient.RectangleRight = 0.5
                .Gradient.RectangleTop = 0.5
                .Gradient.RectangleBottom = 0.5
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorAccent2
                .TintAndShade = -0.498031556138798
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
        Case "bmr2"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 0
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
            End With
        Case "bll2"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 1
                .Gradient.RectangleRight = 1
                .Gradient.RectangleTop = 0
                .Gradient.RectangleBottom = 0
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.498031556138798
            End With
        Case "bll2g"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 1
                .Gradient.RectangleRight = 1
                .Gradient.RectangleTop = 0
                .Gradient.RectangleBottom = 0
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.250984221930601
            End With
        Case "blm2"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 90
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
            End With
        Case "blm2g"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 90
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.250984221930601
            End With
        Case "blr2"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 0
                .Gradient.RectangleRight = 0
                .Gradient.RectangleTop = 0
                .Gradient.RectangleBottom = 0
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.498031556138798
            End With
        Case "blr2g"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 0
                .Gradient.RectangleRight = 0
                .Gradient.RectangleTop = 0
                .Gradient.RectangleBottom = 0
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.250984221930601
            End With
' " " incr + 2
        Case "bul3"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 1
                .Gradient.RectangleRight = 1
                .Gradient.RectangleTop = 1
                .Gradient.RectangleBottom = 1
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.149021881771294
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.498031556138798
            End With
        Case "bum3"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = -90
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.149021881771294
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
            End With
        Case "bur3"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 0
                .Gradient.RectangleRight = 0
                .Gradient.RectangleTop = 1
                .Gradient.RectangleBottom = 1
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.149021881771294
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.498031556138798
            End With
        Case "bml3"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 180
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.149021881771294
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
            End With
        Case "bmm3"
            With targetCell.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.149021881771294
                .PatternTintAndShade = 0
            End With
        Case "bmr3"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 0
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.149021881771294
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
            End With
        Case "bll3"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 1
                .Gradient.RectangleRight = 1
                .Gradient.RectangleTop = 0
                .Gradient.RectangleBottom = 0
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.149021881771294
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.498031556138798
            End With
        Case "bll3g"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 1
                .Gradient.RectangleRight = 1
                .Gradient.RectangleTop = 0
                .Gradient.RectangleBottom = 0
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.149021881771294
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.250984221930601
            End With
        Case "blm3"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 90
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.149021881771294
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
            End With
        Case "blm3g"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 90
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.149021881771294
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.250984221930601
            End With
        Case "blr3"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 0
                .Gradient.RectangleRight = 0
                .Gradient.RectangleTop = 0
                .Gradient.RectangleBottom = 0
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.149021881771294
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.498031556138798
            End With
        Case "blr3g"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 0
                .Gradient.RectangleRight = 0
                .Gradient.RectangleTop = 0
                .Gradient.RectangleBottom = 0
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.149021881771294
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.250984221930601
            End With
' " " incr + 3
        Case "bul4"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 1
                .Gradient.RectangleRight = 1
                .Gradient.RectangleTop = 1
                .Gradient.RectangleBottom = 1
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.349009674367504
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.498031556138798
            End With
        Case "bum4"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = -90
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.349009674367504
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
            End With
        Case "bur4"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 0
                .Gradient.RectangleRight = 0
                .Gradient.RectangleTop = 1
                .Gradient.RectangleBottom = 1
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.349009674367504
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.498031556138798
            End With
        Case "bml4"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 180
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.349009674367504
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
            End With
        Case "bmm4"
            With targetCell.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.349009674367504
                .PatternTintAndShade = 0
            End With
        Case "bmr4"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 0
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.349009674367504
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
            End With
        Case "bll4"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 1
                .Gradient.RectangleRight = 1
                .Gradient.RectangleTop = 0
                .Gradient.RectangleBottom = 0
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.349009674367504
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.498031556138798
            End With
        Case "bll4g"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 1
                .Gradient.RectangleRight = 1
                .Gradient.RectangleTop = 0
                .Gradient.RectangleBottom = 0
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.349009674367504
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.250984221930601
            End With
        Case "blm4"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 90
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.349009674367504
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
            End With
        Case "blm4g"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 90
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.349009674367504
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.250984221930601
            End With
        Case "blr4"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 0
                .Gradient.RectangleRight = 0
                .Gradient.RectangleTop = 0
                .Gradient.RectangleBottom = 0
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.349009674367504
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.498031556138798
            End With
        Case "blr4g"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 0
                .Gradient.RectangleRight = 0
                .Gradient.RectangleTop = 0
                .Gradient.RectangleBottom = 0
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.349009674367504
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.250984221930601
            End With
'Accessory codes if bomb landfall
        Case "bhl2lf"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = -135
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0.1)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
            End With
        Case "bhr2lf"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = -45
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0.1)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
            End With
        Case "bgl2lf1"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 180
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0.5)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.250984221930601
            End With
        Case "bgr2lf1"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 0
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0.5)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.250984221930601
            End With
        Case "bgl2lf2"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 135
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0.4)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.250984221930601
            End With
        Case "bgr2lf2"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 45
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0.4)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.250984221930601
            End With
        Case "bhl3lf"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = -120
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.149021881771294
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0.4)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
            End With
        Case "bhr3lf"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = -60
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.149021881771294
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0.4)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
            End With
        Case "bgl3lf"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 120
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.149021881771294
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0.4)
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.250984221930601
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.250984221930601
            End With
        Case "bgr3lf"
            With targetCell.Interior
                .Pattern = xlPatternLinearGradient
                .Gradient.Degree = 60
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.149021881771294
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0.4)
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.250984221930601
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = -0.250984221930601
            End With
        Case "bpd3lf"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 0.5
                .Gradient.RectangleRight = 0.5
                .Gradient.RectangleTop = 0.5
                .Gradient.RectangleBottom = 0.5
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.149021881771294
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
            End With
        Case "bpd4lf"
            With targetCell.Interior
                .Pattern = xlPatternRectangularGradient
                .Gradient.RectangleLeft = 0.5
                .Gradient.RectangleRight = 0.5
                .Gradient.RectangleTop = 0.5
                .Gradient.RectangleBottom = 0.5
                .Gradient.ColorStops.Clear
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(0)
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0.349009674367504
            End With
            With targetCell.Interior.Gradient.ColorStops.Add(1)
                .ThemeColor = xlThemeColorAccent1
                .TintAndShade = -0.499984740745262
            End With
    End Select
End Sub
