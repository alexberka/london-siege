Attribute VB_Name = "modUDTs"
Public Type BoardSquare
    Icon As String
    BG As String
End Type

Public Enum FlightDir
    DirNone = 0
    DirDownRt = 1
    DirDown = 2
    DirDownLf = 3
    DirLeft = 4
    DirUpLf = 5
    DirUp = 6
    DirUpRt = 7
    DirRight = 8
End Enum

Public Enum RAFFireType
    Base = 1
    Rocket = 2
End Enum

Public Enum RAFStatus
    OnGround = 1
    InAir = 2
End Enum

Public Type RAFFire
    Icon As String
    Row As Integer
    Column As Integer
    Type As RAFFireType
    Direction As FlightDir
End Type

Public Type RAF
    Row As Integer
    Column As Integer
    Speed As Integer
    Direction As FlightDir
    Rockets As Integer
    Status As RAFStatus
    Health As Integer
    Fire As RAFFire
End Type

Public Type EFFire
    Icon As String
    Row As Integer
    Column As Integer
    Direction As FlightDir
End Type

Public Type EFighter
    Row As Integer
    Column As Integer
    Speed As Integer
    Direction As FlightDir
    Fire As EFFire
End Type

Public Type EBFire
    Row As Integer
    Column As Integer
End Type

Public Type EBomber
    Icon As String
    Row As Integer
    Column As Integer
    Speed As Integer
    Direction As FlightDir
    Damage As Integer
    FirePrime As Integer
    Fire As EBFire
End Type

Public Enum TurretFireDir
    DirNone = 0
    Dir27 = 1
    Dir45 = 2
    Dir63 = 3
    Dir72 = 4
    Dir76 = 5
    Dir90 = 6
    Dir104 = 7
    Dir108 = 8
    Dir117 = 9
    Dir135 = 10
    Dir153 = 11
End Enum

Public Type TurretFire
    Row As Integer
    Column As Integer
    Direction As TurretFireDir
    Type As Integer
End Type

Public Type Turret
    Icon As String
    Health As Integer
    Fire As TurretFire
End Type
