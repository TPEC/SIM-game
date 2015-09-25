Attribute VB_Name = "mdl"
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Public GameTime As Currency, SystLast_ms As Integer, FPS As Integer

Public frmMExisted As Boolean
Public Draw_BagF As Boolean
Public Draw_Place As Integer '-1-动画,0-仓库,1-花房,2-商店

Public UserName As String
Public Money As Long

Public Type tFlowerType
    fName As String
    fAge(0 To 3) As Integer
    fPrice As Integer
    fQuality As Integer
    fHealthd As Single '得病率(基准0,+-25MAX)
    fWaterd As Single '耗水率 (基准1)
    fp_Leaf As Integer
    fp_Flower As Integer
    fp_Seed As Integer
    fp_Stem As Integer
End Type
Public Type tPointInt
    Xi As Integer
    Yi As Integer
End Type
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Type tFlowerUnit
    fuHealth As Boolean
    fuStyle As Integer '样式
    fuPosition As tPointInt
End Type
Public Type tFlower
    fPotId As Integer
    fPlace As Integer '0:花房，1:商店
    fType As Integer '0:None
    fColor As Integer
    fBirth As Long
    fAge As Integer
    fAgeP As Integer '处于哪种时期
    fHealth As Integer
    fQualityC As Integer
    fPriceC As Integer
    fLeaf(1 To 8) As tFlowerUnit
    fFlower(1 To 8) As tFlowerUnit
End Type

Public FloT(1 To 128) As tFlowerType
Public Flo(1 To 64) As tFlower

Public Type tFlowerpot
    fp_Pattern As Integer
    fFloId As Integer '0:None
    fWater As Integer
    fSoilQ As Integer '0:None
End Type
Public Flop(1 To 64) As tFlowerpot

Public EleFHT(0 To 31) As Integer  '0=None,1=Flowerpot,2-Ornament
Public EleFHI(0 To 31) As Integer 'ID

Public MouseP As tPointInt, MouseDP As tPointInt, MouseDL As Boolean, MouseDR As Boolean

Public Type tBagUnit
    bType As Integer '1:种子
    bId As Integer
    bAmount As Integer
End Type
Public Bag(1 To 256) As tBagUnit


Sub Main()
    frmM.Show
    frmMExisted = True
    Draw_Place = 1
    Dim syst As SYSTEMTIME
    Dim td As Double
    SystLast_ms = 0
    Do
        GetLocalTime syst
        td = (syst.wMilliseconds - SystLast_ms) / 1000
        SystLast_ms = syst.wMilliseconds
        If td < 0 Then td = td + 1
        FPS = 1 / (td + 0.001)
        GameTime = GameTime + td
        DoEvents
        Call Draw
    Loop Until Not frmMExisted
    End
End Sub

Public Sub Draw()
    With frmM.picM
        frmM.picM.Line (0, 540)-Step(960, 180), RGB(255, 255, 255), BF
        Select Case Draw_Place
            Case 0
            Case 1
                frmP.picL.Picture = LoadPicture(App.Path & "\Graphics\BGFlowerHouse.bmp")
                BitBlt .hDC, 0, 0, 960, 540, frmP.picL.hDC, 0, 0, vbSrcCopy
                For i = 0 To 31
                    If EleFHT(i) = 1 Then
                        
                    ElseIf EleFHT(i) = 2 Then
                        BitBlt .hDC, 40 + (i Mod 8) * 110, 40 + (i \ 8) * 120, 110, 125, frmP.picL.hDC, 0, 0, vbSrcCopy
                    End If
                Next i
                frmP.picL.Picture = LoadPicture(App.Path & "\Graphics\BGTools.bmp")
                BitBlt .hDC, 0, 540, 960, 180, frmP.picL.hDC, 0, 0, vbSrcCopy
            Case 2
            Case -1
        End Select
        .CurrentX = 0
        .CurrentY = 0
        .FontSize = 20
        frmM.picM.Print GameTime, FPS, MouseP.Xi
        .Refresh
    End With
End Sub

Private Sub Logic()
    Dim ti As Integer, ts As Single
    Randomize Timer
    For i = 1 To 64
        If Flo(i).fType > 0 Then
            With Flo(i)
                ti = GameTime - .fBirth
                If ti > .fAge Then
                    If Rnd * 3 < FloT(.fType).fWaterd Then
                        If Flop(.fPotId).fWater > 0 Then
                            Flop(.fPotId).fWater = Flop(.fPotId).fWater - 1
                        End If
                    End If
                    ts = Rnd * 100 - 50 - FloT(.fType).fHealthd
                    If Abs(ts) >= 25 Then
                        If ts < 0 Then
                            If .fHealth > 0 Then .fHealth = .fHealth - 1
                        Else
                            If .fHealth < 100 Then .fHealth = .fHealth + 1
                        End If
                    End If
                    .fAge = ti
                    If .fAge > FloT(.fType).fAge(.fAgeP) Then
                        Call FlowerGrow(i)
                    End If
                End If
            End With
        End If
    Next i
End Sub

Private Sub FlowerGrow(fid As Integer)
    With Flo(fid)
        .fAgeP = .fAgeP + 1
        
    End With
End Sub

Private Sub DeviceInput()
    Select Case Draw_Place
        Case 1
            If MouseP.Xi >= 800 And MouseP.Xi < 940 And MouseP.Yi >= 560 And MouseP.Yi < 700 Then 'bagbutton
                Draw_BagF = True
            End If
    End Select
    If Draw_BagF Then
        If MouseP.Xi >= 800 Or MouseP.Xi < 940 Or MouseP.Yi >= 560 Or MouseP.Yi < 700 Then 'bagpic
            Draw_BagF = False
        End If
    End If
End Sub

