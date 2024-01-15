VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwOverlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule IntegerDataType, ModuleWithoutFolder
'---------------------------------------------------------------------------------------
' Module    : cwOverlay
' Author    : beededea
' Date      : 31/07/2023
' Purpose   : create a fully click-through Widget, which allows for all kind of
'             things to render on top of the PSD-generated Widgets created in cfAlpha
'---------------------------------------------------------------------------------------

Option Explicit

Private Declare Function GetDiskFreeSpace _
    Lib "kernel32" _
    Alias "GetDiskFreeSpaceA" ( _
        ByVal lpRootPathName As String, _
        lpSectorsPerCluster As Long, _
        lpBytesPerSector As Long, _
        lpNumberOfFreeClusters As Long, _
        lpTotalNumberOfClusters As Long) _
        As Long
                
Private storagePercent As Long
Private storageFree As Long

'these Public Properties are set in the Parent (here, cfAlpha)
Public cCollOverlayImages As cCollection
Public cxPerc As Single
Public cyPerc As Single

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
Private WithEvents tmrClock As cTimer
Attribute tmrClock.VB_VarHelpID = -1

Private SDeg As Double

Private mTicking As Boolean ' property Ticking
Private mSmoothSecondHand As Boolean ' property SmoothSecondHand
Private mLocked As Boolean ' property Locked
Private mHidden As Boolean ' property Hidden
Private mThisDriveNo As Integer ' property thisDriveNo
Private mThisDrive As String ' property thisDrive
Private mThisOpacity As String ' property ThisOpacity
Private mSamplingInterval As String ' property SamplingInterval



'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Dim intervalValue As Integer: intervalValue = 0
    
    On Error GoTo Class_Initialize_Error

    Set W = Cairo.WidgetBase
        W.BackColor = -1  'we don't use any BackColor here, transparent
        W.ImplementsHitTest = True 'when at False, no HitTest-Event would be triggered
                    
    intervalValue = 5000

    Set tmrClock = Nothing
    Set tmrClock = New_c.Timer(100, mTicking)

   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cwOverlay"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Widget
' Author    : olaf schmidt
' Date      : 21/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Widget() As cWidgetBase
   On Error GoTo Widget_Error

    Set Widget = W

   On Error GoTo 0
   Exit Property

Widget_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widget of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Procedure : Widgets
' Author    : olaf schmidt
' Date      : 21/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Widgets() As cWidgets
   On Error GoTo Widgets_Error

    Set Widgets = W.Widgets

   On Error GoTo 0
   Exit Property

Widgets_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widgets of Class Module cwOverlay"
End Property


'---------------------------------------------------------------------------------------
' Property : Ticking
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Ticking() As Boolean
   On Error GoTo TickingGet_Error

  Ticking = mTicking

   On Error GoTo 0
   Exit Property

TickingGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Ticking of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : Ticking
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Ticking(ByVal newValue As Boolean)
    'Dim intervalValue As Integer: intervalValue = 0
    On Error GoTo TickingLet_Error

    If mTicking <> newValue Then mTicking = newValue Else Exit Property
    
    If mTicking Then
'        If mSmoothSecondHand Then
'            'intervalValue = 100
'        Else
'            'intervalValue = 500
'        End If

        tmrClock.Enabled = True
    Else
        tmrClock.Enabled = False
    End If

   On Error GoTo 0
   Exit Property

TickingLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Ticking of Class Module cwOverlay"
End Property ' the earth Ticking flag

'---------------------------------------------------------------------------------------
' Property : SmoothSecondHand
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get SmoothSecondHand() As Boolean
   On Error GoTo SmoothSecondHandGet_Error

    SmoothSecondHand = mSmoothSecondHand

   On Error GoTo 0
   Exit Property

SmoothSecondHandGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property SmoothSecondHand of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : SmoothSecondHand
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let SmoothSecondHand(ByVal newValue As Boolean)
    Dim intervalValue As Integer: intervalValue = 0
    On Error GoTo SmoothSecondHandLet_Error

    If mSmoothSecondHand <> newValue Then mSmoothSecondHand = newValue Else Exit Property
    
    If mSmoothSecondHand Then
        intervalValue = 100
    Else
        intervalValue = 500
    End If
            
    tmrClock.Interval = intervalValue
    
   On Error GoTo 0
   Exit Property

SmoothSecondHandLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property SmoothSecondHand of Class Module cwOverlay"
End Property

'---------------------------------------------------------------------------------------
' Property : thisOpacity
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get thisOpacity() As String
   On Error GoTo thisOpacityGet_Error

    thisOpacity = mThisOpacity

   On Error GoTo 0
   Exit Property

thisOpacityGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisOpacity of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : thisOpacity
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let thisOpacity(ByVal newValue As String)
    On Error GoTo thisOpacityLet_Error

    If mThisOpacity <> newValue Then mThisOpacity = newValue Else Exit Property
        
    On Error GoTo 0
    Exit Property

thisOpacityLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisOpacity of Class Module cwOverlay"
End Property


'---------------------------------------------------------------------------------------
' Property : thisDrive
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get thisDrive() As String
   On Error GoTo thisDriveGet_Error

    thisDrive = mThisDrive

   On Error GoTo 0
   Exit Property

thisDriveGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisDrive of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : thisDrive
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let thisDrive(ByVal newValue As String)
    On Error GoTo thisDriveLet_Error

    If mThisDrive <> newValue Then mThisDrive = newValue Else Exit Property
    
    Call obtainStorage
    
    On Error GoTo 0
    Exit Property

thisDriveLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisDrive of Class Module cwOverlay"
End Property


'---------------------------------------------------------------------------------------
' Property : thisDrive
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get thisDriveNo() As Integer
   On Error GoTo thisDriveNoGet_Error

    thisDriveNo = mThisDriveNo

   On Error GoTo 0
   Exit Property

thisDriveNoGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisDriveNo of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : thisDriveNo
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let thisDriveNo(ByVal newValue As Integer)
    On Error GoTo thisDriveNoLet_Error

    If mThisDriveNo <> newValue Then mThisDriveNo = newValue Else Exit Property
            
    'given a drive number select the drive string
    If mThisDriveNo < gblDriveCount Then
        mThisDrive = gblDriveArray(mThisDriveNo)
    End If
    
    Call obtainStorage
    
    On Error GoTo 0
    Exit Property

thisDriveNoLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisDriveNo of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property : samplingInterval
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get samplingInterval() As String
   On Error GoTo samplingIntervalGet_Error

    samplingInterval = mSamplingInterval

   On Error GoTo 0
   Exit Property

samplingIntervalGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property samplingInterval of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : samplingInterval
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let samplingInterval(ByVal newValue As String)
    On Error GoTo samplingIntervalLet_Error

    If mSamplingInterval <> newValue Then mSamplingInterval = newValue Else Exit Property
    
    Call obtainStorage
    
    On Error GoTo 0
    Exit Property

samplingIntervalLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property samplingInterval of Class Module cwOverlay"
End Property

'---------------------------------------------------------------------------------------
' Property : Locked
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Locked() As Boolean
   On Error GoTo LockedGet_Error

   Locked = mLocked

   On Error GoTo 0
   Exit Property

LockedGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : Locked
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Locked(ByVal newValue As Boolean)
   On Error GoTo LockedLet_Error

   If mLocked <> newValue Then mLocked = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

LockedLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwOverlay"
End Property

'---------------------------------------------------------------------------------------
' Property  : Hidden
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Hidden() As Boolean
   On Error GoTo HiddenGet_Error

   Hidden = mHidden

   On Error GoTo 0
   Exit Property

HiddenGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : Hidden
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Hidden(ByVal newValue As Boolean)
   On Error GoTo HiddenLet_Error

    If mHidden <> newValue Then mHidden = newValue Else Exit Property
    If mHidden = True Then
        W.Visible = False
        
        W.Refresh
    Else
        W.Visible = True
        
        W.Refresh
    End If

   On Error GoTo 0
   Exit Property

HiddenLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwOverlay"
End Property
'
'---------------------------------------------------------------------------------------------
' Procedure : W_HitTest
' Author    : olaf schmidt
' Date      : 01/08/2023
' Purpose   : This is run continuously for the whole widget.
'             For this Overlay-Widget alone, we make everything "fully click-through"
'---------------------------------------------------------------------------------------------
'
Private Sub W_HitTest(ByVal x As Single, ByVal y As Single, HitResultHit As Boolean) 'ensure ClickThrough-behaviour in ImagePixels which are "fully Alpha"
   On Error GoTo W_HitTest_Error

   HitResultHit = False

   On Error GoTo 0
   Exit Sub

W_HitTest_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_HitTest of Class Module cwOverlay"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : W_Paint
' Author    : olaf schmidt
' Date      : 01/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub W_Paint(CC As RC6.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
    Dim I As Long: I = 0
    Dim Arr As Variant
    Dim Key As String: Key = vbNullString
    Dim Pat As cCairoPattern
    Dim txtStorageFree As String: txtStorageFree = vbNullString
    Dim txtStoragePercent As String: txtStoragePercent = vbNullString
    Dim textDrawn As Boolean:   textDrawn = False
    Dim someOpacity As Double: someOpacity = 0
    Dim cream As String: cream = vbNullString
    
    cream = RGB(235, 217, 179) ' cannot create a constant from an RGB value
   
    On Error GoTo W_Paint_Error
    
    'set some Font- and ForeColor-Properties and draw some Text on the Widget
    W.FontName = "arial"
    W.ForeColor = cream
    W.FontSize = 18
    W.FontBold = True
    W.SelectFontSettingsInto CC
    W.Refresh
    someOpacity = (Val(mThisOpacity) / 100) * 0.75
    
    txtStoragePercent = CStr(Int(storagePercent))
    txtStorageFree = CStr(Int(storageFree))

    For I = 0 To cCollOverlayImages.Count - 1 ' this is the ExcludedPaths collection populated in cfAlpha by duplicating collPSDNonUIElements.
      Arr = cCollOverlayImages.ItemByIndex(I)
      Key = cCollOverlayImages.KeyByIndex(I)
      
      If IsArray(Arr) Then
            CC.Save
            'derive a Pattern-Obj from each retrieved ImageList-Surface-reference
            Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
                         
            ' for specific keys perform rotation actions, clock hands &c
             Select Case Key

                 Case "secondshadow", "secondhand"
                    
                    'set some Font and ForeColor Properties and draw some Text on the Widget
                    W.FontName = "arial"
                    W.ForeColor = vbBlack
                    W.FontSize = 38
                    W.FontBold = True
                    W.SelectFontSettingsInto CC
                    W.Refresh
                    
                    ' draw the drive in the top window
                    CC.DrawText 335, 67, 200, 200, Left$(mThisDrive, 2), , , 2, True, , someOpacity
                    
                    'set some Font and ForeColor Properties and draw some Text on the Widget
                    W.FontName = "arial"
                    W.ForeColor = cream
                    W.FontSize = 17
                    W.FontBold = True
                    W.SelectFontSettingsInto CC
                    W.Refresh
                    
                    ' draw the drive data in the two black windows
                    CC.DrawText 267, 105, 200, 200, txtStoragePercent, , , 2, True, , someOpacity
                    CC.DrawText 453, 298, 200, 200, txtStorageFree, , , 2, True, , someOpacity
                    textDrawn = True
                    
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(11, 246) 'shift to the "rotation-point" within the Pattern
                    CC.TranslateDrawings W.Width * cxPerc, W.Height * cyPerc 'shift to the center-point of the Overlay-Widget-area
                    
                    CC.RotateDrawingsDeg SDeg
                    CC.Paint Arr(2), Pat 'render the Pattern with the current PSD-LayerAlpha, present in Arr(2)
                    
                 Case "redlamp"
                    
                    If storagePercent >= 80 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint Arr(2), Pat 'render the Pattern
                    End If
                    
                 Case "greenlamp"
                    
                    If storagePercent < 80 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1))
                        CC.Paint Arr(2), Pat ' render the Pattern
                    End If
                    
                 Case Else 'plain output for everything else ... this includes the reflection-layers,
                     
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1))
                    CC.Paint Arr(2), Pat 'render the Pattern
                   
             End Select
             
             CC.Restore
      End If
     Next

   On Error GoTo 0
   Exit Sub

W_Paint_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_Paint of Class Module cwOverlay"
End Sub
 
 

    
'---------------------------------------------------------------------------------------
' Procedure : tmrClock_Timer
' Author    :
' Date      : 30/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub tmrClock_Timer()
    
    On Error GoTo tmrClock_Timer_Error
    
    tmrClock.Interval = Val(mSamplingInterval) * 1000

'    If mSmoothSecondHand = True Then

'    Else
        Call obtainStorage
        
'    End If
    
    W.Parent.Refresh

   On Error GoTo 0
   Exit Sub

tmrClock_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrClock_Timer of Class Module cwOverlay"
End Sub

' ----------------------------------------------------------------
' Procedure Name: obtainStorage
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 13/01/2024
' ----------------------------------------------------------------
Private Sub obtainStorage()

    On Error GoTo obtainStorage_Error
    
    storagePercent = Storage_Usage_Percent
    storageFree = Storage_Usage_Free
    SDeg = (storagePercent * 3) + 30
        
    W.Refresh
    
    On Error GoTo 0
    Exit Sub

obtainStorage_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure obtainStorage, line " & Erl & "."

End Sub
' ----------------------------------------------------------------
' Procedure Name: Storage_Usage_Percent
' Purpose:
' Procedure Kind: Function
' Procedure Access: Private
' Return Type: Long
' Author: https://www.tek-tips.com/userinfo.cfm?member=Hypetia
' Date: 09/01/2024
' ----------------------------------------------------------------
Private Function Storage_Usage_Percent() As Long

    Storage_Usage_Percent = getDrivePercent(mThisDrive)
      
    On Error GoTo 0
    Exit Function

Storage_Usage_Percent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Storage_Usage_Percent, line " & Erl & "."

End Function

' ----------------------------------------------------------------
' Procedure Name: Storage_Usage_Free
' Purpose:
' Procedure Kind: Function
' Procedure Access: Private
' Return Type: Long
' Author: https://www.tek-tips.com/userinfo.cfm?member=Hypetia
' Date: 09/01/2024
' ----------------------------------------------------------------
Private Function Storage_Usage_Free() As Long

    Storage_Usage_Free = getDriveFree(mThisDrive)
      
    On Error GoTo 0
    Exit Function

Storage_Usage_Free_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Storage_Usage_Free, line " & Erl & "."

End Function

' ----------------------------------------------------------------
' Procedure Name: getDrivePercent
' Purpose:
' Procedure Kind: Function
' Procedure Access: Private
' Parameter sRootDrive (String):
' Return Type: Long
' Author: beededea
' Date: 12/01/2024
' ----------------------------------------------------------------
Private Function getDrivePercent(ByVal sRootDrive As String) As Long
    Dim lSecPerClust        As Long
    Dim lBytesPerSec        As Long
    Dim lFreeClusters       As Long
    Dim lTotalClusters      As Long
    
    Dim dTotalBytes         As Double
    Dim dFreeBytes          As Double
    Dim dUsedBytes          As Double
    Dim thePercentage       As Long
    
    On Error GoTo getDrivePercent_Error
    
    Call GetDiskFreeSpace(sRootDrive, lSecPerClust, lBytesPerSec, lFreeClusters, lTotalClusters)
                        
    ' Get Total Disk Space, divided by 10000 to overcome the limitation of VB6 double var type for current and future large drives.
    ' after all, we only need the final percentage...
    
    dTotalBytes = CDbl(lBytesPerSec * CDbl(lSecPerClust * (lTotalClusters / 10000)))
                        
    ' Get Free Disk Space
    dFreeBytes = lBytesPerSec * CDbl(lSecPerClust * lFreeClusters / 10000)
                            
    ' Get Used Disk Space
    dUsedBytes = dTotalBytes - dFreeBytes
    
    ' get percentage of used space
    thePercentage = dUsedBytes / dTotalBytes * 100
    
    getDrivePercent = thePercentage
    
    On Error GoTo 0
    Exit Function

getDrivePercent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getDrivePercent, line " & Erl & "."

End Function


' ----------------------------------------------------------------
' Procedure Name: getDriveFree
' Purpose:
' Procedure Type: Function
' Procedure Access: Private
' Parameter sRootDrive (String):
' Return Type: Long
' Author: beededea
' Date: 12/01/2024
' ----------------------------------------------------------------
Private Function getDriveFree(ByVal sRootDrive As String) As Long
    On Error GoTo getDriveFree_Error
    Dim lSecPerClust        As Long
    Dim lBytesPerSec        As Long
    Dim lFreeClusters       As Long
    Dim lTotalClusters      As Long
    
    Dim dFreeGig            As Double
    Dim dUsedGig            As Double

    Call GetDiskFreeSpace(sRootDrive, lSecPerClust, lBytesPerSec, lFreeClusters, lTotalClusters)
                                                
    ' Get Free Disk Space
    dFreeGig = lBytesPerSec * CDbl(lSecPerClust * lFreeClusters)
    dFreeGig = dFreeGig / 1024 ' KB
    dFreeGig = dFreeGig / 1024 ' MB
    dFreeGig = dFreeGig / 1024 ' GB
                            
    getDriveFree = dFreeGig
    
    On Error GoTo 0
    Exit Function

getDriveFree_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getDriveFree, line " & Erl & "."

End Function
