VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cfMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cfMain
' Author    : beededea
' Date      : 28/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Public WithEvents aboutForm As cWidgetForm
Attribute aboutForm.VB_VarHelpID = -1
Public WithEvents planetForm As cWidgetForm
Attribute planetForm.VB_VarHelpID = -1


'---------------------------------------------------------------------------------------
' Procedure : InitAndShowAsFreeForm
' Author    :
' Date      : 27/04/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub InitAndShowAsFreeForm(ByVal X As Long, ByVal Y As Long, Optional ByVal dx As Long, Optional ByVal dy As Long, Optional ByVal Caption As String)
    
    Const WFFStyleTool As Integer = 7 ' additional styles above the normal five  WFFStyleTool = 7 makes the taskbar button go away
    
    On Error GoTo InitAndShowAsFreeForm_Error
    
    ' create an invisible form using a Cairo Widget Form with the predefined caption and location
    Set planetForm = Cairo.WidgetForms.Create(WFFStyleTool, Caption, , 1200, 1200)
    
    ' make the form transparent
    planetForm.WidgetRoot.BackColor = -1

    ' create a new planet widget with a name, location and width, runs Class_Initialize
    Set planetWidget = planetForm.Widgets.Add(New cwPlanet, "planet widget", 0, 0, 1000, 1000)
        
    ' I do not 'get' the size/position relationship between the form and the widget on the form. Seems peculiar to me.
      
    'NOTE that when you move a widget by dragging, you are moving the invisible form it is drawn upon.
    planetForm.Move X, Y ' position the form & display it
    planetForm.Load
    
    ' set the z-ordering of the main form
    Call setWindowZordering
    
    ' create a second invisible form using a Cairo Widget Form with the predefined caption and sizing
    Set aboutForm = Cairo.WidgetForms.Create(WFFStyleTool, Caption, , 350, 675)
    
    ' make the about form transparent
    aboutForm.WidgetRoot.BackColor = -1
         
    ' create a new planet widget with a name and location, first step of that is to run Class_Initialize
    Set aboutWidget = aboutForm.Widgets.Add(New cwAbout, "about", 0, 0, 470, 670)

   On Error GoTo 0
   Exit Sub

InitAndShowAsFreeForm_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InitAndShowAsFreeForm of Class Module cfMain"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : planetForm_DblClick
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub planetForm_DblClick()
    Dim userprof As String: userprof = vbNullString
    Dim thisCommand As String: thisCommand = vbNullString
    
    On Error GoTo planetForm_DblClick_Error

    If LTrim$(gblPlDblClickCommand) = vbNullString Then Exit Sub
    
    thisCommand = gblPlDblClickCommand
        
    If InStr(thisCommand, "%userprofile%") Then
        userprof = Environ$("USERPROFILE")
        thisCommand = Replace(thisCommand, "%userprofile%", userprof)
    End If
    
    ' .91 DAEB 08/12/2022 frmMain.frm SteamyDock responds to %systemroot% environment variables during runCommand
    If InStr(thisCommand, "%systemroot%") Then
        userprof = Environ$("SYSTEMROOT")
        thisCommand = Replace(thisCommand, "%systemroot%", userprof)
    End If
    
    If SHIFT_1 = True Then
        SHIFT_1 = False
        Call ShellExecute(fMain.planetForm.hwnd, "Open", gblPlOpenFile, vbNullString, App.Path, 1)
    Else
        Call ShellExecute(fMain.planetForm.hwnd, "runas", thisCommand, vbNullString, App.Path, 1)
    End If

    On Error GoTo 0
    Exit Sub

planetForm_DblClick_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure planetForm_DblClick of Class Module cfMain"
            Resume Next
          End If
    End With
End Sub

 
'Private Sub planetForm_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'        'If pinClicked = True Then MsgBox "X = " & x & " Y = " & y & " pinClicked = " & pinClicked
'
'End Sub

'---------------------------------------------------------------------------------------
' Procedure : planetForm_MouseMove
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub planetForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo planetForm_MouseMove_Error

    If planetWidget.Locked = True Then Exit Sub
    If gblPlIgnoreMouse = "1" Then Exit Sub
    
    Static x0 As Single
    Static y0 As Single
    If Button Then planetForm.Move planetForm.Left + X - x0, planetForm.Top + Y - y0 Else x0 = X: y0 = Y

    On Error GoTo 0
    Exit Sub

planetForm_MouseMove_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure planetForm_MouseMove of Class Module cfMain"
            Resume Next
          End If
    End With
End Sub
'---------------------------------------------------------------------------------------
' Procedure : planetForm_KeyDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : get F5 and SHIFT keypresses
'---------------------------------------------------------------------------------------
'
Private Sub planetForm_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
    On Error GoTo planetForm_KeyDown_Error

    Call getKeyPress(KeyCode, Shift)

    On Error GoTo 0
    Exit Sub

planetForm_KeyDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure planetForm_KeyDown of Class Module cfMain"
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : planetForm_MouseUp
' Author    : beededea
' Date      : 05/05/2023
' Purpose   : save the form x,y position when ever the planet/form is dragged
'---------------------------------------------------------------------------------------
'
Private Sub planetForm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo planetForm_MouseUp_Error
    
    If planetWidget.Locked = True Then Exit Sub
    
    ' no point in saving when prefs are visible.
    If planetPrefs.IsVisible = False Then ' checks without activating the form - important!.
       gblPlMaximiseFormX = Str$(planetForm.Left) ' saving in pixels
       gblPlMaximiseFormY = Str$(planetForm.Top)
       
       sPutINISetting softwarePlanet, "maximiseFormX", gblPlMaximiseFormX, gblPlSettingsFile
       sPutINISetting softwarePlanet, "maximiseFormY", gblPlMaximiseFormY, gblPlSettingsFile
    End If
    
    'Static x0, y0: If Button Then planetForm.Move planetForm.Left + X - x0, planetForm.Top + Y - y0 Else x0 = X: y0 = Y

    On Error GoTo 0
    Exit Sub

planetForm_MouseUp_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure planetForm_MouseUp of Class Module cfMain"
            Resume Next
          End If
    End With
End Sub
