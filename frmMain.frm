VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   2115
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const tgVal As Long = 9002

Private iCnt As Long


Private Sub Command1_Click()
    Dim v(26, 1) As Double
    
    'v(1, 0) = 1:    v(1, 1) = 100:     v(1, 2) = 0
    'v(2, 0) = 2:    v(2, 1) = 100:     v(2, 2) = 0
    'v(3, 0) = 3:    v(3, 1) = 1000:     v(3, 2) = 0
    'v(4, 0) = 4:    v(4, 1) = 100:     v(4, 2) = 0
    'v(5, 0) = 5:    v(5, 1) = 100:     v(5, 2) = 0
    'v(6, 0) = 6:    v(6, 1) = 100:     v(6, 2) = 0
    'v(7, 0) = 7:    v(7, 1) = 100:     v(7, 2) = 0
    'v(8, 0) = 8:    v(8, 1) = 100:     v(8, 2) = 0
    'v(9, 0) = 9:    v(9, 1) = 100:     v(9, 2) = 0
    'v(10, 0) = 10:  v(10, 1) = 100:    v(10, 2) = 0
    '
    'v(11, 0) = 11:  v(11, 1) = 1000:    v(11, 2) = 0
    'v(12, 0) = 12:  v(12, 1) = 1000:    v(12, 2) = 0
    'v(13, 0) = 13:  v(13, 1) = 1000:    v(13, 2) = 0
    'v(14, 0) = 14:  v(14, 1) = 1000:    v(14, 2) = 0
    'v(15, 0) = 15:  v(15, 1) = 1000:    v(15, 2) = 0
    'v(16, 0) = 16:  v(16, 1) = 1000:    v(16, 2) = 0
    'v(17, 0) = 17:  v(17, 1) = 1000:    v(17, 2) = 0
    'v(18, 0) = 18:  v(18, 1) = 1000:    v(18, 2) = 0
    'v(19, 0) = 19:  v(19, 1) = 1001:    v(19, 2) = 0
    'v(20, 0) = 20:  v(20, 1) = 1000:    v(20, 2) = 0
    
    
    v(1, 1) = 100:     v(1, 0) = 0
    v(2, 1) = 100:     v(2, 0) = 0
    v(3, 1) = 100:     v(3, 0) = 0
    v(4, 1) = 100:     v(4, 0) = 0
    v(5, 1) = 100:     v(5, 0) = 0
    v(6, 1) = 100:     v(6, 0) = 0
    v(7, 1) = 100:     v(7, 0) = 0
    v(8, 1) = 100:     v(8, 0) = 0
    v(9, 1) = 100:     v(9, 0) = 0
    v(10, 1) = 1000:    v(10, 0) = 0
    
    v(11, 1) = 1000:    v(11, 0) = 0
    v(12, 1) = 1000:    v(12, 0) = 0
    v(13, 1) = 1000:    v(13, 0) = 0
    v(14, 1) = 1000:    v(14, 0) = 0
    v(15, 1) = 1000:    v(15, 0) = 0
    v(16, 1) = 1000:    v(16, 0) = 0
    v(17, 1) = 1000:    v(17, 0) = 0
    v(18, 1) = 1000:    v(18, 0) = 0
    v(19, 1) = 1001:    v(19, 0) = 0
    v(20, 1) = 1000:    v(20, 0) = 0
    
    v(21, 1) = 1000:    v(21, 0) = 0
    v(22, 1) = 1000:    v(22, 0) = 0
    v(23, 1) = 1000:    v(23, 0) = 0
    v(24, 1) = 1000:    v(24, 0) = 0
    v(25, 1) = 1000:    v(25, 0) = 0
    v(26, 1) = 1001:    v(26, 0) = 0
    'v(27, 1) = 1000:    v(27, 0) = 0
    'v(28, 1) = 1000:    v(28, 0) = 0
    'v(29, 1) = 1001:    v(29, 0) = 0
    'v(30, 1) = 1000:    v(30, 0) = 0
    
    
    Dim rtnObj As Variant
    
    Dim ii As Long
    Dim bRtn As Boolean
    
    Debug.Print "begin : " & Now & " : " & Timer
    
    iCnt = 0
    
    For ii = 1 To UBound(v)
        bRtn = False
        Debug.Print "idx = " & ii
        bRtn = add(v, ii, 0, rtnObj)
        If bRtn Then Exit For
    Next ii
    
    Debug.Print "end   : " & Now & " : " & Timer & " :: " & CStr(iCnt)
    
    If bRtn Then
        For ii = 1 To UBound(v)
            If rtnObj(ii, 0) = 1 Then
                Debug.Print CStr(ii) & " : " & rtnObj(ii, 1) & " : " & rtnObj(ii, 0)
            End If
        Next ii
    Else
        Debug.Print "Failed.."
    End If
    
    
End Sub


Private Function add(ByVal v As Variant, ByVal idx As Long, ByVal addval As Long, ByRef sv As Variant) As Boolean
    Dim bRtn As Boolean
    iCnt = iCnt + 1
    
    addval = addval + v(idx, 1)
    
    If addval > tgVal Then
        add = False
        Exit Function
    ElseIf addval = tgVal Then
        v(idx, 0) = 1
        sv = v
        add = True
        Exit Function
    ElseIf addval < tgVal Then
        v(idx, 0) = 1
        sv = v
        Do While idx < UBound(sv)
            idx = idx + 1
            bRtn = add(sv, idx, addval, sv)
            
            If bRtn Then
                add = True
                Exit Function
            Else
                sv = v
            End If
        Loop
        
        
        add = False
        Exit Function
    End If
    
End Function
