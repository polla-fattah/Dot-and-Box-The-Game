VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Five and five"
   ClientHeight    =   6420
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "PG_Helebje Title"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   Icon            =   "FORM1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   11
   ScaleMode       =   0  'User
   ScaleWidth      =   10
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuplay 
      Caption         =   "&play"
      Begin VB.Menu mnunew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnustatus 
         Caption         =   "Status..."
         Shortcut        =   ^S
      End
      Begin VB.Menu space 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private Type coordinate
        x As Single
        y As Single
    End Type
    Dim user As Single
    Dim computer As Single
    Dim pos1 As coordinate
    Dim pos2 As coordinate
    Dim cubs(8 - 1, 8 - 1) As Single
    Dim vertical(8 - 1, 9 - 1) As Integer
    Dim horizantal(9 - 1, 8 - 1) As Integer
    Dim copy_cubs(8 - 1, 8 - 1) As Single
    Dim copy_vertical(8 - 1, 9 - 1) As Integer
    Dim copy_horizantal(9 - 1, 8 - 1) As Integer
    Dim lose As Single
    Dim player As Boolean
    Dim enable As Boolean
Private Sub Form_Load()
    Call mnunew_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub mnuexit_Click()
    End
End Sub

Private Sub mnunew_Click()
    Form1.Cls
    Dim i As Integer
    Dim j As Integer
    For i = 0 To 9 - 1
        For j = 0 To 8 - 1
            vertical(j, i) = 0
            horizantal(i, j) = 0
            If i <= 8 - 1 And j <= 8 - 1 Then cubs(i, j) = 0
        Next j
    Next i
    For i = 1 To 9
        For j = 1 To 9
            Circle (i, j), 0.1
        Next j
    Next i
    pos1.x = 0
    pos1.y = 0
    enable = True
    player = True
    user = 0
    computer = 0
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim x2 As Single
    Dim y2 As Single
    Dim axis As Single
    Dim s1 As Single
    Dim s2 As Single
    x2 = regulate(x)
    y2 = regulate(y)
    If x2 = Int(x2) And y2 = Int(y2) And x2 > 0 And x2 < 10 And y2 > 0 And y2 < 10 Then
        If (pos1.x + pos1.y <> 0) And ((pos1.x - x2) ^ 2 + (pos1.y - y2) ^ 2) ^ (1 / 2) = 1 Then
            If pos1.x = x2 Then axis = 1
            If pos1.y = y2 Then axis = 2
            If pos1.y < y2 Then
                pos2.y = pos1.y
            Else
                pos2.y = y2
            End If
            If pos1.x < x2 Then
                pos2.x = pos1.x
            Else
                pos2.x = x2
            End If
            If pos1.x = x2 And pos1.y < y2 Then Call record_V(pos1.x, pos1.y)
            If pos1.x = x2 And pos1.y > y2 Then Call record_V(x2, y2)
            If pos1.y = y2 And pos1.x < x2 Then Call record_H(pos1.x, pos1.y)
            If pos1.y = y2 And pos1.x > x2 Then Call record_H(x2, y2)
            If enable Then Call user_win(pos2, axis)
                
            If enable And player Then
                Call arbitrate(pos2.x, pos2.y, axis)
                Call computer_play
            End If
            enable = True
            player = True
            pos1.x = 0
            pos1.y = 0
        Else
            pos1.x = x2
            pos1.y = y2
        End If
        
    End If
    Dim ending As Boolean
    ending = True
    For x2 = 0 To 8 - 1
        For y2 = 0 To 8 - 1
            If cubs(x2, y2) = 0 Then
                ending = False
                Exit For
            End If
        Next
        If Not ending Then Exit For
    Next
    If ending Then
        If user > computer Then MsgBox "congragulation you are win" + Chr(13) + "you get " + Str(user) + " and computer gets " + Str(computer)
        If user < computer Then MsgBox "sorry you are lost" + Chr(13) + "you get " + Str(user) + " and computer gets " + Str(computer)
        If user = computer Then MsgBox "congragulation it is equal game" + Chr(13) + "you get " + Str(user) + " and computer gets " + Str(computer)
    End If
End Sub
Private Sub user_win(pos As coordinate, axis As Single)
    Dim s1 As Single
    Dim s2 As Single
    If axis = 1 Then vertical(pos.y - 1, pos.x - 1) = 1
    If axis = 2 Then horizantal(pos.y - 1, pos.x - 1) = 1
    Call sums(pos.y - 1, pos.x - 1, axis, s1, s2)
    If axis = 1 Then
        Line (pos.x, pos.y)-(pos.x, pos.y + 1)
        If s1 = 4 Then
            Line (pos.x + 0.05, pos.y + 0.05)-(pos.x + 0.95, pos.y + 0.95), RGB(0, 120, 220), BF
            cubs(pos.y - 1, pos.x - 1) = 1
            user = user + 1
            player = False
        End If
        If s2 = 4 Then
            Line (pos.x - 0.95, pos.y + 0.95)-(pos.x - 0.05, pos.y + 0.05), RGB(0, 120, 220), BF
            cubs(pos.y - 1, pos.x - 2) = 1
            user = user + 1
            player = False
        End If
    End If
    If axis = 2 Then
        Line (pos.x, pos.y)-(pos.x + 1, pos.y)
        If s1 = 4 Then
            Line (pos.x + 0.05, pos.y + 0.05)-(pos.x + 0.95, pos.y + 0.95), RGB(0, 120, 220), BF
            cubs(pos.y - 1, pos.x - 1) = 1
            user = user + 1
            player = False
        End If
        If s2 = 4 Then
            Line (pos.x + 0.05, pos.y - 0.05)-(pos.x + 0.95, pos.y - 0.95), RGB(0, 120, 220), BF
            cubs(pos.y - 2, pos.x - 1) = 1
            user = user + 1
            player = False
        End If
    End If
End Sub

Private Function regulate(a As Single)
    regulate = a
    If Abs(a - Int(a)) < 0.1 Then regulate = Int(a)
    If Abs(a - Int(a)) > 0.9 Then regulate = Int(a) + 1
End Function
Private Sub record_V(x As Single, y As Single)
    If vertical(y - 1, x - 1) = 1 Then enable = False
End Sub
Private Sub record_H(x As Single, y As Single)
    If horizantal(y - 1, x - 1) = 1 Then enable = False
End Sub
Private Sub arbitrate(x As Single, y As Single, axis As Single)
    Dim a As Double
    Dim temp As Single
    For a = 0 To 1000000: Next a
    If axis = 1 Then
        vertical(y - 1, x - 1) = 1
        Line (x, y)-(x, y + 1)
            If x > 0 And y > 0 And x < 9 And y < 9 Then
                temp = vertical(y - 1, x - 1) + vertical(y - 1, x) + horizantal(y - 1, x - 1) + horizantal(y, x - 1)
                If (temp = 3 Or (temp = 4 And cubs(y - 1, x - 1) = 0)) Then
                    Line (x + 0.05, y + 0.05)-(x + 0.95, y + 0.95), RGB(255, 120, 220), BF
                    cubs(y - 1, x - 1) = 1
                    computer = computer + 1
                    If vertical(y - 1, x) = 0 Then Call arbitrate(x + 1, y, 1)
                    If horizantal(y - 1, x - 1) = 0 Then Call arbitrate(x, y, 2)
                    If horizantal(y, x - 1) = 0 Then Call arbitrate(x, y + 1, 2)
                End If
            End If
            If y >= 1 And x > 1 And y < 9 And x <= 9 Then
                temp = vertical(y - 1, x - 2) + vertical(y - 1, x - 1) + horizantal(y - 1, x - 2) + horizantal(y, x - 2)
                If (temp = 3 Or (temp = 4 And cubs(y - 1, x - 2) = 0)) Then
                    Line (x - 0.95, y + 0.95)-(x - 0.05, y + 0.05), RGB(255, 120, 220), BF
                    cubs(y - 1, x - 2) = 1
                    computer = computer + 1
                    If vertical(y - 1, x - 2) = 0 Then Call arbitrate(x - 1, y, 1)
                    If horizantal(y - 1, x - 2) = 0 Then Call arbitrate(x - 1, y, 2)
                    If horizantal(y, x - 2) = 0 Then Call arbitrate(x - 1, y + 1, 2)
                End If
            End If
    End If
    If axis = 2 Then
        horizantal(y - 1, x - 1) = 1
        Line (x, y)-(x + 1, y)
            If y > 0 And x > 0 And y < 9 And x < 9 Then
                temp = vertical(y - 1, x - 1) + vertical(y - 1, x) + horizantal(y - 1, x - 1) + horizantal(y, x - 1)
                If (temp = 3 Or (temp = 4 And cubs(y - 1, x - 1) = 0)) Then
                    Line (x + 0.05, y + 0.05)-(x + 0.95, y + 0.95), RGB(255, 120, 220), BF
                    cubs(y - 1, x - 1) = 1
                    computer = computer + 1
                    If vertical(y - 1, x - 1) = 0 Then Call arbitrate(x, y, 1)
                    If vertical(y - 1, x) = 0 Then Call arbitrate(x + 1, y, 1)
                    If horizantal(y, x - 1) = 0 Then Call arbitrate(x, y + 1, 2)
                End If
            End If
            If y > 1 And x > 0 And x < 9 And y <= 9 Then
                temp = vertical(y - 2, x - 1) + vertical(y - 2, x) + horizantal(y - 1, x - 1) + horizantal(y - 2, x - 1)
                If (temp = 3 Or (temp = 4 And cubs(y - 2, x - 1) = 0)) Then
                    Line (x + 0.05, y - 0.05)-(x + 0.95, y - 0.95), RGB(255, 120, 220), BF
                    cubs(y - 2, x - 1) = 1
                    computer = computer + 1
                    If vertical(y - 2, x - 1) = 0 Then Call arbitrate(x, y - 1, 1)
                    If vertical(y - 2, x) = 0 Then Call arbitrate(x + 1, y - 1, 1)
                    If horizantal(y - 2, x - 1) = 0 Then Call arbitrate(x, y - 1, 2)
                End If
            End If
    End If
End Sub
Private Sub computer_play()
    Randomize
    Dim i As Single
    Dim j As Single
    Dim axis As Single
    Dim s1 As Single
    Dim s2 As Single
    Dim lop As Integer
    Dim a As String
    For i = 0 To 9 - 1
        For j = 0 To 8 - 1
            If vertical(j, i) = 0 Then
                Call sums(j, i, 1, s1, s2)
                If s1 = 3 Or s2 = 3 Then Call arbitrate(i + 1, j + 1, 1)
            End If
            If horizantal(i, j) = 0 Then
                Call sums(i, j, 2, s1, s2)
                If s1 = 3 Or s2 = 3 Then Call arbitrate(j + 1, i + 1, 2)
            End If
        Next j
    Next i
    For lop = 1 To 30
        axis = Int(Rnd * 2) + 1
        i = Int(Rnd * 8) + 1
        j = Int(Rnd * 8) + 1
        If axis = 1 Then Call record_V(i, j)
        If axis = 2 Then Call record_H(i, j)
        Call sums(i - 1, j - 1, axis, s1, s2)
        If enable And s1 < 2 And s2 < 2 Then
            Call arbitrate(j, i, axis)
            enable = False
            Exit For
        End If
        enable = True
    Next lop
    If enable Then Call min_lost
    enable = True
End Sub
Private Sub sums(i As Single, j As Single, axis As Single, ByRef sum1 As Single, ByRef sum2 As Single)
    If axis = 1 Then
        If j < 9 - 1 Then
            sum1 = vertical(i, j) + horizantal(i, j) + vertical(i, j + 1) + horizantal(i + 1, j)
        Else
            sum1 = 0
        End If
        If j > 0 Then
            sum2 = vertical(i, j - 1) + horizantal(i, j - 1) + vertical(i, j) + horizantal(i + 1, j - 1)
        Else
            sum2 = 0
        End If
    End If
    If axis = 2 Then
        If i < 9 - 1 Then
            sum1 = vertical(i, j) + horizantal(i, j) + vertical(i, j + 1) + horizantal(i + 1, j)
        Else
            sum1 = 0
        End If
        If i > 0 Then
             sum2 = vertical(i - 1, j) + horizantal(i - 1, j) + vertical(i - 1, j + 1) + horizantal(i, j)
        Else
            sum2 = 0
        End If
    End If
End Sub
Private Sub min_lost()
    Dim array1(500) As coordinate
    Dim array2(500) As Single
    Dim array3(500) As Single
    Dim Index As Integer
    Dim i, i1, j1, j As Single
    Dim s As String
    
    For i1 = 0 To 9 - 1
        For j1 = 0 To 8 - 1
            copy_vertical(j1, i1) = vertical(j1, i1)
            copy_horizantal(i1, j1) = horizantal(i1, j1)
            If i <= 8 - 1 And j <= 8 - 1 Then copy_cubs(i, j) = cubs(i, j)
        Next j1
    Next i1

    For i = 0 To 9 - 1
        For j = 0 To 8 - 1
            If copy_horizantal(i, j) = 0 Then
            lose = 0
            Call copy_arbitrate(j + 1, i + 1, 2)
            array1(Index).x = j + 1
            array1(Index).y = i + 1
            array2(Index) = lose
            array3(Index) = 2
            Index = Index + 1
            
            For i1 = 0 To 9 - 1
                For j1 = 0 To 8 - 1
                    copy_vertical(j1, i1) = vertical(j1, i1)
                    copy_horizantal(i1, j1) = horizantal(i1, j1)
                    If i1 <= 8 - 1 Then copy_cubs(i1, j1) = cubs(i1, j1)
                Next j1
            Next i1
            
        End If
    Next j: Next i
    
    For i = 0 To 8 - 1
        For j = 0 To 9 - 1
            If copy_vertical(i, j) = 0 Then
            lose = 0
            Call copy_arbitrate(j + 1, i + 1, 1)
            array1(Index).x = j + 1
            array1(Index).y = i + 1
            array2(Index) = lose
            array3(Index) = 1
            Index = Index + 1
            
            For i1 = 0 To 9 - 1
                For j1 = 0 To 8 - 1
                    copy_vertical(j1, i1) = vertical(j1, i1)
                    copy_horizantal(i1, j1) = horizantal(i1, j1)
                    If i1 <= 8 - 1 Then copy_cubs(i1, j1) = cubs(i1, j1)
                Next j1
            Next i1
            
        End If
    Next j: Next i
    Dim min As Integer
    Dim min_index As Integer
    min = array2(0)
    min_index = 0
    For i = 0 To Index - 1
        s = s + Str(array2(i)) + "  "
        If array2(i) < min Then
            min = array2(i)
            min_index = i
        End If
    Next i
    If array3(min_index) = 1 Then
        vertical(array1(min_index).y - 1, array1(min_index).x - 1) = 1
        Line (array1(min_index).x, array1(min_index).y)-(array1(min_index).x, array1(min_index).y + 1)
    End If
    If array3(min_index) = 2 Then
        horizantal(array1(min_index).y - 1, array1(min_index).x - 1) = 1
        Line (array1(min_index).x, array1(min_index).y)-(array1(min_index).x + 1, array1(min_index).y)
    End If
End Sub
Private Sub copy_arbitrate(x As Single, y As Single, axis As Single)
    Dim temp As Single
    If axis = 1 Then
        copy_vertical(y - 1, x - 1) = 1
            If x > 0 And y > 0 And x < 9 And y < 9 Then
                temp = copy_vertical(y - 1, x - 1) + copy_vertical(y - 1, x) + copy_horizantal(y - 1, x - 1) + copy_horizantal(y, x - 1)
                If (temp = 3 Or (temp = 4 And copy_cubs(y - 1, x - 1) = 0)) Then
                    copy_cubs(y - 1, x - 1) = 1
                    lose = lose + 1
                    If copy_vertical(y - 1, x) = 0 Then Call copy_arbitrate(x + 1, y, 1)
                    If copy_horizantal(y - 1, x - 1) = 0 Then Call copy_arbitrate(x, y, 2)
                    If copy_horizantal(y, x - 1) = 0 Then Call copy_arbitrate(x, y + 1, 2)
                End If
            End If
            If y >= 1 And x > 1 And y < 9 And x <= 9 Then
                temp = copy_vertical(y - 1, x - 2) + copy_vertical(y - 1, x - 1) + copy_horizantal(y - 1, x - 2) + copy_horizantal(y, x - 2)
                If (temp = 3 Or (temp = 4 And copy_cubs(y - 1, x - 2) = 0)) Then
                    copy_cubs(y - 1, x - 2) = 1
                    lose = lose + 1
                    If copy_vertical(y - 1, x - 2) = 0 Then Call copy_arbitrate(x - 1, y, 1)
                    If copy_horizantal(y - 1, x - 2) = 0 Then Call copy_arbitrate(x - 1, y, 2)
                    If copy_horizantal(y, x - 2) = 0 Then Call copy_arbitrate(x - 1, y + 1, 2)
                End If
            End If
    End If
    If axis = 2 Then
        copy_horizantal(y - 1, x - 1) = 1
            If y > 0 And x > 0 And y < 9 And x < 9 Then
                temp = copy_vertical(y - 1, x - 1) + copy_vertical(y - 1, x) + copy_horizantal(y - 1, x - 1) + copy_horizantal(y, x - 1)
                If (temp = 3 Or (temp = 4 And copy_cubs(y - 1, x - 1) = 0)) Then
                    copy_cubs(y - 1, x - 1) = 1
                    lose = lose + 1
                    If copy_vertical(y - 1, x - 1) = 0 Then Call copy_arbitrate(x, y, 1)
                    If copy_vertical(y - 1, x) = 0 Then Call copy_arbitrate(x + 1, y, 1)
                    If copy_horizantal(y, x - 1) = 0 Then Call copy_arbitrate(x, y + 1, 2)
                End If
            End If
            If y > 1 And x > 0 And x < 9 And y <= 9 Then
                temp = copy_vertical(y - 2, x - 1) + copy_vertical(y - 2, x) + copy_horizantal(y - 1, x - 1) + copy_horizantal(y - 2, x - 1)
                If (temp = 3 Or (temp = 4 And copy_cubs(y - 2, x - 1) = 0)) Then
                    copy_cubs(y - 2, x - 1) = 1
                    lose = lose + 1
                    If copy_vertical(y - 2, x - 1) = 0 Then Call copy_arbitrate(x, y - 1, 1)
                    If copy_vertical(y - 2, x) = 0 Then Call copy_arbitrate(x + 1, y - 1, 1)
                    If copy_horizantal(y - 2, x - 1) = 0 Then Call copy_arbitrate(x, y - 1, 2)
                End If
            End If
    End If
End Sub


Private Sub mnustatus_Click()
MsgBox "Till now you get " + Str(user) + " and Computer gets " + Str(computer)
End Sub
