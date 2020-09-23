VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFFF&
   Caption         =   "New Pythagoras Program"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7695
   FillColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox List2 
      Height          =   4545
      Left            =   3840
      TabIndex        =   7
      Top             =   1560
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00008000&
      Caption         =   "New Calculate"
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00008000&
      Caption         =   "Old Calculate"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Enter max value (max=12500)"
      Height          =   495
      Left            =   3960
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblFound2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Number found"
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Enter max value (max=1000)"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblFound1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Number found"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "Richard Hall 2004 (revised by Shark)"
      Height          =   495
      Left            =   6240
      TabIndex        =   3
      Top             =   6240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program is to illustrate a more efficient way to calculate Pythagorean Triples
'than what was originally posted on PSC.
'The original version used a triple-nested loop which makes it an order 3 solution.
'The new version only uses a double-nested loop which makes it an order 2 solution.
'when max = 1000, the original version does approximately a billion calculations (1000*1000*1000).
'while the new version does approximately a million calculations (1000*1000), about 1000 times faster.

Private Sub Command1_Click()
'This sub illustrates an inefficient way to calculate Pythagorean Triples
'This version was from a previous post to PSC
'Although it comes up with the correct answers, it takes 'forever' to run
'It was left in its original form, with only minor modifications
Dim x
Dim y
Dim z
Dim max$
Dim ans As Long
Dim start_time As Date
Dim stop_time As Date
Dim elapsed As Double

max$ = Text1.Text

If max$ > 1000 Then
    ans = MsgBox(max$ & " is too high for this method...please revise", , "Please Revise Input")
    Exit Sub
End If

List1.Clear
start_time = Timer
For x = 1 To max$

    For y = 1 To max$ 'to eliminate duplicates, change '1' to 'x + 1'
    
        For z = 1 To max$
        
        If (x * x + y * y = z * z) Then
        List1.AddItem x & ", " & y & ", hypotenuse = " & z
        lblFound1.Caption = "Found=" & Str(List1.ListCount)
        lblFound1.Refresh
        End If
              
        Next z
        
    z = 1
        
    Next y

y = 1
    
Next x
stop_time = Timer
elapsed = Round(stop_time - start_time, 2)
lblFound1.Caption = "Total found=" & Str$(List1.ListCount) & vbCrLf & _
"Elapsed: " & Str$(elapsed) & " seconds"
End Sub

Private Sub Command2_Click()
Dim a As Long
Dim b As Long
Dim c As Double
Dim max As Long
Dim start_time As Date
Dim stop_time As Date
Dim elapsed As Double
max = Val(Text2.Text)
List2.Clear
start_time = Timer
For a = 1 To max
    For b = 1 To max 'to eliminate duplicates, change '1' to 'a + 1'
        c = Sqr(a * a + b * b)
        If c > max Then 'escape early if c is too big
                        'changing to If b > max if will allow c to be bigger
                        'than max but still have a & b <= max
            Exit For 'b
        End If
        If Int(c) = c Then 'if c is an integer, we have found a Pythagorean Triple
            'If c <= max Then 'limit results to match other calc method
                If List2.ListCount = 32736 Then 'list full...exit
                    lblFound2.Caption = "Max of 32736 reached"
                    Exit Sub
                 End If
                List2.AddItem a & ", " & b & ", hypotenuse = " & c
                lblFound2.Caption = "Found=" & Str(List2.ListCount)
                lblFound2.Refresh
            'End If
        End If

    Next 'b
Next 'a
stop_time = Timer
elapsed = Round(stop_time - start_time, 2)
lblFound2.Caption = "Total found=" & Str$(List2.ListCount) & vbCrLf & _
"Elapsed: " & Str$(elapsed) & " seconds"
End Sub


