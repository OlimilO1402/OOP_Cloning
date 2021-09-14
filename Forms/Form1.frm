VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12975
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   12975
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnInfo 
      Caption         =   "?"
      Height          =   495
      Left            =   12360
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton BtnGoBack 
      Caption         =   "<"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton BtnGoAhead 
      Caption         =   "Command1"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   11535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Top             =   720
      Width           =   12975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim State     As Long
Dim Strings() As String

Dim Person1 As Person
Dim Person2 As Person
Dim Person3 As Person
Dim Person4 As Person

'Person1 {Name: Sam; BirthD: 01.01.1900; Brain: Brain {smartness: Single; Value: 50}; City: Amsterdam}
'Person2 {Name: Sam; BirthD: 01.01.1900; Brain: Brain {smartness: Single; Value: 50}; City: Amsterdam}
'No,  Person1 and Person2 do not share the same brain.
'Yes, Person1 and Person2 are living in the same city.
'Person3 {Name: Sami; BirthD: 31.12.2000; Brain: BrainSmart {smartness: Double; Value: 100}; City: New York}
'Person4 {Name: Sami; BirthD: 31.12.2000; Brain: BrainSmart {smartness: Double; Value: 100}; City: New York}
'No,  Person3 and Person4 do not share the same brain.
'Person {Name: Sami; BirthD: 31.12.2000; Brain: BrainSmart {smartness: Double; Value: 100}; City: Tokio}
'No,  Person3 and Person4 do not live in the same city.

Private Sub BtnGoAhead_Click()
    
    Dim s As String
    Dim b As Boolean
    
    Select Case State
    
    Case 0: Set Person1 = Mnew.Person("01.01.1900", "Sam", Mnew.Brain, Mnew.City("Amsterdam"))
            s = "1. " & Person1.ToStr
    
    Case 1: Set Person2 = Person1.Clone
            s = "2. " & Person2.ToStr
    
    Case 2: b = Person1.Brain.IsSame(Person2.Brain)
            s = IIf(b, "Yes, ", "No, ") & "Person1 and Person2 " & IIf(b, "are sharing ", "do not share ") & "the same brain."
    
    Case 3: b = Person1.City.IsSame(Person2.City)
            s = IIf(b, "Yes, ", "No, ") & "Person1 and Person2 " & IIf(b, "are living in ", "do not live in ") & "the same city."
    
    
    
    Case 4: Set Person3 = Mnew.Person("31.12.2000", "Sami", Mnew.BrainSmart, Mnew.City("New York"))
            s = "3. " & Person3.ToStr
    
    Case 5: Set Person4 = Person3.Clone
            s = "4. " & Person4.ToStr
        
    Case 6: b = Person3.Brain.IsSame(Person4.Brain)
            s = IIf(b, "Yes, ", "No, ") & "Person3 and Person4 " & IIf(b, "are sharing ", "do not share ") & "the same brain."
    
    Case 7: Set Person3.City = Mnew.City("Tokio")
            s = Person3.ToStr
    
    Case 8: b = Person3.City.IsSame(Person4.City)
            s = IIf(b, "Yes, ", "No, ") & "Person3 and Person4 " & IIf(b, "are living in ", "do not live in ") & "the same city."
    
    Case 9: s = ""
    End Select
    
    Text1.Text = Text1.Text & s & vbCrLf
    State = State + 1
    If State = 10 Then
        State = 0
    End If
    BtnGoBack.Enabled = State > 0
    BtnGoAhead.Caption = Strings(State)
End Sub

Private Sub BtnGoBack_Click()
    'If State > 1 Then
    State = State - 2
    BtnGoBack.Enabled = State > 0
    BtnGoAhead_Click
End Sub

Private Sub BtnInfo_Click()
    MsgBox App.CompanyName & " " & App.ProductName & vbCrLf & _
           App.FileDescription & vbCrLf & _
           "Version: " & App.Major & "." & App.Minor & "." & App.Revision, vbInformation
End Sub

Private Sub Form_Load()
    ReDim Strings(0 To 20)
    Dim i As Long
    
    Strings(i) = "Create person1 Sam":                            i = i + 1
    Strings(i) = "Create person2 Sam by cloning Sam":             i = i + 1
    Strings(i) = "Do person1 and person2 share the same brain?":  i = i + 1
    Strings(i) = "Do person1 and person2 live in the same city?": i = i + 1
    Strings(i) = "Create person3 Sami":                           i = i + 1
    Strings(i) = "Create person4 Sami by cloning person3":        i = i + 1
    Strings(i) = "Do person3 and person4 share the same brain?":  i = i + 1
    Strings(i) = "Move person4 to Tokio":                         i = i + 1
    Strings(i) = "Do person3 and person4 live in the same city?": i = i + 1
    Strings(i) = "Start New?":                                    i = i + 1
    
    BtnGoAhead.Caption = Strings(0)
    BtnGoBack.Enabled = State > 0
'    Dim Sam1 As Person: Set Sam1 = Mnew.Person("01.01.1900", "Sam", Mnew.Brain, Mnew.City("Amsterdam"))
'    Dim Sam2 As Person: Set Sam2 = Sam1.Clone
'
'    Debug.Print Sam1.ToStr
'    Debug.Print Sam2.ToStr
'    Debug.Print "Sam1 and its clone Sam2 are sharing the same brain: " & Sam1.Brain.IsSame(Sam2.Brain)
'    Debug.Print "Sam1 and its clone Sam2 are living in the same city: " & Sam1.City.IsSame(Sam2.City)
'
'    Dim Sam3 As Person: Set Sam3 = Mnew.Person("31.12.2000", "Sami", Mnew.BrainSmart, Mnew.City("New York"))
'    Dim Sam4 As Person: Set Sam4 = Sam3.Clone: Set Sam4.City = Mnew.City("Tokio")
'
'    Debug.Print Sam3.ToStr
'    Debug.Print Sam4.ToStr
'    Debug.Print "Sam3 and its clone Sam4 are sharing the same brain: " & Sam3.Brain.IsSame(Sam4.Brain)
'    Debug.Print "Sam3 and its clone Sam4 are living in the same city: " & Sam3.City.IsSame(Sam4.City)
    
End Sub

Private Sub Form_Resize()
    Dim b As Single: b = 8 * Screen.TwipsPerPixelX
    Dim L As Single, T As Single, W As Single, H As Single
    'first put BtnInfo to the right
    H = BtnInfo.Height
    W = BtnInfo.Width
    T = BtnInfo.Top
    L = Me.ScaleWidth - W - b
    If W > 0 Then BtnInfo.Move L, T, W, H
    'then put BtnGoAhead in between
    
    W = Me.ScaleWidth - 4 * b - BtnGoBack.Width - BtnInfo.Width 'Left - b - L
    If W > 0 Then BtnGoAhead.Width = W
    W = Me.ScaleWidth '- 2 * b
    H = Me.ScaleHeight - Text1.Top '- b
    If W > 0 And H > 0 Then Text1.Move 0, Text1.Top, W, H
End Sub
