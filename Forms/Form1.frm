VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Object Constructors And Cloning"
   ClientHeight    =   9780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12975
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9780
   ScaleWidth      =   12975
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnInfo 
      Caption         =   "?"
      Height          =   495
      Left            =   12480
      TabIndex        =   3
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton BtnExampleModalDialog 
      Caption         =   "Example Modal Dialog"
      Height          =   495
      Left            =   9960
      TabIndex        =   4
      Top             =   0
      Width           =   2535
   End
   Begin VB.CommandButton BtnGoAhead 
      Caption         =   "Command1"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   0
      Width           =   9495
   End
   Begin VB.CommandButton BtnGoBack 
      Caption         =   "<"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   495
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
      Height          =   9015
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Top             =   480
      Width           =   12975
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim State     As Long
Dim Commands() As String
Dim Response() As String
Dim forw      As Boolean

Dim Person1 As Person
Dim Person2 As Person
Dim Person3 As Person
Dim Person4 As Person

Private Sub Form_Load()
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    PrepareNCreateStrings
    BtnGoAhead.Caption = Commands(1)
    BtnGoBack.Enabled = State > 0
End Sub

Sub PrepareNCreateStrings()
    Set Person1 = MNew.Person("01.01.1900", MNew.Brain, MNew.City("Amsterdam"), 1, "Sam")
    Set Person2 = Person1.Clone: Person2.IndexInc
    Set Person3 = MNew.Person("31.12.2000", MNew.BrainSmart, MNew.City("New York"), 3, "Sami")
    Set Person4 = Person3.Clone: Person4.IndexInc
    
    ReDim Commands(0 To 20)
    ReDim Response(0 To 20)
    
    Dim b As Boolean
    Dim i As Long: i = 1
    Commands(i) = i & ". Create Person1 Sam:"
    Response(i) = "   " & Person1.ToStr:                                 i = i + 1
    
    Commands(i) = i & ". Create Person2 Sam by cloning Sam:"
    Response(i) = "   " & Person2.ToStr:                                 i = i + 1
    
    Commands(i) = i & ". Do Person1 and Person2 share the same brain?"
    b = Person1.Brain.IsSame(Person2.Brain)
    Response(i) = "   " & IIf(b, "Yes, ", "No, ") & "Person1 and Person2 " & IIf(b, "are sharing ", "do not share ") & "the same brain.":    i = i + 1
    
    Commands(i) = i & ". Do Person1 and Person2 live in the same city?"
    b = Person1.City.IsSame(Person2.City)
    Response(i) = "   " & IIf(b, "Yes, ", "No, ") & "Person1 and Person2 " & IIf(b, "are living in ", "do not live in ") & "the same city.": i = i + 1
    
    Commands(i) = i & ". Create Person3 Sami:"
    Response(i) = "   " & Person3.ToStr:                                 i = i + 1
    
    Commands(i) = i & ". Create Person4 Sami by cloning Person3:"
    Response(i) = "   " & Person4.ToStr:                                 i = i + 1
    
    Commands(i) = i & ". Do Person3 and Person4 share the same brain?"
    b = Person3.Brain.IsSame(Person4.Brain)
    Response(i) = "   " & IIf(b, "Yes, ", "No, ") & "Person3 and Person4 " & IIf(b, "are sharing ", "do not share ") & "the same brain.":    i = i + 1
    
    Commands(i) = i & ". Move Person4 to Tokio:"
    Set Person4.City = MNew.City("Tokio")
    Response(i) = "   " & Person4.ToStr:                                 i = i + 1
    
    Commands(i) = i & ". Do Person3 and Person4 live in the same city?"
    b = Person3.City.IsSame(Person4.City)
    Response(i) = "   " & IIf(b, "Yes, ", "No, ") & "Person3 and Person4 " & IIf(b, "are living in ", "do not live in ") & "the same city.": i = i + 1
    
    Commands(i) = "Start New?":                                          i = i + 1
    Set Person4.City = MNew.City("New York")
    Response(i) = "   "
End Sub

Sub State_Inc()
    State = State + 1
    forw = True
End Sub

Sub State_Dec()
    State = State - 1
    If forw Then State = State - 1
    forw = False
End Sub

Private Sub Form_Resize()
    Dim L As Single, T As Single, W As Single, H As Single
    T = Text1.Top
    W = Me.ScaleWidth
    H = Me.ScaleHeight - Text1.Top
    If W > 0 And H > 0 Then Text1.Move L, T, W, H
End Sub

Private Sub BtnExampleModalDialog_Click()
    If MData.Persons.Count = 0 Then
        MData.Persons_Add Person1.Clone
        MData.Persons_Add Person2.Clone
        MData.Persons_Add Person3.Clone
        MData.Persons_Add Person4.Clone
    End If
    FPersons.Show
End Sub

'1. Create Person1 Sam:
'   Person1 {Sam; 01.01.1900; Brain: Brain {smartness: Single; Value: 50}; City: Amsterdam}
'2. Create Person2 Sam by cloning Sam:
'   Person2 {Sam; 01.01.1900; Brain: Brain {smartness: Single; Value: 50}; City: Amsterdam}
'3. Do Person1 and Person2 share the same brain?
'   No, Person1 and Person2 do not share the same brain.
'4. Do Person1 and Person2 live in the same city?
'   Yes, Person1 and Person2 are living in the same city.
'5. Create Person3 Sami:
'   Person3 {Sami; 31.12.2000; Brain: BrainSmart {smartness: Double; Value: 100}; City: New York}
'6. Create Person4 Sami by cloning Person3:
'   Person4 {Sami; 31.12.2000; Brain: BrainSmart {smartness: Double; Value: 100}; City: New York}
'7. Do Person3 and Person4 share the same brain?
'   No, Person3 and Person4 do not share the same brain.
'8. Move Person4 to Tokio:
'   Person4 {Sami; 31.12.2000; Brain: BrainSmart {smartness: Double; Value: 100}; City: Tokio}
'9. Do Person3 and Person4 live in the same city?
'   No, Person3 and Person4 do not live in the same city.
'Start New?

Private Sub BtnGoAhead_Click()
    
    DoItNLogIt
    State_Inc
    
    If State = 10 Then
        PrepareNCreateStrings
        State = 0
    End If
    BtnGoBack.Enabled = State > 0
    BtnGoAhead.Caption = Commands(State)
End Sub

Private Sub BtnGoBack_Click()
    
    DoItNLogIt
    State_Dec
    
    If State = 10 Then State = 0
    BtnGoBack.Enabled = State > 0
    BtnGoAhead.Caption = Commands(State)
End Sub

Private Sub DoItNLogIt()
    
    Dim s As String
    Dim b As Boolean

'    Select Case State
'
'    'Case 0: s = "   " & Person1.ToStr
'
'    'Case 1: s = "   " & Person2.ToStr
'
'    'Case 2: b = Person1.Brain.IsSame(Person2.Brain)
'    '        s = "   " & IIf(b, "Yes, ", "No, ") & "Person1 and Person2 " & IIf(b, "are sharing ", "do not share ") & "the same brain."
'
'    'Case 3: b = Person1.City.IsSame(Person2.City)
'    '        s = "   " & IIf(b, "Yes, ", "No, ") & "Person1 and Person2 " & IIf(b, "are living in ", "do not live in ") & "the same city."
'
'    'Case 4: s = "   " & Person3.ToStr
'
'    'Case 5: s = "   " & Person4.ToStr
'
'    'Case 6: b = Person3.Brain.IsSame(Person4.Brain)
'    '        s = "   " & IIf(b, "Yes, ", "No, ") & "Person3 and Person4 " & IIf(b, "are sharing ", "do not share ") & "the same brain."
'
'    'Case 7
'
'    'Case 8: b = Person3.City.IsSame(Person4.City)
'    '        s = "   " & IIf(b, "Yes, ", "No, ") & "Person3 and Person4 " & IIf(b, "are living in ", "do not live in ") & "the same city."
'
'    'Case 9:
'    End Select
    s = Response(State)
    
    Text1.Text = Text1.Text & BtnGoAhead.Caption & vbCrLf & s & vbCrLf

End Sub

Private Sub BtnInfo_Click()
    MsgBox App.CompanyName & " " & App.ProductName & vbCrLf & _
           App.FileDescription & vbCrLf & _
           "Version: " & App.Major & "." & App.Minor & "." & App.Revision, vbInformation
End Sub
