VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12975
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   12975
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   12735
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
      Height          =   4575
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Top             =   720
      Width           =   12735
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

Dim Sam1 As Person
Dim Sam2 As Person
Dim Sam3 As Person
Dim Sam4 As Person

'Person {Name: Sam; BirthD: 01.01.1900; Brain: Brain {smartness: Single; Value: 50}; City: Amsterdam}
'Person {Name: Sam; BirthD: 01.01.1900; Brain: Brain {smartness: Single; Value: 50}; City: Amsterdam}
'Sam1 and its clone Sam2 are sharing the same brain: Falsch
'Sam1 and its clone Sam2 are living in the same city: Wahr
'Person {Name: Sami; BirthD: 31.12.2000; Brain: BrainSmart {smartness: Double; Value: 100}; City: New York}
'Person {Name: Sami; BirthD: 31.12.2000; Brain: BrainSmart {smartness: Double; Value: 100}; City: New York}
'Sam3 and its clone Sam4 are sharing the same brain: Falsch
'Sam3 and its clone Sam4 are living in the same city: Wahr

Private Sub Command1_Click()
    
    Dim s As String
    Dim b As Boolean
    
    Select Case State
    
    Case 0: Set Sam1 = Mnew.Person("01.01.1900", "Sam", Mnew.Brain, Mnew.City("Amsterdam"))
            s = Sam1.ToStr
    
    Case 1: Set Sam2 = Sam1.Clone
            s = Sam2.ToStr
    
    Case 2: b = Sam1.Brain.IsSame(Sam2.Brain)
            s = IIf(b, "Yes ", "No ") & "Sam1 and Sam2 " & IIf(b, "are sharing ", "do not share ") & "the same brain."
    
    Case 3: b = Sam1.City.IsSame(Sam2.City)
            s = IIf(b, "Yes ", "No ") & "Sam1 and Sam2 " & IIf(b, "are living in ", "do not live in ") & "the same city."
    
    
    
    Case 4: Set Sam3 = Mnew.Person("31.12.2000", "Sami", Mnew.BrainSmart, Mnew.City("New York"))
            s = Sam3.ToStr
    
    Case 5: Set Sam4 = Sam3.Clone
            s = Sam4.ToStr
        
    Case 6: b = Sam3.Brain.IsSame(Sam4.Brain)
            s = IIf(b, "Yes ", "No ") & "Sam3 and Sam4 " & IIf(b, "are sharing ", "do not share ") & "the same brain."
    
    Case 7: b = Sam3.City.IsSame(Sam4.City)
            s = IIf(b, "Yes ", "No ") & "Sam3 and Sam4 " & IIf(b, "are living in ", "do not live in ") & "the same city."
    
    
    End Select
    
    State = State + 1
    If State = 8 Then State = 0
    Command1.Caption = Strings(State)
    Text1.Text = Text1.Text & s & vbCrLf
End Sub

Private Sub Form_Load()
    ReDim Strings(0 To 20)
    Dim i As Long
    
    Strings(i) = "Create person Sam1":                  i = i + 1
    Strings(i) = "Create person Sam2 by cloning Sam1":  i = i + 1
    Strings(i) = "Do Sam1 Sam2 share the same brain?":  i = i + 1
    Strings(i) = "Do Sam1 Sam2 live in the same city?": i = i + 1
    Strings(i) = "Create Person Sam3":                  i = i + 1
    Strings(i) = "Create Person Sam4 by Cloning Sam3":  i = i + 1
    Strings(i) = "Do Sam3 Sam4 share the same brain?":  i = i + 1
    Strings(i) = "Do Sam3 Sam4 live in the same city?": i = i + 1
    
    Command1.Caption = Strings(0)
    
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
