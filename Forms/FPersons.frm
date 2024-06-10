VERSION 5.00
Begin VB.Form FPersons 
   Caption         =   "Persons"
   ClientHeight    =   2865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14145
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   14145
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      ItemData        =   "FPersons.frx":0000
      Left            =   0
      List            =   "FPersons.frx":0002
      TabIndex        =   0
      ToolTipText     =   "DoubleClick to Edit"
      Top             =   0
      Width           =   13815
   End
End
Attribute VB_Name = "FPersons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    MData.Persons_ToListBox List1
End Sub

Public Sub Persons_Add(Person As Person)
    MData.Persons_Add Person
End Sub

Private Sub List1_DblClick()
    Dim i   As Long:     i = List1.ListIndex
    Dim key As String: key = List1.ItemData(i)
    
    Dim p As Person: Set p = MData.Persons_Item(key)
    
    If FPerson.ShowDialog(p, Me) = vbCancel Then Exit Sub
    List1.List(i) = p.ToStr
End Sub
