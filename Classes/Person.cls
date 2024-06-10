VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_BirthD As Date
Private m_Brain  As Brain
Private m_City   As City
Private m_Index  As Long
Private m_Name   As String

Friend Sub New_(BirthDay As Date, Brain As Brain, City As City, Index As Long, Name As String)
    m_BirthD = BirthDay: Set m_Brain = Brain: Set m_City = City: m_Index = Index: m_Name = Name
End Sub

Friend Sub NewC(other As Person)
    With other
        m_BirthD = .BirthDay
        Set m_Brain = .Brain.Clone
        Set m_City = .City.Clone
        m_Name = .Name
        m_Index = .Index
    End With
End Sub

Friend Function Clone() As Person
    Set Clone = New Person: Clone.NewC Me
End Function

Friend Function IsEqual(other As Person) As Boolean
    IsEqual = Me.BirthDay = other.BirthDay: If Not IsEqual Then Exit Function
    IsEqual = Me.Name = other.Name:         If Not IsEqual Then Exit Function
    IsEqual = Brain.IsEqual(other.Brain):   If Not IsEqual Then Exit Function
    IsEqual = City.IsEqual(other.City):     If Not IsEqual Then Exit Function
End Function

Friend Function IsSame(other As Person) As Boolean
    If Me Is other Then IsSame = ObjPtr(Me) = ObjPtr(other)
End Function

Public Property Get BirthDay() As Date
    BirthDay = m_BirthD
End Property

Public Property Get Brain() As Brain
    Set Brain = m_Brain
End Property

Public Property Get City() As City
    Set City = m_City
End Property
Public Property Set City(ByVal Value As City)
    Set m_City = Value
End Property

Public Property Get Index() As Long
    Index = m_Index
End Property
Public Property Let Index(ByVal Value As Long)
    m_Index = Value
End Property

Public Property Get Key() As String
    Key = ObjPtr(Me)
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Function ToStr() As String
    ToStr = "Person" & m_Index & " {Name: " & Me.Name & "; Birhtday: " & Me.BirthDay & "; Brain: " & Brain.ToStr & "; City: " & Me.City.Name & "}"
End Function
