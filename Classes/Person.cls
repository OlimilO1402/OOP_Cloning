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

Private Sub Class_Initialize()
    m_Name = "Max Mustermann"
    m_BirthD = DateSerial(Year(Now), Month(Now), Day(Now))
    Set m_Brain = MNew.Brain
    Set m_City = MNew.City("Musterstadt")
    m_Index = MData.Persons.Count
End Sub

Friend Sub New_(BirthDay As Date, Brain As Brain, City As City, Index As Long, Name As String)
    m_BirthD = BirthDay: Set m_Brain = Brain: Set m_City = City: m_Index = Index: m_Name = Name
End Sub

Public Sub NewC(other)
    Dim Obj As Person: Set Obj = other
    With Obj
        m_BirthD = .BirthDay
        Set m_Brain = .Brain.Clone
        Set m_City = .City.Clone
        m_Name = .Name
        m_Index = .Index
    End With
End Sub

Public Function Clone() As Person
    Set Clone = New Person: Clone.NewC Me
End Function

Public Property Get Ptr() As LongPtr
    Ptr = ObjPtr(Me)
End Property

Public Property Get Key() As String
    Key = Me.Ptr
End Property

Public Function IsEqual(other As Person) As Boolean
    IsEqual = Me.BirthDay = other.BirthDay: If Not IsEqual Then Exit Function
    IsEqual = Me.Name = other.Name:         If Not IsEqual Then Exit Function
    IsEqual = Brain.IsEqual(other.Brain):   If Not IsEqual Then Exit Function
    IsEqual = City.IsEqual(other.City):     If Not IsEqual Then Exit Function
End Function

Public Function IsSame(other As Person) As Boolean
    If Me Is other Then IsSame = Me.Ptr = other.Ptr
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
Public Property Set City(ByVal value As City)
    Set m_City = value
End Property

Public Property Get Index() As Long
    Index = m_Index
End Property
Public Sub IndexInc()
    m_Index = m_Index + 1
End Sub

Public Property Get Name() As String
    Name = m_Name
End Property

Public Function ToStr() As String
    ToStr = "Person" & m_Index & " {Name: " & Me.Name & "; Birthday: " & Me.BirthDay & "; Brain: " & Brain.ToStr & "; City: " & Me.City.Name & "}"
End Function
