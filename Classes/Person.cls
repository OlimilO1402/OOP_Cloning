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
Public birthD As Date
Public Name   As String
Public Brain  As Brain
Public City   As City

Friend Sub New_(bd As Date, nm As String, br As Brain, ct As City)
    birthD = bd: Name = nm: Set Brain = br: Set City = ct
End Sub
Friend Sub NewC(other As Person)
    Me.birthD = other.birthD
    Me.Name = other.Name
    Set Me.Brain = other.Brain.Clone
    Set Me.City = other.City.Clone
End Sub

Public Function Clone() As Person
    Set Clone = Mnew.PersonC(Me)
End Function

Public Function IsEqual(other As Person) As Boolean
    IsEqual = Me.birthD = other.birthD:   If Not IsEqual Then Exit Function
    IsEqual = Me.Name = other.Name:       If Not IsEqual Then Exit Function
    IsEqual = Brain.IsEqual(other.Brain): If Not IsEqual Then Exit Function
    IsEqual = City.IsEqual(other.City):   If Not IsEqual Then Exit Function
End Function

Public Function IsSame(other As Person) As Boolean
    If Me Is other Then IsSame = ObjPtr(Me) = ObjPtr(other)
End Function

Public Function ToStr() As String
    ToStr = "Person {" & Me.Name & "; " & Me.birthD & "; Brain: " & Brain.ToStr & "; City: " & Me.City.Name & "}"
End Function
