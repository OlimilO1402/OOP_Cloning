Attribute VB_Name = "Mnew"
Option Explicit

Public Function Person(BirthDay As Date, Brain As Brain, City As City, Index As Long, Name As String) As Person
    Set Person = New Person: Person.New_ BirthDay, Brain, City, Index, Name
End Function

Public Function Brain() As Brain
    Set Brain = New Brain
End Function

Public Function BrainSmart() As BrainSmart
    Set BrainSmart = New BrainSmart
End Function

Public Function City(ByVal Name As String) As City
    Set City = MData.Cities_Add(Name)
'    If MData.Cities_Contains(Name) Then
'        Set City = m_Cities.Item(Name)
'    Else
'        Set City = New City: City.New_ Name
'        m_Cities.Add City, Name
'    End If
End Function

Sub Main()
    Init
    FMain.Show
End Sub
