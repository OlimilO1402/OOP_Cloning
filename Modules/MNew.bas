Attribute VB_Name = "MNew"
Option Explicit

Public Function Person(ByVal BirthDay As Date, ByVal Brain As Brain, ByVal City As City, ByVal Index As Long, ByVal Name As String) As Person
    Set Person = New Person: Person.New_ BirthDay, Brain, City, Index, Name
End Function

Public Function Brain() As Brain
    Set Brain = New Brain
End Function

Public Function BrainSmart() As BrainSmart
    Set BrainSmart = New BrainSmart
End Function

Public Function City(ByVal Name As String) As City
    Set City = New City: City.New_ Name
    Set City = MData.Cities_Add(City)
End Function

Sub Main()
    Init
    FMain.Show
End Sub
