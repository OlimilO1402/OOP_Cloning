Attribute VB_Name = "Mnew"
Option Explicit
Private m_Cities As New Collection

Private Function CitiesContains(key As String) As Boolean
    On Error Resume Next
    If IsEmpty(m_Cities(key)) Then: 'DoNothing
    CitiesContains = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function Person(bd As Date, nm As String, br As Brain, ct As City) As Person
    Set Person = New Person: Person.New_ bd, nm, br, ct
End Function
Public Function PersonC(other As Person) As Person
    Set PersonC = New Person: PersonC.NewC other
End Function

Public Function Brain() As Brain
    Set Brain = New Brain: Brain.New_
End Function
Public Function BrainC(other As Brain) As Brain
    Set BrainC = New Brain: BrainC.NewC other
End Function

Public Function BrainSmart() As BrainSmart
    Set BrainSmart = New BrainSmart: BrainSmart.New_
End Function
Public Function BrainSmartC(other As BrainSmart) As BrainSmart
    Set BrainSmartC = New BrainSmart: BrainSmartC.NewC other
End Function

Public Function City(nm As String) As City
    If CitiesContains(nm) Then
        Set City = m_Cities.Item(nm)
    Else
        Set City = New City: City.New_ nm
        m_Cities.Add City, nm
    End If
End Function
Public Function CityC(other As City) As City
    Set CityC = other.Clone
End Function
