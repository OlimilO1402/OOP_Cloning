# OOP_Cloning
## Correct creating and cloning of objects with a constructor and a clone function  

[![GitHub](https://img.shields.io/github/license/OlimilO1402/OOP_Cloning?style=plastic)](https://github.com/OlimilO1402/OOP_Cloning/blob/master/LICENSE) 
[![GitHub release (latest by date)](https://img.shields.io/github/v/release/OlimilO1402/OOP_Cloning?style=plastic)](https://github.com/OlimilO1402/OOP_Cloning/releases/latest)
[![Github All Releases](https://img.shields.io/github/downloads/OlimilO1402/OOP_Cloning/total.svg)](https://github.com/OlimilO1402/OOP_Cloning/releases/download/v2024.06.14/CtorCloning_v2024.06.14.zip)
[![Follow](https://img.shields.io/github/followers/OlimilO1402.svg?style=social&label=Follow&maxAge=2592000)](https://github.com/OlimilO1402/OOP_Cloning/watchers)  

and asking  
 * if one object is  equal   to another object (IsEqual) and  
 * if one object is the same as another object (IsSame)  
  
Project started around may 2005.  
  
In most OOP-languages nowadays you have an object-constructor (aka ctor) for first creation of an object and for first initialization of necessary properties of an object in just one line of code.  
And also a copy-constructor for cloning resp copying the complete state of all properties of one object into another new object also in just one line of code.  
In Classic Visual Basic we do not have the convenience of a language built-in constructor, but this does not mean we have to disclaim about it. Of course we can write our own object initialization constructor and object cloning functions for every class.  

1. Example of Using an **Object Constructor**:  
```vb6
Private Sub BntOpenFile_Click()
    Dim file As PathFileName
	Set file = MNew.PathFileName("C:\MyPath\MyFile.ext", FileAccess.BinaryWrite, FileMode.OpenOrCreate)  
    If file.Open Then
        '
    End If
End Sub
```

1.1. To achieve this do the following:  
1.1a) In every Class we need one Sub of the same name, it's name makes clear it is meant for creating a New Object e.g. "New_" 
Making it "Friend", has the benefit it has not to be implemented in derived classes and every derived class can again have it's own ctor-function:  
```vb6  
Friend Sub New_(aPathFileName As String, aFileAccess As FileAccess, aFileMode As FileMode)  
    'inside the class we have access to all private variables  
    '. . . init and assign all private variables  
End Sub  
```  
1.1b) We have a Modul as a factory for object-creation, with Public functions with the Name of every class
    and every function Parameter is the same as in the Friend Sub New_:  
```vb6  
Public Function PathFileName(aPathFileName As String, aFileAccess As FileAccess, aFileMode As FileMode)  
    Set PathFileName = New PathFileName: PathFileName.New_ aPathFileName, aFileAccess, aFileMode  
End Function  
```  
But pay attention to the following, if things begin to change, you have to synchronize all function parameters between 
"Friend Sub New_(<all function parameters>)" and "Public Function MyClass(<all function parameters>) As MyClass  

2. Example of Using **Cloning** Of Objects  
```vb6  
Private Sub BtnPerson_Click()  
    Dim simon As Person: Set simon = peter.Clone  
    Dim dolly As Sheep:  Set dolly = shaun.Clone  
    '. . .  
End Sub  
```  
2.1 To achieve this do the following:  
2.1a) In every cloneable class we need a function Clone:  
```vb6  
Friend Function Clone() As Person
    Set Clone = MNew.Person(Me.Name, Me.EyeColor, Me.HairColor)
End Function
```  
2.1.b) And if we need direct private access to class members we can do a Friend Sub NewC where we just give the other object  
```vb6  
Friend Sub NewC(other As Person) As Person
    With other
        m_Name = .Name
        m_BirthDate = .BirthDate
        m_EyeColor = .EyeColor
        m_HairColor = .HairColor
    End With
End Function
```  
Then in the Public Function Clone you just use it like this:  
```vb6  
Public Function Clone() As Person  
    Set Clone = New Person: Clone.NewC Me
End Function  
```  
Maybe in the first place, you somewhat have to wrap your brain around it before you dig how it works. So I would give the advice: just write it down yourself in an example class, and step through the code so you will get used to it quickly.  

![OOP_Cloning Image](Resources/PCloningIsEqualOrSame.png "OOP-Cloning Image")