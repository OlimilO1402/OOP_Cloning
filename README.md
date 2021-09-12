# OOP_Cloning
## Correct cloning of objects with copy-constructor or function clone  

and asking  
 * if one object is  equal   to another object (IsEqual) and  
 * if one object is the same as another object (IsSame)  
 
in most OOP-languages you have a object-constructor for first creation of an object and for first initialization of necessary properties of an object in one line of code,
and a copy-constructor for cloning resp copying the complete state of all properties of one object into another new object also in one line of code.
In Classic Visual Basic we do not have the convenience of a language built-in constructor, but this does not mean we have to disclaim about it.
Of course we can write our own object initialization constructor and object cloning functions for every class.

1. Example of Using an Object Constructor:  
```vb6
    Private Sub BntOpenFile_Click()
        Dim file As PathFileName: Set file = MNew.PathFileName(aPFN As String, FileAccess.BinaryWrite, FileMode.OpenOrCreate)
	    If file.Open Then
    	    '
    	End If
    End Sub
```

1.1. To achieve this do the following:  
1.1a) In every Class we need one Sub of the same name, it's name makes clear it is meant for creating a New Object  
    We make it "Friend", so it has not to be implemented in derived classes and every derived class can have it's own  
    ctor-function:  
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
We must admit, what can be annoying if things begin to change, you have to synchronize <all function parameters> 
between "Friend Sub New_(<all function parameters>)" and "Public Function MyClass(<all function parameters>) As MyClass
	
2. Example of Using Cloning Of Objects
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
2.1.b) And if we need more direct private access to class members we need a Friend Sub NewC
```vb6  
    Friend Sub NewC(other As Person) As Person
        m_Name = other.Name: m_BirthDate = other.BirthDate: m_EyeColor = other.EyeColor		
    End Function
```  
  
![OOP_Cloning Image](Resources/PCloningIsEqualOrSame.png "OOP-Cloning Image")