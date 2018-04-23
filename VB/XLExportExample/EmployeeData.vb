Imports System.Collections.Generic
Imports System.Linq

Namespace XLExportExample
    Friend Class EmployeeData

        Public Sub New(ByVal id As Integer, ByVal name As String, ByVal salary As Double, ByVal bonus As Double, ByVal department As String)
            Me.Id = id
            Me.Name = name
            Me.Salary = salary
            Me.Bonus = bonus
            Me.Department = department
        End Sub

        Private privateId As Integer
        Public Property Id() As Integer
            Get
                Return privateId
            End Get
            Private Set(ByVal value As Integer)
                privateId = value
            End Set
        End Property
        Private privateName As String
        Public Property Name() As String
            Get
                Return privateName
            End Get
            Private Set(ByVal value As String)
                privateName = value
            End Set
        End Property
        Private privateSalary As Double
        Public Property Salary() As Double
            Get
                Return privateSalary
            End Get
            Private Set(ByVal value As Double)
                privateSalary = value
            End Set
        End Property
        Private privateBonus As Double
        Public Property Bonus() As Double
            Get
                Return privateBonus
            End Get
            Private Set(ByVal value As Double)
                privateBonus = value
            End Set
        End Property
        Private privateDepartment As String
        Public Property Department() As String
            Get
                Return privateDepartment
            End Get
            Private Set(ByVal value As String)
                privateDepartment = value
            End Set
        End Property
    End Class

    Friend NotInheritable Class EmployeesRepository

        Private Sub New()
        End Sub
        Private Shared departments() As String = { "Accounting", "Logistics", "IT", "Management", "Manufacturing", "Marketing" }

        Public Shared Function CreateEmployees() As List(Of EmployeeData)
            Dim result As New List(Of EmployeeData)()
            result.Add(New EmployeeData(10115, "Augusta Delono", 1100.0, 50.0, "Accounting"))
            result.Add(New EmployeeData(10501, "Berry Dafoe", 1650.0, 150.0, "IT"))
            result.Add(New EmployeeData(10709, "Chris Cadwell", 2000.0, 180.0, "Management"))
            result.Add(New EmployeeData(10356, "Esta Mangold", 1400.0, 75.0, "Logistics"))
            result.Add(New EmployeeData(10401, "Frank Diamond", 1750.0, 100.0, "Marketing"))
            result.Add(New EmployeeData(10202, "Liam Bell", 1200.0, 80.0, "Manufacturing"))
            result.Add(New EmployeeData(10205, "Simon Newman", 1250.0, 80.0, "Manufacturing"))
            result.Add(New EmployeeData(10403, "Wendy Underwood", 1100.0, 50.0, "Marketing"))
            Return result
        End Function

        Public Shared Function CreateDepartments() As List(Of String)
            Return departments.ToList()
        End Function
    End Class
End Namespace
