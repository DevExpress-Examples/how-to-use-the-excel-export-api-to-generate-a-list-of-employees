Imports System.Collections.Generic
Imports System.Linq

Namespace XLExportExample

    Friend Class EmployeeData

        Private _Id As Integer, _Name As String, _Salary As Double, _Bonus As Double, _Department As String

        Public Sub New(ByVal id As Integer, ByVal name As String, ByVal salary As Double, ByVal bonus As Double, ByVal department As String)
            Me.Id = id
            Me.Name = name
            Me.Salary = salary
            Me.Bonus = bonus
            Me.Department = department
        End Sub

        Public Property Id As Integer
            Get
                Return _Id
            End Get

            Private Set(ByVal value As Integer)
                _Id = value
            End Set
        End Property

        Public Property Name As String
            Get
                Return _Name
            End Get

            Private Set(ByVal value As String)
                _Name = value
            End Set
        End Property

        Public Property Salary As Double
            Get
                Return _Salary
            End Get

            Private Set(ByVal value As Double)
                _Salary = value
            End Set
        End Property

        Public Property Bonus As Double
            Get
                Return _Bonus
            End Get

            Private Set(ByVal value As Double)
                _Bonus = value
            End Set
        End Property

        Public Property Department As String
            Get
                Return _Department
            End Get

            Private Set(ByVal value As String)
                _Department = value
            End Set
        End Property
    End Class

    Friend Module EmployeesRepository

        Private departments As String() = New String() {"Accounting", "Logistics", "IT", "Management", "Manufacturing", "Marketing"}

        Public Function CreateEmployees() As List(Of EmployeeData)
            Dim result As List(Of EmployeeData) = New List(Of EmployeeData)()
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

        Public Function CreateDepartments() As List(Of String)
            Return departments.ToList()
        End Function
    End Module
End Namespace
