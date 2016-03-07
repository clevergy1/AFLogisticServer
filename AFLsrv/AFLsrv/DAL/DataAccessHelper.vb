Imports System
Imports System.Configuration


Namespace SCP.DAL

    Public Class DataAccessHelper
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetDataAccess() As DataAccess
            Dim dataAccessStringType As String = ConfigurationManager.AppSettings("SCP_DataAccessLayerType")
            If String.IsNullOrEmpty(dataAccessStringType) Then
                Return Nothing
            Else

                Dim dataAccessType As Type = Type.GetType(dataAccessStringType)
                If dataAccessType Is Nothing Then
                    Throw New NullReferenceException("DataAccessType can not be found")
                End If

                Dim tp As Type = Type.GetType("SCP.DAL.SQLDataAccess")
                If (Not tp.IsAssignableFrom(dataAccessType)) Then
                    Throw (New ArgumentException("DataAccessType does not inherits from SCP.DAL.SQLDataAccess "))
                End If

                Dim dc As DataAccess = CType(Activator.CreateInstance(dataAccessType), DataAccess)
                Return dc
            End If

        End Function


    End Class

End Namespace

