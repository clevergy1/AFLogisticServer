'
' Copyright (c) Clevergy 
'
' The SOFTWARE, as well as the related copyrights and intellectual property rights, are the exclusive property of Clevergy srl. 
' Licensee acquires no title, right or interest in the SOFTWARE other than the license rights granted herein.
'
' conceived and developed by Marco Fagnano (D.R.T.C.)
' il software è ricorsivo, nel tempo rigenera se stesso.
'

Imports System
Imports System.Collections.Generic
Imports SCP.DAL

Namespace SCP.BLL
    Public Class Impianti_Contatti
        Private _IdContatto As Integer
        Private _IdImpianto As String
        Private _Descrizione As String
        Private _Indirizzo As String
        Private _Nome As String
        Private _telFisso As String
        Private _telMobile As String
        Private _emailaddress As String

#Region "constructor"
        Public Sub New(IdContatto As Integer, _
                       IdImpianto As String, _
                       Descrizione As String, _
                       Indirizzo As String, _
                       Nome As String, _
                       telFisso As String, _
                       telMobile As String, _
                       emailaddress As String)
            _IdContatto = IdContatto
            _IdImpianto = IdImpianto
            _Descrizione = Descrizione
            _Indirizzo = Indirizzo
            _Nome = Nome
            _telFisso = telFisso
            _telMobile = telMobile
            _emailaddress = emailaddress
        End Sub
        Public Sub New()
            Me.New(0, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty)
        End Sub
#End Region

#Region "methods"
        Public Shared Function Add(IdImpianto As String, _
                                   Descrizione As String, _
                                   Indirizzo As String, _
                                   Nome As String, _
                                   telFisso As String, _
                                   telMobile As String, _
                                   emailaddress As String) As Boolean
            If String.IsNullOrEmpty(IdImpianto) Then
                Return False
            End If
            If String.IsNullOrEmpty(Descrizione) Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.Impianti_Contatti_Add(IdImpianto, Descrizione, Indirizzo, Nome, telFisso, telMobile, emailaddress)
        End Function

        Public Shared Function Del(IdContatto As Integer) As Boolean
            If IdContatto <= 0 Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.Impianti_Contatti_Del(IdContatto)
        End Function

        Public Shared Function List(IdImpianto As String) As IList(Of Impianti_Contatti)
            If String.IsNullOrEmpty(IdImpianto) Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.Impianti_Contatti_List(IdImpianto)
        End Function

        Public Shared Function Read(IdContatto As Integer) As Impianti_Contatti
            If IdContatto <= 0 Then
                Return Nothing
            End If
            Return DataAccessHelper.GetDataAccess.Impianti_Contatti_Read(IdContatto)
        End Function

        Public Shared Function Update(IdContatto As Integer, _
                                      Descrizione As String, _
                                      Indirizzo As String, _
                                      Nome As String, _
                                      telFisso As String, _
                                      telMobile As String, _
                                      emailaddress As String) As Boolean
            If IdContatto <= 0 Then
                Return False
            End If
            If String.IsNullOrEmpty(Descrizione) Then
                Return False
            End If
            Return DataAccessHelper.GetDataAccess.Impianti_Contatti_Update(IdContatto, Descrizione, Indirizzo, Nome, telFisso, telMobile, emailaddress)
        End Function
#End Region

#Region "public properties"
        Public Property IdContatto As Integer
            Get
                Return Me._IdContatto
            End Get
            Set(value As Integer)
                Me._IdContatto = value
            End Set
        End Property
        Public Property IdImpianto As String
            Get
                Return Me._IdImpianto
            End Get
            Set(value As String)
                Me._IdImpianto = value
            End Set
        End Property
        Public Property Descrizione As String
            Get
                Return Me._Descrizione
            End Get
            Set(value As String)
                Me._Descrizione = value
            End Set
        End Property
        Public Property Indirizzo As String
            Get
                Return Me._Indirizzo
            End Get
            Set(value As String)
                Me._Indirizzo = value
            End Set
        End Property
        Public Property Nome As String
            Get
                Return Me._Nome
            End Get
            Set(value As String)
                Me._Nome = value
            End Set
        End Property
        Public Property telFisso As String
            Get
                Return Me._telFisso
            End Get
            Set(value As String)
                Me._telFisso = value
            End Set
        End Property
        Public Property telMobile As String
            Get
                Return Me._telMobile
            End Get
            Set(value As String)
                Me._telMobile = value
            End Set
        End Property
        Public Property emailaddress As String
            Get
                Return Me._emailaddress
            End Get
            Set(value As String)
                Me._emailaddress = value
            End Set
        End Property
#End Region
    End Class
End Namespace
