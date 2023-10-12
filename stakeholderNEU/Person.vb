﻿Imports System.ComponentModel
Public Class Person
    Implements INotifyPropertyChanged

    Public Event PropertyChanged(ByVal sender As Object, ByVal e As PropertyChangedEventArgs) _
     Implements INotifyPropertyChanged.PropertyChanged

    Protected Sub OnPropertyChanged(ByVal prop As String)
        anychange = True
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(prop))
    End Sub
    Public anychange As Boolean
    'Public Kontakte As new List(Of Kontaktdaten) 


    Property ExpandHeaderInSachgebiet$
    Private _kontakt As New Kontaktdaten
    Public Property Kontakt() As Kontaktdaten
        Get
            Return _kontakt
        End Get
        Set(ByVal Value As Kontaktdaten)
            _kontakt = Value
        End Set
    End Property
    ''' <summary>
    ''' Herr Frau Frollein Firma Eheleute, Familie
    ''' </summary>
    ''' <remarks></remarks>
    Private _anrede As String
    Public Property Anrede() As String
        Get
            Return _anrede
        End Get
        Set(ByVal Value As String)
            _anrede = Value
            OnPropertyChanged("Anrede")
        End Set
    End Property


    Private _Bezirk As String
    Public Property Bezirk() As String
        Get
            Return _Bezirk
        End Get
        Set(ByVal Value As String)
            _Bezirk = Value
            OnPropertyChanged("Bezirk")
        End Set
    End Property

    Private _name As String
    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(ByVal Value As String)
            _name = Value
            OnPropertyChanged("Name")
        End Set
    End Property
    Private _vorname As String
    Public Property Vorname() As String
        Get
            Return _vorname
        End Get
        Set(ByVal Value As String)
            _vorname = Value
            OnPropertyChanged("Vorname")
        End Set
    End Property
    Private _bemerkung As String
    Public Property Bemerkung() As String
        Get
            Return _bemerkung
        End Get
        Set(ByVal Value As String)
            _bemerkung = Value
            OnPropertyChanged("Bemerkung")
        End Set
    End Property

    Private _Status As Integer
    Public Property Status() As Integer
        Get
            Return _Status%
        End Get
        Set(ByVal Value As Integer)
            _Status% = Value
            OnPropertyChanged("Status")
        End Set
    End Property



    Private _personenID As Integer
    Public Property PersonenID() As Integer
        Get
            Return _personenID
        End Get
        Set(ByVal Value As Integer)
            _personenID = Value
            OnPropertyChanged("PersonenID")
        End Set
    End Property
    Private _rolle As String
    ''' <summary>
    ''' rolle im workflow
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Rolle() As String
        Get
            Return _rolle
        End Get
        Set(ByVal Value As String)
            _rolle = Value
            OnPropertyChanged("Rolle")
        End Set
    End Property
    Private _namenszusatz As String
    Public Property Namenszusatz() As String
        Get
            Return _namenszusatz
        End Get
        Set(ByVal Value As String)
            _namenszusatz = Value
            OnPropertyChanged("Namenszusatz")
        End Set
    End Property
    Public Sub clear()
        PersonenID = 0
        Name = ""
        Vorname = ""
        Anrede = ""
        Bemerkung = ""
        Namenszusatz = ""
        Quelle = ""
        Kontakt.clear()
        Kassenkonto = ""
        Status = 0
        Rolle = ""
        Bezirk = ""
        ' Kontakt = New Kontaktdaten
    End Sub
    Sub New()
        clear()
    End Sub
    'Private _Raum As String
    'Public Property Raum() As String
    '    Get
    '        Return _Raum
    '    End Get
    '    Set(ByVal Value As String)
    '        _Raum = Value
    '        OnPropertyChanged("Raum")
    '    End Set
    'End Property

    Private _kassenkonto As String
    Public Property Kassenkonto() As String
        Get
            Return _kassenkonto
        End Get
        Set(ByVal Value As String)
            _kassenkonto = Value
            OnPropertyChanged("Kassenkonto")
        End Set
    End Property
    Public Property Quelle() As String

    'Private Module1.sitzung.aktperson.changed_Anschrift, Module1.sitzung.aktperson.changed_Org, Module1.sitzung.aktperson.changed_Kontakt As Boolean
    Public Property changed_Bankkonto() As Boolean
    Public Property changed_Anschrift() As Boolean
    Public Property changed_Kontakt() As Boolean
    Public Property changed_Org() As Boolean

    Public Overrides Function tostring() As String
        Dim a$ = String.Format("Name: {0}{1}", Name, vbCrLf)
        a$ = String.Format("{0}Vorname: {1}{2}", a$, Vorname, vbCrLf)
        a$ = String.Format("{0}Anrede: {1}{2}", a$, Anrede, vbCrLf)
        a$ = String.Format("{0}Zusatz: {1}{2}", a$, Namenszusatz, vbCrLf)
        a$ = String.Format("{0}Bemerkung: {1}{2}", a$, Bemerkung, vbCrLf)
        Return a$
    End Function

 

End Class