
Imports System.Runtime.InteropServices
Imports ZoneId.Enums
Imports ZoneId.Interfaces

Public Class ZoneIdentifier
    Implements IDisposable

    Public Sub New(fileName As String)
        If Not IO.File.Exists(fileName) Then Throw New IO.FileNotFoundException()
        PersistFileB = CType(New PersistFile(), IPersistFile)
        PersistFileB.Load(fileName, StorageMode.ReadWrite Or StorageMode.ShareExclusive)
        ZoneIdentifierB = CType(PersistFileB, IZoneIdentifier)
        ZoneIdentifierB.GetId(UrlZoneB)
    End Sub

    Private Property UrlZoneB as UrlZone

    Public Property Zone() As UrlZone
        Get
            If PersistFileB Is Nothing Then Throw New ObjectDisposedException([GetType]().Name)
            Return UrlZoneB
        End Get
        Set 
            If PersistFileB Is Nothing Then Throw New ObjectDisposedException([GetType]().Name)
            If UrlZoneB = value Then Return
            If Not [Enum].IsDefined(GetType(UrlZone), value) OrElse value = UrlZone.Invalid Then Throw New ArgumentOutOfRangeException() 
            UrlZoneB = value
            ZoneIdentifierB.SetId(UrlZoneB)
            PersistFileB.Save(Nothing, True)
        End Set
    End Property

    Private Property PersistFileB as IPersistFile
    Private Property ZoneIdentifierB as IZoneIdentifier


    Public Sub Remove()
        If PersistFileB IsNot Nothing Then
            If UrlZoneB = UrlZone.LocalMachine Then Exit Sub
            UrlZoneB = UrlZone.LocalMachine
            ZoneIdentifierB.Remove()
            PersistFileB.Save(Nothing, True)
        End If
    End Sub


    Public Sub Dispose()
        If PersistFileB IsNot Nothing Then
            Marshal.ReleaseComObject(PersistFileB)
            Marshal.ReleaseComObject(ZoneIdentifierB)
        End If
    End Sub

    Private Sub IDisposable_Dispose() Implements IDisposable.Dispose
        Dispose
    End Sub
End Class