' http://msdn.microsoft.com/en-us/library/ms537032(v=vs.85).aspx
Imports System.Runtime.InteropServices
Imports ZoneId.Enums

Namespace Interfaces

    <ComImport>
    <Guid("cd45f185-1b21-48e2-967b-ead743a8914e")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IZoneIdentifier
        <PreserveSig>
        Sub GetId(<Out> ByRef dwZone As UrlZone)

        <PreserveSig>
        Sub SetId(dwZone As UrlZone)

        <PreserveSig>
        Sub Remove()
    End Interface
End NameSpace