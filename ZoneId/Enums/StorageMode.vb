' http://msdn.microsoft.com/en-us/library/ie/aa380337(v=vs.85).aspx
Namespace  Enums
    <Flags>
    Public Enum StorageMode
        Read = 0
        Write = 1
        ReadWrite = 2
        ShareDenyNone = &H40
        ShareDenyRead = &H30
        ShareDenyWrite = &H20
        ShareExclusive = &H10
        Priority = &H40000
        Create = &H1000
        Convert = &H20000
        FailIfThere = 0
        Direct = 0
        Transacted = &H10000
        NoScratch = &H100000
        NoSnapshot = &H200000
        Simple = &H8000000
        DirectSWMR = &H400000
        DeleteOnRelease = &H4000000
    End Enum
End NameSpace