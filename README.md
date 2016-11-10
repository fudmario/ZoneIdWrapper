Zone Identifier
===================
 

When you download a file via Internet, the file is tagged with a little bit of information known as a zone identifier which remembers where the file was downloaded from. This is what tells Explorer to put up the "Yo, did you really want to run this program?" prompt and which is taken into account by applications so that they can do things like disable scripting and macros when they open the document, just in case the file is malicious. *
 

#### <i class="icon-file"></i> Example Code
``` .vb
   Using zi As New ZoneIdentifier("D:\Sample.exe")
            If zi.Zone = UrlZone.Internet Then
                zi.Remove
            End If
   End Using
```

