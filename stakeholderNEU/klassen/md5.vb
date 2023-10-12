﻿Imports System.Security.Cryptography
public Class md5

    Public shared Function GetMD5FromString(ByVal sText As String) As String
        ' MD5-Hash eines Strings ermitteln
        ' Der String-Inhalt muss hierbei als Byte-Array 
        ' übergeben werden. Hierzu verweden wir einfach 
        ' System.Text.Encoding-Klasse
        Using MD5 As New MD5CryptoServiceProvider()
            MD5.ComputeHash(System.Text.Encoding.Default.GetBytes(sText))
            ' als Ergebnis erhalten wir wieder ein Byte-Array, 
            ' das mittels der BitConverter-Klasse zurück in 
            ' einen String konvertiert wird.
            Return BitConverter.ToString(MD5.Hash)
        End Using
    End Function
End class
