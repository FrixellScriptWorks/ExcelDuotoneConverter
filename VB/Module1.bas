Attribute VB_Name = "Module1"
Option Explicit

Sub Tampilkan_Borang()
    Dim borang As New Borang1
    
    
    With borang
        .Show
    End With
End Sub
Sub aaaaa() 'Ngetes doang
    Dim s As String
    s = "C:\Picture\Dua\Image.jpg"
    Debug.Print LeftUntilLastBackslash(s)
    
    
    'Kumpulan Data Channel Grayscale
    
    'lngChannelValue = CLng(0.299 * lngRed + 0.587 * lngGreen + 0.114 * lngBlue) 'YUV
    'lngChannelValue = CLng(0.2126 * lngBlue + 0.7152 * lngBlue + 0.0722 * lngBlue) 'HDTV/ATSC
    'lngChannelValue = CLng(0.2627 * lngBlue + 0.678 * lngBlue + 0.0593 * lngBlue) 'HDR
    'lngChannelValue = CLng(0.333 * lngBlue + 0.333 * lngBlue + 0.333 * lngBlue)
    'lngChannelValue = clng(0.22 * lngBlue + 0.44 * lngBlue + 0.34 * lngBlue)
    'lngChannelValue = CLng(0.12 * lngBlue + 0.34 * lngBlue + 0.54 * lngBlue) 'darker grayscale
End Sub
Function LeftUntilLastBackslash(ByVal strInput As String) As String
    Dim lastBackslashPos As Long
    ' Find the position of the last backslash
    lastBackslashPos = InStrRev(strInput, "\")
    
    ' Extract the left part of the string up to the last backslash
    If lastBackslashPos > 0 Then
        LeftUntilLastBackslash = Left(strInput, lastBackslashPos - 1)
    Else
        ' If no backslash is found, return the original string
        LeftUntilLastBackslash = strInput
    End If
End Function
