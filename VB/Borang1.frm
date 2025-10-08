VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Borang1 
   Caption         =   "Borang Gambar"
   ClientHeight    =   5580
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9852.001
   OleObjectBlob   =   "Borang1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Borang1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'#=============================================================================
'# Dualtone Effect Converter
'# Version      : 1.0
'# Author       : Bayu R. Jati
'# Copyright    : MIT
'#=============================================================================
Private strFileAsli As String
Private strBuatSave As String

Private Sub btnGambarAsli_Click()
    Dim strPathFile As String
    
    'Buat menampilkan gambar asal
    strPathFile = pathGambar1.Text
    If InStr(1, strPathFile, "/") >= 1 Then
        MsgBox "Jangan pakai tanda / untuk path. Pakai \ saja." & vbCrLf & "Jika nama file ada /, tolong diganti atau dihapus."
        strPathFile = ""
    ElseIf strPathFile = "" Then
        MsgBox "Path file kosong!"
        strPathFile = ""
    Else
        Gambar1.Picture = LoadPicture(strPathFile)
        strFileAsli = strPathFile
        pathGambar1.Visible = False
    End If
End Sub

Private Sub btnKonversi_Click()
    
    If (pathGambar1.Text = "") Then
        MsgBox "Path kosong, tidak bisa melanjutkan proses."
        Exit Sub
    End If
    
    'paksa ambil teks dari textbox kalo lupa insert
    If strFileAsli = "" Then
        MsgBox "Input Gambar..."
        btnGambarAsli_Click
    End If
    
    Dim objImageFile    As Object
    Dim objImagePicture As Object
    Dim objVektor       As Object
    
    'Ambil gambar pakai Windows Image Aquisition Library
    Set objImageFile = CreateObject("WIA.ImageFile")
    Set objImagePicture = CreateObject("WIA.ImageProcess")
    
    Dim strSaveFileNoExt        As String
    Dim strFileAsliExt          As String
    Dim i                       As Long
    Dim lngAlpha                As Long    '0-255
    Dim lngRed                  As Long    '0-255
    Dim lngGreen                As Long    '0-255
    Dim lngBlue                 As Long    '0-255
    Dim lngChannelValue         As Long    '0-255
    
    Dim lngWarna1               As Long
    Dim lngAlpha1               As Long    '0-255
    Dim lngRed1                 As Long    '0-255
    Dim lngGreen1               As Long    '0-255
    Dim lngBlue1                As Long    '0-255
    
    Dim lngWarna2               As Long
    Dim lngAlpha2               As Long    '0-255
    Dim lngRed2                 As Long    '0-255
    Dim lngGreen2               As Long    '0-255
    Dim lngBlue2                As Long    '0-255
    Dim lngNormalizedGray       As Long    '0-1
    
    
    'Bikin file output memiliki ekstensi yang sama dengan original
    strSaveFileNoExt = Left(strFileAsli, InStrRev(strFileAsli, ".") - 1) & "1"
    strFileAsliExt = Right(strFileAsli, Len(strFileAsli) - InStrRev(strFileAsli, "."))
    strBuatSave = strSaveFileNoExt & "." & strFileAsliExt
    
    'Buat ngecek kalau file sudah ada di tempat save atau belum.
    'Kalau sudah ada, timpa.
    If Len(Dir(strBuatSave)) > 0 Then
        Kill strBuatSave
    End If
    
    'Ambil data ARGB dari gambar
    objImageFile.LoadFile strFileAsli
    Set objVektor = objImageFile.ARGBData
    
    'Ambil Warna DualTone
    lngWarna1 = GetDecFromHex(teksWarna1.Text)
    lngWarna2 = GetDecFromHex(teksWarna2.Text)
    
    'Get Warna Baru
    Call GetARGB(lngWarna1, lngAlpha1, lngRed1, lngGreen1, lngBlue1)
    Call GetARGB(lngWarna2, lngAlpha2, lngRed2, lngGreen2, lngBlue2)
    
    For i = 1 To objVektor.Count
        Call GetARGB(objVektor(i), lngAlpha, lngRed, lngGreen, lngBlue)

        If (lngRed = 0 And lngGreen = 0 And lngBlue = 0) Or (lngRed = 255 And lngGreen = 255 And lngBlue = 255) Then
        
            'Ini jika warna sudah hitam atau putih
            objVektor(i) = BuildARGB(lngAlpha, lngRed, lngGreen, lngBlue)
        Else
            'Bikin channel buat interpolasi grayscale (channel yang lain ada di Module1)
            lngChannelValue = CLng(0.299 * lngRed + 0.587 * lngGreen + 0.114 * lngBlue) 'YUV
            
            'Normalisasi persentase abu-abu
            lngNormalizedGray = lngChannelValue / 255
            
            'pewarnaan dual tone pada channel grayscale
            lngRed = (1 - lngNormalizedGray) * lngRed1 + lngNormalizedGray * lngRed2
            lngGreen = (1 - lngNormalizedGray) * lngGreen1 + lngNormalizedGray * lngBlue2
            lngBlue = (1 - lngNormalizedGray) * lngBlue1 + lngNormalizedGray * lngBlue2
            
            'Jadikan vektor gambar
            objVektor(i) = BuildARGB(lngAlpha, lngRed, lngGreen, lngBlue)
        End If
    Next
    
    'Timpa dualtone ke gambar
    objImagePicture.Filters.Add objImagePicture.FilterInfos("ARGB").FilterID
    Set objImagePicture.Filters(1).Properties("ARGBData") = objVektor
    Set objImageFile = objImagePicture.Apply(objImageFile)
   
    'Save file
    objImageFile.SaveFile strBuatSave
    Gambar2.Picture = LoadPicture(strBuatSave)
    
Error_Handler_Exit:
    On Error Resume Next
    Set objVektor = Nothing
    Set objImagePicture = Nothing
    Set objImageFile = Nothing
    pathGambar1.Visible = True
    Exit Sub

Error_Handler:
    MsgBox "Terjadi error guys" & vbCrLf & vbCrLf & _
           "Dari: Button Konversi" & vbCrLf & _
           "Nomor: " & Err.Number & vbCrLf & _
           "Deskripsi: " & Err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "Terjadi Error!"
    Resume Error_Handler_Exit
End Sub


'Ga tau ini nemu di mana buat konversi decimal, hexadecimal, dan ARGB
Private Function GetDecFromHex(val) As String
    Dim s                     As String
    
    s = val
    If InStr(1, s, "#") Then
        s = Right(s, Len(s) - 1)
    End If
    s = CLng("&h" & s)
    GetDecFromHex = s
End Function

Private Function Get4ByteHex(val) As String
    Dim s                     As String

    s = Hex(val)
    Do While Len(s) < 8
        s = "0" & s
    Loop
    Get4ByteHex = Right(s, 8)
End Function

Private Function Get1ByteHex(val) As String
    Dim s                     As String

    s = Hex(val)
    Do While Len(s) < 2
        s = "0" & s
    Loop
    Get1ByteHex = Right(s, 2)
End Function

Private Function BuildARGB(a, r, g, b) As Long
    Dim s                     As String

    s = "&h" & Get1ByteHex(a) & Get1ByteHex(r) & Get1ByteHex(g) & Get1ByteHex(b)
    BuildARGB = CLng(s)
End Function

Private Function GetARGB(val, ByRef lApha As Long, ByRef lRed As Long, ByRef lGreen As Long, _
                         ByRef lBlue As Long)
    Dim s                     As String

    s = Get4ByteHex(val)
    lApha = CLng("&h" & Left(s, 2))
    lRed = CLng("&h" & Mid(s, 3, 2))
    lGreen = CLng("&h" & Mid(s, 5, 2))
    lBlue = CLng("&h" & Right(s, 2))
End Function

Private Sub CommandButton2_Click()
    'Buka folder
    Dim folderPath As String
    folderPath = LeftUntilLastBackslash(strBuatSave)
    Shell "explorer.exe " & folderPath, vbNormalFocus
End Sub

'Dari AI
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
