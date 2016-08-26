Attribute VB_Name = "mod_Juniper_Decrypt"
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : juniper_decrypt
' Author    : Boyan Kolev
' Github    : https://github.com/boyank
' Purpose   : Decrypt Juniper $9$ Type password. Ported to VBA from Python https://github.com/mhite/junosdecode
'---------------------------------------------------------------------------------------

Function juniper_decrypt(strPassword) As Variant

    Dim intCounter As Integer
    Dim intGap As Integer
    Dim intLen As Integer
    Dim intDiff As Integer
    Dim strChars As String
    Dim strFirst As String
    Dim strToss As String
    Dim strPrev As String
    Dim strDecrypt As String
    Dim strNibble As String
    Dim strChar1 As String
    Dim strChar2 As String
    Dim vChar As Variant
    Dim vFamily As Variant
    Dim vNum_Alpha As Variant
    Dim vNibble As Variant
    Dim vDecode As Variant
    Dim objEncoding As Object
    Dim objAlpha_Num As Object
    Dim objExtra As Object
    Dim objGaps As Object


    On Error GoTo juniper_decrypt_error_handler

    vNum_Alpha = Array("Q", "z", "F", "3", "n", "6", "/", "9", "C", "A", "t", "p", "u", "0", "O", _
                       "B", "1", "I", "R", "E", "h", "c", "S", "y", "r", "l", "e", "K", "v", "M", "W", "8", "L", "X", "x", _
                       "7", "N", "-", "d", "V", "b", "w", "s", "Y", "2", "g", "4", "o", "a", "J", "Z", "G", "U", "D", "j", _
                       "i", "H", "k", "q", ".", "m", "P", "f", "5", "T")

    vFamily = Array(Array("Q", "z", "F", "3", "n", "6", "/", "9", "C", "A", "t", "p", "u", "0", "O"), _
                    Array("B", "1", "I", "R", "E", "h", "c", "S", "y", "r", "l", "e", "K", "v", "M", "W", "8", "L", "X", "x"), _
                    Array("7", "N", "-", "d", "V", "b", "w", "s", "Y", "2", "g", "4", "o", "a", "J", "Z", "G", "U", "D", "j"), _
                    Array("i", "H", "k", "q", ".", "m", "P", "f", "5", "T"))


    Set objEncoding = CreateObject("Scripting.Dictionary")
    objEncoding.Add objEncoding.Count, Array(1, 4, 32)
    objEncoding.Add objEncoding.Count, Array(1, 16, 32)
    objEncoding.Add objEncoding.Count, Array(1, 8, 32)
    objEncoding.Add objEncoding.Count, Array(1, 64)
    objEncoding.Add objEncoding.Count, Array(1, 32)
    objEncoding.Add objEncoding.Count, Array(1, 4, 16, 128)
    objEncoding.Add objEncoding.Count, Array(1, 32, 64)

    Set objAlpha_Num = CreateObject("Scripting.Dictionary")
    For intCounter = LBound(vNum_Alpha) To UBound(vNum_Alpha):
        objAlpha_Num.Add vNum_Alpha(intCounter), intCounter
    Next intCounter


    Set objExtra = CreateObject("Scripting.Dictionary")
    For intCounter = LBound(vFamily) To UBound(vFamily)
        For Each vChar In vFamily(intCounter):
            objExtra.Add vChar, 3 - intCounter
        Next vChar
    Next intCounter

    If Left(strPassword, 3) = "$9$" Then
        strChars = Right(strPassword, Len(strPassword) - 3)
        vNibble = Nibble(strChars, 1)
        strFirst = vNibble(0)
        strChars = vNibble(1)
        vNibble = Nibble(strChars, objExtra(strFirst))
        strToss = vNibble(0)
        strChars = vNibble(1)
        strPrev = strFirst
        strDecrypt = vbNullString
        Do While strChars <> vbNullString:
            vDecode = objEncoding(Len(strDecrypt) Mod objEncoding.Count)
            vNibble = Nibble(strChars, (UBound(vDecode) - LBound(vDecode) + 1))
            strNibble = vNibble(0)
            strChars = vNibble(1)
            Set objGaps = CreateObject("Scripting.Dictionary")
            For intCounter = 1 To Len(strNibble):
                strChar1 = strPrev
                strChar2 = Mid(strNibble, intCounter, 1)
                intDiff = objAlpha_Num(strChar2) - objAlpha_Num(strChar1)
                intLen = UBound(vNum_Alpha) - LBound(vNum_Alpha) + 1
                intGap = (objAlpha_Num(strChar2) - objAlpha_Num(strChar1) Mod intLen) - 1
                If intDiff < 0 Then
                    intGap = intGap + intLen
                End If
                strPrev = strChar2
                objGaps.Add objGaps.Count, intGap
            Next intCounter
            strDecrypt = strDecrypt & Gap_Decode(objGaps, vDecode)
        Loop
        juniper_decrypt = strDecrypt
    Else
        juniper_decrypt = CVErr(xlErrValue)
    End If
    Exit Function
    
juniper_decrypt_error_handler:
    MsgBox Err.Description
End Function

Function Nibble(strCref, intLength) As Variant
    Dim strNib As String
    Dim strRest As String

    If Len(strCref) < intLength Then
        Err.Raise Number:=1 + vbObjectError, Description:="Ran out of characters: hit " & strCref & ", expecting " & intLength & " chars."
    Else
        strNib = Left(strCref, intLength)
        strRest = Right(strCref, Len(strCref) - intLength)
        Nibble = Array(strNib, strRest)
    End If
End Function

Function Gap_Decode(objGaps As Object, vDec As Variant) As String
    Dim num As Integer
    Dim intIndex As Integer

    If objGaps.Count <> (UBound(vDec) - LBound(vDec) + 1) Then
        Err.Raise Number:=2 + vbObjectError, Description:="Nibble and decode size not the same!"
    Else
        For intIndex = LBound(vDec) To UBound(vDec)
            num = num + objGaps(intIndex) * vDec(intIndex)
        Next intIndex
        Gap_Decode = Chr(num Mod 256)
    End If
End Function

