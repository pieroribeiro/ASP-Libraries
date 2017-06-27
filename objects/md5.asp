<%
Class MD5

	'-----------------------------------------------------
	'Atributos/Constantes da Classe
	'-----------------------------------------------------
	Private dblCenterY
	Private dblCenterX
	Private LastResult
	Private LastErrDes
	Private LastErrNum 
	Private errInvalidKeyLength
	Private errInvalidKey
	Private errInvalidSize
	Private errKeyMissing
	Private errClearTextMissing
	Private errCipherTextMissing
	Private A
	Private B
	Private C
	Private D
	Private E
	Private F

	'-----------------------------------------------------
	'Procedimentos de Inicialização de Destruição da Classe
	'-----------------------------------------------------
	Private Sub Class_Initialize()
		'Inivializando as variáveis
		errInvalidKeyLength = 65101
		errInvalidKey = 65102
		errInvalidSize = 65103
		errKeyMissing = 65303
		errClearTextMissing = 65304
		errCipherTextMissing = 65305
		A = 10
		B = 11
		C = 12
		D = 13
		E = 14
		F = 15
	End Sub
	Private Sub Class_Terminate()
	End Sub

	Public Function QuickEncrypt(strClear, strKey)
		Dim intRet
		intRet = EncryptText(strClear, strKey)
		If intRet = -1 Then
			QuickEncrypt = "ERROR"
		Else
			QuickEncrypt = LastResult
		End If
	End Function

	Public Function QuickDecrypt(strCipher, strKey)
		Dim intRet
		intRet = DecryptText(strCipher, strKey)
		If intRet = -1 Then
			QuickDecrypt = "ERROR"
		Else
			QuickDecrypt = LastResult
		End If
	End Function

	Private Function GetStrength(strPassword)
		strPassword = CStr(strPassword)
		GetStrength = (Len(strPassword) * 8) + (Len(GetSerial) * 8)
	End Function

	Private Function GetSerial()
		GetSerial = Now
	End Function

	Private Function GetHash(strKey)
	    Dim strCipher
	    Dim byKey()
	    ReDim byKey(Len(strKey))
	    For i = 1 To Len(strKey)
	        byKey(i) = Asc(Mid(strKey, i, 1))
	    Next

	    For i = 1 To UBound(byKey) / 2
	        strCipher = strCipher & NumToHex(byKey(i) Xor byKey((UBound(byKey) - i) + 1))
	    Next
	    GetHash = strCipher
	End Function

	Private Function CreatePassword(strSeed, lngLength)
	    Dim bySeed()
	    Dim bySerial()
	    Dim strTimeSerial
	    Dim Random
	    Dim lngPosition
	    Dim lngSerialPosition
	    strCipher = ""
	    lngPosition = 1
	    lngSerialPosition = 1
	    ReDim bySeed(Len(strSeed))
	    For i = 1 To Len(strSeed)
	      bySeed(i) = Asc(Mid(strSeed, i, 1))
	    Next
	    strTimeSerial = GetSerial()
	    ReDim bySerial(Len(strTimeSerial))
	    For i = 1 To Len(strTimeSerial)
	      bySerial(i) = Asc(Mid(strTimeSerial, i, 1))
	    Next
	    ReCenter CDbl(bySeed(lngPosition)), CDbl(bySerial(lngSerialPosition))
	    lngPosition = lngPosition + 1
	    lngSerialPosition = lngSerialPosition + 1
	    For i = 1 To (lngLength / 2)
			Generate CDbl(bySeed(lngPosition)), CDbl(bySerial(lngSerialPosition)), False
			strCipher = strCipher & NumToHex(MakeRandom(dblCenterX, 255))
			strCipher = strCipher & NumToHex(MakeRandom(dblCenterY, 255))
			If lngPosition = Len(strSeed) Then
				lngPosition = 1
	        Else
				lngPosition = lngPosition + 1
	        End If
	        If lngSerialPosition = Len(strTimeSerial) Then
				lngSerialPosition = 1
	        Else
				lngSerialPosition = lngSerialPosition + 1
	        End If
	    Next
	    CreatePassword = Left(strCipher, lngLength)
	End Function

	Private Sub ReCenter(mdblCenterY, mdblCenterX)
	    dblCenterY = mdblCenterY
	    dblCenterX = mdblCenterX
	End Sub

	Private Sub Generate(dblRadius, dblTheta, blnRad)
	    Const Pi = 3.14159265358979
	    Const sngMaxUpper = 2147483647
	    Const sngMaxLower = -2147483648
	    If blnRad = False Then
	        If (dblRadius * Cos((dblTheta / 180) * Pi)) + dblCenterX > sngMaxUpper Or (dblRadius * Cos((dblTheta / 180) * Pi)) + dblCenterX < sngMaxLower Then
	            ReCenter dblCenterY, 0
	        Else
	            dblCenterX = (dblRadius * Cos((dblTheta / 180) * Pi)) + dblCenterX
	        End If
	        
	        If (dblRadius * Sin((dblTheta / 180) * Pi)) + dblCenterY > sngMaxUpper Or (dblRadius * Sin((dblTheta / 180) * Pi)) + dblCenterY < sngMaxLower Then
	            ReCenter 0, dblCenterX
	        Else
	            dblCenterY = (dblRadius * Sin((dblTheta / 180) * Pi)) + dblCenterY
	        End If
	    Else
	        If (dblRadius * Cos(dblTheta)) + dblCenterX > sngMaxUpper Or (dblRadius * Cos(dblTheta)) + dblCenterX < sngMaxLower Then
	            ReCenter dblCenterY, 0
	        Else
	            dblCenterX = (dblRadius * Cos(dblTheta)) + dblCenterX
	        End If
	    
	        If (dblRadius * Sin(dblTheta)) + dblCenterY > sngMaxUpper Or (dblRadius * Sin(dblTheta)) + dblCenterY < sngMaxLower Then
	            ReCenter 0, dblCenterX
	        Else
	            dblCenterY = (dblRadius * Sin(dblTheta)) + dblCenterY
	        End If
	    End If
	End Sub

	Private Function MakeRandom(dblValue, lngMax)
	    Dim lngRandom
	    lngRandom = Int(dblValue Mod (lngMax + 1))
	    If lngRandom > lngMax Then
	        lngRandom = 0
	    End If
	    MakeRandom = Abs(lngRandom)
	End Function

	Private Sub RaiseError(lngErrNum, strErrDes)
	    LastErrDes = strErrDes
	    LastErrNum = lngErrNum
	End Sub

	Private Function EncryptText(strClear, strKey)
	    Dim byClear()
	    Dim byKey()
	    Dim byCipher()
	    Dim lngPosition
	    Dim lngSerialPosition
	    Dim strTimeSerial
	    Dim blnSecondValue
	    Dim strCipher
	    Dim i
	    strKey = CStr(strKey)
	    strClear = CStr(strClear)
	    If strKey = "" Then
	        RaiseError errKeyMissing, "Key Missing"
			EncryptText = -1
			Exit Function
	    End If
	    If Len(strKey) <= 1 Then
	        RaiseError errInvalidKeyLength, "Invalid Key Length"
			EncryptText = -1
			Exit Function
	    End If
	    strTimeSerial = GetSerial()
	    ReDim byKey(Len(strKey))
	    For i = 1 To Len(strKey)
	        byKey(i) = Asc(Mid(strKey, i, 1))
	    Next
	    If Len(strClear) = 0 Then
	        RaiseError errInvalidSize, "Text Has Zero Length"
			EncryptText = -1
			Exit Function
	    End If
	    ReDim byClear(Len(strClear))
	    For i = 1 To Len(strClear)
	        byClear(i) = Asc(Mid(strClear, i, 1))
	    Next
	    lngPosition = 1
	    lngSerialPosition = 1
	    For i = 1 To UBound(byKey) / 2
	        strCipher = strCipher & NumToHex(byKey(i) Xor byKey((UBound(byKey) - i) + 1))
	    Next
	    lngPosition = 1
	    strCipher = strCipher & NumToHex(Len(strTimeSerial) Xor byKey(lngPosition))
	    lngPosition = lngPosition + 1
	    For i = 1 To Len(strTimeSerial)
	        strCipher = strCipher & NumToHex(byKey(lngPosition) Xor Asc(Mid(strTimeSerial, i, 1)))
            If lngPosition = UBound(byKey) Then
                lngPosition = 1
            Else
                lngPosition = lngPosition + 1
            End If
	    Next
	    lngPosition = 1
	    lngSerialPosition = 1
	    ReCenter CDbl(byKey(lngPosition)), Asc(Mid(strTimeSerial, lngSerialPosition, 1))
	    lngPosition = lngPosition + 1
	    lngSerialPosition = lngSerialPosition + 1
	    blnSecondValue = False     
	    For i = 1 To UBound(byClear)    
			If blnSecondValue = False Then
				Generate CDbl(byKey(lngPosition)), Asc(Mid(strTimeSerial, lngSerialPosition, 1)), False
				strCipher = strCipher & NumToHex(byClear(i) Xor MakeRandom(dblCenterX, 255))
	            blnSecondValue = True
            Else
				strCipher = strCipher & NumToHex(byClear(i) Xor MakeRandom(dblCenterY, 255))
                blnSecondValue = False
            End If
            If lngPosition = UBound(byKey) Then
				lngPosition = 1
			Else
				lngPosition = lngPosition + 1
			End If
			If lngSerialPosition = Len(strTimeSerial) Then
				lngSerialPosition = 1
			Else
				lngSerialPosition = lngSerialPosition + 1
			End If
	    Next
	    LastResult = strCipher
	    EncryptText = 1
	    Exit Function
	End Function

	Public Function DecryptText(strCipher, strKey)
	    Dim strClear
	    Dim byCipher()
	    Dim byKey()
	    Dim strTimeSerial
	    Dim strCheckKey
	    Dim lngPosition
	    Dim lngSerialPosition
	    Dim lngCipherPosition
	    Dim bySerialLength
	    Dim blnSecondValue
	    Dim i
	    strCipher = CStr(strCipher)
	    strKey = CStr(strKey)
	    If Len(strCipher) = 0 Then
			RaiseError errCipherTextMissing, "Cipher Text Missing"
			DecryptText = -1
			Exit Function
	    End If
	    If Len(strCipher) < 10 Then
			RaiseError errInvalidSize, "Bad Text Length"
			DecryptText = -1
			Exit Function
	    End If
	    If Len(strKey) = 0 Then
			RaiseError errKeyMissing, "Key Missing"
			DecryptText = -1
			Exit Function
	    End If
	    If Len(strKey) <= 1 Then
			RaiseError errInvalidKeyLength, "Invalid Key Length"
			DecryptText = -1
			Exit Function
	    End If
	    ReDim byKey(Len(strKey))
	    For i = 1 To Len(strKey)
	        byKey(i) = Asc(Mid(strKey, i, 1))
	    Next
	    ReDim byCipher(Len(strCipher) / 2)
	    lngCipherPosition = 1
	    For i = 1 To Len(strCipher) Step 2
	        byCipher(lngCipherPosition) = HexToNum(Mid(strCipher, i, 2))
	        lngCipherPosition = lngCipherPosition + 1
	    Next
	    lngCipherPosition = 1
	    For i = 1 To UBound(byKey) / 2
	        strCheckKey = strCheckKey & NumToHex(byKey(i) Xor byKey((UBound(byKey) - i) + 1))
	    Next
	    If Left(strCipher, Len(strCheckKey)) <> strCheckKey Then
	        RaiseError errInvalidKey, "Invalid Key"
			DecryptText = -1
			Exit Function
	    Else
	        lngCipherPosition = (Len(strCheckKey) / 2) + 1
	    End If
	    lngPosition = 1
	    bySerialLength = byCipher(lngCipherPosition) Xor byKey(lngPosition)
	    lngCipherPosition = lngCipherPosition + 1
	    lngPosition = lngPosition + 1
	    For i = 1 To bySerialLength
	        strTimeSerial = strTimeSerial & Chr(byCipher(lngCipherPosition) Xor byKey(lngPosition))
	        If lngPosition = UBound(byKey) Then
				lngPosition = 1
	        Else
	            lngPosition = lngPosition + 1
	        End If
	        lngCipherPosition = lngCipherPosition + 1
	    Next
	    lngPosition = 1
	    lngSerialPosition = 1
	    ReCenter CDbl(byKey(lngPosition)), Asc(Mid(strTimeSerial, lngSerialPosition, 1))
	    lngPosition = lngPosition + 1
	    lngSerialPosition = lngSerialPosition + 1
	    blnSecondValue = False
	    For i = 1 To UBound(byCipher) - lngCipherPosition + 1
			If blnSecondValue = False Then
				Generate CDbl(byKey(lngPosition)), Asc(Mid(strTimeSerial, lngSerialPosition, 1)), False
				strClear = strClear & Chr(byCipher(lngCipherPosition) Xor MakeRandom(dblCenterX, 255))
				blnSecondValue = True
            Else
				strClear = strClear & Chr(byCipher(lngCipherPosition) Xor MakeRandom(dblCenterY, 255))
				blnSecondValue = False
            End If
            If lngPosition = UBound(byKey) Then
				lngPosition = 1
            Else
				lngPosition = lngPosition + 1
            End If
            If lngSerialPosition = Len(strTimeSerial) Then
                lngSerialPosition = 1
            Else
                lngSerialPosition = lngSerialPosition + 1
            End If
            lngCipherPosition = lngCipherPosition + 1
	    Next
	    LastResult = strClear
	    DecryptText = 1
	    Exit Function
	End Function


	Private Function NumToHex(bByte)
	    Dim strOne
	    Dim strTwo
	    strOne = CStr(Int((bByte / 16)))
	    strTwo = bByte - (16 * strOne)
	    If CDbl(strOne) > 9 Then
	        If CDbl(strOne) = 10 Then
	            strOne = "A"
	        ElseIf CDbl(strOne) = 11 Then
	            strOne = "B"
	        ElseIf CDbl(strOne) = 12 Then
	            strOne = "C"
	        ElseIf CDbl(strOne) = 13 Then
	            strOne = "D"
	        ElseIf CDbl(strOne) = 14 Then
	            strOne = "E"
	        ElseIf CDbl(strOne) = 15 Then
	            strOne = "F"
	        End If
	    End If
	    
	    If CDbl(strTwo) > 9 Then
	        If strTwo = "10" Then
	            strTwo = "A"
	        ElseIf strTwo = "11" Then
	            strTwo = "B"
	        ElseIf strTwo = "12" Then
	            strTwo = "C"
	        ElseIf strTwo = "13" Then
	            strTwo = "D"
	        ElseIf strTwo = "14" Then
	            strTwo = "E"
	        ElseIf strTwo = "15" Then
	            strTwo = "F"
	        End If
	    End If
	    NumToHex = Right(strOne & strTwo, 2)
	End Function

	Private Function HexToNum(hexnum)
		Dim X
		Dim y
		Dim cur
	    hexnum = UCase(hexnum)
	    cur = CStr(Right(hexnum, 1))
		Select Case cur
	        Case "A"
	            y = A
	        Case "B"
	            y = B
	        Case "C"
	            y = C
	        Case "D"
	            y = D
	        Case "E"
	            y = E
	        Case "F"
	            y = F
			Case Else
	            y = CDbl(cur)
		End Select    
	    Select Case Left(hexnum, 1)
	        Case "0"
	            X = (16 * CInt(Left(hexnum, 1))) + y
	        Case "1"
	            X = (16 * CInt(Left(hexnum, 1))) + y
	        Case "2"
	            X = (16 * CInt(Left(hexnum, 1))) + y
	        Case "3"
	            X = (16 * CInt(Left(hexnum, 1))) + y
	        Case "4"
	            X = (16 * CInt(Left(hexnum, 1))) + y
	        Case "5"
	            X = (16 * CInt(Left(hexnum, 1))) + y
	        Case "6"
	            X = (16 * CInt(Left(hexnum, 1))) + y
	        Case "7"
	            X = (16 * CInt(Left(hexnum, 1))) + y
	        Case "8"
	            X = (16 * CInt(Left(hexnum, 1))) + y
	        Case "9"
	            X = (16 * CInt(Left(hexnum, 1))) + y
	        Case "A"
	            X = 160 + y
	        Case "B"
	            X = 176 + y
	        Case "C"
	            X = 192 + y
	        Case "D"
	            X = 208 + y
	        Case "E"
	            X = 224 + y
	        Case "F"
	            X = 240 + y
	    End Select
		HexToNum = X
	End Function
End Class
%>
