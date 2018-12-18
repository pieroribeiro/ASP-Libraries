<%
'######################################################################
'#  Classe para manipula��o de Array
'#  Autor: Piero Lino Ribeiro
'#  Email: piero.webmaster@gmail.com      
'#####################################################################
Class Arrays

    Private currentArray
    Private tempArray	
	
    'Retorna o Tamanho do Array atual
    Public Property Get Count()
        Count = Ubound(currentArray)
    End Property

    'Pega um valor pelo indice
    Public Property Get Value(arrKey)
        Value = currentArray(arrKey)
    End Property
    
    Private Sub Class_Initialize()
        If Not IsArray(currentArray) Then 
            currentArray = Array()
        End If
    End Sub
    
    Private Sub Class_Terminate()
        If IsArray(currentArray) Then 
            currentArray = Empty
        End If
    End Sub
    
    'Cria um novo array no Array Corrente com valores passados nos parametros separados por virgulas
    Public Sub Create(arrValues)
        If Isempty(arrValues) Then Exit Sub
        tempArray = Split(arrValues,",")
        If IsArray(tempArray) Then 
            currentArray = tempArray
        End If
        tempArray = Empty
    End Sub
    
    'Configura um Array externo para ser operado na classe
    Public Sub setArray(vArr)
        If IsArray(vArr) Then 
            currentArray = vArr
        End If
    End Sub
    
    'Retorna o Array 
    Public Function getArray()
        If IsArray(currentArray) Then 
            getArray = currentArray
        End If
    End Function
    
    'Adiciona um item ao array corrente
    Public Sub Add(arrValue) 
        If IsEmpty(arrValue) Or Not IsArray(currentArray) Then Exit Sub
        ReDim preserve currentArray(UBound(currentArray)+1)
        currentArray(Ubound(currentArray)) = arrValue
    End Sub
    
    'Remove um item do array corrente
    Public Sub Remove(arrValue)
        If Not IsArray(currentArray) OR IsNull(arrValue) Then Exit Sub
        tempArray = Array()
        For xy = 0 To Count
            If currentArray(xy) <> arrValue Then
                ReDim preserve tempArray (UBound(tempArray)+1)
                tempArray(Ubound(tempArray)) = currentArray(xy)
            End If
        Next
        currentArray = tempArray
        tempArray = Empty
    End Sub
        
    'Verifica se existe o valor repassado dentro do array corrente
    Public Function Exists(arrValue)
        If Not IsArray(currentArray) OR IsNull(arrValue) Then Exit Function
        Exists = False
        For xy = 0 To Count
            If trim(currentArray(xy)) = arrValue Then
                Exists = True : Exit Function
            End If
        Next    
    End Function
    
    'Aplica a fun��o Join
    Public Function JoinArray(char)
		JoinArray = Join(currentArray, char)
    End Function
    
	Public Function sortArray(varArray)
		For i = UBound(varArray) - 1 To 1 Step - 1
			MaxVal = varArray(i)
			MaxIndex = i			
			For j = 0 To i
				If varArray(j) > MaxVal Then
					MaxVal = varArray(j)
					MaxIndex = j
				End If
			Next			
			If MaxIndex < i Then
				varArray(MaxIndex) = varArray(i)
				varArray(i) = MaxVal
			End If
		Next 
	End Function
    
    ' 'Retirado de http://www.asptutorial.info/sscript/sortarray.html
    ' Public Function sortArray(varArray, typeOrder)
        ' Dim max, i, j, TemporalVariable
        
        ' Select Case typeOrder
            ' Case "ASC"
                ' max=ubound(varArray)
                ' For i = 0 To max  
                   ' For j = i + 1 To max  
                      ' If varArray(i) > varArray(j) Then 
                          ' TemporalVariable=varArray(i) 
                          ' varArray(i)=varArray(j) 
                          ' varArray(j)=TemporalVariable 
                     ' End If 
                   ' Next  
                ' Next
            ' Case "DESC"
                ' max = UBound(varArray)
                ' For i = 0 To max  
                   ' For j = i + 1 To max  
                      ' If varArray(i) < varArray(j) then 
                          ' TemporalVariable = varArray(i) 
                          ' varArray(i) = varArray(j) 
                          ' varArray(j) = TemporalVariable 
                     ' End If 
                   ' Next
                ' Next 
            ' Case Else
            
        ' End Select
        
        ' sortArray = varArray        
	' End Function
	
    'Escreve o array separado por v�rgulas
    Public Sub Print()
        Response.Write(JoinArray(","))
    End Sub
    
End Class
%>