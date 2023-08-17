'define array range
Public Const element = 0
Public Const timeVal = 1
Public Const rowRanInit = 5
Public Const rowRanFin = 25

'define print range
Public Const printColumnRanInit = 9 'A column
Public Const printRowRanInit = 5

Public Const ran = printRowRanInit + rowRanFin - rowRanInit

'-------------------------------------------------------

Function showArr(arr, col1, col2) 'show arr
	'* arr에 들어있는 값들을 for문으로 정해진 셀 범위에 출력   
    For y = rowRanInit To rowRanFin
        Range(col1 & y).Value = arr(element,y) 
    Next
    
    For y = rowRanInit To rowRanFin
        Range(col2 & y).Value = arr(timeVal,y) 
    Next
End Function

Function showArr2(arrPrintLoc, personNum) 'arrPrintLoc
	'* arrPrintLoc에 들어있는 값들을 for문으로 정해진 셀 범위에 출력
    For x = 1 To 4
    	For y = 1 To 4
        	Cells(y+30 , x).Value = arrPrintLoc(x, y)
        Next
    Next
End Function

Function delCol(arr, col) 'delete column
	'* arr에서 정해진 column의 row값들을 제거   
	arr(element,col) = Empty
	arr(timeVal,col) = 0
End Function

Function findAm(arr, amTime, col) 'find 3:30, 시간 3개의 합까지 찾을 수 있는데 추가하면 더 가능	
	'* [시간 값] 또는 [시간 값들의 합]이 210(3:30)인 애들을 찾고 print, 3:30 뒤에 "점심시간"까지 print
	'* 3:30을 만족하는 애가 있으면 걔를 최우선적으로 넣어줄려고 for문을 따로 만듦
   	For i = rowRanInit To rowRanFin  
   		If arr(timeVal,i) = amTime Then
   	    	am1 = i
        	Cells(printRowRanInit , col).Value = arr(element,am1) '<Cells>는 <Range>와 x,y 순서가 반대  
        	Cells(printRowRanInit , col).Offset(0, 1).Value = arr(timeVal,am1) '<Offset>은 y,x 순서
        	Cells(printRowRanInit+1 , col).Value = "점심시간"   
        	
        	Call delCol(arr, am1)	    
			'Call showArr(arr, "R", "S")
			
		    Exit Function
		End If
	Next
	
	For i = rowRanInit To rowRanFin
    	For j = i+1 To rowRanFin
    		x = timeVal
      		If arr(x,i)+arr(x,j) = amTime and arr(x,i) <> 0 and arr(x,j) <> 0 Then
      			am1 = i
      			am2 = j	
      			Cells(printRowRanInit , col).Value = arr(element,am1)
      			Cells(printRowRanInit , col).Offset(0, 1).Value = arr(timeVal,am1)    
        		Cells(printRowRanInit+1 , col).Value = arr(element,am2) 
        		Cells(printRowRanInit+1 , col).Offset(0, 1).Value = arr(timeVal,am2)
        		Cells(printRowRanInit+2 , col).Value = "점심시간"    
        		
        		Call delCol(arr, am1)
        		Call delCol(arr, am2)	    
				'Call showArr(arr, "R", "S")  
				
				Exit Function			
      		End If
      	Next
	Next
	
	For i = rowRanInit To rowRanFin
    	For j = i+1 To rowRanFin
      		For k = j+1 To rowRanFin
	      		If arr(x,i)+arr(x,j)+arr(x,k) = amTime and arr(x,i) <> 0 and arr(x,j) <> 0 and arr(x,k) <> 0 Then
	      			am1 = i
	      			am2 = j	
	      			am3 = k
	      			Cells(printRowRanInit , col).Value = arr(element,am1)
	      			Cells(printRowRanInit , col).Offset(0, 1).Value = arr(timeVal,am1)    
	        		Cells(printRowRanInit+1 , col).Value = arr(element,am2) 
	        		Cells(printRowRanInit+1 , col).Offset(0, 1).Value = arr(timeVal,am2)
	        		Cells(printRowRanInit+2 , col).Value = arr(element,am3) 
	        		Cells(printRowRanInit+2 , col).Offset(0, 1).Value = arr(timeVal,am3)
	        		Cells(printRowRanInit+3 , col).Value = "점심시간"    
	        		
	        		Call delCol(arr, am1)
	        		Call delCol(arr, am2)	 
	        		Call delCol(arr, am3)   
					'Call showArr(arr, "R", "S")  
					
					Exit Function			
	      		End If
	      	Next
      	Next
    Next
End Function

Function maxfn(arr, maxVal, maxIdx) 'max
	'* arr에서 최댓값을 찾아주는 함수
	For y = rowRanInit To rowRanFin
        If arr(timeVal, y) > maxVal Then
        	maxVal = arr(timeVal, y)
        	maxIdx = y
        End If
    Next
End Function

Function countfn(arr)
	'* arr에 남아있는 값의 개수
    countfn = 0
    For i = rowRanInit To rowRanFin
    	If arr(timeVal, i) <> 0 Then
    		countfn = countfn + 1
    	End If
    Next
End Function

Function timeSum(arrPrintLoc, personNum)
	'* arrPrintLoc의 각 column의 [시간 값]을 다 더하고 출력 & arrPrintLoc에 입력
	For i = 1 To personNum	
    	timeSum = 0
    	For j = printRowRanInit To arrPrintLoc(3,i)
    		timeSum = timeSum + Cells(j, arrPrintLoc(2,i)+1).Value
    		Cells(ran, arrPrintLoc(2,i)).Value = "timeSum: "
    		Cells(ran, arrPrintLoc(2,i)+1).Value = timeSum
    	Next
    	arrPrintLoc(4,i) = Cells(ran, arrPrintLoc(2,i)+1).Value  '왜 바로 timeSum을 넣으면 값이 이상하게 들어가는 지 모르겠는데 이렇게 한 번 경유하면 잘 됨
    Next
End Function

Function minWho(arrPrintLoc, personNum)
	'[timeSum]이 최소인 사람을 찾는 함수
	'Call showArr2(arrPrintLoc, personNum)
	minVal = 1000
	For i = 1 To personNum
        If arrPrintLoc(4,i) < minVal Then
        	minVal = arrPrintLoc(4,i)
        	minWho = i
        End If
    Next
End Function

'---------------------------------------------------

Sub TimetableWizard()
	personNum = Range("C26").Value 'set person number
	If personNum <> 2 and personNum <> 3 and personNum <> 4 Then
		 MsgBox ("Error" & vbCrLf & "사람 수는 2, 3, 4만 가능합니다.")
		 Exit Sub
	End If
		 
	Dim arr(element To timeVal, rowRanInit To rowRanFin) 'define 2D array
	
	'2D array LD
	'* 정해진 범위에서 [요소]와 [시간 값]을 로드에서 arr에 입력		
    For y = rowRanInit To rowRanFin
        arr(element, y) = Range("A" & y).Value
    Next
    
    For y = rowRanInit To rowRanFin '3:30 -> 210
        Range("G" & y).Value = (Hour(Range("F" & y).Value)*60)+Minute(Range("F" & y).Value)
    Next
    
    For y = rowRanInit To rowRanFin
        arr(timeVal, y) = Range("G" & y).Value
    Next
    
    'Call showArr(arr, "R", "S")
    
    'print clear
    '* print 할 범위의 셀들을 모두 clear
    For i = printColumnRanInit To printColumnRanInit + 8
    	For j = printRowRanInit To printRowRanInit + rowRanFin - rowRanInit
    		Cells(j , i).Font.Color = RGB(0, 0, 0)
    		Cells(j , i).Value =  Empty
    	Next
    Next
    
    'find 3:30, distribute AM
    '* <findAm> 함수로 오전을 채움
    For i = 1 To personNum
    	Call findAm(arr, 210, printColumnRanInit+((i-1)*2))
    Next
    
    '* 오전을 3:30으로 딱맞게 못채우는 error
    lunchCount = 0
    For i = printColumnRanInit To printColumnRanInit + 8
    	For j = printRowRanInit To printRowRanInit + 15
    		If Cells(j , i).Value = "점심시간" Then
    			lunchCount = lunchCount + 1
    		End If
    	Next
    Next
    If lunchCount <> personNum Then
    	MsgBox ("Error" & vbCrLf & "모든 인원의 오전을 3:30로 딱맞게 채울 수가 없습니다." & vbCrLf & "따라서 오늘은 시간표마법사를 사용할 수 없습니다.")
    	Exit Sub
    End If
        
    'distribute PM, 여기서부터 오후를 채우는 알고리즘
    
    '각각의 사람들의 [시간 값]을 print할 위치를 넣는 array
    Dim arrPrintLoc(1 To 4, 1 To 4)
    For i = 1 To personNum
    	arrPrintLoc(1, i) = "person" & i
    Next   
    
    '점심시간의 위치를 각각 찾기 위한 알고리즘 - 배열에 점심시간 다음 셀의 위치를 넣음
    iNum = printColumnRanInit - 1 '배열의 알맞은 위치에 넣기 위한 변수
    For i = printColumnRanInit To printColumnRanInit+((personNum-1)*2) Step 2
    	arrPrintLoc(2, (i-iNum+1)/2) = i
    	For j = printRowRanInit To printRowRanInit+5
    		If Cells(j, i).Value = "점심시간" Then
    			arrPrintLoc(3, (i-iNum+1)/2) = j+1
    			Exit For
    		End If
    	Next
    Next

    count = countfn(arr)  
    timeNumQ = count \ personNum
    
    'distribute PM & print
    '* alpha, bravo, charlie, charlie, bravo, alpha, alpha ... 순으로 [시간 값]을 큰 거부터 차례로 넣기
    maxVal = 0
    maxIdx = 0
    
    For j = 1 To personNum '[시간 값]을 일단 모든 사람에게 하나씩 각각 넣어줌
	    maxVal = 0
	   	Call maxfn(arr, maxVal, maxIdx)
	   	Cells(arrPrintLoc(3,j), arrPrintLoc(2,j)).Value = arr(element, maxIdx)
	  	Cells(arrPrintLoc(3,j), arrPrintLoc(2,j)+1).Value = maxVal
	  	arrPrintLoc(3,j) = arrPrintLoc(3,j) + 1
	   	Call delcol(arr, maxIdx)
	Next
    
    For i = 1 To 100 '[timeVal]이 최소인 사람에게 최우선적으로 [시간 값]을 하나씩 넣어줌
    	count = countfn(arr) '무의미한 값들을 없애기 위해 계속 count하면서 arr에 [시간 값]이 있어야만 출력
    	If count = 0 Then
    		Call timeSum(arrPrintLoc, personNum)
	  		Exit For
    	End If
	    Call timeSum(arrPrintLoc, personNum)
	    j = minWho(arrPrintLoc, personNum)
	    'Call showArr2(arrPrintLoc, personNum)
	    maxVal = 0
	   	Call maxfn(arr, maxVal, maxIdx)
	   	Cells(arrPrintLoc(3,j), arrPrintLoc(2,j)).Value = arr(element, maxIdx)
	  	Cells(arrPrintLoc(3,j), arrPrintLoc(2,j)+1).Value = maxVal
	  	arrPrintLoc(3,j) = arrPrintLoc(3,j) + 1
	  	Call delcol(arr, maxIdx)	
    Next
    
	'Call showArr(arr, "R", "S")
End Sub
