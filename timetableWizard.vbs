'define array range
Public Const element = 0
Public Const timeVal = 1
Public Const rowRanInit = 2
Public Const rowRanFin = 10

'define print range
Public Const printColumnRanInit = 1 'A
Public Const printRowRanInit = 15

'define etc.
Public Const personNum = 3

Function showArr(arr, col1, col2) 'show arr   
    x = element
    For y = rowRanInit To rowRanFin
        Range(col1 & y).Value = arr(x,y) 
    Next
    
    x = timeVal
    For y = rowRanInit To rowRanFin
        Range(col2 & y).Value = arr(x,y) 
    Next
End Function

Function delCol(arr, col) 'delete column   
	For x = element To timeVal
		arr(x,col) = 0
	Next
End Function

Function findAm(arr, amTime, col) 'find 3:30, 3개의 합까지 찾을 수 있는데 추가하면 더 가능
   	x = timeVal   	
   	For i = rowRanInit To rowRanFin '3:30을 만족하는 애가 있으면 걔를 최우선적으로 넣어줄려고 for문을 따로 만듦 
   		If arr(x,i) = amTime Then
   	    	am1 = i
        	Cells(printRowRanInit , col).Value = arr(element,am1) '<Cells>는 <Range>와 x,y 순서가 반대  
        	Cells(printRowRanInit , col).Offset(0, 1).Value = arr(timeVal,am1) '<Offset>은 y,x 순서
        	Cells(printRowRanInit+1 , col).Value = "점심시간"   
        	
        	Call delCol(arr, am1)	    
			Call showarr(arr, "D", "E")
			
		    Exit Function
		End If
	Next
	
	For i = rowRanInit To rowRanFin
    	For j = i+1 To rowRanFin
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
				Call showarr(arr, "D", "E")  
				
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
					Call showarr(arr, "D", "E")  
					
					Exit Function			
	      		End If
	      	Next
      	Next
    Next
End Function

Function maxfn(arr, maxVal, maxIdx) 'max
	For y = rowRanInit To rowRanFin
        If arr(timeVal, y) > maxVal Then
        	maxVal = arr(timeVal, y)
        	maxIdx = y
        End If
    Next
End Function

Sub TimetableWizard()
	Dim arr(element To timeVal, rowRanInit To rowRanFin) 'define 2D array
	
	'2D array LD		
	x = element
    For y = rowRanInit To rowRanFin
        arr(x, y) = Range("A" & y).Value
    Next
    
    For y = rowRanInit To rowRanFin '3:30 -> 210
        Range("C" & y).Value = (Hour(Range("B" & y).Value)*60)+Minute(Range("B" & y).Value)
    Next
    
    x = timeVal
    For y = rowRanInit To rowRanFin
        arr(x, y) = Range("C" & y).Value
    Next
    
    Call showArr(arr, "D", "E")
    
    'time Sum
    timeSum = 0
    
    x = timeVal
    For y = rowRanInit To rowRanFin
        timeSum = arr(x,y) + timeSum 
    Next
    
    Range("F" & 1).Value = timeSum
    
    'find 3:30, distribute AM
    For i = 1 To personNum
    	Call findAm(arr, 210, printColumnRanInit+((i-1)*2))
    Next
    
    'distribute PM
    '점심시간의 위치를 각각 찾기 위한 알고리즘 - 배열에 점심시간 다음 셀의 위치를 넣음
    Dim arrPrintLoc(1 To 3, 1 To 4)
    For i = 1 To personNum
    	arrPrintLoc(1, i) = "person" & i
    Next   
    For i = printColumnRanInit To printColumnRanInit+((personNum-1)*2) Step 2
    	arrPrintLoc(2, (i+1)/2) = i
    	For j = printRowRanInit To printRowRanInit+5
    		If Cells(j, i).Value = "점심시간" Then
    			arrPrintLoc(3, (i+1)/2) = j+1
    			Exit For
    		End If
    	Next
    Next 

    'arr에 남아있는 값의 개수
    count = 0
    For i = rowRanInit To rowRanFin
    	If arr(timeVal, i) <> 0 Then
    		count = count + 1
    	End If
    Next
    
    timeNumQ = count \ personNum
    timeNumR = count Mod personNum
    
    'distribute PM
    maxVal = 0
    maxIdx = 0
    
    For i = 1 To timeNumQ
    	For j = 1 To personNum
    		maxVal = 0
    		Call maxfn(arr, maxVal, maxIdx)
    		Cells(arrPrintLoc(3,j), arrPrintLoc(2,j)).Value = arr(element, maxIdx)
    		Cells(arrPrintLoc(3,j), arrPrintLoc(2,j)+1).Value = maxVal
    		arrPrintLoc(3,j) = arrPrintLoc(3,j) + 1
    		Call delcol(arr, maxIdx)
    	Next
    	For j = personNum To 1 Step -1
    		maxVal = 0
    		Call maxfn(arr, maxVal, maxIdx)
    		Cells(arrPrintLoc(3,j), arrPrintLoc(2,j)).Value = arr(element, maxIdx)
    		Cells(arrPrintLoc(3,j), arrPrintLoc(2,j)+1).Value = maxVal
    		arrPrintLoc(3,j) = arrPrintLoc(3,j) + 1
    		Call delcol(arr, maxIdx)
    	Next
    Next
    
    'timeSum
    For i = 1 To personNum	
    	timeSum = 0
    	For j = printRowRanInit To arrPrintLoc(3,i)
    		timeSum = timeSum + Cells(j, arrPrintLoc(2,i)+1).Value
    		Cells(arrPrintLoc(3,i)+2, arrPrintLoc(2,i)+1).Value = timeSum	
    	Next
    Next
    
    Call showArr(arr, "D", "E")   
End Sub