# timetableWizard

# Algorithm
- [ ] 예방정비 복붙
​
- [ ] 3:30이 있으면 그걸 불러오고 없으면 이거보다 작은 것들 중 제일 큰 걸 불러오기
​
- [ ] {3:30보다 작은 걸 불러왔으면  3:30-[불러온 것] 로드} 3:30 될 때까지 반복
​
- [ ] 남은 것들 나누기 인원 수대로
​
- [ ] 그거에 맞게 또 로드해서 하기(큰 것부터)

---

# Excel Fn.
## 3:30 -> 210분
=(HOUR(A2)*60)+MINUTE(A2)

# VBA
## 기본 구조
```
Sub helloWorld()
    'Sheet1의 "A1"에 "hello world"를 출력 하시오.
    Sheet1.Range("A1").Value = "hello world"
End Sub
```

## 주석
'로 주석 사용

## Range
`Range("a1").Value = "hello Wolrd"`

## 셀 선택
```
Sub 셀에내용추가하기()
	행 = 2
    열 = 3

    '셀 선택하기
    Cells(행, 열).Select

    '셀에 내용 추가하기
    Selection.Value = "2행 3열"
End Sub
```

## 글자 연결
```
Sub 글자연결하기()
	무엇 = Range("A1").Value
    MsgBox ("나는" & 무엇 & "(이)다.")
End Sub
```

## 변수
`a=1`

## 반복
```
Sub For문배우기()
    For 반복범위 = 1 To 10
        Range("F" & 반복범위).Value = "반복" & 반복범위
    Next
End Sub
```

## 조건
```
Sub if문배우기()
	사원명 = Range("b2").Value
	부서 = Range("c2").Value

	If 사원명 = "김경록" Then
    	MsgBox ("해당 사원명은 김경록이 맞습니다.")
	Else
    	MsgBox ("해당 사원명은 김경록이 아닙니다.")
	End If
End Sub
```

## 셀주소
```
Sub 선택셀주소가져오기()
	Range("b2").Select
	Range("a10").Value = Selection.Address
End Sub
```

## 3:30 -> 210분
`Range("a1").Value = (Hour(Range("a2").Value)*60)+Minute(Range("a2").Value)`

---

# Program
## 2차원 배열로 로드해서 최대값 2개 찾기
```
Sub Test()
	'2D array LD
	Dim arr(0 To 1, 2 To 10)
	
	x = 0
	y = 0
    For y = 2 To 10
        arr(x, y) = Range("A" & y).Value
    Next
    
    y = 0
    For y = 2 To 10
        Range("C" & y).Value = (Hour(Range("B" & y).Value)*60)+Minute(Range("B" & y).Value)
    Next
    
    x = 1
	y = 0
    For y = 2 To 10
        arr(x, y) = Range("C" & y).Value
    Next
    
    'show arr   
    x = 0
	y = 0
    For y = 2 To 10
        Range("D" & y).Value = arr(x,y) 
    Next
    
    x = 1
	y = 0
    For y = 2 To 10
        Range("E" & y).Value = arr(x,y) 
    Next
    
    'time Sum
    timeSum = 0
    
    x = 1
	y = 0 
    For y = 2 To 10
        timeSum = arr(x,y) + timeSum 
    Next
    
    Range("F" & 1).Value = timeSum
    
    'max
    maxIdx = 0
    
    x = 1
	y = 0
    For y = 2+1 To 10
        If arr(x,y) > arr(x,y-1) Then
        	maxIdx = y       	
        End If
    Next
    
    Range("G" & 2).Value = arr(0,maxIdx)
    Range("H" & 2).Value = arr(1,maxIdx)
    
    'delete max
    y = maxIdx
    For x = 0 To 1
    	arr(x,y) = 0
    Next
    
    x = 0 'show arr
	y = 0
    For y = 2 To 10
        Range("D" & y).Value = arr(x,y) 
    Next
    
    x = 1
	y = 0
    For y = 2 To 10
        Range("E" & y).Value = arr(x,y) 
    Next 
    
    'max 2nd
    maxIdx = 0
    
    x = 1
	y = 0
    For y = 2+1 To 10
        If arr(x,y) > arr(x,y-1) Then
        	maxIdx = y       	
        End If
    Next
    
    Range("G" & 3).Value = arr(0,maxIdx)
    Range("H" & 3).Value = arr(1,maxIdx)
       
    y = maxIdx 'delete
    For x = 0 To 1 
    	arr(x,y) = 0
    Next
    
    x = 0 'show arr
	y = 0
    For y = 2 To 10
        Range("D" & y).Value = arr(x,y) 
    Next
    
    x = 1
	y = 0
    For y = 2 To 10
        Range("E" & y).Value = arr(x,y) 
    Next
      
End Sub
```


## find 3:30 미완성
```
Sub Test()
	'2D array LD
	Dim arr(0 To 1, 2 To 10)
	
	x = 0
	y = 0
    For y = 2 To 10
        arr(x, y) = Range("A" & y).Value
    Next
    
    y = 0
    For y = 2 To 10
        Range("C" & y).Value = (Hour(Range("B" & y).Value)*60)+Minute(Range("B" & y).Value)
    Next
    
    x = 1
	y = 0
    For y = 2 To 10
        arr(x, y) = Range("C" & y).Value
    Next
    
    'show arr   
    x = 0
	y = 0
    For y = 2 To 10
        Range("D" & y).Value = arr(x,y) 
    Next
    
    x = 1
	y = 0
    For y = 2 To 10
        Range("E" & y).Value = arr(x,y) 
    Next
    
    'time Sum
    timeSum = 0
    
    x = 1
	y = 0 
    For y = 2 To 10
        timeSum = arr(x,y) + timeSum 
    Next
    
    Range("F" & 1).Value = timeSum
    
	'find 3:30
	am = 0
	
	x = 1
	y = 0 
    For y = 2 To 10
        If arr(x,y) = 210 Then
        	am = y
        	Range("G" & 2).Value = arr(0,maxIdx)
    
        	Range("H" & 2).Value = arr(1,maxIdx)
        	   	
    		y = am 'delete
		    For x = 0 To 1
		    	arr(x,y) = 0
		    Next
		    
		    x = 0 'show arr
			y = 0
		    For y = 2 To 10
		        Range("D" & y).Value = arr(x,y) 
		    Next
		    
		    x = 1
			y = 0
		    For y = 2 To 10
		        Range("E" & y).Value = arr(x,y) 
		    Next
		End If
	Next
    
    If am = 0 Then
    	x = 1   	
    	i = 0
    	j = 0
    	am1 = 0
    	am2 = 0
    	For i = 2 To 10-1
    		For j = i+1 To 10
      			If arr(x,i)+arr(x,j) = 210
End Sub
```
      





