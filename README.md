# Description 
하루의 업무를 상번 인원에게 시간을 공평하게 분배하여 시간표를 제작하는 데 도움을 주는 프로그램이다. 업무들의 요소와 시간을 입력하면 모두에게 시간이 공평하게 돌아가도록 업무를 분배해주며 점심시간 전 오전에 3:30 업무 분배까지 수행한다. 공군에 맞게 제작되었기에 엑셀이 아닌 한셀 기반으로 제작되었으며 VBA가 아닌 VBS로 제작되어 엑셀과는 호환이 안 될 수 있다.

---

# Algorithm
## 예방정비 복붙
- [x] 예방정비 복붙 - ctrl+c -> ctrl+alt+v -> 값 및 숫자 서식 붙여넣기

## 배열에 예방정비 집어 넣기

## 3:30 찾기
- [x] 3:30이 있으면 그걸 불러오기
​
- [x] 3:30이 없으면 합이 3:30 되는 애들 불러오기

## main
### 큰 거부터 차례로 넣기 - 단순
- [x] alpha, bravo, charlie, charlie, bravo, alpha, alpha ... 순으로 큰 거부터 차례로 넣기
- [x] 시간 합을 다 더해서 "최댓값에서부터의 차이"(차이)를 보기

\timetableWizard beta.vbs

### 시간 합이 제일 작은 사람부터 하나씩 넣기
- [x] 각 사람별로 시간 합을 더해서 시간 합이 제일 작은 사람 찾기
- [x] 그 사람부터 하나씩 큰 값을 넣기

\timetableWizard.vbs

## 기타
- [x] 설명이랑 주석 달 필요가 있음
- [x] 그리고 vba 함수들 정리할 필요도 있음
---

# VBA (VBS)
## 기본 구조
```vbs
Sub helloWorld()
    'Sheet1의 "A1"에 "hello world"를 출력 하시오.
    Sheet1.Range("A1").Value = "hello world"
End Sub
```

## 주석
'로 주석 사용

## 셀 선택 & 값 입력, 셀 값 로드
```vbs
Sub 셀에내용추가하기()
	row = 2
    col = 3 'A열 = 1, B열 = 2, ...

    '셀 선택하기
	Range("C" & 2).Select
	Range("C2").Select
    Cells(row, col).Select

    '셀에 내용 추가하기
    Selection.Value = "2행 3열"

	'값 입력
	Range("C2").Value = "hello Wolrd"
	Cells(row, col).Value = "hello Wolrd"
	Cells(row, col).Value = Empty '빈칸

	'셀 값 로드
	 row = Range("C2").Value
End Sub
```

## 글자 연결
```vbs
Sub 글자연결하기()
	무엇 = Range("A1").Value
    MsgBox ("나는" & 무엇 & "(이)다.")
End Sub
```

## 변수
`a = 1`

* VBS는 VBA와 다르게 자료형 선언하면 안 됨

## 반복
```vbs
Sub For문배우기()
    For 반복범위 = 1 To 10 '* 1~10, 1~9 X
        Range("F" & 반복범위).Value = "반복" & 반복범위
    Next

	For i = 1 To 10 Step 2
    	arrPrintLoc(2, (i+1)/2) = i
    Next
End Sub
```

## 조건
```vbs
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

## 상수 define
`Public Const element = 0`

## 함수 선언, 호출, 값 반환
```vbs
'선언
Function delCol(arr, col) 'delete column
	'* arr에서 정해진 column의 row값들을 제거   
	arr(element,col) = Empty
	arr(timeVal,col) = 0
End Function

Function countfn(arr)
	'* arr에 남아있는 값의 개수
    countfn = 0 '값을 반환하려면 함수 이름과 변수명이 같아야 함 
    For i = rowRanInit To rowRanFin
    	If arr(timeVal, i) <> 0 Then
    		countfn = countfn + 1
    	End If
    Next
End Function

'호출
Call delCol(arr, am1)

'값 반환
count = countfn(arr)
```

## 메세지창
```vbs
Sub TimetableWizard()
	MsgBox ("Error" & vbCrLf & "사람 수는 1, 2, 3, 4만 가능합니다.") 'vba에서 \n은 vbCrLf로 코드처럼 사용해야함
End Sub
```

## 배열
`Dim arr(1 To 4, 2 To 10)`

`arr(1, 2) = Range("A1").Value`

## 글자색
`Cells(j , i).Font.Color = RGB(0, 0, 0)`

## 연산자
![image](https://github.com/Arduriz/timetableWizard/assets/65582244/d45ab85e-caf3-4dfc-b5ee-d77f446c97fc)


* 셀주소
```vbs
Sub 선택셀주소가져오기()
	Range("b2").Select
	Range("a10").Value = Selection.Address
End Sub
```

---

## 3:30 -> 210분
`Range("a1").Value = (Hour(Range("a2").Value)*60)+Minute(Range("a2").Value)`

## 매크로 실행 버튼
`입력` - `양식 개체` - `명령 단추`
왜인지 모르겠는데 버튼을 생성하면 vbs를 수정할 수 없는 버그가 있는 거 같음







