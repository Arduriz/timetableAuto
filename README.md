# timetableAuto

# Algorithm
- 예방정비 복붙
​
- 3:30이 있으면 그걸 불러오고 없으면 이거보다 작은 것들 중 제일 큰 걸 불러오기
​
- {3:30보다 작은 걸 불러왔으면  3:30-[불러온 것] 로드} 3:30 될 때까지 반복
​
- 남은 것들 나누기 인원 수대로
​
- 그거에 맞게 또 로드해서 하기(큰 것부터)

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



