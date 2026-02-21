Sub RawDataToFormat()
    ' ==========================================================
    ' [변수 선언부] - 용도별로 그룹화하여 가독성 향상
    ' ==========================================================
    
    ' 1. 시트 및 범위 제어 변수
    Dim ws As Worksheet          ' 작업할 메인 시트
    Dim refWs As Worksheet       ' 참조할 "호수-위치" 데이터베이스 시트
    Dim rng As Range             ' 데이터가 있는 전체 범위 (반복문용)
    Dim cell As Range            ' 반복문에서 현재 작업 중인 셀 하나
    
    ' 2. [D열] 데이터 가공 변수 (숫자 분리 및 줄바꿈)
    Dim rawStr As String         ' 셀에 있는 원본 문자열 (예: "63077063")
    Dim newStr As String         ' 가공된 결과 문자열 (예: "63(줄바꿈)77...")
    Dim i As Integer             ' 반복문 카운터 (Loop Index)
    
    ' 3. [E열] 위치 매칭 변수 (VLOOKUP)
    Dim arrCodes As Variant      ' 줄바꿈으로 쪼갠 제품번호들을 담을 배열
    Dim codeItem As Variant      ' 배열 내의 제품번호 하나
    Dim locResult As String      ' 최종적으로 합쳐진 위치 결과값
    Dim singleLoc As Variant     ' VLOOKUP 검색 결과 (에러 처리를 위해 Variant 사용)

    ' 4. [F, G열] 텍스트 파싱 변수 (괄호 추출)
    Dim strG As String           ' G열(공역)의 전체 텍스트
    Dim startPos As Integer      ' 여는 괄호 "(" 의 위치
    Dim endPos As Integer        ' 닫는 괄호 ")" 의 위치
    Dim extractedText As String  ' 괄호 안에서 추출한 최종 텍스트

    ' ==========================================================
    ' [초기 설정] 시트 지정 및 화면 설정
    ' ==========================================================
    Set ws = ActiveSheet
    
    ' 참조 시트(xx-xx)가 실제로 존재하는지 확인
    On Error Resume Next
    Set refWs = Sheets("xx-xx")
    On Error GoTo 0
    
    If refWs Is Nothing Then
        MsgBox "오류: 'xx-xx' 시트가 없습니다!", vbCritical
        Exit Sub
    End If

    ' 작업 속도 향상 및 과정 시각화를 위한 설정
    Application.ScreenUpdating = True
    
    ' [0단계] 시작 알림
    Application.StatusBar = "매크로 시작..."
    Call DelayTime(0.5)

    ' ==========================================================
    ' [1단계] 불필요한 열 삭제 (G -> E -> D 순서로 안전하게)
    ' ==========================================================
    ws.Columns("G").Delete Shift:=xlToLeft   ' G열 삭제
    ws.Columns("E").Delete Shift:=xlToLeft   ' E열 삭제
    ws.Columns("D").Delete Shift:=xlToLeft   ' D열 삭제
    
    Application.StatusBar = "1단계: 열 삭제 완료"
    Call DelayTime(0.5)

    ' ==========================================================
    ' [2~3단계] 새로운 열 추가 (위치, 조)
    ' ==========================================================
    ' E열 삽입 (위치)
    ws.Columns("E").Insert Shift:=xlToRight
    ws.Range("E1").Value = "위치"
    
    ' I열 삽입 (조)
    ws.Columns("I").Insert Shift:=xlToRight
    ws.Range("I1").Value = "조"
    
    Application.StatusBar = "2-3단계: 열 추가 완료"
    Call DelayTime(0.5)

    ' ==========================================================
    ' [4단계] D열 데이터 가공 (0 제거 및 줄바꿈 처리)
    ' ==========================================================
    Application.StatusBar = "4단계: D열 데이터 가공 중..."
    
    ' D2부터 데이터가 있는 끝까지 범위 설정
    Set rng = ws.Range("D2", ws.Cells(ws.Rows.Count, "D").End(xlUp))
    
    For Each cell In rng
        rawStr = cell.Value
        newStr = ""
        
        ' 데이터가 2자리 이상일 때만 처리
        If Len(rawStr) >= 2 Then
            ' [핵심 로직] 1부터 시작해서 3칸씩 점프 (1, 4, 7...) -> 중간의 '0'을 건너뜀
            For i = 1 To Len(rawStr) Step 3
                
                ' 첫 번째 줄이 아닐 때만 줄바꿈(Enter) 추가
                If i > 1 Then newStr = newStr & vbLf
                
                ' 현재 위치(i)에서 딱 2글자만 가져옴 (제품번호 추출)
                newStr = newStr & Mid(rawStr, i, 2)
            Next i
            
            cell.Value = newStr
            cell.WrapText = True   ' 셀 줄바꿈 속성 켜기
        End If
    Next cell
    
    Call DelayTime(0.5)

    ' ==========================================================
    ' [5단계] E열 위치 매칭 (VLOOKUP: 문자/숫자 형식 모두 대응)
    ' ==========================================================
    Application.StatusBar = "5단계: 위치 데이터 매칭 중..."
    
    For Each cell In rng
        ' 셀 값을 줄바꿈(Enter) 기준으로 쪼개서 배열로 만듦
        arrCodes = Split(cell.Value, vbLf)
        locResult = ""
        
        For Each codeItem In arrCodes
            codeItem = Trim(codeItem)  ' 공백 제거
            
            ' 1차 시도: 텍스트 형식으로 찾기
            singleLoc = Application.VLookup(codeItem, refWs.Range("A:B"), 2, 0)
            
            ' 2차 시도: 못 찾았고(Error) + 내용이 숫자라면? -> 숫자(Val)로 변환해 다시 찾기
            If IsError(singleLoc) And IsNumeric(codeItem) Then
                singleLoc = Application.VLookup(Val(codeItem), refWs.Range("A:B"), 2, 0)
            End If
            
            ' 그래도 못 찾았으면 "확인불가" 처리
            If IsError(singleLoc) Then singleLoc = "확인불가"
            
            ' 결과 이어붙이기 (줄바꿈 포함)
            If locResult = "" Then
                locResult = singleLoc
            Else
                locResult = locResult & vbLf & singleLoc
            End If
        Next codeItem
        
        ' 완성된 값을 오른쪽(E열)에 입력
        cell.Offset(0, 1).Value = locResult
        cell.Offset(0, 1).WrapText = True
    Next cell
    
    Call DelayTime(0.5)

    ' ==========================================================
    ' [6단계] F열(T/O) 데이터 병합: G열의 괄호(...) 내용 가져오기
    ' ==========================================================
    Application.StatusBar = "6단계: x/x 및 xxx 데이터 병합 중..."
    
    ' F열 데이터 범위 재설정 (혹시 데이터 길이가 다를 수 있으므로)
    Set rng = ws.Range("F2", ws.Cells(ws.Rows.Count, "F").End(xlUp))
    
    For Each cell In rng
        strG = cell.Offset(0, 1).Value       ' 바로 오른쪽(G열) 값을 가져옴
        startPos = InStr(1, strG, "(")       ' 여는 괄호 "(" 위치 찾기
        endPos = InStr(startPos, strG, ")")  ' 닫는 괄호 ")" 위치 찾기
        
        ' 괄호 쌍이 정상적으로 존재할 때만 실행
        If startPos > 0 And endPos > startPos Then
            
            ' 괄호 시작부터 끝까지 문자열 추출
            extractedText = Mid(strG, startPos, endPos - startPos + 1)
            
            ' 기존 값 + 줄바꿈 + 추출한 괄호 내용
            cell.Value = cell.Value & vbLf & extractedText
            cell.WrapText = True
        End If
    Next cell
    
    Call DelayTime(1)

    ' ==========================================================
    ' [7단계] G열 잘라내서 K열로 이동 (Cut & Insert)
    ' ==========================================================
    Application.StatusBar = "7단계: G열 -> K열 이동 중..."
    
    ws.Columns("G").Cut                         ' G열 잘라내기 (클립보드 저장)
    ws.Columns("K").Insert Shift:=xlToRight     ' K열 위치에 끼워넣기 (기존 열은 밀려남)
    Application.CutCopyMode = False             ' 클립보드 점선 해제
    
    Application.StatusBar = "7단계: 이동 완료"
    Call DelayTime(1)

    ' ==========================================================
    ' [완료] 종료 알림
    ' ==========================================================
    MsgBox "모든 작업이 성공적으로 완료되었습니다!", vbInformation, "작업 완료"
    Application.StatusBar = False
End Sub

' [보조 함수] 지정된 시간(초)만큼 대기하는 함수
Sub DelayTime(Seconds As Double)
    Dim StartTime As Double: StartTime = Timer
    Do While Timer < StartTime + Seconds: DoEvents: Loop
End Sub
