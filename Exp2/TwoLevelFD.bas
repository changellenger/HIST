Attribute VB_Name = "TwoLevelFD"

'완전요인설계
Sub fullFD(a1, nfact, wsheet)
    '요인
    For k = 1 To Frm_CreateFD.ComboBox2.value
        For j = 1 To nfact
            For i = 0 To (a1 - 1)
                If i Mod (2 ^ nfact) / (2 ^ (j - 1)) < (2 ^ nfact) / (2 * (2 ^ (j - 1))) Then
                        Selection.Offset(a1 * (k - 1) + (i + 1), j) = 1
                Else
                    Selection.Offset((i + 1) + a1 * (k - 1), j) = -1
                End If
            Next i
        Next j
    Next k
    '중심점 추가
    If Frm_CreateFD.ComboBox4.value >= 1 Then
        For j = 1 To nfact
            For i = 1 To Frm_CreateFD.ComboBox4.value * Frm_CreateFD.ComboBox3.value
                Selection.Offset(Frm_CreateFD.ComboBox2.value * a1 + i, j) = 0
            Next i
        Next j
    End If
    '블록 설정
    block a1, 1
    
End Sub

'1/2요인설계
Sub halfFD(a1, nfact, wsheet)
    '요인
    For k = 1 To Frm_CreateFD.ComboBox2.value
        For j = 1 To (nfact - 1)
            For i = 0 To (a1 / 2 - 1)
                If i Mod (2 ^ (nfact - 1)) / (2 ^ (j - 1)) < (2 ^ (nfact - 1)) / (2 * (2 ^ (j - 1))) Then
                    Selection.Offset((a1 / 2) * (k - 1) + (i + 1), j) = 1
                Else
                    Selection.Offset((a1 / 2) * (k - 1) + (i + 1), j) = -1
                End If
            Next i
        Next j

        value = 1
        For i = 1 To (a1 / 2) * k
            For j = 1 To (nfact - 1)
                value = value * Selection.Offset(i, j)
            Next j
            Selection.Offset(i, nfact) = value
            value = 1
        Next i
    Next k
    '중심점 추가
    If Frm_CreateFD.ComboBox4.value >= 1 Then
        For j = 1 To nfact
            For i = 1 To Frm_CreateFD.ComboBox4.value * Frm_CreateFD.ComboBox3.value
                Selection.Offset(Frm_CreateFD.ComboBox2.value * (a1 / 2) + i, j) = 0
            Next i
        Next j
    End If
    '블록 설정
    block a1, 2
    
End Sub

'1/4요인설계
Sub quarterFD(a1, nfact, wsheet)
    '요인
    For k = 1 To Frm_CreateFD.ComboBox2.value
        For j = 1 To (nfact - 2)
            For i = 0 To (a1 / 4 - 1)
                If i Mod (2 ^ (nfact - 2)) / (2 ^ (j - 1)) < (2 ^ (nfact - 2)) / (2 * (2 ^ (j - 1))) Then
                    Selection.Offset((a1 / 4) * (k - 1) + (i + 1), j) = 1
                Else
                    Selection.Offset((a1 / 4) * (k - 1) + (i + 1), j) = -1
                End If
            Next i
        Next j

        value = 1
        For i = 1 To (a1 / 4) * k
            For j = 1 To (nfact - 3)
                value = value * Selection.Offset(i, j)
            Next j
            Selection.Offset(i, nfact - 1) = value
            
            value = 1
            If Frm_CreateFD.TextBox1.value = 5 Then
                value = Selection.Offset(i, 1) * Selection.Offset(i, 3)
            End If
            If Frm_CreateFD.TextBox1.value = 6 Then
                value = Selection.Offset(i, 2) * Selection.Offset(i, 3) * Selection.Offset(i, 4)
            End If
            If Frm_CreateFD.TextBox1.value = 7 Then
                value = Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 4) * Selection.Offset(i, 5)
            End If
            Selection.Offset(i, nfact) = value
            value = 1
        Next i
    Next k
    '중심점 추가
    If Frm_CreateFD.ComboBox4.value >= 1 Then
        For j = 1 To nfact
            For i = 1 To Frm_CreateFD.ComboBox4.value * Frm_CreateFD.ComboBox3.value
                Selection.Offset(Frm_CreateFD.ComboBox2.value * (a1 / 4) + i, j) = 0
            Next i
        Next j
    End If
    '블록 설정
    block a1, 4

End Sub

'블록 설정
Sub block(n, key)
    '반복우선
    If Frm_CreateFD.ComboBox3.value = 1 Then    '블록수 1
        For i = 1 To (n / key) * Frm_CreateFD.ComboBox2.value
            Selection.Offset(i, 0) = 1
        Next i
        Selection.Offset(i, 0).Select
        
    Else
        '블록수=반복수인 경우
        If Frm_CreateFD.ComboBox3.value = Frm_CreateFD.ComboBox2 Then
            For i = 1 To Frm_CreateFD.ComboBox2.value '반복수
                For j = 1 To (n / key)
                    Selection.Offset((n / key) * (i - 1) + j, 0) = i
                Next j
            Next i
            Selection.Offset((n / key) * (i - 2) + j, 0).Select
        End If
        '블록수 2, 반복수 4인 경우
        If Frm_CreateFD.ComboBox3.value = 2 And Frm_CreateFD.ComboBox2 = 4 Then
            For i = 1 To 2
                For j = 1 To 2 * (n / key)
                    Selection.Offset(2 * (n / key) * (i - 1) + j, 0) = i
                Next j
            Next i
            Selection.Offset(2 * (n / key) * (i - 2) + j, 0).Select
        End If
    End If
    
    
    
    
    
    
    
    
    '블록우선:블록수 2,4
    If (Frm_CreateFD.ComboBox2.value = 1 And Frm_CreateFD.ComboBox3.value = 2) Or _
       (Frm_CreateFD.ComboBox2.value = 2 And Frm_CreateFD.ComboBox3.value = 4) Or _
       (Frm_CreateFD.ComboBox2.value = 3 And Frm_CreateFD.ComboBox3.value = 2) Or _
       (Frm_CreateFD.ComboBox2.value = 5 And Frm_CreateFD.ComboBox3.value = 2) Then
        '요인수 2, 완전
        If Frm_CreateFD.TextBox1.value = 2 Then
            If Frm_CreateFD.ComboBox2.value = 2 Then
                For i = 1 To (n / key)
                    If Selection.Offset(i, 1) * Selection.Offset(i, 2) = -1 Then
                        Selection.Offset(i, 0) = 1
                    Else
                        Selection.Offset(i, 0) = 2
                    End If
                Next i
                For i = (n / key) + 1 To 2 * (n / key)
                    If Selection.Offset(i, 1) * Selection.Offset(i, 2) = -1 Then
                        Selection.Offset(i, 0) = 3
                    Else
                        Selection.Offset(i, 0) = 4
                    End If
                Next i
            Else
                For k = 1 To Frm_CreateFD.ComboBox2.value
                    For i = 1 To (n / key)
                        If Selection.Offset(i, 1) * Selection.Offset(i, 2) = -1 Then
                            Selection.Offset((n / key) * (k - 1) + i, 0) = 1
                        Else
                            Selection.Offset((n / key) * (k - 1) + i, 0) = 2
                        End If
                    Next i
                Next k
            End If
        End If
        '요인수 3, 완전
        If (Frm_CreateFD.TextBox1.value = 3) And (Frm_CreateFD.ListBox1.Selected(0) = True) Then
            If Frm_CreateFD.ComboBox2.value = 2 Then
                For i = 1 To (n / key)
                    If Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 3) = -1 Then
                        Selection.Offset(i, 0) = 1
                    Else
                        Selection.Offset(i, 0) = 2
                    End If
                Next i
                For i = (n / key) + 1 To 2 * (n / key)
                    If Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(k, 3) = -1 Then
                        Selection.Offset(i, 0) = 3
                    Else
                        Selection.Offset(i, 0) = 4
                    End If
                Next i
            Else
                For k = 1 To Frm_CreateFD.ComboBox2.value
                    For i = 1 To (n / key)
                        If Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 3) = -1 Then
                            Selection.Offset((n / key) * (k - 1) + i, 0) = 1
                        Else
                            Selection.Offset((n / key) * (k - 1) + i, 0) = 2
                        End If
                    Next i
                Next k
            End If
        End If
        '요인수 4, 완전
        If (Frm_CreateFD.TextBox1.value = 4) And (Frm_CreateFD.ListBox1.Selected(0) = True) Then
            If Frm_CreateFD.ComboBox2.value = 2 Then
                For i = 1 To (n / key)
                    If Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 3) * Selection.Offset(i, 4) = -1 Then
                        Selection.Offset(i, 0) = 1
                    Else
                        Selection.Offset(i, 0) = 2
                    End If
                Next i
                For i = (n / key) + 1 To 2 * (n / key)
                    If Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 3) * Selection.Offset(i, 4) = -1 Then
                        Selection.Offset(i, 0) = 3
                    Else
                        Selection.Offset(i, 0) = 4
                    End If
                Next i
            Else
                For k = 1 To Frm_CreateFD.ComboBox2.value
                    For i = 1 To (n / key)
                        If Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 3) * Selection.Offset(i, 4) = -1 Then
                            Selection.Offset((n / key) * (k - 1) + i, 0) = 1
                        Else
                            Selection.Offset((n / key) * (k - 1) + i, 0) = 2
                        End If
                    Next i
                Next k
            End If
        End If
        '요인수 4, 1/2
        If (Frm_CreateFD.TextBox1.value = 4) And (Frm_CreateFD.ListBox1.Selected(1) = True) Then
            If Frm_CreateFD.ComboBox2.value = 2 Then
                For i = 1 To (n / key)
                    If Selection.Offset(i, 1) * Selection.Offset(i, 2) = -1 Then
                        Selection.Offset(i, 0) = 1
                    Else
                        Selection.Offset(i, 0) = 2
                    End If
                Next i
                For i = (n / key) + 1 To 2 * (n / key)
                    If Selection.Offset(i, 1) * Selection.Offset(i, 2) = -1 Then
                        Selection.Offset(i, 0) = 3
                    Else
                        Selection.Offset(i, 0) = 4
                    End If
                Next i
            Else
                For k = 1 To Frm_CreateFD.ComboBox2.value
                    For i = 1 To (n / key)
                        If Selection.Offset(i, 1) * Selection.Offset(i, 2) = -1 Then
                            Selection.Offset((n / key) * (k - 1) + i, 0) = 1
                        Else
                            Selection.Offset((n / key) * (k - 1) + i, 0) = 2
                        End If
                    Next i
                Next k
            End If
        End If
        '요인수 5, 완전
        If (Frm_CreateFD.TextBox1.value = 5) And (Frm_CreateFD.ListBox1.Selected(0) = True) Then
            If Frm_CreateFD.ComboBox2.value = 2 Then
                For i = 1 To (n / key)
                    If Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 3) * Selection.Offset(i, 4) * Selection.Offset(i, 5) = -1 Then
                        Selection.Offset(i, 0) = 1
                    Else
                        Selection.Offset(i, 0) = 2
                    End If
                Next i
                For i = (n / key) + 1 To 2 * (n / key)
                    If Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 3) * Selection.Offset(i, 4) * Selection.Offset(i, 5) = -1 Then
                        Selection.Offset(i, 0) = 3
                    Else
                        Selection.Offset(i, 0) = 4
                    End If
                Next i
            Else
                For k = 1 To Frm_CreateFD.ComboBox2.value
                    For i = 1 To (n / key)
                        If Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 3) * Selection.Offset(i, 4) * Selection.Offset(i, 5) = -1 Then
                            Selection.Offset((n / key) * (k - 1) + i, 0) = 1
                        Else
                            Selection.Offset((n / key) * (k - 1) + i, 0) = 2
                        End If
                    Next i
                Next k
            End If
        End If
        '요인수 5, 1/2
        If (Frm_CreateFD.TextBox1.value = 5) And (Frm_CreateFD.ListBox1.Selected(1) = True) Then
            If Frm_CreateFD.ComboBox2.value = 2 Then
                For i = 1 To (n / key)
                    If Selection.Offset(i, 1) * Selection.Offset(i, 2) = -1 Then
                        Selection.Offset(i, 0) = 1
                    Else
                        Selection.Offset(i, 0) = 2
                    End If
                Next i
                For i = (n / key) + 1 To 2 * (n / key)
                    If Selection.Offset(i, 1) * Selection.Offset(i, 2) = -1 Then
                        Selection.Offset(i, 0) = 3
                    Else
                        Selection.Offset(i, 0) = 4
                    End If
                Next i
            Else
                For k = 1 To Frm_CreateFD.ComboBox2.value
                    For i = 1 To (n / key)
                        If Selection.Offset(i, 1) * Selection.Offset(i, 2) = -1 Then
                            Selection.Offset((n / key) * (k - 1) + i, 0) = 1
                        Else
                            Selection.Offset((n / key) * (k - 1) + i, 0) = 2
                        End If
                    Next i
                Next k
            End If
        End If
        '요인수 5, 1/4
        If (Frm_CreateFD.TextBox1.value = 5) And (Frm_CreateFD.ListBox1.Selected(2) = True) Then
            If Frm_CreateFD.ComboBox2.value = 2 Then
                For i = 1 To (n / key)
                    If Selection.Offset(i, 2) * Selection.Offset(i, 3) = -1 Then
                        Selection.Offset(i, 0) = 1
                    Else
                        Selection.Offset(i, 0) = 2
                    End If
                Next i
                For i = (n / key) + 1 To 2 * (n / key)
                    If Selection.Offset(i, 2) * Selection.Offset(i, 3) = -1 Then
                        Selection.Offset(i, 0) = 3
                    Else
                        Selection.Offset(i, 0) = 4
                    End If
                Next i
            Else
                For k = 1 To Frm_CreateFD.ComboBox2.value
                    For i = 1 To (n / key)
                        If Selection.Offset(i, 2) * Selection.Offset(i, 3) = -1 Then
                            Selection.Offset((n / key) * (k - 1) + i, 0) = 1
                        Else
                            Selection.Offset((n / key) * (k - 1) + i, 0) = 2
                        End If
                    Next i
                Next k
            End If
        End If
        '요인수 6, 완전
        If (Frm_CreateFD.TextBox1.value = 6) And (Frm_CreateFD.ListBox1.Selected(0) = True) Then
            If Frm_CreateFD.ComboBox2.value = 2 Then
                For i = 1 To (n / key)
                    If Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 3) * Selection.Offset(i, 4) * Selection.Offset(i, 5) * Selection.Offset(i, 6) = -1 Then
                        Selection.Offset(i, 0) = 1
                    Else
                        Selection.Offset(i, 0) = 2
                    End If
                Next i
                For i = (n / key) + 1 To 2 * (n / key)
                    If Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(k, 3) * Selection.Offset(i, 4) * Selection.Offset(i, 5) * Selection.Offset(i, 6) = -1 Then
                        Selection.Offset(i, 0) = 3
                    Else
                        Selection.Offset(i, 0) = 4
                    End If
                Next i
            Else
                For k = 1 To Frm_CreateFD.ComboBox2.value
                    For i = 1 To (n / key)
                        If Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 3) * Selection.Offset(i, 4) * Selection.Offset(i, 5) * Selection.Offset(i, 6) = -1 Then
                            Selection.Offset((n / key) * (k - 1) + i, 0) = 1
                        Else
                            Selection.Offset((n / key) * (k - 1) + i, 0) = 2
                        End If
                    Next i
                Next k
            End If
        End If
        '요인수 6, 1/2
        If (Frm_CreateFD.TextBox1.value = 6) And (Frm_CreateFD.ListBox1.Selected(1) = True) Then
            If Frm_CreateFD.ComboBox2.value = 2 Then
                For i = 1 To (n / key)
                    If Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 6) = -1 Then
                        Selection.Offset(i, 0) = 1
                    Else
                        Selection.Offset(i, 0) = 2
                    End If
                Next i
                For i = (n / key) + 1 To 2 * (n / key)
                    If Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 6) = -1 Then
                        Selection.Offset(i, 0) = 3
                    Else
                        Selection.Offset(i, 0) = 4
                    End If
                Next i
            Else
                For k = 1 To Frm_CreateFD.ComboBox2.value
                    For i = 1 To (n / key)
                        If Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 6) = -1 Then
                            Selection.Offset((n / key) * (k - 1) + i, 0) = 1
                        Else
                            Selection.Offset((n / key) * (k - 1) + i, 0) = 2
                        End If
                    Next i
                Next k
            End If
        End If
        '요인수 6, 1/4
        If (Frm_CreateFD.TextBox1.value = 6) And (Frm_CreateFD.ListBox1.Selected(2) = True) Then
            If Frm_CreateFD.ComboBox2.value = 2 Then
                For i = 1 To (n / key)
                    If Selection.Offset(i, 1) * Selection.Offset(i, 3) * Selection.Offset(i, 4) = -1 Then
                        Selection.Offset(i, 0) = 1
                    Else
                        Selection.Offset(i, 0) = 2
                    End If
                Next i
                For i = (n / key) + 1 To 2 * (n / key)
                    If Selection.Offset(i, 1) * Selection.Offset(i, 3) * Selection.Offset(i, 4) = -1 Then
                        Selection.Offset(i, 0) = 3
                    Else
                        Selection.Offset(i, 0) = 4
                    End If
                Next i
            Else
                For k = 1 To Frm_CreateFD.ComboBox2.value
                    For i = 1 To (n / key)
                        If Selection.Offset(i, 1) * Selection.Offset(i, 3) * Selection.Offset(i, 4) = -1 Then
                            Selection.Offset((n / key) * (k - 1) + i, 0) = 1
                        Else
                            Selection.Offset((n / key) * (k - 1) + i, 0) = 2
                        End If
                    Next i
                Next k
            End If
        End If
        '요인수 7, 완전
        If (Frm_CreateFD.TextBox1.value = 7) And (Frm_CreateFD.ListBox1.Selected(0) = True) Then
            If Frm_CreateFD.ComboBox2.value = 2 Then
                For i = 1 To (n / key)
                    If Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 3) * Selection.Offset(i, 4) * Selection.Offset(i, 5) * Selection.Offset(i, 6) * Selection.Offset(i, 7) = -1 Then
                        Selection.Offset(i, 0) = 1
                    Else
                        Selection.Offset(i, 0) = 2
                    End If
                Next i
                For i = (n / key) + 1 To 2 * (n / key)
                    If Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(k, 3) * Selection.Offset(i, 4) * Selection.Offset(i, 5) * Selection.Offset(i, 6) * Selection.Offset(i, 7) = -1 Then
                        Selection.Offset(i, 0) = 3
                    Else
                        Selection.Offset(i, 0) = 4
                    End If
                Next i
            Else
                For k = 1 To Frm_CreateFD.ComboBox2.value
                    For i = 1 To (n / key)
                        If Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 3) * Selection.Offset(i, 4) * Selection.Offset(i, 5) * Selection.Offset(i, 6) * Selection.Offset(i, 7) = -1 Then
                            Selection.Offset((n / key) * (k - 1) + i, 0) = 1
                        Else
                            Selection.Offset((n / key) * (k - 1) + i, 0) = 2
                        End If
                    Next i
                Next k
            End If
        End If
        '요인수 7, 1/2, 1/4
        If (Frm_CreateFD.TextBox1.value = 7) And (Frm_CreateFD.ListBox1.Selected(1) = True Or Frm_CreateFD.ListBox1.Selected(2) = True) Then
            If Frm_CreateFD.ComboBox2.value = 2 Then
                For i = 1 To (n / key)
                    If Selection.Offset(i, 3) * Selection.Offset(i, 4) * Selection.Offset(i, 5) = -1 Then
                        Selection.Offset(i, 0) = 1
                    Else
                        Selection.Offset(i, 0) = 2
                    End If
                Next i
                For i = (n / key) + 1 To 2 * (n / key)
                    If Selection.Offset(i, 3) * Selection.Offset(i, 4) * Selection.Offset(i, 5) = -1 Then
                        Selection.Offset(i, 0) = 3
                    Else
                        Selection.Offset(i, 0) = 4
                    End If
                Next i
            Else
                For k = 1 To Frm_CreateFD.ComboBox2.value
                    For i = 1 To (n / key)
                        If Selection.Offset(i, 3) * Selection.Offset(i, 4) * Selection.Offset(i, 5) = -1 Then
                            Selection.Offset((n / key) * (k - 1) + i, 0) = 1
                        Else
                            Selection.Offset((n / key) * (k - 1) + i, 0) = 2
                        End If
                    Next i
                Next k
            End If
        End If
        Selection.Offset((n / key) * Frm_CreateFD.ComboBox2.value + 1, 0).Select
    End If
    
    '블록우선:블록수 4
    If (Frm_CreateFD.ComboBox2.value = 1 And Frm_CreateFD.ComboBox3.value = 4) Or _
       (Frm_CreateFD.ComboBox2.value = 3 And Frm_CreateFD.ComboBox3.value = 4) Or _
       (Frm_CreateFD.ComboBox2.value = 5 And Frm_CreateFD.ComboBox3.value = 4) Then
        '요인수 3, 완전 or 요인수 4, 1/2 or 요인수 5, 1/2
        If (Frm_CreateFD.TextBox1.value = 3 And Frm_CreateFD.ListBox1.Selected(0) = True) Or (Frm_CreateFD.TextBox1.value = 4 And Frm_CreateFD.ListBox1.Selected(1) = True) Or (Frm_CreateFD.TextBox1.value = 5 And Frm_CreateFD.ListBox1.Selected(1) = True) Then
            For k = 1 To Frm_CreateFD.ComboBox2.value
                For i = 1 To (n / key)
                    If (Selection.Offset(i, 1) * Selection.Offset(i, 2) = -1) And (Selection.Offset(i, 1) * Selection.Offset(i, 3) = -1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 1
                    ElseIf (Selection.Offset(i, 1) * Selection.Offset(i, 2) = -1) And (Selection.Offset(i, 1) * Selection.Offset(i, 3) = 1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 2
                    ElseIf (Selection.Offset(i, 1) * Selection.Offset(i, 2) = 1) And (Selection.Offset(i, 1) * Selection.Offset(i, 3) = -1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 3
                    Else
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 4
                    End If
                Next i
            Next k
        End If
        '요인수 4, 완전
        If (Frm_CreateFD.TextBox1.value = 4 And Frm_CreateFD.ListBox1.Selected(0) = True) Then
            For k = 1 To Frm_CreateFD.ComboBox2.value
                For i = 1 To (n / key)
                    If (Selection.Offset(i, 2) * Selection.Offset(i, 3) = -1) And (Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 4) = -1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 1
                    ElseIf (Selection.Offset(i, 2) * Selection.Offset(i, 3) = -1) And (Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 4) = 1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 2
                    ElseIf (Selection.Offset(i, 2) * Selection.Offset(i, 3) = 1) And (Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 4) = -1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 3
                    Else
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 4
                    End If
                Next i
            Next k
        End If
         '요인수 5, 완전
        If (Frm_CreateFD.TextBox1.value = 5 And Frm_CreateFD.ListBox1.Selected(0) = True) Then
            For k = 1 To Frm_CreateFD.ComboBox2.value
                For i = 1 To (n / key)
                    If (Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 3) = -1) And (Selection.Offset(i, 3) * Selection.Offset(i, 4) * Selection.Offset(i, 5) = -1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 1
                    ElseIf (Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 3) = -1) And (Selection.Offset(i, 3) * Selection.Offset(i, 4) * Selection.Offset(i, 5) = 1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 2
                    ElseIf (Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 3) = 1) And (Selection.Offset(i, 3) * Selection.Offset(i, 4) * Selection.Offset(i, 5) = -1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 3
                    Else
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 4
                    End If
                Next i
            Next k
        End If
        '요인수 6, 완전
        If (Frm_CreateFD.TextBox1.value = 6 And Frm_CreateFD.ListBox1.Selected(0) = True) Then
            For k = 1 To Frm_CreateFD.ComboBox2.value
                For i = 1 To (n / key)
                    If (Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 3) * Selection.Offset(i, 6) = -1) And (Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 4) * Selection.Offset(i, 5) = -1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 1
                    ElseIf (Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 3) * Selection.Offset(i, 6) = -1) And (Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 4) * Selection.Offset(i, 5) = 1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 2
                    ElseIf (Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 3) * Selection.Offset(i, 6) = 1) And (Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 4) * Selection.Offset(i, 5) = -1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 3
                    Else
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 4
                    End If
                Next i
            Next k
        End If
        '요인수 6, 1/2
        If (Frm_CreateFD.TextBox1.value = 6 And Frm_CreateFD.ListBox1.Selected(1) = True) Then
            For k = 1 To Frm_CreateFD.ComboBox2.value
                For i = 1 To (n / key)
                    If (Selection.Offset(i, 2) * Selection.Offset(i, 3) = -1) And (Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 6) = -1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 1
                    ElseIf (Selection.Offset(i, 2) * Selection.Offset(i, 3) = -1) And (Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 6) = 1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 2
                    ElseIf (Selection.Offset(i, 2) * Selection.Offset(i, 3) = 1) And (Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 6) = -1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 3
                    Else
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 4
                    End If
                Next i
            Next k
        End If
        '요인수 6, 1/4
        If (Frm_CreateFD.TextBox1.value = 6 And Frm_CreateFD.ListBox1.Selected(2) = True) Then
            For k = 1 To Frm_CreateFD.ComboBox2.value
                For i = 1 To (n / key)
                    If (Selection.Offset(i, 1) * Selection.Offset(i, 5) = -1) And (Selection.Offset(i, 1) * Selection.Offset(i, 3) * Selection.Offset(i, 4) = -1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 1
                    ElseIf (Selection.Offset(i, 1) * Selection.Offset(i, 5) = -1) And (Selection.Offset(i, 1) * Selection.Offset(i, 3) * Selection.Offset(i, 4) = 1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 2
                    ElseIf (Selection.Offset(i, 1) * Selection.Offset(i, 5) = 1) And (Selection.Offset(i, 1) * Selection.Offset(i, 3) * Selection.Offset(i, 4) = -1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 3
                    Else
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 4
                    End If
                Next i
            Next k
        End If
        '요인수 7, 완전
        If (Frm_CreateFD.TextBox1.value = 7 And Frm_CreateFD.ListBox1.Selected(0) = True) Then
            For k = 1 To Frm_CreateFD.ComboBox2.value
                For i = 1 To (n / key)
                    If (Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 4) * Selection.Offset(i, 5) = -1) And (Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 3) * Selection.Offset(i, 6) * Selection.Offset(i, 7) = -1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 1
                    ElseIf (Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 4) * Selection.Offset(i, 5) = -1) And (Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 3) * Selection.Offset(i, 6) * Selection.Offset(i, 7) = 1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 2
                    ElseIf (Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 4) * Selection.Offset(i, 5) = 1) And (Selection.Offset(i, 1) * Selection.Offset(i, 2) * Selection.Offset(i, 3) * Selection.Offset(i, 6) * Selection.Offset(i, 7) = -1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 3
                    Else
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 4
                    End If
                Next i
            Next k
        End If
        '요인수 7, 1/2
        If (Frm_CreateFD.TextBox1.value = 7 And Frm_CreateFD.ListBox1.Selected(1) = True) Then
            For k = 1 To Frm_CreateFD.ComboBox2.value
                For i = 1 To (n / key)
                    If (Selection.Offset(i, 1) * Selection.Offset(i, 3) * Selection.Offset(i, 6) = -1) And (Selection.Offset(i, 3) * Selection.Offset(i, 4) * Selection.Offset(i, 5) = -1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 1
                    ElseIf (Selection.Offset(i, 1) * Selection.Offset(i, 3) * Selection.Offset(i, 6) = -1) And (Selection.Offset(i, 3) * Selection.Offset(i, 4) * Selection.Offset(i, 5) = 1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 2
                    ElseIf (Selection.Offset(i, 1) * Selection.Offset(i, 3) * Selection.Offset(i, 6) = 1) And (Selection.Offset(i, 3) * Selection.Offset(i, 4) * Selection.Offset(i, 5) = -1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 3
                    Else
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 4
                    End If
                Next i
            Next k
        End If
        '요인수 7, 1/4
        If (Frm_CreateFD.TextBox1.value = 7 And Frm_CreateFD.ListBox1.Selected(2) = True) Then
            For k = 1 To Frm_CreateFD.ComboBox2.value
                For i = 1 To (n / key)
                    If (Selection.Offset(i, 3) * Selection.Offset(i, 6) = -1) And (Selection.Offset(i, 3) * Selection.Offset(i, 4) * Selection.Offset(i, 5) = -1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 1
                    ElseIf (Selection.Offset(i, 3) * Selection.Offset(i, 6) = -1) And (Selection.Offset(i, 3) * Selection.Offset(i, 4) * Selection.Offset(i, 5) = 1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 2
                    ElseIf (Selection.Offset(i, 3) * Selection.Offset(i, 6) = 1) And (Selection.Offset(i, 3) * Selection.Offset(i, 4) * Selection.Offset(i, 5) = -1) Then
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 3
                    Else
                        Selection.Offset((n / key) * (k - 1) + i, 0) = 4
                    End If
                Next i
            Next k
        End If
        Selection.Offset((n / key) * Frm_CreateFD.ComboBox2.value + 1, 0).Select
    End If
    
    '중심점에 대한 블록 설정
    If Frm_CreateFD.ComboBox4.value > 0 Then
        For i = 1 To Frm_CreateFD.ComboBox3.value
            For j = 1 To Frm_CreateFD.ComboBox4.value
                Selection.Offset((i - 1) * Frm_CreateFD.ComboBox4.value + (j - 1), 0) = i
            Next j
        Next i
    End If

End Sub
