Attribute VB_Name = "Module1"
Option Base 1
Dim p() As Single           '������ ������������'
Dim d() As Single           '������ ��������� ��� ���������'
Dim gs() As Single           '������ ��������� ��� ������'
Dim R() As Single           '������ ��� �����������'
Dim Vid() As Integer          '������ ����� �����'
Dim W() As Single           '������ ����� �����'
Dim s() As Single           '������ ��������� �����������'
Dim grMass() As Single      '������ ��������� ������� - ���������'
Dim G As Single, a As Single  '������������ ��������� �������, ��������������� �����.'
Dim C As Single, CP As Single, ch As Single, E As Single ' ��������� �����, ��������� ��������� �������, ��������� ��������� �������, ������������ �����,'
Dim T As Single, Cl As Single, Tk As Single '������ ������ ������������, ������ ��������� ������, ������ ������ ������'
Dim aJ() As Single '������ ����� ���������� �����'
Dim U() As Single, V() As Single '������� ������� ������ � ������� ���������� �������, �������������� ���� ���� �������'
Dim Par() As Integer    '������ ������� ������������ ����� ��� �������������� ����� ���� �������'
Dim Comp() As Single   '������ �������� ��������� ����� ��� �������������� ����� ���� �������'
Dim Number() As Integer  '������  �������, �������������� ������ ����� � �������'
Dim NP() As Integer '������ ������� ����� � ����� ��� �������������� ����� ���� �������'
Dim UStage() As Single, VStage() As Single '������� ������� ������ � ������� ���������� ������� �� ������� ������'
Dim ParStage() As Integer    '������ ������� ������������ ����� �� ������� ������'
Dim CompStage() As Single   '������ �������� ��������� ����� �� ������� ������'
Dim NPStage() As Integer '������ ������� ����� � ����� �� ������� ������'
Dim InstSheet As Worksheet



Sub Utility1()
      Dim Alfa As Integer, Beta As Integer, Gamma As Integer, Delta As Integer
      Dim Temp As String
      Dim tNapr As Single, U0 As Single, V0 As Single, K0 As Single, tGamma As Single, tGor As Single
      Dim sngGurv As Single '����������� ������ �������'
      Dim TreeStage() As Integer '������ ������� ��������������� ����� �� ������� ������, �������������� ����� ������ ������ � �������� ��� ���� �������   '
      Dim VarNumber() As Integer '������ ������� ���������  '
      Dim MetCurStage() As Integer '������ �����  ����� �������� ������ '
      Dim Shift As Integer, Job As Integer
      Dim UG() As Single, VG() As Single '������� ������� ���������� �� �������� ���������'
      Dim SU() As Single, SV() As Single, maxS() As Single  '������� ��������� �� �������� ���������'
      Dim Result As String, ResultS As String, ResultG As String         ' �������� ��������'
      Dim ResSort As String         '��������� ����������'
      Dim intN As Integer '���������� ����� � ����'
      Dim intU As Integer '����������� ����������  �����  ����� ��������� '
      Dim intNodes As Integer '������ ����������  �����  ����� ��������� '
      Dim UpperLeft1 As String, LowerRight1 As String, UpperLeft2 As String, LowerRight2 As String
      Dim maxU As Single, minU As Single, maxV As Single, minV As Single, maxGur As Single, minGur As Single
      Dim Exec As Single
      Set InstSheet = Application.ActiveSheet
      Data_input.Show       '����� ������� ����� '
      Temp = Cells(1, "J").Value
       Do While Range(Temp).Row < 4     '�������������� ������� �������� ������������ �����'
         MsgBox "����� ������ ������ ����� ������������"
         Data_input.Show
         Temp = Cells(1, "J").Value            '������������ ����� �����'
       Loop
      Temp = Cells(1, "L").Value
       Do While Range(Temp).Row < 4     '�������������� ������� �������� ������������ �����������'
         MsgBox "����� ������ ������ ����������� ������������"
         Data_input.Show
         Temp = Cells(1, "L").Value            '������������ ����� �����������'
       Loop
      If Cells(1, "Q").Value = 1 Then     '�������������� ������'
       Exit Sub
      End If
      
  UpperLeft1 = Cells(1, "J").Value '��� ������ ������ ������ �����
  LowerRight1 = Cells(1, "K").Value '��� ��������� ������ ������ �����
  UpperLeft2 = Cells(1, "M").Value '��� ������ ������ ��������� �����������
  LowerRight2 = Cells(1, "L").Value '��� ��������� ������ ��������� �����������
  
  a = Cells(2, "B").Value '��������������� �����������'
  K0 = Cells(2, "H").Value   '��������� ��������� �� ��� ������'
  C = Cells(2, "J").Value '��������� �����'
  ch = Cells(2, "P").Value '��������� ��������� �������'
  CP = Cells(2, "R").Value '��������� ��������� �������'
  E = Cells(2, "L").Value '���������� ������� ����� � ����� ���'
  G = E * Cells(2, "D").Value '������������ ������� � �����'
  Beta = Range(UpperLeft2, LowerRight2).Rows.Count
  Delta = Range(UpperLeft2, LowerRight2).Columns.Count
    
  ReDim s(Beta, Delta)
      For j = 1 To Delta                          '���������� ������� �����������'
       Alfa = Range(UpperLeft2, LowerRight2).Row
       Gamma = Range(UpperLeft2, LowerRight2).Column + j - 1
        For i = 1 To Beta
         s(i, j) = Cells(Alfa, Gamma).Value
         Alfa = Alfa + 1
        Next i
      Next j
  Alfa = Range(UpperLeft1, LowerRight1).Row
  Gamma = Range(UpperLeft1, LowerRight1).Column
  Beta = Range(UpperLeft1, LowerRight1).Rows.Count '���������� �����'
  
  ReDim p(Beta)
  ReDim d(Beta)
  ReDim gs(Beta)
  ReDim R(Beta)
  ReDim Vid(Beta)
  ReDim W(Beta)
  ReDim aJ(Beta)                                    '������ ������������� �����'
  
  For i = 1 To Beta                               '���������� �������� �����'
   p(i) = Cells(Alfa, Gamma).Value
   d(i) = Cells(Alfa, Gamma + 1).Value
   gs(i) = d(i) - p(i) + 1
   R(i) = Cells(Alfa, Gamma + 2).Value
   Vid(i) = Cells(Alfa, Gamma + 3).Value
   W(i) = Cells(Alfa, Gamma + 4).Value
   If p(i) = 0 Or Vid(i) = 0 Then
     MsgBox "������ ����� ������"
     Exit Sub
   End If
   aJ(i) = 1
   Alfa = Alfa + 1
  Next i
     '��������� ���������'
      T = 0
      Cl = 0
      tNapr = 0
      U0 = 0
      V0 = 0
      
      For i = 1 To Beta              '������ ������������� � ���� 0'
       If d(i) <= 0 Then
        tNapr = tNapr + Napr2(W(i), p(i), d(i))
       Else
       tNapr = tNapr + Napr1(W(i), p(i), d(i))
       End If
       V0 = V0 + p(i) / G
      Next i
       V0 = V0 - tNapr
    
      ReDim Number(Beta)
      ReDim MetCurStage(Beta)
      ReDim UStage(Beta)                 '��������� �������� ������� ������'
      ReDim VStage(Beta)
      ReDim ParStage(Beta)
      ReDim CompStage(Beta)
      ReDim NPStage(Beta)
      For i = 1 To Beta           '����������  ����� ��� ����� ������� ������'
        MetCurStage(i) = 0
      Next i
      
      For i = 1 To Beta      '������ ���������� ������� �� ������ ������'
        Tk = s(Vid(i), K0)
        aJ(i) = 0.5
        VStage(i) = VZak(Beta, i, 0, 0)
        aJ(i) = 1
        UStage(i) = UZak(i, 0, K0)
        
        ParStage(i) = 0
        CompStage(i) = p(i)
        NPStage(i) = i
      Next i
      Number(1) = 0            ' ���� ����� ������'
      intN = Beta                '��������� ��������� ���������� ����� ������������ �� ������ ������'
      intU = 0
       
      For j = Beta To 1 Step -1    '���������� �������������� ����� ������� ������'
        If R(j) > Tk Then        '������� �������� ����������� ������'
         intN = intN - 1         '���������� ������������� ����� ������� ������ '
         MetCurStage(j) = 1     '����� ������������� ��� ������� �������'
        Else
        For i = 1 To Beta
         If R(i) <= Tk Then        '������� �������������� ����������� ������'
          If i <> j Then
           If UStage(j) > UStage(i) And VStage(j) <= VStage(i) And gs(j) > gs(i) Then  '������� �������������'
            intN = intN - 1         '���������� ������������� ����� ������� ������ '
            MetCurStage(j) = 1     '����� ������������� ��� ������� �������'
            i = Beta               '����� �� �����'
           End If
          End If
         End If
        Next i
       End If
      Next j
      
      ReDim TreeStage(intN)
      ReDim U(intN)                 '��������� ��������� ��������'
      ReDim V(intN)
      ReDim Par(intN)
      ReDim Comp(intN)
      ReDim NP(intN)
      n = 0
      For i = 1 To Beta
       If MetCurStage(i) = 0 Then
        n = n + 1
        TreeStage(n) = i           '������ ������������� ����� ������� ������'
       End If
      Next i
      For i = 1 To intN
       U(i) = UStage(TreeStage(i))   '��������� �������������� ����� ������� ������ � ������ ������� �������������� �����'
       V(i) = VStage(TreeStage(i))
       Par(i) = ParStage(TreeStage(i))
       Comp(i) = CompStage(TreeStage(i))
       NP(i) = NPStage(TreeStage(i))
       TreeStage(i) = i           '������  ����� ������ ��������, ��������������� �������������� ����� ������� ������'
      Next i
      intNodes = intN  '������ ���������� �������������� ����� �� ������ �� �������� ������ ������������'
       
   For intY = 2 To Beta      '���� �� �������'
       'MsgBox intY
       intU = intN * (Beta - intY + 1)      '���������� ����������� ����� �� ������ ������'
       Number(intY) = intNodes            '����������� ������ ���������� ���� ����������� ������ � ������ ��������'
       ReDim UStage(intU)                 '��������� ��������� ��������'
       ReDim VStage(intU)
       ReDim ParStage(intU)
       ReDim CompStage(intU)
       ReDim NPStage(intU)
       For j = 1 To intN                  '���� �� ����� ������������ �� ������'
          Cl = Comp(TreeStage(j))        '������ ���������  ������ � ������������ ����'
          For i = 1 To Beta                '��������� ����� ����� �� ������������'
           aJ(i) = 1
          Next i
          k = TreeStage(j)                   '������� ���� ������������'
          aJ(NP(k)) = 0                      '�����  ������ � ���� ������������'
          Do Until Par(k) = 0                '����������� ����� ����� ����������� �����'
            k = Par(k)
            aJ(NP(k)) = 0
          Loop
          n = 1
          For i = 1 To Beta                '���� �� ���� (�������������) �������'
           If aJ(i) > 0 Then   '���������� ������������� ������'
             Pr = 0 '��������� �������� �������'
             Tk = Cl + s(Vid(i), Vid(NP(TreeStage(j)))) '������ ������ ����� ������ � ������ ������� ����������� �� ������� �������� � s �� ������'
             If R(i) > Tk Then
              Pr = R(i) - Tk         '����� �������'
              Tk = R(i)
             End If
             k = (j - 1) * (Beta - intY + 1) + n         '������� ����� ���� �� ������� ������'
             aJ(i) = 0.5          '�������� ��� ����������� ������'
             UStage(k) = U(TreeStage(j)) + UZak(i, NP(TreeStage(j)), 0) + Pr * CP / C
             VStage(k) = V(TreeStage(j)) * Cl / (Tk + p(i)) + VZak(Beta, i, Cl, Tk)
             aJ(i) = 1
             ParStage(k) = TreeStage(j)
             CompStage(k) = Tk + p(i) '������ ��������� ������'
             NPStage(k) = i
             n = n + 1                     '���������� ����� �������������� ���� �� ������� ������'
           End If
          Next i
       Next j
       intN = intU         '���������� ��������� ���������� ������������� ����� �� ����� ������'
       ReDim MetCurStage(intU) '����� �����, �����������  � �������� ������ '
       
      For j = intU To 1 Step -1        '���������� �������������� ����� �� ������� ������'
       MetCurStage(j) = 0
       If R(NPStage(j)) > Tk Then    '���������� ������ � ������� ���������'
         intN = intN - 1         '���������� ������������� ����� �������� ������ '
         MetCurStage(j) = 1     '����� ������������� ��� ������� �������'
       Else
         For i = 1 To intU  '���������� ������������ ����� �� ������� ������'
           If i <> j Then
            If R(NPStage(i)) <= Tk Then '��� ��������� �����'
             If intY < Beta Then '�������� ��� �������, ����� ����������'
              'If (UStage(j) > UStage(i) And VStage(j) <= (1 + CompStage(i) / CompStage(j)) / 2 * VStage(i) Or UStage(j) = UStage(i) And VStage(j) < (1 + CompStage(i) / CompStage(j)) / 2 * VStage(i)) And gs(NPStage(j)) > gs(NPStage(i)) Then
              If (UStage(j) > UStage(i) And VStage(j) <= VStage(i) Or UStage(j) = UStage(i) And VStage(j) <= VStage(i)) And gs(NPStage(j)) > gs(NPStage(i)) Then
                intN = intN - 1
                MetCurStage(j) = 1     '����� ������������� ��� ������� �������'
                i = intU          '����� �� �����'
              End If
             Else                 '�������� ��� ���������� ������'
              If UStage(j) >= UStage(i) And VStage(j) <= VStage(i) Then
                intN = intN - 1
                MetCurStage(j) = 1     '����� ������������� ��� ������� �������'
                i = intU          '����� �� �����'
              End If
             End If
           End If
          End If
         Next i
       End If
      Next j
      If intN = 0 Then
       intN = intU
      End If
      ReDim TreeStage(intN)     '������ �������������� ����� � ������'
      n = 0
      If intN < intU Then
       For i = 1 To intU
         If MetCurStage(i) = 0 Then
          n = n + 1               '�������������� ���� ��  ������� ������'
          TreeStage(n) = i
         End If
        Next i
       Else
        For i = 1 To intU
         n = n + 1               '�������������� ���� ��  ������� ������'
         TreeStage(n) = i
        Next i
       End If
       intNodes = intNodes + n        '������ ���������� ����� �� ���� ������'
       ReDim Preserve U(intNodes)             '�������������� ��������'
       ReDim Preserve V(intNodes)
       ReDim Preserve Par(intNodes)
       ReDim Preserve Comp(intNodes)
       ReDim Preserve NP(intNodes)
       For i = 1 To intN               '������� ������ � ������ ������� ��� �������������� �����'
        U(Number(intY) + i) = UStage(TreeStage(i))
        V(Number(intY) + i) = VStage(TreeStage(i))
        Par(Number(intY) + i) = ParStage(TreeStage(i))
        Comp(Number(intY) + i) = CompStage(TreeStage(i))
        NP(Number(intY) + i) = NPStage(TreeStage(i))
        TreeStage(i) = Number(intY) + i    '������ ������� ����� ������ �������� ��� ������������� ������'
       Next i
    Next intY         '����� ����� �� �������'
   
    Temp = Cells(1, "S").Value
    Alfa = Range(Temp).Row
    Gamma = Range(Temp).Column
    k = Cells(1, "T").Value
    For i = 1 To k + 1
     Cells(Alfa + i - 1, Gamma + 1).Value = " " '������� �� ����������� ������'
    Next i
    ReDim VarNumber(n)   '���������� ��������� ����� ���������� �������������� ����� ���������� ������'
   
    Temp = Cells(1, "N").Value
    Alfa = Range(Temp).Row
    Gamma = Range(Temp).Column
    Cells(1, "S").Value = Cells(1, "N").Value '������ ������ ��������� ����������'
    Cells(1, "T").Value = n
    For i = 1 To n
     Cells(Alfa + i, Gamma).Value = -V(TreeStage(i))  '��� ���������� � ������� �������� V'
     Cells(Alfa + i, Gamma - 1).Value = i
    Next i
    
    Temp = Left(Temp, 1) & CStr(Alfa + n)
    Shift = Alfa + n       '��������� ��������� ����� ������'
    InstSheet.Range(Temp).Sort _
     Key1:=InstSheet.Columns(Gamma)
     Temp = LTrim(Cells(Shift, Gamma).Value)  '�������� �������� � ��������� ������ ����������'
       Do While Len(Temp) = 0
         Shift = Shift - 1                       '����� ��������� ��������� ����������� ������'
         Temp = LTrim(Cells(Shift, Gamma).Value) '�������� �������� � ������ ������� ����������'
       Loop
     For i = 1 To n
       VarNumber(i) = Cells(Shift - n + i, Gamma - 1).Value '����� �������� �� ����������'
       Cells(Shift - n + i, Gamma).Value = " "
       Cells(Shift - n + i, Gamma - 1).Value = " "
     Next i
   
      tGamma = Comp(TreeStage(1))  '������ ��������� ������ ������� ��������'
      Temp = Left(Temp, 1) & CStr(Alfa + Beta)  '��������� ������ ��� ����������'
      For i = 1 To n                       '���� �� ���� ���������'
        k = TreeStage(VarNumber(i))                   '��������� ���� ��������'
        Result = " "                    '�������� ������ ������� ������ � ��������� ����'
        ResSort = " "                   ' ������ ������� ������ � � ����� ������'
        l = Beta
        M = Vid(NP(k))                  '������� ��� ������'
        q = 1                               '���������� ����� ������ ���� � ������'
        Cells(Alfa + l, Gamma).Value = NP(k)
        Do Until Par(k) = 0
            k = Par(k)               '���� �� ������ ������ ������'
            If Vid(NP(k)) = M Then   '�������� ������� ���������� � ����� ������'
             l = l - 1                '����� ������ ������'
             Cells(Alfa + l, Gamma).Value = NP(k)  '���������� ������ �����  ������ ����'
             q = q + 1                        '���������� ����� ������ ���� '
            Else                       '������� �� ������ ������'
             Shift = Alfa + Beta        '��������� ��������� ����� ������'
             For j = 1 To q    '���� �� ������'
                ResSort = ResSort & CStr(Cells(Shift - q + j, Gamma).Value) & "," '���������� ������ ������� ������'
                Cells(Shift - q + j, Gamma).Value = " "   '������� ����� ����������'
             Next j
             Result = ResSort & " " & Result    '������� ������ � �������� ������'
             ResSort = " "                      '���������� ������ �� ��������� �������'
             q = 1
             l = Beta
             M = Vid(NP(k)) '����������� ������ ���� �����'
             Cells(Alfa + l, Gamma).Value = NP(k)
            End If
         Loop
           Shift = Alfa + Beta
           For j = 1 To q                   '������� ������ ������ �����'
             ResSort = ResSort & CStr(Cells(Shift - q + j, Gamma).Value) & ","
             Cells(Shift - q + j, Gamma).Value = " "
           Next j
           
           Result = ResSort & " " & Result        '������������ � ������ ��������'
           Result = Left(Result, Len(Result) - 3)
           Result = "������� " & i & ": " & Result
           Cells(Alfa + i, Gamma + 1).Value = Result
      Next i
      tGamma = Round(tGamma, 1)
      Cells(Alfa, Gamma + 1).Value = "�������������� �������� � ���������� ������� �� ��������� ����������� �����, ������ " & tGamma
      Cells(Alfa + n + 1, Gamma).Value = " "   '��������� ������ ����� �������� ������'
      
      Compute.Show           '����� ����� �������'
      Temp = Cells(1, "R").Value
       Do While Range(Temp).Row <= Alfa + n + 5 '�������������� ������� �������� ������������ �������'
         MsgBox "����� ������ ������ ����������� ������������"
         Compute.Show
         Temp = Cells(1, "R").Value            '������������ ����� ������������� ���������'
       Loop
      tGor = Cells(1, "O").Value  '�������� �������� ������ �������'
      sngGurv = Cells(1, "P").Value '�������� ����������� ������ �������'
      ReDim UG(n)    '������� ����������� �� ������� � �������'
      ReDim VG(n)
      ReDim SU(n)
      ReDim SV(n)
      ReDim maxS(n)
      For i = 1 To n                '���� �� �������������� ���������'
       If tGor < tGamma Then       '������ ������� ���������� ��� �������� ���������, ������� �������������'
        k = TreeStage(i)            '��������� ��������� �� ��������� ���� ������'
        l = TreeStage(i)
        Do Until Comp(k) <= tGor   '���������� ������ ��������� ������ � ����, �������� ��������� ���������'
         k = Par(k)                '���������� ������������� ����'
         If Comp(k) <= tGor Then
           If Comp(k) = tGor Then
            UG(i) = U(k)
            VG(i) = V(k)
           Else
            UG(i) = U(k) + (tGor - Comp(k)) / (Comp(l) - Comp(k)) * (U(l) - U(k)) '������������'
            VG(i) = V(k) + (tGor - Comp(k)) / (Comp(l) - Comp(k)) * (V(l) - V(k))
           End If
         End If
          l = k                '����������� ��������� ���� '
        Loop
       Else
        UG(i) = U(TreeStage(i))
        VG(i) = V(TreeStage(i))
       End If
      Next i
      maxU = Application.Max(UG())
      minU = Application.Min(UG())
      maxV = Application.Max(VG())
      minV = Application.Min(VG())
      If maxU = minU Then            '�������������� �� ������, ����� �� ��������� ��������� ��� �������� ���������'
       l = 1
       M = 1
      Else
       For i = 1 To n                       '���������� ���������� ��������� �� ������ �������'
        SU(i) = (UG(i) - minU) / (maxU - minU)
        SV(i) = 1 - (VG(i) - minV) / (maxV - minV)
        maxS(i) = Application.Max(SU(i), SV(i))
       Next i
       l = 1
       For i = 2 To n                       '����������� �������� � ����������� ����������'
        If maxS(i) < maxS(l) Then
         l = i
        End If
       Next i
       
       For i = 1 To n                       '���������� ������������� ��������� �� ������ �������'
        SU(i) = 1 - (UG(i) - minU) / (maxU - minU)
        SV(i) = (VG(i) - minV) / (maxV - minV)
        maxGur = Application.Max(SU(i), SV(i))
        minGur = Application.Min(SU(i), SV(i))
        maxS(i) = sngGurv * minGur + (1 - sngGurv) * maxGur
       Next i
       M = 1
       For i = 2 To n                       '����������� �������� � ������������ ��������� �������'
        If maxS(i) > maxS(M) Then
         M = i
        End If
       Next i
      End If
       ResultS = " �� ������ ������� - ������� " & l
       ResultG = " �� ������ ������� - ������� " & M
       Temp = Cells(1, "R").Value            '������������ ����� ������������� ���������'
       Alfa = Range(Temp).Row
       Gamma = Range(Temp).Column
       Cells(Alfa, Gamma + 2).Value = "������������� ��������  �� ��������� ���������, ������ " & tGor
       Cells(Alfa + 1, Gamma + 2).Value = ResultS '������ �������� �� ������ �������'
       Cells(Alfa + 2, Gamma + 2).Value = ResultG '������ �������� �� ������ �������'
       Cells(Alfa + 4, Gamma + 3).Value = "������� ������� �� ������� � ������� ���������� ������� �� ��������� ���������, ������ " & tGor
       For i = 1 To n                '���� �� �������� ���������'
         k = TreeStage(i)            '��������� ��������� �� ��������� ���� ��������'
         If tGor < tGamma Then      '������ ������� ���������� ��� �������� ���������, ������� �������������'
           Do Until Comp(k) <= tGor   '���������� ������ ��������� ������ � ����, �������� ��������� ���������'
             k = Par(k)                '���������� ������������� ����'
           Loop
         End If
         U0 = Round(U(k), 3)
         V0 = Round(V(k), 3)
         Result = "������� " & i & ": " & "U = " & U0 & ";  V = " & V0
         Cells(Alfa + 4 + i, Gamma + 2).Value = Result
       Next i
       Cells(Alfa + 5 + n, Gamma + 2).Value = " "
       Shift = Alfa + n + 5
       Plan.Show           '����� ����� �����'
       Temp = Cells(1, "V").Value
      
       Do While Range(Temp).Row <= Shift     '�������������� ������� �������� ������������ �������'
         MsgBox "����� ������ ������ ����������� ������������"
         Plan.Show
         Temp = Cells(1, "V").Value            '������������ ����� ������������� ���������'
       Loop
       If Cells(1, "Q").Value = 1 Then     '�������������� ������'
         Exit Sub
       End If
       Do While Cells(1, "Q").Value = 0
         Temp = Cells(1, "N").Value   '������������ ���������'
         Alfa = Range(Temp).Row
         Gamma = Range(Temp).Column
         k = Cells(1, "U").Value      '����� ��������'
        
         Temp = Cells(Alfa + k, Gamma + 1).Value '������ ������������������ ����� � ��������� ��������'
         Shift = InStr(Temp, ":")
         Temp = Right(Temp, Len(Temp) - Shift) '��������� ���������� ������������������'
         Result = " "
         Exec = 0
         Tk = 0
         Shift = InStr(Temp, ",")
         
         ReDim grMass(Beta*4 + 1, 3)         '����������� ������� ��� ���������� (������ �������: [Betta]*4 x 3)
         Cells(1, "H").Value = UBound(grMass) '���������� � ������ ���������� ����� �� �������
         grMass(1, 1) = 0                  '������ ����� ������� ������'
         grMass(1, 2) = 0
		 grMass(1, 3) = 0
         i = 1
         Do While Shift > 0
           ResSort = Left(Temp, Shift - 1) '��������� ������ ������'
           Job = CInt(ResSort)             '����� ����������� ��������� ������
           Tk = R(Job)                     '������ ���������� ������ ��������� ����� ������� ������� ������ �� ������
           If Exec = 0 Then                '���� ������ ������ ���������� ������ ����� ����
            Exec = Exec + s(Vid(Job), K0)  '������ ���������� ������ ������������ �������� �� ����������� (����� ������� ������� �� ���� ����������� ������ � ���� �������������� ���������
           Else                            '
            Exec = Exec + s(Vid(Job), Vid(l)) '�  ������� ������ ����������� ����� (����� ������� ������� �� ����� ����������� ������ � ���������� ������
           End If
           If Tk > Exec Then                 '���� ������ ���������� ������ � ��������� ������, ������� ���������� ������
            Exec = Tk                        '���������� ������ ������������ �� ������� ���������� ������
           End If
		   
           i = i + 1
           grMass(i, 1) = Round(Exec, 1)   '������ � ������ ������� ������ ������'
		   grMass(i, 2) = Job              '������ ������ ������ � ������
		   grMass(i, 3) = 0
		   i = i + 1
           grMass(i, 1) = Round(Exec, 1)   '������ � ������ ������� ������ ������'  
           grMass(i, 2) = Job              '������ ������ ������ � ������
		   grMass(i, 3) = 1
		   
           Result = Result & Exec & " " & "(" & Job & ") "
        
           
           l = Job                          '��������� ������ ����������� ������
           Exec = Round(Exec + p(Job), 1)   '��������� ������  ������� ���������� ����� ������ �� ������ (��������� ������������ ������� ������)
		  
		   i = i + 1
           grMass(i, 1) = Round(Exec, 1)    '������ � ������ ������� ����� ������'
		   grMass(i, 2) = Job
		   grMass(i, 3) = 1
		   
		   i = i + 1
           grMass(i, 1) = Round(Exec, 1)    '������ � ������ ������� ����� ������'
		   grMass(i, 2) = Job
		   grMass(i, 3) = 0
		   
           Temp = Right(Temp, Len(Temp) - Shift) '��������� ���������� �����  ������������������ �����'
           Shift = InStr(Temp, ",")
         Loop
         Job = CInt(Temp) '����� ��������� �����������  ������
         Exec = Exec + s(Vid(Job), Vid(l)) '������ ��������� ����������� ������
		 
         i = i + 1
         grMass(i, 1) = Round(Exec, 1)    '������ � ������ ������� ������ ��������� ������'
         grMass(i, 2) = Job
		 grMass(i, 3) = 1
          
         Result = Result & Exec & " " & "(" & Job & ") "
         Exec = Round(Exec + p(Job), 1)
		 
		 i = i + 1
         grMass(i, 1) = Round(Exec, 1)    '������ � ������ ������� ����� ��������� ������'
         grMass(i, 2) = Job 
		 grMass(i, 3) = 0
		 
         Result = Result & Exec
         Temp = Cells(1, "V").Value            '������������ ����� ������������� ���������'
         Alfa = Range(Temp).Row
         Gamma = Range(Temp).Column
         Cells(Alfa, Gamma + 1).Value = "����  ��������� �� �������� " & k
         Cells(Alfa + 1, Gamma + 1).Value = Result
         
         Temp = Cells(1, "I").Value            '������������ ����� ����� ��������� �����'
         Alfa = Range(Temp).Row
         Gamma = Range(Temp).Column
         Cells(Alfa, Gamma) = "���������� ������"
         Cells(Alfa, Gamma + 2) = "������ ��� ����������"
         
         Alfa = Alfa + 1
         Cells(Alfa, Gamma).Value = "� ������"
         Cells(Alfa, Gamma + 1).Value = "����� ������ � �����"
         Cells(Alfa, Gamma + 2).Value = "����� ������"
         Cells(Alfa, Gamma + 3).Value = "����� ������"
		 
		 For i = 1 To UBound(grMass) 
		  Cells(Alfa, Gamma).Value = grMass(i, 2) 
		 Next i 
		 
         For i = 1 To UBound(grMass) '������ �������� ��� ���������� ��������� �����'
          Cells(Alfa, Gamma + 2).Value = Round(grMass(i, 1), 1)
		  Cells(Alfa, Gamma + 3).Value = Round(grMass(i, 3), 0)
         Next i
         
         Temp = 0
         For i = 1 To UBound(grMass) '���������� ������� ��������� ���� ����� (������������ �������� �������)
          If Temp < grMass(i, 2) Then
           Temp = grMass(i, 2)
          End If
         Next i
         
         Cells(1, "W").Value = Temp     '������ ������������� �������� �������
         Graph.Show                     '����� ����� � ��������
       Loop
   End Sub


Function Napr1(snW, snP, snD) As Single           '������������� � �������� �������'
Napr1 = snW * snP / G / ((snD - T) / a / G + 1)
End Function

Function Napr2(snW, snP, snD) As Single            '������������� ��� ������� �������'
Napr2 = snW * snP / G * ((T - snD) / a / G + 1)
End Function

Function H1(snW, snPi, snPk, snD, snCl, snTk)        '������ �������'
H1 = a * snW * snPi * Log(((snD - snCl) / a / G + 1) / ((snD - snPk - snTk) / a / G + 1))
End Function
Function H2(snW, snPk, snD, snCl, snTk)         '������ �������'
H2 = a * snW * (snPk - (snD - snTk - snPk + a * G) * Log(((snD - snTk) / a / G + 1) / ((snD - snPk - snTk) / a / G + 1))) _
+ a * snW * snPk * Log(((snD - snCl) / a / G + 1) / ((snD - snTk) / a / G + 1))
End Function
Function H3(snW, snPi, snPk, snD, snCl, snTk)        '������ �������'
H3 = a * snW * snPi / 2 * (((snTk + snPk - snD) / a / G + 1) ^ 2 - 1) _
+ a * snW * snPi * Log(((snD - snCl) / a / G + 1))
End Function
Function H4(snW, snPk, snD, snCl, snTk)         '��������� �������'
H4 = a * snW * (snD - snTk + (snTk + snPk - snD - a * G) * Log((snD - snTk) / a / G + 1)) _
+ a * snW * (snTk + snPk) / 2 * (((snTk + snPk - snD) / a / G + 1) ^ 2 - 1) _
- snW / 2 / G * (1 - snD / a / G) * ((snTk + snPk) ^ 2 - snD ^ 2) _
- snW / 3 / a / G ^ 2 * ((snTk + snPk) ^ 3 - snD ^ 3) _
+ a * snW * snPk * Log(((snD - snCl) / a / G + 1) / ((snD - snTk) / a / G + 1))
End Function
 Function H5(snW, snPi, snPk, snD, snCl, snTk)        '����� �������'
   H5 = a * snW * snPi / 2 * (((snTk + snPk - snD) / a / G + 1) ^ 2 - ((snCl - snD) / a / G + 1) ^ 2)
 End Function
 Function H6(snW, snPk, snD, snCl, snTk)            '������ �������'
 H6 = a * snW * snPk * Log(((snD - snCl) / a / G + 1)) _
 + a * snW * snPk / 2 * (((snTk - snD) / a / G + 1) ^ 2 - 1) _
 + a * snW * (snTk + snPk) / 2 * (((snTk + snPk - snD) / a / G + 1) ^ 2 - ((snTk - snD) / a / G + 1) ^ 2) _
 - snW / 2 / G * (1 - snD / a / G) * ((snTk + snPk) ^ 2 - snTk ^ 2) _
 - snW / 3 / a / G ^ 2 * ((snTk + snPk) ^ 3 - snTk ^ 3)
 End Function
Function H7(snW, snPk, snD, snCl, snTk)               '������� �������'
H7 = a * snW * snPk / 2 * (((snTk - snD) / a / G + 1) ^ 2 - ((snCl - snD) / a / G + 1) ^ 2) _
+ a * snW * (snTk + snPk) / 2 * (((snTk + snPk - snD) / a / G + 1) ^ 2 - ((snTk - snD) / a / G + 1) ^ 2) _
 - snW / 2 / G * (1 - snD / a / G) * ((snTk + snPk) ^ 2 - snTk ^ 2) _
 - snW / 3 / a / G ^ 2 * ((snTk + snPk) ^ 3 - snTk ^ 3)
End Function
Function Vybor(snaJ, snW, snPi, snPk, snD, snCl, snTk)
 Dim snGam As Single
 If snaJ = 0 Then
  Vybor = 0
 Else
   If snaJ = 0.5 Then
     snGam = snPk * (snPk / 2 + snTk - snCl) / G
     If snD - snPk - snTk >= 0 Then
      Vybor = (snGam - H2(snW, snPk, snD, snCl, snTk)) / (snTk + snPk)
     Else
       If snD - snTk >= 0 Then
         Vybor = (snGam - H4(snW, snPk, snD, snCl, snTk)) / (snTk + snPk)
       Else
         If snD - snCl >= 0 Then
          Vybor = (snGam - H6(snW, snPk, snD, snCl, snTk)) / (snTk + snPk)
         Else
          Vybor = (snGam - H7(snW, snPk, snD, snCl, snTk)) / (snTk + snPk)
         End If
       End If
     End If
   Else
     snGam = snPi * (snPk + snTk - snCl) / G
     If snD - snPk - snTk >= 0 Then
      Vybor = (snGam - H1(snW, snPi, snPk, snD, snCl, snTk)) / (snTk + snPk)
     Else
      If snD - snTk >= 0 Then
        Vybor = (snGam - H3(snW, snPi, snPk, snD, snCl, snTk)) / (snTk + snPk)
      Else
        Vybor = (snGam - H5(snW, snPi, snPk, snD, snCl, snTk)) / (snTk + snPk)
      End If
     End If
   End If
 End If
End Function

Function VZak(inBeta, ink, snCl, snTk)
 VZak = 0
 For i = 1 To inBeta
   
   VZak = VZak + Vybor(aJ(i), W(i), p(i), p(ink), d(i), snCl, snTk)
   
  Next i

End Function
Function UZak(ink, inl, inK0)             'ink - ����� ������, inl - ���������� ������, inK0  - ��������� ��������� ���� ������'
  If inK0 > 0 Then
   UZak = ch * s(Vid(ink), inK0) / C    '������������� ������� ������� �����������'
  Else
   UZak = ch * s(Vid(ink), Vid(inl)) / C
 End If
End Function




