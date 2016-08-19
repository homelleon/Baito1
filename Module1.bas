Attribute VB_Name = "Module1"
Option Base 1
Dim p() As Single           'массив трудоемкости'
Dim d() As Single           'массив требуемых дат окончания'
Dim gs() As Single           'массив требуемых дат начала'
Dim R() As Single           'массив дат поступления'
Dim Vid() As Integer          'массив видов работ'
Dim W() As Single           'массив весов работ'
Dim s() As Single           'массив стоимости переналадок'
Dim grMass() As Single      'массив координат графика - добавлено'
Dim G As Single, a As Single  'длительность планового периода, психологический коэфф.'
Dim C As Single, CP As Single, ch As Single, E As Single ' стоимость смены, стоимость нормочаса простоя, стоимость нормочаса наладки, длительность смены,'
Dim T As Single, Cl As Single, Tk As Single 'момент начала планирования, момент окончания работы, момент начала работы'
Dim aJ() As Single 'массив меток выполнения работ'
Dim U() As Single, V() As Single 'массивы функции затрат и функции полезности заказов, недоминируемые узлы всех уровней'
Dim Par() As Integer    'массив номеров родительских узлов для недоминируемых узлов всех уровней'
Dim Comp() As Single   'массив моментов окончания работ для недоминируемых узлов всех уровней'
Dim Number() As Integer  'массив  номеров, предшествующих первым узлам в уровнях'
Dim NP() As Integer 'массив номеров работ в узлах для недоминируемых узлов всех уровней'
Dim UStage() As Single, VStage() As Single 'массивы функции затрат и функции полезности заказов на текущем уровне'
Dim ParStage() As Integer    'массив номеров родительских узлов на текущем уровне'
Dim CompStage() As Single   'массив моментов окончания работ на текущем уровне'
Dim NPStage() As Integer 'массив номеров работ в узлах на текущем уровне'
Dim InstSheet As Worksheet



Sub Utility1()
      Dim Alfa As Integer, Beta As Integer, Gamma As Integer, Delta As Integer
      Dim Temp As String
      Dim tNapr As Single, U0 As Single, V0 As Single, K0 As Single, tGamma As Single, tGor As Single
      Dim sngGurv As Single 'коэффициент метода Гурвица'
      Dim TreeStage() As Integer 'массив номеров разветвляющихся узлов на текущем уровне, недоминируемых узлов одного уровня в массивах для всех уровней   '
      Dim VarNumber() As Integer 'массив номеров вариантов  '
      Dim MetCurStage() As Integer 'массив меток  узлов текущего уровня '
      Dim Shift As Integer, Job As Integer
      Dim UG() As Single, VG() As Single 'массивы функций полезности на заданном горизонте'
      Dim SU() As Single, SV() As Single, maxS() As Single  'массивы сожалений на заданном горизонте'
      Dim Result As String, ResultS As String, ResultG As String         ' выходные варианты'
      Dim ResSort As String         'результат сортировки'
      Dim intN As Integer 'количество веток в узле'
      Dim intU As Integer 'добавляемое количество  узлов  после ветвления '
      Dim intNodes As Integer 'полное количество  узлов  после ветвления '
      Dim UpperLeft1 As String, LowerRight1 As String, UpperLeft2 As String, LowerRight2 As String
      Dim maxU As Single, minU As Single, maxV As Single, minV As Single, maxGur As Single, minGur As Single
      Dim Exec As Single
      Set InstSheet = Application.ActiveSheet
      Data_input.Show       'вызов входной формы '
      Temp = Cells(1, "J").Value
       Do While Range(Temp).Row < 4     'предохранитель слишком близкого расположения работ'
         MsgBox "Номер первой строки работ недостаточен"
         Data_input.Show
         Temp = Cells(1, "J").Value            'расположение строк работ'
       Loop
      Temp = Cells(1, "L").Value
       Do While Range(Temp).Row < 4     'предохранитель слишком близкого расположения переналадок'
         MsgBox "Номер первой строки переналадок недостаточен"
         Data_input.Show
         Temp = Cells(1, "L").Value            'расположение строк переналадок'
       Loop
      If Cells(1, "Q").Value = 1 Then     'предохранитель выхода'
       Exit Sub
      End If
      
  UpperLeft1 = Cells(1, "J").Value 'Имя первой ячейки списка работ
  LowerRight1 = Cells(1, "K").Value 'Имя последней ячейки списка работ
  UpperLeft2 = Cells(1, "M").Value 'Имя первой ячейки стоимости переналадок
  LowerRight2 = Cells(1, "L").Value 'Имя последней ячейки стоимости переналадок
  
  a = Cells(2, "B").Value 'Психологический коэффициент'
  K0 = Cells(2, "H").Value   'Начальная настройка на вид работы'
  C = Cells(2, "J").Value 'Стоимость смены'
  ch = Cells(2, "P").Value 'Стоимость нормочаса наладки'
  CP = Cells(2, "R").Value 'Стоимость нормочаса простоя'
  E = Cells(2, "L").Value 'Количество рабочих часов в одном дне'
  G = E * Cells(2, "D").Value 'Длительность периода в часах'
  Beta = Range(UpperLeft2, LowerRight2).Rows.Count
  Delta = Range(UpperLeft2, LowerRight2).Columns.Count
    
  ReDim s(Beta, Delta)
      For j = 1 To Delta                          'наполнение массива переналадок'
       Alfa = Range(UpperLeft2, LowerRight2).Row
       Gamma = Range(UpperLeft2, LowerRight2).Column + j - 1
        For i = 1 To Beta
         s(i, j) = Cells(Alfa, Gamma).Value
         Alfa = Alfa + 1
        Next i
      Next j
  Alfa = Range(UpperLeft1, LowerRight1).Row
  Gamma = Range(UpperLeft1, LowerRight1).Column
  Beta = Range(UpperLeft1, LowerRight1).Rows.Count 'количество работ'
  
  ReDim p(Beta)
  ReDim d(Beta)
  ReDim gs(Beta)
  ReDim R(Beta)
  ReDim Vid(Beta)
  ReDim W(Beta)
  ReDim aJ(Beta)                                    'массив невыполненных работ'
  
  For i = 1 To Beta                               'наполнение массивов работ'
   p(i) = Cells(Alfa, Gamma).Value
   d(i) = Cells(Alfa, Gamma + 1).Value
   gs(i) = d(i) - p(i) + 1
   R(i) = Cells(Alfa, Gamma + 2).Value
   Vid(i) = Cells(Alfa, Gamma + 3).Value
   W(i) = Cells(Alfa, Gamma + 4).Value
   If p(i) = 0 Or Vid(i) = 0 Then
     MsgBox "Ошибки ввода данных"
     Exit Sub
   End If
   aJ(i) = 1
   Alfa = Alfa + 1
  Next i
     'начальные настройки'
      T = 0
      Cl = 0
      tNapr = 0
      U0 = 0
      V0 = 0
      
      For i = 1 To Beta              'расчет напряженности в узле 0'
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
      ReDim UStage(Beta)                 'установка массивов первого уровня'
      ReDim VStage(Beta)
      ReDim ParStage(Beta)
      ReDim CompStage(Beta)
      ReDim NPStage(Beta)
      For i = 1 To Beta           'подготовка  меток для узлов первого уровня'
        MetCurStage(i) = 0
      Next i
      
      For i = 1 To Beta      'расчет полезности заказов на первом уровне'
        Tk = s(Vid(i), K0)
        aJ(i) = 0.5
        VStage(i) = VZak(Beta, i, 0, 0)
        aJ(i) = 1
        UStage(i) = UZak(i, 0, K0)
        
        ParStage(i) = 0
        CompStage(i) = p(i)
        NPStage(i) = i
      Next i
      Number(1) = 0            ' узел перед первым'
      intN = Beta                'начальное возможное количество узлов разветвления на первом уровне'
      intU = 0
       
      For j = Beta To 1 Step -1    'нахождение недоминируемых узлов первого уровня'
        If R(j) > Tk Then        'условие позднего поступления работы'
         intN = intN - 1         'количество разветвляемых узлов первого уровня '
         MetCurStage(j) = 1     'метка доминирования над текущей работой'
        Else
        For i = 1 To Beta
         If R(i) <= Tk Then        'условие своевременного поступления работы'
          If i <> j Then
           If UStage(j) > UStage(i) And VStage(j) <= VStage(i) And gs(j) > gs(i) Then  'условие доминирования'
            intN = intN - 1         'количество разветвляемых узлов первого уровня '
            MetCurStage(j) = 1     'метка доминирования над текущей работой'
            i = Beta               'выход из цикла'
           End If
          End If
         End If
        Next i
       End If
      Next j
      
      ReDim TreeStage(intN)
      ReDim U(intN)                 'начальная установка массивов'
      ReDim V(intN)
      ReDim Par(intN)
      ReDim Comp(intN)
      ReDim NP(intN)
      n = 0
      For i = 1 To Beta
       If MetCurStage(i) = 0 Then
        n = n + 1
        TreeStage(n) = i           'номера разветвляемых узлов первого уровня'
       End If
      Next i
      For i = 1 To intN
       U(i) = UStage(TreeStage(i))   'включение недоминируемых узлов первого уровня в полные массивы недоминируемых узлов'
       V(i) = VStage(TreeStage(i))
       Par(i) = ParStage(TreeStage(i))
       Comp(i) = CompStage(TreeStage(i))
       NP(i) = NPStage(TreeStage(i))
       TreeStage(i) = i           'номера  узлов полных массивов, соответствующих недоминируемым узлам первого уровня'
      Next i
      intNodes = intN  'полное количество недоминируемых узлов от начала до текущего уровня включительно'
       
   For intY = 2 To Beta      'цикл по уровням'
       'MsgBox intY
       intU = intN * (Beta - intY + 1)      'количество добавляемых узлов на каждом уровне'
       Number(intY) = intNodes            'запоминание номера последнего узла предыдущего уровня в полных массивах'
       ReDim UStage(intU)                 'начальная установка массивов'
       ReDim VStage(intU)
       ReDim ParStage(intU)
       ReDim CompStage(intU)
       ReDim NPStage(intU)
       For j = 1 To intN                  'цикл по узлам разветвления на уровне'
          Cl = Comp(TreeStage(j))        'момент окончания  работы в родительском узле'
          For i = 1 To Beta                'установка меток работ на невыполнение'
           aJ(i) = 1
          Next i
          k = TreeStage(j)                   'текущий узел разветвления'
          aJ(NP(k)) = 0                      'метка  работы в узле разветвления'
          Do Until Par(k) = 0                'простановка меток ранее выполненных работ'
            k = Par(k)
            aJ(NP(k)) = 0
          Loop
          n = 1
          For i = 1 To Beta                'цикл по всем (невыполненным) работам'
           If aJ(i) > 0 Then   'отбираются невыполненные работы'
             Pr = 0 'начальное значение простоя'
             Tk = Cl + s(Vid(i), Vid(NP(TreeStage(j)))) 'момент начала новой работы с учетом времени переналадки со второго элемента в s на первый'
             If R(i) > Tk Then
              Pr = R(i) - Tk         'время простоя'
              Tk = R(i)
             End If
             k = (j - 1) * (Beta - intY + 1) + n         'текущий номер узла на текущем уровне'
             aJ(i) = 0.5          'параметр для выполняемой работы'
             UStage(k) = U(TreeStage(j)) + UZak(i, NP(TreeStage(j)), 0) + Pr * CP / C
             VStage(k) = V(TreeStage(j)) * Cl / (Tk + p(i)) + VZak(Beta, i, Cl, Tk)
             aJ(i) = 1
             ParStage(k) = TreeStage(j)
             CompStage(k) = Tk + p(i) 'момент окончания работы'
             NPStage(k) = i
             n = n + 1                     'порядковый номер разветвляемого узла на текущем уровне'
           End If
          Next i
       Next j
       intN = intU         'наибольшее возможное количество разветвляемых узлов на новом уровне'
       ReDim MetCurStage(intU) 'метки узлов, относящихся  к текущему уровню '
       
      For j = intU To 1 Step -1        'нахождение недоминируемых узлов на текущем уровне'
       MetCurStage(j) = 0
       If R(NPStage(j)) > Tk Then    'отключение работы с поздним прибытием'
         intN = intN - 1         'количество разветвляемых узлов текущего уровня '
         MetCurStage(j) = 1     'метка доминирования над текущей работой'
       Else
         For i = 1 To intU  'количество сравниваемых узлов на текущем уровне'
           If i <> j Then
            If R(NPStage(i)) <= Tk Then 'для прибывших работ'
             If intY < Beta Then 'проверка для уровней, кроме последнего'
              'If (UStage(j) > UStage(i) And VStage(j) <= (1 + CompStage(i) / CompStage(j)) / 2 * VStage(i) Or UStage(j) = UStage(i) And VStage(j) < (1 + CompStage(i) / CompStage(j)) / 2 * VStage(i)) And gs(NPStage(j)) > gs(NPStage(i)) Then
              If (UStage(j) > UStage(i) And VStage(j) <= VStage(i) Or UStage(j) = UStage(i) And VStage(j) <= VStage(i)) And gs(NPStage(j)) > gs(NPStage(i)) Then
                intN = intN - 1
                MetCurStage(j) = 1     'метка доминирования над текущей работой'
                i = intU          'выход из цикла'
              End If
             Else                 'проверка для последнего уровня'
              If UStage(j) >= UStage(i) And VStage(j) <= VStage(i) Then
                intN = intN - 1
                MetCurStage(j) = 1     'метка доминирования над текущей работой'
                i = intU          'выход из цикла'
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
      ReDim TreeStage(intN)     'номера разветвляющися узлов в уровне'
      n = 0
      If intN < intU Then
       For i = 1 To intU
         If MetCurStage(i) = 0 Then
          n = n + 1               'недоминируемые узлы на  текущем уровне'
          TreeStage(n) = i
         End If
        Next i
       Else
        For i = 1 To intU
         n = n + 1               'недоминируемые узлы на  текущем уровне'
         TreeStage(n) = i
        Next i
       End If
       intNodes = intNodes + n        'полное количество узлов во всем дереве'
       ReDim Preserve U(intNodes)             'переназначение массивов'
       ReDim Preserve V(intNodes)
       ReDim Preserve Par(intNodes)
       ReDim Preserve Comp(intNodes)
       ReDim Preserve NP(intNodes)
       For i = 1 To intN               'перенос данных в полные массивы для недоминируемых узлов'
        U(Number(intY) + i) = UStage(TreeStage(i))
        V(Number(intY) + i) = VStage(TreeStage(i))
        Par(Number(intY) + i) = ParStage(TreeStage(i))
        Comp(Number(intY) + i) = CompStage(TreeStage(i))
        NP(Number(intY) + i) = NPStage(TreeStage(i))
        TreeStage(i) = Number(intY) + i    'массив номеров узлов полных массивов для рассчитанного уровня'
       Next i
    Next intY         'конец цикла по уровням'
   
    Temp = Cells(1, "S").Value
    Alfa = Range(Temp).Row
    Gamma = Range(Temp).Column
    k = Cells(1, "T").Value
    For i = 1 To k + 1
     Cells(Alfa + i - 1, Gamma + 1).Value = " " 'очистка от предыдущего текста'
    Next i
    ReDim VarNumber(n)   'количество вариантов равно количеству недоминируемых узлов последнего уровня'
   
    Temp = Cells(1, "N").Value
    Alfa = Range(Temp).Row
    Gamma = Range(Temp).Column
    Cells(1, "S").Value = Cells(1, "N").Value 'запись нового положения результата'
    Cells(1, "T").Value = n
    For i = 1 To n
     Cells(Alfa + i, Gamma).Value = -V(TreeStage(i))  'для сортировки в порядке убывания V'
     Cells(Alfa + i, Gamma - 1).Value = i
    Next i
    
    Temp = Left(Temp, 1) & CStr(Alfa + n)
    Shift = Alfa + n       'начальное положение ячеек записи'
    InstSheet.Range(Temp).Sort _
     Key1:=InstSheet.Columns(Gamma)
     Temp = LTrim(Cells(Shift, Gamma).Value)  'удаление пробелов в последней строке сортировки'
       Do While Len(Temp) = 0
         Shift = Shift - 1                       'поиск положения последней заполненной ячейки'
         Temp = LTrim(Cells(Shift, Gamma).Value) 'удаление пробелов в пустых строках сортировки'
       Loop
     For i = 1 To n
       VarNumber(i) = Cells(Shift - n + i, Gamma - 1).Value 'номер варианта до сортировки'
       Cells(Shift - n + i, Gamma).Value = " "
       Cells(Shift - n + i, Gamma - 1).Value = " "
     Next i
   
      tGamma = Comp(TreeStage(1))  'Момент окончания работы первого варианта'
      Temp = Left(Temp, 1) & CStr(Alfa + Beta)  'начальная ячейка для сортировок'
      For i = 1 To n                       'цикл по всем вариантам'
        k = TreeStage(VarNumber(i))                   'последний узел варианта'
        Result = " "                    'выходная строка номеров работы в последнем узле'
        ResSort = " "                   ' строка номеров работы в в одной группе'
        l = Beta
        M = Vid(NP(k))                  'текущий вид работы'
        q = 1                               'количество работ одного вида в группе'
        Cells(Alfa + l, Gamma).Value = NP(k)
        Do Until Par(k) = 0
            k = Par(k)               'цикл до начала дерева поиска'
            If Vid(NP(k)) = M Then   'проверка условия нахождения в одной группе'
             l = l - 1                'сдвиг ячейки записи'
             Cells(Alfa + l, Gamma).Value = NP(k)  'накопление разных работ  одного вида'
             q = q + 1                        'количество работ одного вида '
            Else                       'переход на другую группу'
             Shift = Alfa + Beta        'начальное положение ячеек записи'
             For j = 1 To q    'цикл по группе'
                ResSort = ResSort & CStr(Cells(Shift - q + j, Gamma).Value) & "," 'наполнение строки номеров группы'
                Cells(Shift - q + j, Gamma).Value = " "   'очистка ячеек сортировки'
             Next j
             Result = ResSort & " " & Result    'перенос группы в выходную строку'
             ResSort = " "                      'подготовка работы со следующей группой'
             q = 1
             l = Beta
             M = Vid(NP(k)) 'запоминание нового вида работ'
             Cells(Alfa + l, Gamma).Value = NP(k)
            End If
         Loop
           Shift = Alfa + Beta
           For j = 1 To q                   'добавка первой группы работ'
             ResSort = ResSort & CStr(Cells(Shift - q + j, Gamma).Value) & ","
             Cells(Shift - q + j, Gamma).Value = " "
           Next j
           
           Result = ResSort & " " & Result        'формирование и запись варианта'
           Result = Left(Result, Len(Result) - 3)
           Result = "Вариант " & i & ": " & Result
           Cells(Alfa + i, Gamma + 1).Value = Result
      Next i
      tGamma = Round(tGamma, 1)
      Cells(Alfa, Gamma + 1).Value = "Недоминируемые варианты в результате расчета на горизонте планируемых работ, равном " & tGamma
      Cells(Alfa + n + 1, Gamma).Value = " "   'обнуление строки после выходных данных'
      
      Compute.Show           'вызов формы решения'
      Temp = Cells(1, "R").Value
       Do While Range(Temp).Row <= Alfa + n + 5 'предохранитель слишком близкого расположения решения'
         MsgBox "Номер первой строки результатов недостаточен"
         Compute.Show
         Temp = Cells(1, "R").Value            'расположение строк рекомендуемых вариантов'
       Loop
      tGor = Cells(1, "O").Value  'Заданный горизонт поиска решения'
      sngGurv = Cells(1, "P").Value 'Заданный коэффициент метода Гурвица'
      ReDim UG(n)    'массивы полезностей по Гурвицу и Сэвиджу'
      ReDim VG(n)
      ReDim SU(n)
      ReDim SV(n)
      ReDim maxS(n)
      For i = 1 To n                'цикл по недоминируемым вариантам'
       If tGor < tGamma Then       'расчет функций полезности при заданном горизонте, меньшем максимального'
        k = TreeStage(i)            'начальные установки на последнем узле дерева'
        l = TreeStage(i)
        Do Until Comp(k) <= tGor   'нахождение момент окончания работы в узле, меньшего заданному горизонту'
         k = Par(k)                'нахождение родительского узла'
         If Comp(k) <= tGor Then
           If Comp(k) = tGor Then
            UG(i) = U(k)
            VG(i) = V(k)
           Else
            UG(i) = U(k) + (tGor - Comp(k)) / (Comp(l) - Comp(k)) * (U(l) - U(k)) 'интерполяция'
            VG(i) = V(k) + (tGor - Comp(k)) / (Comp(l) - Comp(k)) * (V(l) - V(k))
           End If
         End If
          l = k                'запоминание дочернего узла '
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
      If maxU = minU Then            'предохранитель на случай, когда на расчетном горизонте все варианты одинаковы'
       l = 1
       M = 1
      Else
       For i = 1 To n                       'нахождение наибольших сожалений по методу Сэвиджа'
        SU(i) = (UG(i) - minU) / (maxU - minU)
        SV(i) = 1 - (VG(i) - minV) / (maxV - minV)
        maxS(i) = Application.Max(SU(i), SV(i))
       Next i
       l = 1
       For i = 2 To n                       'определение варианта с минимаксным сожалением'
        If maxS(i) < maxS(l) Then
         l = i
        End If
       Next i
       
       For i = 1 To n                       'нахождение характеристик вариантов по методу Гурвица'
        SU(i) = 1 - (UG(i) - minU) / (maxU - minU)
        SV(i) = (VG(i) - minV) / (maxV - minV)
        maxGur = Application.Max(SU(i), SV(i))
        minGur = Application.Min(SU(i), SV(i))
        maxS(i) = sngGurv * minGur + (1 - sngGurv) * maxGur
       Next i
       M = 1
       For i = 2 To n                       'определение варианта с максимальным критерием Гурвица'
        If maxS(i) > maxS(M) Then
         M = i
        End If
       Next i
      End If
       ResultS = " по методу Сэвиджа - вариант " & l
       ResultG = " по методу Гурвица - вариант " & M
       Temp = Cells(1, "R").Value            'расположение строк рекомендуемых вариантов'
       Alfa = Range(Temp).Row
       Gamma = Range(Temp).Column
       Cells(Alfa, Gamma + 2).Value = "Рекомендуемые варианты  на расчетном горизонте, равном " & tGor
       Cells(Alfa + 1, Gamma + 2).Value = ResultS 'запись варианта по методу Сэвиджа'
       Cells(Alfa + 2, Gamma + 2).Value = ResultG 'запись варианта по методу Гурвица'
       Cells(Alfa + 4, Gamma + 3).Value = "Средние затраты на наладку и средняя полезность заказов на расчетном горизонте, равном " & tGor
       For i = 1 To n                'цикл по выходным вариантам'
         k = TreeStage(i)            'начальные установки на последнем узле варианта'
         If tGor < tGamma Then      'расчет функций полезности при заданном горизонте, меньшем максимального'
           Do Until Comp(k) <= tGor   'нахождение момент окончания работы в узле, меньшего заданному горизонту'
             k = Par(k)                'нахождение родительского узла'
           Loop
         End If
         U0 = Round(U(k), 3)
         V0 = Round(V(k), 3)
         Result = "Вариант " & i & ": " & "U = " & U0 & ";  V = " & V0
         Cells(Alfa + 4 + i, Gamma + 2).Value = Result
       Next i
       Cells(Alfa + 5 + n, Gamma + 2).Value = " "
       Shift = Alfa + n + 5
       Plan.Show           'вызов формы плана'
       Temp = Cells(1, "V").Value
      
       Do While Range(Temp).Row <= Shift     'предохранитель слишком близкого расположения решения'
         MsgBox "Номер первой строки результатов недостаточен"
         Plan.Show
         Temp = Cells(1, "V").Value            'расположение строк рекомендуемых вариантов'
       Loop
       If Cells(1, "Q").Value = 1 Then     'предохранитель выхода'
         Exit Sub
       End If
       Do While Cells(1, "Q").Value = 0
         Temp = Cells(1, "N").Value   'расположение вариантов'
         Alfa = Range(Temp).Row
         Gamma = Range(Temp).Column
         k = Cells(1, "U").Value      'номер варианта'
        
         Temp = Cells(Alfa + k, Gamma + 1).Value 'строка последовательности работ в выбранном варианте'
         Shift = InStr(Temp, ":")
         Temp = Right(Temp, Len(Temp) - Shift) 'выделение собственно последовательности'
         Result = " "
         Exec = 0
         Tk = 0
         Shift = InStr(Temp, ",")
         
         ReDim grMass(Beta*4 + 1, 3)         'Определение массива для построения (массив размера: [Betta]*4 x 3)
         Cells(1, "H").Value = UBound(grMass) 'Сохранение в ячейку количества работ на графике
         grMass(1, 1) = 0                  'Момент перед началом работы'
         grMass(1, 2) = 0
		 grMass(1, 3) = 0
         i = 1
         Do While Shift > 0
           ResSort = Left(Temp, Shift - 1) 'выделение номера работы'
           Job = CInt(ResSort)             'номер планируемой следующей работы
           Tk = R(Job)                     'момент возможного начала обработки равен моменту прихода работы на машину
           If Exec = 0 Then                'если момент начала подготовки машины равен нулю
            Exec = Exec + s(Vid(Job), K0)  'момент подготовки машины определяется временем на переналадку (время наладки зависит от вида планируемой работы и вида первоначальной настройки
           Else                            '
            Exec = Exec + s(Vid(Job), Vid(l)) 'в  моменты начала последующих работ (время наладки зависит от видов планируемой работы и предыдущей работы
           End If
           If Tk > Exec Then                 'если момент готовности работы к обработке больше, момента готовности машины
            Exec = Tk                        'готовность машины определяется по моменту готовности работы
           End If
		   
           i = i + 1
           grMass(i, 1) = Round(Exec, 1)   'Запись в массив момента начала работы'
		   grMass(i, 2) = Job              'Запись номера работы в массив
		   grMass(i, 3) = 0
		   i = i + 1
           grMass(i, 1) = Round(Exec, 1)   'Запись в массив момента начала работы'  
           grMass(i, 2) = Job              'Запись номера работы в массив
		   grMass(i, 3) = 1
		   
           Result = Result & Exec & " " & "(" & Job & ") "
        
           
           l = Job                          'установка номера выполненной работы
           Exec = Round(Exec + p(Job), 1)   'возможное начало  момента выполнения новой работы на машине (учитывает длительность текущей работы)
		  
		   i = i + 1
           grMass(i, 1) = Round(Exec, 1)    'Запись в массив момента конца работы'
		   grMass(i, 2) = Job
		   grMass(i, 3) = 1
		   
		   i = i + 1
           grMass(i, 1) = Round(Exec, 1)    'Запись в массив момента конца работы'
		   grMass(i, 2) = Job
		   grMass(i, 3) = 0
		   
           Temp = Right(Temp, Len(Temp) - Shift) 'выделение остающейся части  последовательности работ'
           Shift = InStr(Temp, ",")
         Loop
         Job = CInt(Temp) 'номер последней планируемой  работы
         Exec = Exec + s(Vid(Job), Vid(l)) 'начало последней планируемой работы
		 
         i = i + 1
         grMass(i, 1) = Round(Exec, 1)    'запись в массив момента начала последней работы'
         grMass(i, 2) = Job
		 grMass(i, 3) = 1
          
         Result = Result & Exec & " " & "(" & Job & ") "
         Exec = Round(Exec + p(Job), 1)
		 
		 i = i + 1
         grMass(i, 1) = Round(Exec, 1)    'запись в массив момента конца последней работы'
         grMass(i, 2) = Job 
		 grMass(i, 3) = 0
		 
         Result = Result & Exec
         Temp = Cells(1, "V").Value            'расположение строк рекомендуемых вариантов'
         Alfa = Range(Temp).Row
         Gamma = Range(Temp).Column
         Cells(Alfa, Gamma + 1).Value = "План  обработки по варианту " & k
         Cells(Alfa + 1, Gamma + 1).Value = Result
         
         Temp = Cells(1, "I").Value            'расположение меток точек диаграммы Ганта'
         Alfa = Range(Temp).Row
         Gamma = Range(Temp).Column
         Cells(Alfa, Gamma) = "Справочные данные"
         Cells(Alfa, Gamma + 2) = "Данные для построения"
         
         Alfa = Alfa + 1
         Cells(Alfa, Gamma).Value = "№ работы"
         Cells(Alfa, Gamma + 1).Value = "Время работы в часах"
         Cells(Alfa, Gamma + 2).Value = "Точки работы"
         Cells(Alfa, Gamma + 3).Value = "Точки работы"
		 
		 For i = 1 To UBound(grMass) 
		  Cells(Alfa, Gamma).Value = grMass(i, 2) 
		 Next i 
		 
         For i = 1 To UBound(grMass) 'запись значений для построения диаграммы Ганта'
          Cells(Alfa, Gamma + 2).Value = Round(grMass(i, 1), 1)
		  Cells(Alfa, Gamma + 3).Value = Round(grMass(i, 3), 0)
         Next i
         
         Temp = 0
         For i = 1 To UBound(grMass) 'Нахождение времени окончания всех работ (максимальное значение времени)
          If Temp < grMass(i, 2) Then
           Temp = grMass(i, 2)
          End If
         Next i
         
         Cells(1, "W").Value = Temp     'Запись максимального значения времени
         Graph.Show                     'Вывод формы с графиком
       Loop
   End Sub


Function Napr1(snW, snP, snD) As Single           'напряженность с резервом времени'
Napr1 = snW * snP / G / ((snD - T) / a / G + 1)
End Function

Function Napr2(snW, snP, snD) As Single            'напряженность без резерва времени'
Napr2 = snW * snP / G * ((T - snD) / a / G + 1)
End Function

Function H1(snW, snPi, snPk, snD, snCl, snTk)        'первый вариант'
H1 = a * snW * snPi * Log(((snD - snCl) / a / G + 1) / ((snD - snPk - snTk) / a / G + 1))
End Function
Function H2(snW, snPk, snD, snCl, snTk)         'второй вариант'
H2 = a * snW * (snPk - (snD - snTk - snPk + a * G) * Log(((snD - snTk) / a / G + 1) / ((snD - snPk - snTk) / a / G + 1))) _
+ a * snW * snPk * Log(((snD - snCl) / a / G + 1) / ((snD - snTk) / a / G + 1))
End Function
Function H3(snW, snPi, snPk, snD, snCl, snTk)        'третий вариант'
H3 = a * snW * snPi / 2 * (((snTk + snPk - snD) / a / G + 1) ^ 2 - 1) _
+ a * snW * snPi * Log(((snD - snCl) / a / G + 1))
End Function
Function H4(snW, snPk, snD, snCl, snTk)         'четвертый вариант'
H4 = a * snW * (snD - snTk + (snTk + snPk - snD - a * G) * Log((snD - snTk) / a / G + 1)) _
+ a * snW * (snTk + snPk) / 2 * (((snTk + snPk - snD) / a / G + 1) ^ 2 - 1) _
- snW / 2 / G * (1 - snD / a / G) * ((snTk + snPk) ^ 2 - snD ^ 2) _
- snW / 3 / a / G ^ 2 * ((snTk + snPk) ^ 3 - snD ^ 3) _
+ a * snW * snPk * Log(((snD - snCl) / a / G + 1) / ((snD - snTk) / a / G + 1))
End Function
 Function H5(snW, snPi, snPk, snD, snCl, snTk)        'пятый вариант'
   H5 = a * snW * snPi / 2 * (((snTk + snPk - snD) / a / G + 1) ^ 2 - ((snCl - snD) / a / G + 1) ^ 2)
 End Function
 Function H6(snW, snPk, snD, snCl, snTk)            'шестой вариант'
 H6 = a * snW * snPk * Log(((snD - snCl) / a / G + 1)) _
 + a * snW * snPk / 2 * (((snTk - snD) / a / G + 1) ^ 2 - 1) _
 + a * snW * (snTk + snPk) / 2 * (((snTk + snPk - snD) / a / G + 1) ^ 2 - ((snTk - snD) / a / G + 1) ^ 2) _
 - snW / 2 / G * (1 - snD / a / G) * ((snTk + snPk) ^ 2 - snTk ^ 2) _
 - snW / 3 / a / G ^ 2 * ((snTk + snPk) ^ 3 - snTk ^ 3)
 End Function
Function H7(snW, snPk, snD, snCl, snTk)               'седьмой вариант'
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
Function UZak(ink, inl, inK0)             'ink - новая работа, inl - предыдущая работа, inK0  - начальная установка вида работы'
  If inK0 > 0 Then
   UZak = ch * s(Vid(ink), inK0) / C    'относительные затраты времени переналадки'
  Else
   UZak = ch * s(Vid(ink), Vid(inl)) / C
 End If
End Function




