VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Graph 
   Caption         =   "Построение диаграммы"
   ClientHeight    =   2505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3060
   OleObjectBlob   =   "Graph.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Graph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK1_Click()
 Set InstSheet = Application.ActiveSheet
 InstSheet.Cells(1, "Q").Value = 1
 Graph.Hide
End Sub

Private Sub cmdOK2_Click()
 Dim Diag As Chart                 'Объект "Диаграмма"
 Dim Ser As Series                 'Объект "Ряд"
 Dim Temp, DiagFileName As String  'Переменная ячеек таблицы
 Dim Alfa, Gamma As Integer        'Переменные для столбцов и строк таблицы
 Dim W, H As Single                'Переменные длины и высоты
 
 DiagFileName = "DiagGant.bmp"                     'Имя файла с изображением диаграммы
 Set InstSheet = Application.Sheets(2)             'Установка текущим лист с номером 2 (УБРАТЬ)
 'Set InstSheet = Application.ActiveSheet          'Установка текущим выбранный лист (ВОССТАНОВИТЬ)
 Application.DisplayAlerts = False                 'Отключение уведомлений Excel (нужно при удалении листа)
 On Error Resume Next
  Application.Sheets("Диаграмма Ганта").Delete     'Удаление предыдущего листа с диаграммой
 Set Diag = ActiveWorkbook.Charts.Add(, InstSheet) 'Добавление пустой диаграммы на лист
 
 'Задание входящих данных из листа
 Temp = InstSheet.Range("I1").Value
 Alfa = InstSheet.Range(Temp).Row
 Gamma = InstSheet.Range(Temp).Column
 
 With Diag
  .SetSourceData (InstSheet.Range(InstSheet.Cells(Alfa + 2, Gamma + 2), _
  InstSheet.Cells(Alfa + 88, Gamma + 2)))       'Создание диаграммы и задание области ячеек для диаграммы (1-ый ряд)
  .ChartType = xlXYScatterLinesNoMarkers        'Задание типа диаграммы (линейчатая без маркеров)
  .Name = "Диаграмма Ганта"                     'Название листа с диаграммой
  .HasTitle = True                              'Вывод названия диаграммы
  .ChartTitle.Text = "Диаграмма Ганта"          'Название диаграммы
  .HasLegend = False                            'Скрытие легенды
  .Axes(xlCategory).HasTitle = True             'Вывод названия вертикальной оси
  .Axes(xlCategory).AxisTitle.Text = InstSheet.Cells(Alfa + 1, _
  Gamma).Value                                  'Присвоение названия вертикальной оси
  .Axes(xlValue).HasTitle = True                'Вывод названия горизонтальной оси
  .Axes(xlValue).AxisTitle.Text = InstSheet.Cells(Alfa + 1, _
  Gamma + 1).Value                              'Присвоение названия горизонтальной оси

  
  
  'Настройка Ряда № 1:'
  .SeriesCollection(1).Name = "Начало работы"                'Задание имени 1-го ряда
  .SeriesCollection(1).Border.LineStyle = xlLineStyleNone    'Задание окантовки для 1-го ряда (нет линии)
  .SeriesCollection(1).Interior.Color = 16777215             'Задание цвета 1-го ряда (нет цвета)
  .FullSeriesCollection(1).XValues = _
  InstSheet.Range(InstSheet.Cells(Alfa + 2, Gamma), _
  InstSheet.Cells(Alfa + 88, Gamma))                     'Задание названия элементов 1-го ряда
  .Axes(xlValue).TickLabelPosition = xlHigh              'Смещение линейки оси OX вниз
  
  'Настройка Ряда № 2:'
  '.SeriesCollection.NewSeries     'Создание нового ряда
  '.SeriesCollection(2).Values = InstSheet.Range(InstSheet.Cells(Alfa + 2, Gamma + 1), _
  'InstSheet.Cells(Alfa + 22, Gamma + 1))                 'Задание значений 2-го ряда для диаграммы:
  '.SeriesCollection(2).Name = _
  'InstSheet.Range(InstSheet.Cells(Alfa + 1, Gamma + 1))  'Задание имени 2-го ряда
  '.ChartGroups(1).GapWidth = 20                          'Увеличение размера линий графика
  '.SetElement (msoElementPrimaryValueGridLinesNone)      'Убрать линии по вертикали
  '.SetElement (msoElementPrimaryCategoryGridLinesMajor)  'Добавить линии по горизонтали
  
  
  
  
  .Export Filename:=ActiveWorkbook.Path & "\" & DiagFileName, FilterName:="BMP" 'Экспортирования диаграммы в виде картинки
 End With
 
 Set Diag = Nothing                               'Удаление объекта диаграммы
 Application.Sheets("Диаграмма Ганта").Delete     'Удаление листа с диаграммой
 Application.DisplayAlerts = True                 'Отключение уведомлений Excel


 Image1.Picture = LoadPicture(ActiveWorkbook.Path _
 & "\" & DiagFileName)                            'Загрузка изображения диаграммы на форму
 
 cmdOK1.Top = 50                                  'Задание начальных положений кнопок и надписей
 cmdOK1.Left = 50
 cmdOK2.Top = 50
 cmdOK2.Left = -50
 Label1.Top = 24
 Label1.Left = -50
 Label2.Top = 24
 Label2.Left = 50
 
 Graph.Height = 150                               'Задание начального положения формы и начальные размеры
 Graph.Width = 165
 Graph.Left = 0
 Graph.Top = 0
 H = InstSheet.Cells(1, "H").Value             'Передача переменной количества строк на графике
 W = InstSheet.Cells(1, "W").Value             'Передача переменной максимального значения времени на графике
 
 Image1.Height = H * 23                         'Задание размеров рамки для картинки
 Image1.Width = W * 15
 
 Graph.Height = Graph.Height + Image1.Height      'Смена масштаба формы
 Graph.Width = Graph.Width + Image1.Width - 120

 cmdOK1.Top = cmdOK1.Top + Image1.Height          'Смещение кнопок и текста в соответствии с размерами картинки
 cmdOK1.Left = cmdOK1.Left + Image1.Width / 2
 cmdOK2.Top = cmdOK2.Top + Image1.Height
 cmdOK2.Left = cmdOK2.Left + Image1.Width / 2
 Label1.Top = Label1.Top + Image1.Height
 Label1.Left = Label1.Left + Image1.Width / 2
 Label2.Top = Label2.Top + Image1.Height
 Label2.Left = Label2.Left + Image1.Width / 2
 
 InstSheet.Activate                               'Возвращение на предыдущий активный лист
End Sub

Private Sub Image1_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

End Sub
