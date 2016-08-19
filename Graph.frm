VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Graph 
   Caption         =   "���������� ���������"
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
 Dim Diag As Chart                 '������ "���������"
 Dim Ser As Series                 '������ "���"
 Dim Temp, DiagFileName As String  '���������� ����� �������
 Dim Alfa, Gamma As Integer        '���������� ��� �������� � ����� �������
 Dim W, H As Single                '���������� ����� � ������
 
 DiagFileName = "DiagGant.bmp"                     '��� ����� � ������������ ���������
 Set InstSheet = Application.Sheets(2)             '��������� ������� ���� � ������� 2 (������)
 'Set InstSheet = Application.ActiveSheet          '��������� ������� ��������� ���� (������������)
 Application.DisplayAlerts = False                 '���������� ����������� Excel (����� ��� �������� �����)
 On Error Resume Next
  Application.Sheets("��������� �����").Delete     '�������� ����������� ����� � ����������
 Set Diag = ActiveWorkbook.Charts.Add(, InstSheet) '���������� ������ ��������� �� ����
 
 '������� �������� ������ �� �����
 Temp = InstSheet.Range("I1").Value
 Alfa = InstSheet.Range(Temp).Row
 Gamma = InstSheet.Range(Temp).Column
 
 With Diag
  .SetSourceData (InstSheet.Range(InstSheet.Cells(Alfa + 2, Gamma + 2), _
  InstSheet.Cells(Alfa + 88, Gamma + 2)))       '�������� ��������� � ������� ������� ����� ��� ��������� (1-�� ���)
  .ChartType = xlXYScatterLinesNoMarkers        '������� ���� ��������� (���������� ��� ��������)
  .Name = "��������� �����"                     '�������� ����� � ����������
  .HasTitle = True                              '����� �������� ���������
  .ChartTitle.Text = "��������� �����"          '�������� ���������
  .HasLegend = False                            '������� �������
  .Axes(xlCategory).HasTitle = True             '����� �������� ������������ ���
  .Axes(xlCategory).AxisTitle.Text = InstSheet.Cells(Alfa + 1, _
  Gamma).Value                                  '���������� �������� ������������ ���
  .Axes(xlValue).HasTitle = True                '����� �������� �������������� ���
  .Axes(xlValue).AxisTitle.Text = InstSheet.Cells(Alfa + 1, _
  Gamma + 1).Value                              '���������� �������� �������������� ���

  
  
  '��������� ���� � 1:'
  .SeriesCollection(1).Name = "������ ������"                '������� ����� 1-�� ����
  .SeriesCollection(1).Border.LineStyle = xlLineStyleNone    '������� ��������� ��� 1-�� ���� (��� �����)
  .SeriesCollection(1).Interior.Color = 16777215             '������� ����� 1-�� ���� (��� �����)
  .FullSeriesCollection(1).XValues = _
  InstSheet.Range(InstSheet.Cells(Alfa + 2, Gamma), _
  InstSheet.Cells(Alfa + 88, Gamma))                     '������� �������� ��������� 1-�� ����
  .Axes(xlValue).TickLabelPosition = xlHigh              '�������� ������� ��� OX ����
  
  '��������� ���� � 2:'
  '.SeriesCollection.NewSeries     '�������� ������ ����
  '.SeriesCollection(2).Values = InstSheet.Range(InstSheet.Cells(Alfa + 2, Gamma + 1), _
  'InstSheet.Cells(Alfa + 22, Gamma + 1))                 '������� �������� 2-�� ���� ��� ���������:
  '.SeriesCollection(2).Name = _
  'InstSheet.Range(InstSheet.Cells(Alfa + 1, Gamma + 1))  '������� ����� 2-�� ����
  '.ChartGroups(1).GapWidth = 20                          '���������� ������� ����� �������
  '.SetElement (msoElementPrimaryValueGridLinesNone)      '������ ����� �� ���������
  '.SetElement (msoElementPrimaryCategoryGridLinesMajor)  '�������� ����� �� �����������
  
  
  
  
  .Export Filename:=ActiveWorkbook.Path & "\" & DiagFileName, FilterName:="BMP" '��������������� ��������� � ���� ��������
 End With
 
 Set Diag = Nothing                               '�������� ������� ���������
 Application.Sheets("��������� �����").Delete     '�������� ����� � ����������
 Application.DisplayAlerts = True                 '���������� ����������� Excel


 Image1.Picture = LoadPicture(ActiveWorkbook.Path _
 & "\" & DiagFileName)                            '�������� ����������� ��������� �� �����
 
 cmdOK1.Top = 50                                  '������� ��������� ��������� ������ � ��������
 cmdOK1.Left = 50
 cmdOK2.Top = 50
 cmdOK2.Left = -50
 Label1.Top = 24
 Label1.Left = -50
 Label2.Top = 24
 Label2.Left = 50
 
 Graph.Height = 150                               '������� ���������� ��������� ����� � ��������� �������
 Graph.Width = 165
 Graph.Left = 0
 Graph.Top = 0
 H = InstSheet.Cells(1, "H").Value             '�������� ���������� ���������� ����� �� �������
 W = InstSheet.Cells(1, "W").Value             '�������� ���������� ������������� �������� ������� �� �������
 
 Image1.Height = H * 23                         '������� �������� ����� ��� ��������
 Image1.Width = W * 15
 
 Graph.Height = Graph.Height + Image1.Height      '����� �������� �����
 Graph.Width = Graph.Width + Image1.Width - 120

 cmdOK1.Top = cmdOK1.Top + Image1.Height          '�������� ������ � ������ � ������������ � ��������� ��������
 cmdOK1.Left = cmdOK1.Left + Image1.Width / 2
 cmdOK2.Top = cmdOK2.Top + Image1.Height
 cmdOK2.Left = cmdOK2.Left + Image1.Width / 2
 Label1.Top = Label1.Top + Image1.Height
 Label1.Left = Label1.Left + Image1.Width / 2
 Label2.Top = Label2.Top + Image1.Height
 Label2.Left = Label2.Left + Image1.Width / 2
 
 InstSheet.Activate                               '����������� �� ���������� �������� ����
End Sub

Private Sub Image1_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

End Sub
