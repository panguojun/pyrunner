VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000006&
   Caption         =   "python"
   ClientHeight    =   7935
   ClientLeft      =   8145
   ClientTop       =   2370
   ClientWidth     =   12735
   BeginProperty Font 
      Name            =   "@Fixedsys"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form4"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   7935
   ScaleWidth      =   12735
   Begin VB.CommandButton Command6 
      BackColor       =   &H00008080&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "������"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00008080&
      Caption         =   "Ŀ¼"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00008080&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00008080&
      Caption         =   "�༭"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   7575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000C000&
      Caption         =   "����>"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10560
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      DragIcon        =   "Form1.frx":10CA
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   7215
      Left            =   0
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   720
      Width           =   12615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private objStream As Object
Public lastSaveTime As Date ' ���ڼ�¼��󱣴�ʱ��

Public Sub Command1_Click()
    Dim pythonDir As String
    Dim pythonScriptPath As String
    Dim pythonScriptFile As String
    Dim objStream As Object
    
    pythonDir = Text2.Text
    pythonScriptPath = App.Path & "\script.py" ' ����Ϊscript.py�ļ�
    pythonScriptFile = "script.py"
    
     ' ����ļ����޸�ʱ��
    If FileDateTime(pythonScriptPath) > lastSaveTime Then
        ' ����ļ����޸ģ����¶�ȡ�ļ�����
        ReadScriptFile pythonScriptPath
    Else
        ' ���ı����е����ݱ���ΪUTF-8�����ʽ��Python�ű��ļ�
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Charset = "utf-8" ' ȷ��ʹ��UTF-8����
        objStream.Open
        objStream.WriteText Text1.Text
        objStream.SaveToFile pythonScriptPath, 2 ' 2 ��ʾ adSaveCreateOverWrite
        objStream.Close
    End If
     ' ������󱣴�ʱ��
    lastSaveTime = Now
    
    ' ʹ�����Ž�·�����ļ�����������
    pythonScriptPath = Chr(34) & pythonScriptPath & Chr(34)
    pythonDir = Chr(34) & pythonDir & Chr(34) ' ȷ��pythonDirҲ�����Ű���
    
    ' ������ڲ�ִ��Python�ű�
    Shell "cmd /k cd " & App.Path & " && " & pythonDir & "\python.exe " & pythonScriptPath, vbNormalFocus
End Sub

Public Sub Command2_Click()
    Dim filePath As String
    Dim stream As Object

    ' ���ýű��ļ�·��
    filePath = App.Path & "\script.py"

    ' ���� ADODB.Stream ����
    Set stream = CreateObject("ADODB.Stream")

    ' ������������Ϊ�ı�
    stream.Type = 2 ' adTypeText
    stream.Charset = "UTF-8" ' �����ַ���Ϊ UTF-8

    ' ����
    stream.Open

    ' �� Text1 ������д����
    stream.WriteText Text1.Text

    ' ���������ļ�
    stream.SaveToFile filePath, 2 ' adSaveCreateOverWrite

    ' �ر���
    stream.Close

    ' ʹ�� Notepad++ �򿪽ű��ļ�
    Shell "C:\Program Files\Notepad++\notepad++.exe " & Chr(34) & filePath & Chr(34), vbNormalFocus

    ' �ͷŶ���
    Set stream = Nothing
End Sub

Private Sub Command3_Click()
    ' ��������ʾ��������ָ��·��
    Shell "cmd.exe /K cd """ & Text2.Text & """", vbNormalFocus
End Sub

Private Sub ReadScriptFile(filePath As String)
    Dim objStream As Object
    
    ' ����ļ��Ƿ����
    If Dir(filePath) <> "" Then
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Type = 2 ' ����������Ϊ�ı�
        objStream.Charset = "utf-8" ' ȷ��ʹ��UTF-8����
        objStream.Open
        objStream.LoadFromFile filePath ' �����ļ�����
        Text1.Text = objStream.ReadText ' ��ȡ�ı����ݵ��ı���
        objStream.Close
    Else
        MsgBox "�ļ�δ�ҵ�: " & filePath, vbExclamation
    End If
End Sub

Private Sub Command4_Click()
    Dim currentPath As String
    currentPath = App.Path ' ��ȡ��ǰĿ¼
    Shell "explorer.exe " & currentPath, vbNormalFocus ' ���ļ���Դ������
End Sub

Private Sub Command6_Click()
Form2.Show
End Sub

Public Sub Form_Activate()
    Dim pythonScriptPath As String
    
    pythonScriptPath = App.Path & "\script.py" ' ����Ϊscript.py�ļ�
 ' ����ļ����޸�ʱ��
    If FileDateTime(pythonScriptPath) > lastSaveTime Then
        ' ����ļ����޸ģ����¶�ȡ�ļ�����
        ReadScriptFile pythonScriptPath
    End If
End Sub

Private Sub Form_Load()
    Dim pythonScriptPath As String
    Dim configFilePath As String
    Dim objStream As Object
    Dim fso As Object

    pythonScriptPath = App.Path & "\script.py" ' �ű��ļ�·��
    configFilePath = App.Path & "\config.txt" ' �����ļ�·��

    ' �����ļ�ϵͳ����
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' ���script.py�Ƿ���ڣ�����������򴴽�һ���µĿ��ļ�
    If Not fso.FileExists(pythonScriptPath) Then
        Dim newFile As Object
        Set newFile = fso.CreateTextFile(pythonScriptPath, True) ' �������ļ�
        newFile.Close ' �ر��ļ�
    End If

    ' ��ȡUTF-8�����Python�ű��ļ����ݲ���ʾ���ı�����
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 2 ' adTypeText
    objStream.Charset = "utf-8" ' ȷ��ʹ��UTF-8����
    objStream.Open
    objStream.LoadFromFile pythonScriptPath
    Text1.Text = objStream.ReadText
    objStream.Close

    ' ���config.txt�Ƿ���ڣ��������������ʾ������һ���µ������ļ�
    If Not fso.FileExists(configFilePath) Then
        MsgBox "�����ļ������ڣ�������һ���µ������ļ���", vbInformation, "��ʾ"
        Dim newConfigFile As Object
        Set newConfigFile = fso.CreateTextFile(configFilePath, True) ' �����������ļ�
        newConfigFile.WriteLine "C:\Python3x\" ' ����д��һЩĬ������
        newConfigFile.Close ' �ر��ļ�
    End If

    ' ��ȡ�����ļ����ݲ���ʾ��Text2�ı�����
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 2 ' adTypeText
    objStream.Charset = "utf-8" ' ȷ��ʹ��UTF-8����
    objStream.Open
    objStream.LoadFromFile configFilePath
    Text2.Text = objStream.ReadText ' �������ļ����ݸ�ֵ��Text2
    objStream.Close

    Form2.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Unload Form2
End Sub

Private Sub Text1_Change()
lastSaveTime = Now
End Sub

Private Sub Text1_Click()
Form_Activate
End Sub

Private Sub Text1_DblClick()
    ' ����Ƿ�ס Ctrl ��
    If (GetAsyncKeyState(vbKeyControl) And &H8000) <> 0 Then
        ' ���¼����ı������ݣ�ȷ�������ⲿ�ļ�������
        Dim filePath As String
        Dim objStream As Object
        
        filePath = App.Path & "\script.py" ' ָ��Python�ű��ļ�
        
        ' ����ļ��Ƿ����
        If Dir(filePath) <> "" Then
            Set objStream = CreateObject("ADODB.Stream")
            objStream.Type = 2 ' ����������Ϊ�ı�
            objStream.Charset = "utf-8" ' ȷ��ʹ��UTF-8����
            objStream.Open
            objStream.LoadFromFile filePath ' �����ļ�����
            Text1.Text = objStream.ReadText ' ��ȡ�ı����ݵ��ı���
            objStream.Close
        Else
            MsgBox "�ļ�δ�ҵ�: " & filePath, vbExclamation
        End If
    End If
End Sub

Private Sub Text1_DragDrop(Source As Control, X As Single, Y As Single)
    Dim fileName As String
    Dim fileContent As String
    
    ' �����ק�������Ƿ�����ļ�
    If Data.GetFormat(vbCFFiles) Then
        ' ��ȡ��ק���ļ���
        fileName = Data.Files(1)
        
        ' ��ȡ�ļ�����
        Open fileName For Input As #1
        fileContent = Input(LOF(1), #1)
        Close #1
        
        ' ���ļ�������ʾ��Text1�ؼ���
        Text1.Text = fileContent
    End If
End Sub

Private Sub Text1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
     Effect = vbDropEffectCopy
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    ' ����Ƿ��� Ctrl ���� A ��
    If KeyCode = vbKeyA And (Shift And vbCtrlMask) <> 0 Then
        ' ȫѡ Text1 ������
        Text1.SelStart = 0 ' ����ѡ�е���ʼλ��Ϊ 0
        Text1.SelLength = Len(Text1.Text) ' ����ѡ�еĳ���Ϊ�ı��ĳ���
        KeyCode = 0 ' ��ֹ��һ�����������
    End If
    
    If KeyCode = vbKeyF5 Then
        Command1_Click
    End If
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim fileName As String
    Dim fileContent As String
    
    ' �����ק�������Ƿ�����ļ�
    If Data.GetFormat(vbCFFiles) Then
        ' ��ȡ��ק���ļ���
        fileName = Data.Files(1)
        
        ' ʹ��ADODB.Stream�����ȡ�ļ�����
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Type = 2 ' adTypeText
        objStream.Charset = "utf-8" ' ȷ��ʹ��UTF-8����
        objStream.Open
        objStream.LoadFromFile fileName
        fileContent = objStream.ReadText
        objStream.Close
        
        ' ���ļ�������ʾ��Text1�ؼ���
        Text1.Text = fileContent
    End If
End Sub

Private Sub Text1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Effect = vbDropEffectCopy
End Sub

Private Sub Form_Resize()
    Dim formWidth As Integer
    Dim formHeight As Integer
    
    formWidth = ScaleWidth
    formHeight = ScaleHeight
    If (formWidth > 0) Then
        ' ����Text1��λ�úʹ�С
        Text1.Move 10, 500, formWidth - 20, formHeight - 500
        
        ' ����Text2��λ�úʹ�С
        Text2.Move 600, 60, formWidth - 5500, 300
        
        ' ����Command1��λ�úʹ�С
        Command1.Move formWidth - 2200, 60, 2200, 400
        
        ' ����Command2��λ�úʹ�С
        Command2.Move formWidth - 4600, 60, 600, 400
        
        ' ����Command3��λ�úʹ�С
        Command3.Move formWidth - 4000, 60, 600, 400
        
         ' ����Command4��λ�úʹ�С
        Command4.Move formWidth - 3300, 60, 600, 400
        
    End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    ' ����Ƿ����� Enter ��
    If KeyCode = vbKeyReturn Then
        Dim command As String
        ' ��ȡ Text2 �е��ı�
        Dim folderPath As String
        folderPath = Text2.Text
        
        ' ��������
        command = "cmd.exe /k cd """ & folderPath & """"
        
        ' ִ������
        Shell command, vbNormalFocus
    End If
End Sub
