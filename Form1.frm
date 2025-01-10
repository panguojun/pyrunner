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
      Caption         =   "工具"
      BeginProperty Font 
         Name            =   "新宋体"
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
      Caption         =   "目录"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "命令"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "编辑"
      BeginProperty Font 
         Name            =   "宋体"
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
         Name            =   "宋体"
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
      Caption         =   "运行>"
      BeginProperty Font 
         Name            =   "隶书"
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
         Name            =   "楷体"
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
Public lastSaveTime As Date ' 用于记录最后保存时间

Public Sub Command1_Click()
    Dim pythonDir As String
    Dim pythonScriptPath As String
    Dim pythonScriptFile As String
    Dim objStream As Object
    
    pythonDir = Text2.Text
    pythonScriptPath = App.Path & "\script.py" ' 保存为script.py文件
    pythonScriptFile = "script.py"
    
     ' 检查文件的修改时间
    If FileDateTime(pythonScriptPath) > lastSaveTime Then
        ' 如果文件被修改，重新读取文件内容
        ReadScriptFile pythonScriptPath
    Else
        ' 将文本框中的内容保存为UTF-8编码格式的Python脚本文件
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Charset = "utf-8" ' 确保使用UTF-8编码
        objStream.Open
        objStream.WriteText Text1.Text
        objStream.SaveToFile pythonScriptPath, 2 ' 2 表示 adSaveCreateOverWrite
        objStream.Close
    End If
     ' 更新最后保存时间
    lastSaveTime = Now
    
    ' 使用引号将路径和文件名包裹起来
    pythonScriptPath = Chr(34) & pythonScriptPath & Chr(34)
    pythonDir = Chr(34) & pythonDir & Chr(34) ' 确保pythonDir也被引号包裹
    
    ' 打开命令窗口并执行Python脚本
    Shell "cmd /k cd " & App.Path & " && " & pythonDir & "\python.exe " & pythonScriptPath, vbNormalFocus
End Sub

Public Sub Command2_Click()
    Dim filePath As String
    Dim stream As Object

    ' 设置脚本文件路径
    filePath = App.Path & "\script.py"

    ' 创建 ADODB.Stream 对象
    Set stream = CreateObject("ADODB.Stream")

    ' 设置流的类型为文本
    stream.Type = 2 ' adTypeText
    stream.Charset = "UTF-8" ' 设置字符集为 UTF-8

    ' 打开流
    stream.Open

    ' 将 Text1 的内容写入流
    stream.WriteText Text1.Text

    ' 保存流到文件
    stream.SaveToFile filePath, 2 ' adSaveCreateOverWrite

    ' 关闭流
    stream.Close

    ' 使用 Notepad++ 打开脚本文件
    Shell "C:\Program Files\Notepad++\notepad++.exe " & Chr(34) & filePath & Chr(34), vbNormalFocus

    ' 释放对象
    Set stream = Nothing
End Sub

Private Sub Command3_Click()
    ' 打开命令提示符并进入指定路径
    Shell "cmd.exe /K cd """ & Text2.Text & """", vbNormalFocus
End Sub

Private Sub ReadScriptFile(filePath As String)
    Dim objStream As Object
    
    ' 检查文件是否存在
    If Dir(filePath) <> "" Then
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Type = 2 ' 设置流类型为文本
        objStream.Charset = "utf-8" ' 确保使用UTF-8编码
        objStream.Open
        objStream.LoadFromFile filePath ' 加载文件内容
        Text1.Text = objStream.ReadText ' 读取文本内容到文本框
        objStream.Close
    Else
        MsgBox "文件未找到: " & filePath, vbExclamation
    End If
End Sub

Private Sub Command4_Click()
    Dim currentPath As String
    currentPath = App.Path ' 获取当前目录
    Shell "explorer.exe " & currentPath, vbNormalFocus ' 打开文件资源管理器
End Sub

Private Sub Command6_Click()
Form2.Show
End Sub

Public Sub Form_Activate()
    Dim pythonScriptPath As String
    
    pythonScriptPath = App.Path & "\script.py" ' 保存为script.py文件
 ' 检查文件的修改时间
    If FileDateTime(pythonScriptPath) > lastSaveTime Then
        ' 如果文件被修改，重新读取文件内容
        ReadScriptFile pythonScriptPath
    End If
End Sub

Private Sub Form_Load()
    Dim pythonScriptPath As String
    Dim configFilePath As String
    Dim objStream As Object
    Dim fso As Object

    pythonScriptPath = App.Path & "\script.py" ' 脚本文件路径
    configFilePath = App.Path & "\config.txt" ' 配置文件路径

    ' 创建文件系统对象
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 检查script.py是否存在，如果不存在则创建一个新的空文件
    If Not fso.FileExists(pythonScriptPath) Then
        Dim newFile As Object
        Set newFile = fso.CreateTextFile(pythonScriptPath, True) ' 创建新文件
        newFile.Close ' 关闭文件
    End If

    ' 读取UTF-8编码的Python脚本文件内容并显示在文本框中
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 2 ' adTypeText
    objStream.Charset = "utf-8" ' 确保使用UTF-8编码
    objStream.Open
    objStream.LoadFromFile pythonScriptPath
    Text1.Text = objStream.ReadText
    objStream.Close

    ' 检查config.txt是否存在，如果不存在则提示并创建一个新的配置文件
    If Not fso.FileExists(configFilePath) Then
        MsgBox "配置文件不存在，将创建一个新的配置文件。", vbInformation, "提示"
        Dim newConfigFile As Object
        Set newConfigFile = fso.CreateTextFile(configFilePath, True) ' 创建新配置文件
        newConfigFile.WriteLine "C:\Python3x\" ' 可以写入一些默认内容
        newConfigFile.Close ' 关闭文件
    End If

    ' 读取配置文件内容并显示在Text2文本框中
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 2 ' adTypeText
    objStream.Charset = "utf-8" ' 确保使用UTF-8编码
    objStream.Open
    objStream.LoadFromFile configFilePath
    Text2.Text = objStream.ReadText ' 将配置文件内容赋值给Text2
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
    ' 检查是否按住 Ctrl 键
    If (GetAsyncKeyState(vbKeyControl) And &H8000) <> 0 Then
        ' 重新加载文本框内容，确保更新外部文件的内容
        Dim filePath As String
        Dim objStream As Object
        
        filePath = App.Path & "\script.py" ' 指向Python脚本文件
        
        ' 检查文件是否存在
        If Dir(filePath) <> "" Then
            Set objStream = CreateObject("ADODB.Stream")
            objStream.Type = 2 ' 设置流类型为文本
            objStream.Charset = "utf-8" ' 确保使用UTF-8编码
            objStream.Open
            objStream.LoadFromFile filePath ' 加载文件内容
            Text1.Text = objStream.ReadText ' 读取文本内容到文本框
            objStream.Close
        Else
            MsgBox "文件未找到: " & filePath, vbExclamation
        End If
    End If
End Sub

Private Sub Text1_DragDrop(Source As Control, X As Single, Y As Single)
    Dim fileName As String
    Dim fileContent As String
    
    ' 检查拖拽的数据是否包含文件
    If Data.GetFormat(vbCFFiles) Then
        ' 获取拖拽的文件名
        fileName = Data.Files(1)
        
        ' 读取文件内容
        Open fileName For Input As #1
        fileContent = Input(LOF(1), #1)
        Close #1
        
        ' 将文件内容显示在Text1控件中
        Text1.Text = fileContent
    End If
End Sub

Private Sub Text1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
     Effect = vbDropEffectCopy
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    ' 检查是否按下 Ctrl 键和 A 键
    If KeyCode = vbKeyA And (Shift And vbCtrlMask) <> 0 Then
        ' 全选 Text1 的内容
        Text1.SelStart = 0 ' 设置选中的起始位置为 0
        Text1.SelLength = Len(Text1.Text) ' 设置选中的长度为文本的长度
        KeyCode = 0 ' 阻止进一步处理这个键
    End If
    
    If KeyCode = vbKeyF5 Then
        Command1_Click
    End If
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim fileName As String
    Dim fileContent As String
    
    ' 检查拖拽的数据是否包含文件
    If Data.GetFormat(vbCFFiles) Then
        ' 获取拖拽的文件名
        fileName = Data.Files(1)
        
        ' 使用ADODB.Stream对象读取文件内容
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Type = 2 ' adTypeText
        objStream.Charset = "utf-8" ' 确保使用UTF-8编码
        objStream.Open
        objStream.LoadFromFile fileName
        fileContent = objStream.ReadText
        objStream.Close
        
        ' 将文件内容显示在Text1控件中
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
        ' 设置Text1的位置和大小
        Text1.Move 10, 500, formWidth - 20, formHeight - 500
        
        ' 设置Text2的位置和大小
        Text2.Move 600, 60, formWidth - 5500, 300
        
        ' 设置Command1的位置和大小
        Command1.Move formWidth - 2200, 60, 2200, 400
        
        ' 设置Command2的位置和大小
        Command2.Move formWidth - 4600, 60, 600, 400
        
        ' 设置Command3的位置和大小
        Command3.Move formWidth - 4000, 60, 600, 400
        
         ' 设置Command4的位置和大小
        Command4.Move formWidth - 3300, 60, 600, 400
        
    End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    ' 检查是否按下了 Enter 键
    If KeyCode = vbKeyReturn Then
        Dim command As String
        ' 获取 Text2 中的文本
        Dim folderPath As String
        folderPath = Text2.Text
        
        ' 构建命令
        command = "cmd.exe /k cd """ & folderPath & """"
        
        ' 执行命令
        Shell command, vbNormalFocus
    End If
End Sub
