VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000006&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "py lib"
   ClientHeight    =   7950
   ClientLeft      =   2940
   ClientTop       =   2325
   ClientWidth     =   5250
   BeginProperty Font 
      Name            =   "����"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   5
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "View"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00008000&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "@Fixedsys"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "del"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   2
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "save"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   2520
      Width           =   495
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   7740
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Form1.Visible = False Then
        Form1.Form_Activate
    End If
    
    Dim fileName As String
    Dim userInput As String
    Dim objStream As Object
    Dim folderPath As String
    
    ' ��ȡ��ǰ���ںͺ�����
    Dim currentDate As String
    currentDate = Format(Now, "mm-dd-(hhmmss)")
    
    ' ������������û������ļ���
    If List1.ListIndex = -1 Then
        userInput = InputBox("�������ļ�����Ĭ��Ϊ��ǰ���ڼ���������:", "�����ļ�", currentDate)
    Else
        userInput = InputBox("�������ļ�����Ĭ��Ϊ��ǰ���ڼ���������:", "�����ļ�", List1.List(List1.ListIndex))
    End If
    
    ' ����ļ����Ƿ��� .py ��β������������Զ����
    If Right(userInput, 3) <> ".py" Then
        userInput = userInput & ".py"
    End If
    
    ' �����û�������ļ��������ļ�
    If userInput <> "" Then
        On Error GoTo ErrorHandler ' ���ô�����
        
        ' �����ļ���·��
        folderPath = App.Path & "\saved\"
        
        ' ����ļ����Ƿ���ڣ�����������򴴽�
        If Dir(folderPath, vbDirectory) = "" Then
            MkDir folderPath
        End If
        
        fileName = folderPath & userInput
        
        ' ʹ��ADODB.Stream��֧��UTF-8����
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Type = 2 ' adTypeText
        objStream.Charset = "utf-8" ' �����ַ���ΪUTF-8
        objStream.Open
        
        ' д������
        objStream.WriteText Form1.Text1.Text
        
        ' �����ļ�
        objStream.SaveToFile fileName, 2
        objStream.Close
    End If
    
    RefreshList
    Exit Sub ' ȷ�������˳����������������

ErrorHandler:
    MsgBox "��������: " & Err.Description, vbExclamation, "����"
    If Not objStream Is Nothing Then
        objStream.Close
    End If
End Sub

Public Sub RefreshList()
    Dim fileName As String
    Dim fileNames() As String
    Dim i As Integer
    
    ' ���List1�ؼ�
    List1.Clear
    
    ' ��ȡsaved�ļ����µ������ļ���
    fileName = Dir(App.Path & "\saved\*.py")
    If fileName = "" Then
        Exit Sub
    End If
    
    Do While fileName <> ""
        ReDim Preserve fileNames(i)
        fileNames(i) = fileName
        i = i + 1
        fileName = Dir
    Loop
    
    ' ���ļ�����ӵ�List1�ؼ�
    For i = LBound(fileNames) To UBound(fileNames)
        List1.AddItem fileNames(i)
    Next i
End Sub

Private Sub Command2_Click()
Dim fileName As String

If List1.ListIndex <> -1 Then
    fileName = App.Path & "\saved\" & List1.List(List1.ListIndex)
    
    ' ɾ���ļ�
    Kill fileName
    
    ' ˢ���б�
    RefreshList
End If
End Sub

Private Sub Command3_Click()
'Form1.Show
If Form1.Visible = False Then
List1_DblClick
End If

Form1.Command1_Click
Form1.Hide
End Sub

Private Sub Command4_Click()
Form1.Show
End Sub

Private Sub Command5_Click()
Form1.Command2_Click
End Sub

Private Sub Form_Load()
    ' ���ش���ʱˢ��List1�ؼ�
    RefreshList
    Form1.Show
End Sub
Private Sub List1_DblClick()
    Dim fileName As String
    Dim fileContent As String
    
    If List1.ListIndex <> -1 Then
        fileName = App.Path & "\saved\" & List1.List(List1.ListIndex)
        
        Dim objStream As Object
        
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Type = 2 ' adTypeText
        objStream.Charset = "utf-8"
        objStream.Open
        objStream.LoadFromFile fileName
        fileContent = objStream.ReadText
        objStream.Close
        
        ' �滻Form1.Text1������
        Form1.Text1.Text = fileContent
        Form1.lastSaveTime = Now
    End If
End Sub
