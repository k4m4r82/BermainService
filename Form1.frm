VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Bermain-main dengan windows service"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2640
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4455
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   7858
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdRefreshService 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   3405
      TabIndex        =   3
      Top             =   4695
      Width           =   975
   End
   Begin VB.CommandButton cmdDeleteService 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2310
      TabIndex        =   2
      Top             =   4695
      Width           =   975
   End
   Begin VB.CommandButton cmdStopService 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1215
      TabIndex        =   1
      Top             =   4695
      Width           =   975
   End
   Begin VB.CommandButton cmdStartService 
      Caption         =   "Start"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   4695
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
' MMMM  MMMMM  OMMM   MMMO    OMMM    OMMM    OMMMMO     OMMMMO    OMMMMO  '
'  MM    MM   MM MM    MMMO  OMMM    MM MM    MM   MO   OM    MO  OM    MO '
'  MM  MM    MM  MM    MM  OO  MM   MM  MM    MM   MO   OM    MO       OMO '
'  MMMM     MMMMMMMM   MM  MM  MM  MMMMMMMM   MMMMMO     OMMMMO      OMO   '
'  MM  MM        MM    MM      MM       MM    MM   MO   OM    MO   OMO     '
'  MM    MM      MM    MM      MM       MM    MM    MO  OM    MO  OM   MM  '
' MMMM  MMMM    MMMM  MMMM    MMMM     MMMM  MMMM  MMMM  OMMMMO   MMMMMMM  '
'                                                                          '
' K4m4r82's Laboratory                                                     '
' http://coding4ever.wordpress.com                                         '
'***************************************************************************

Option Explicit

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const SYNCHRONIZE       As Long = &H100000
Private Const INFINITE          As Long = &HFFFF

Private Const REC_SPR           As String = "|" 'separator baris
Private Const COL_SPR           As String = "#"  'separator kolom
Private Const FILE_SERVICE      As String = "c:\services.txt" 'file untuk menampung output perintah SC

Dim serviceName                 As String
Dim statusService               As String
Dim row                         As Long
    
Private Sub tunggu(ByVal detik As Long)
    Dim pos As Long
    Dim h
    
    On Error Resume Next
    
    h = Second(Time)
    While pos < detik
        DoEvents
        If h <> Second(Time) Then
           h = Second(Time)
           pos = pos + 1
        End If
    Wend
End Sub

Private Sub execCommand(ByVal cmd As String)
    Dim result  As Long
    Dim lPid    As Long
    Dim lHnd    As Long
    Dim lRet    As Long
    
    cmd = "cmd /c " & cmd
    result = Shell(cmd, vbHide)
   
    lPid = result
    If lPid <> 0 Then
        lHnd = OpenProcess(SYNCHRONIZE, 0, lPid)
        If lHnd <> 0 Then
            lRet = WaitForSingleObject(lHnd, INFINITE)
            CloseHandle (lHnd)
        End If
    End If
End Sub

Private Sub execService(ByVal cmd As String)
    row = ListView1.SelectedItem.Index
    serviceName = ListView1.ListItems(row).SubItems(2)
    
    Screen.MousePointer = vbHourglass
    DoEvents
    Call execCommand("sc " & cmd & " " & serviceName)
    Call tunggu(3) 'tunggu sekitar 3 detik
    statusService = getStatusService(serviceName)
    
    ListView1.ListItems(row).SubItems(1) = statusService
    Screen.MousePointer = vbDefault
    
    If Not (Len(statusService) > 0) Then ListView1.ListItems.Remove row
End Sub

Private Sub showService()
    Dim i               As Long
    Dim fileHandler     As Integer
    
    Dim tmp1            As String
    Dim tmp2            As String
    Dim arrCol()        As String
    Dim arrRec()        As String
    
    Call execCommand("sc query state= all > " & FILE_SERVICE)
    
    fileHandler = FreeFile
    
    Open FILE_SERVICE For Input As fileHandler
    Do While Not EOF(fileHandler)
        Input #fileHandler, tmp1
        If Len(tmp1) > 0 Then
            'ambil nama service
            'nama service dibutuhkan untuk proses start, stop dan delete service
            If Left(tmp1, 12) = "SERVICE_NAME" Then tmp2 = tmp2 & Mid$(tmp1, 15) & COL_SPR
            
            'ambil informasi lengkap service
            If Left(tmp1, 12) = "DISPLAY_NAME" Then tmp2 = tmp2 & Mid$(tmp1, 15) & COL_SPR
            
            'state-> status service: stopped, running dan lain-lain
            If InStr(1, tmp1, "STATE") > 0 Then tmp2 = tmp2 & Mid$(tmp1, 25) & REC_SPR
        End If
    Loop
    Close fileHandler
    
    'contoh hasil perulangan diatas :
    'postgresql-8.4#postgresql-8.4 - PostgreSQL Server 8.4#RUNNING|MySQL#MySQL#RUNNING
    'SERVICE_NAME   DISPLAY_NAME                           STATE
    
    If Len(tmp2) > 0 Then
        tmp2 = Left(tmp2, Len(tmp2) - 1)
        
        arrRec = Split(tmp2, REC_SPR) 'pecah var tmp2 menjadi beberapa baris, REC_SPR = |
        With ListView1
            .ListItems.Clear
            For i = 0 To UBound(arrRec)
                If Len(arrRec(i)) > 0 Then
                    arrCol = Split(arrRec(i), COL_SPR) 'pecah var arrRec menjadi beberapa kolom, COL_SPR = #
                    
                    'tampilkan ke listview
                    .ListItems.Add , , arrCol(1), , 1 'DISPLAY_NAME
                    .ListItems(i + 1).SubItems(1) = StrConv(arrCol(2), vbProperCase) 'STATE
                    .ListItems(i + 1).SubItems(2) = arrCol(0) 'SERVICE_NAME
                End If
            Next i
        End With
    End If
    
    cmdStartService.Enabled = False
    cmdStopService.Enabled = False
    cmdDeleteService.Enabled = False
End Sub

Private Function getStatusService(ByVal serviceName As String) As String
    Dim fileHandler     As Integer
    Dim tmp             As String
    
    Call execCommand("sc query " & serviceName & " > " & FILE_SERVICE)
    
    fileHandler = FreeFile
    
    Open FILE_SERVICE For Input As fileHandler
    Do While Not EOF(fileHandler)
        Input #fileHandler, tmp
        If Len(tmp) > 0 Then
            If InStr(1, tmp, "STATE") > 0 Then getStatusService = StrConv(Mid$(tmp, 25), vbProperCase): Exit Do
        End If
    Loop
    Close fileHandler
End Function

Private Sub cmdDeleteService_Click()
    If MsgBox("Apakah service ini ingin dihapus ???", vbExclamation + vbYesNo, "Konfirmasi") = vbYes Then
        Call execService("delete")
    End If
End Sub

Private Sub cmdStartService_Click()
    Call execService("start")
End Sub

Private Sub cmdStopService_Click()
    Call execService("stop")
End Sub

Private Sub cmdRefreshService_Click()
    Call showService
End Sub

Private Sub Form_Load()
    'inisialisasi listview
    With ListView1
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .SmallIcons = ImageList1 'inisialisasi ImageList
       
        .ColumnHeaders.Add , , "Service", 5000
        .ColumnHeaders.Add , , "Status", 1500
        
        'kolom ini dibutuhkan untuk melakukan aksi terhadap service (start/stop service)
        'widthny diset = 0
        .ColumnHeaders.Add , , "ServiceName", 0
    End With
    
    cmdRefreshService_Click
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    row = ListView1.SelectedItem.Index
    statusService = ListView1.ListItems(row).SubItems(1)
    Select Case LCase$(statusService)
        Case "running"
            cmdStartService.Enabled = False
            cmdDeleteService.Enabled = False
            
            cmdStopService.Enabled = True
            
            
        Case "stopped"
            cmdStartService.Enabled = True
            cmdDeleteService.Enabled = True
            
            cmdStopService.Enabled = False
    End Select
End Sub
