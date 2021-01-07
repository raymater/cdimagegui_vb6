VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CDIMAGE (GUI)"
   ClientHeight    =   8385
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   15240
   Icon            =   "cdimagegui.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton formRefresh 
      Caption         =   "q"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   26.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11640
      TabIndex        =   49
      Top             =   240
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   10380
      Top             =   2580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame8 
      Caption         =   "Informations"
      Height          =   6795
      Left            =   11610
      TabIndex        =   45
      Top             =   1350
      Width           =   3375
      Begin VB.TextBox formInfos 
         Enabled         =   0   'False
         Height          =   6165
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   46
         Text            =   "cdimagegui.frx":030A
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.CommandButton formCreateButton 
      Caption         =   "Créer l'image"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12660
      TabIndex        =   44
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton formBoot_Parcourir 
      Caption         =   "Parcourir"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9540
      TabIndex        =   40
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Paramètres"
      Height          =   6795
      Left            =   240
      TabIndex        =   5
      Top             =   1350
      Width           =   11085
      Begin VB.Frame Frame7 
         Caption         =   "Options de démarrage"
         Height          =   2535
         Left            =   5640
         TabIndex        =   37
         Top             =   3990
         Width           =   5205
         Begin VB.TextBox formBoot_InputID 
            Enabled         =   0   'False
            Height          =   345
            Left            =   2640
            MaxLength       =   2
            TabIndex        =   43
            Text            =   "80"
            Top             =   1530
            Width           =   435
         End
         Begin VB.CheckBox formBoot_ID 
            Caption         =   "Spécifier l'ID de plateforme :"
            Height          =   375
            Left            =   240
            TabIndex        =   42
            Top             =   1530
            Width           =   2385
         End
         Begin VB.CheckBox formBoot_Floppy 
            Caption         =   "Désactiver l'émulation disquette"
            Height          =   405
            Left            =   240
            TabIndex        =   41
            Top             =   1170
            Width           =   4035
         End
         Begin VB.TextBox formBoot_FileBoot 
            Enabled         =   0   'False
            Height          =   375
            Left            =   540
            TabIndex        =   39
            Top             =   660
            Width           =   3015
         End
         Begin VB.CheckBox formBoot_Secteur 
            Caption         =   "Spécifier le fichier de secteur de démarrage"
            Height          =   315
            Left            =   240
            TabIndex        =   38
            Top             =   300
            Width           =   4155
         End
      End
      Begin VB.Frame groupUDF 
         Caption         =   "UDF Options"
         Enabled         =   0   'False
         Height          =   2535
         Left            =   240
         TabIndex        =   20
         Top             =   3990
         Width           =   5205
         Begin VB.CheckBox formUDF_VideoZone 
            Caption         =   "Compatibilité UDF Video Zone (DVD-Video et DVD-Audio)"
            Height          =   345
            Left            =   180
            TabIndex        =   27
            Top             =   2040
            Width           =   4815
         End
         Begin VB.CheckBox formUDF_Allocation 
            Caption         =   "Descripteurs d'allocation longs"
            Height          =   435
            Left            =   180
            TabIndex        =   26
            Top             =   1650
            Width           =   3945
         End
         Begin VB.CheckBox formUDF_Clairseme 
            Caption         =   "Fichiers UDF clairsemés"
            Enabled         =   0   'False
            Height          =   405
            Left            =   180
            TabIndex        =   25
            Top             =   1320
            Width           =   3105
         End
         Begin VB.CheckBox formUDF_FID 
            Caption         =   "Incorporer les entrées FID UDF"
            Enabled         =   0   'False
            Height          =   345
            Left            =   180
            TabIndex        =   24
            Top             =   990
            Width           =   4815
         End
         Begin VB.CheckBox formUDF_Input 
            Caption         =   "Incorporer les données de fichier dans l'entrée d'étendue UDF"
            Enabled         =   0   'False
            Height          =   315
            Left            =   180
            TabIndex        =   23
            Top             =   660
            Width           =   4845
         End
         Begin VB.OptionButton formUDF_UDF 
            Caption         =   "UDF Uniquement"
            Height          =   285
            Left            =   2640
            TabIndex        =   22
            Top             =   330
            Width           =   2205
         End
         Begin VB.OptionButton formUDF_ISO 
            Caption         =   "Compatibilité ISO 9660"
            Height          =   285
            Left            =   180
            TabIndex        =   21
            Top             =   330
            Value           =   -1  'True
            Width           =   2355
         End
      End
      Begin VB.Frame groupJoliet 
         Caption         =   "Options Joliet"
         Enabled         =   0   'False
         Height          =   735
         Left            =   240
         TabIndex        =   17
         Top             =   3150
         Width           =   5205
         Begin VB.OptionButton formJoliet_Unicode 
            Caption         =   "Unicode"
            Height          =   315
            Left            =   2640
            TabIndex        =   19
            Top             =   300
            Width           =   1455
         End
         Begin VB.OptionButton formJoliet_ISO 
            Caption         =   "Compatibilité ISO 9660"
            Height          =   315
            Left            =   180
            TabIndex        =   18
            Top             =   300
            Value           =   -1  'True
            Width           =   2325
         End
      End
      Begin VB.Frame groupISO 
         Caption         =   "Options ISO 9660"
         Height          =   1425
         Left            =   240
         TabIndex        =   12
         Top             =   1620
         Width           =   5205
         Begin VB.CheckBox formISO_Maj 
            Caption         =   "Autoriser les noms de fichiers en minuscule"
            Height          =   315
            Left            =   180
            TabIndex        =   16
            Top             =   990
            Width           =   4035
         End
         Begin VB.OptionButton formISO_LongNT 
            Caption         =   "Long (30) - compatibilité Windows NT 3.51"
            Height          =   285
            Left            =   180
            TabIndex        =   15
            Top             =   660
            Width           =   4365
         End
         Begin VB.OptionButton formISO_Long 
            Caption         =   "Long (30)"
            Height          =   285
            Left            =   2010
            TabIndex        =   14
            Top             =   300
            Width           =   1725
         End
         Begin VB.OptionButton formISO_DOS 
            Caption         =   "DOS (8.3)"
            Height          =   285
            Left            =   180
            TabIndex        =   13
            Top             =   300
            Value           =   -1  'True
            Width           =   1395
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Système de fichiers"
         Height          =   705
         Left            =   240
         TabIndex        =   8
         Top             =   810
         Width           =   5205
         Begin VB.OptionButton formFileUDF 
            Caption         =   "UDF"
            Height          =   345
            Left            =   3720
            TabIndex        =   11
            Top             =   270
            Width           =   1215
         End
         Begin VB.OptionButton formFileJoliet 
            Caption         =   "Joliet"
            Height          =   285
            Left            =   2010
            TabIndex        =   10
            Top             =   300
            Width           =   1515
         End
         Begin VB.OptionButton formFileISO9660 
            Caption         =   "ISO 9660"
            Height          =   285
            Left            =   180
            TabIndex        =   9
            Top             =   300
            Value           =   -1  'True
            Width           =   1605
         End
      End
      Begin VB.TextBox formLabel 
         Height          =   315
         Left            =   870
         TabIndex        =   7
         Top             =   360
         Width           =   4575
      End
      Begin VB.Frame Frame6 
         Caption         =   "Général"
         Height          =   3615
         Left            =   5640
         TabIndex        =   28
         Top             =   270
         Width           =   5205
         Begin MSComCtl2.DTPicker formGen_Hour 
            Height          =   315
            Left            =   1260
            TabIndex        =   48
            Top             =   1530
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   22740994
            CurrentDate     =   44198
         End
         Begin MSComCtl2.DTPicker formGen_Date 
            Height          =   315
            Left            =   1260
            TabIndex        =   47
            Top             =   1170
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   22740993
            CurrentDate     =   44198
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   4500
            Top             =   270
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CheckBox formGen_Force 
            Caption         =   "Forcer la création de l'image"
            Height          =   345
            Left            =   270
            TabIndex        =   36
            Top             =   3090
            Width           =   3255
         End
         Begin VB.CheckBox formGen_ANSI 
            Caption         =   "Utiliser ANSI pour les noms de fichiers"
            Height          =   315
            Left            =   270
            TabIndex        =   35
            Top             =   2760
            Width           =   3825
         End
         Begin VB.CheckBox formGen_Hide 
            Caption         =   "Inclure les fichiers/dossiers cachés"
            Height          =   315
            Left            =   270
            TabIndex        =   34
            Top             =   2400
            Width           =   3735
         End
         Begin VB.CheckBox formGen_GMT 
            Caption         =   "Utiliser le fuseau horaire GMT"
            Height          =   345
            Left            =   270
            TabIndex        =   33
            Top             =   2010
            Width           =   3315
         End
         Begin VB.CheckBox formGen_DateCheck 
            Caption         =   "Définir date/heure pour les fichiers/dossiers :"
            Height          =   465
            Left            =   270
            TabIndex        =   30
            Top             =   780
            Width           =   4575
         End
         Begin VB.CheckBox formGen_limit 
            Caption         =   "Ignorer la limitation à 650 Mio"
            Height          =   345
            Left            =   270
            TabIndex        =   29
            Top             =   330
            Width           =   4635
         End
         Begin VB.Label Label6 
            Caption         =   "Heure :"
            Height          =   285
            Left            =   630
            TabIndex        =   32
            Top             =   1590
            Width           =   645
         End
         Begin VB.Label Label5 
            Caption         =   "Date :"
            Height          =   255
            Left            =   630
            TabIndex        =   31
            Top             =   1230
            Width           =   525
         End
      End
      Begin VB.Label Label3 
         Caption         =   "Label"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.CommandButton formSaveButton 
      Caption         =   "Parcourir"
      Height          =   375
      Left            =   9990
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox formSaveAs 
      Enabled         =   0   'False
      Height          =   345
      Left            =   1530
      TabIndex        =   3
      Top             =   720
      Width           =   8385
   End
   Begin VB.ComboBox formListOpticalDisks 
      Height          =   315
      Left            =   1530
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   210
      Width           =   9795
   End
   Begin VB.Label Label2 
      Caption         =   "Enregistrer sous :"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   750
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "Lecteur optique :"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetVolumeInformation Lib "kernel32.dll" _
Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, _
lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, _
lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, _
ByVal nFileSystemNameSize As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
Private Function isExist(letter) As Boolean
    On Error GoTo noExist:
    Dim driveExist, d, fso
    
    isExist = False

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set d = fso.GetDrive(fso.GetDriveName(fso.GetAbsolutePathName(letter)))
    
    isExist = True
    
    On Error GoTo 0
    Exit Function
    
noExist:
    isExist = False
    
End Function

Private Function FileExists(ByVal Fname As String) As Boolean
        FileExists = IIf(Dir(Fname) <> "", True, False)
End Function

Private Sub Form_Load()
    Dim i As Integer
    
    Dim letters As Variant
    letters = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    
    For i = 0 To 25 Step 1
        Dim existDisk
        
        Dim letter As String
        letter = letters(i) & ":\"
        
        existDisk = isExist(letter)
        
        
        If (existDisk = True) Then
        
            Dim fso, d
            
            Set fso = CreateObject("Scripting.FileSystemObject")
            Set d = fso.GetDrive(fso.GetDriveName(fso.GetAbsolutePathName(letter)))
                
            If (d.DriveType = 4) Then
                formListOpticalDisks.AddItem (letter)
            End If
            
        End If
    Next
    
    If (formListOpticalDisks.ListCount > 0) Then
        formListOpticalDisks.ListIndex = 0
    End If
    
End Sub

Private Sub formBoot_FileBoot_Change()
    Dim pos, rMes
    pos = InStr(formBoot_FileBoot.Text, " ")
    
    If (pos > 0) Then
        rMes = MsgBox("Le chemin du fichier ne doit pas comporter d'espaces", vbExclamation + vbOKOnly, "Erreur")
        formBoot_FileBoot.Text = ""
    End If
End Sub

Private Sub formBoot_ID_Click()
    If (formBoot_ID.Value = Checked) Then
        formBoot_InputID.Enabled = True
    Else
        formBoot_InputID.Enabled = False
    End If
End Sub

Private Sub formBoot_Parcourir_Click()
    CommonDialog2.Filter = "Boot sector file (*.bin;*.com)|*.bin;*.com"
    CommonDialog2.DialogTitle = "Ouvrir fichier Boot sector"
    CommonDialog2.ShowOpen
    
    Dim f
    f = CommonDialog2.FileName
    
    formBoot_FileBoot.Text = f
End Sub

Private Sub formBoot_Secteur_Click()
    If (formBoot_Secteur.Value = Checked) Then
        formBoot_Parcourir.Enabled = True
    Else
        formBoot_Parcourir.Enabled = False
    End If
End Sub

'--------------------
' BOUTON CREER IMAGE
'--------------------
Private Sub formCreateButton_Click()
    Dim commandOptions As String
    Dim rMes
    Dim commandPath
    commandOptions = ""
    
    If (formListOpticalDisks.ListCount > 0) Then
        If (formSaveAs.Text <> "") Then
            If (formLabel.Text <> "") Then
                formLabel.Text = UCase(formLabel.Text)
                formLabel.Text = Replace(formLabel.Text, " ", "_")
                
                commandOptions = "-"
                
                If (formFileISO9660.Value = True) Then
                    If (formISO_Long.Value = True) Then
                        commandOptions = commandOptions & "n"
                    End If
                    If (formISO_LongNT.Value = True) Then
                        commandOptions = commandOptions & "nt"
                    End If
                    If (formISO_Maj.Value = Checked) Then
                        commandOptions = commandOptions & "d"
                    End If
                End If
                
                If (formFileJoliet.Value = True) Then
                    If (formJoliet_ISO.Value = True) Then
                        commandOptions = commandOptions & "j1"
                        
                        If (formISO_Long.Value = True) Then
                            commandOptions = commandOptions & "n"
                        End If
                        If (formISO_LongNT.Value = True) Then
                            commandOptions = commandOptions & "nt"
                        End If
                        If (formISO_Maj.Value = Checked) Then
                            commandOptions = commandOptions & "d"
                        End If
                        
                    End If
                    If (formJoliet_Unicode.Value = True) Then
                        commandOptions = commandOptions & "j2"
                    End If
                End If
                
                If (formFileUDF.Value = True) Then
                    If (formUDF_ISO.Value = True) Then
                        commandOptions = commandOptions & "u1"
                        
                        If (formISO_Long.Value = True) Then
                            commandOptions = commandOptions & "n"
                        End If
                        If (formISO_LongNT.Value = True) Then
                            commandOptions = commandOptions & "nt"
                        End If
                        If (formISO_Maj.Value = Checked) Then
                            commandOptions = commandOptions & "d"
                        End If
                    End If
                    
                    If (formUDF_UDF.Value = True) Then
                        commandOptions = commandOptions & "u2"
                        
                        If (formUDF_Clairseme.Value = Checked) Then
                            commandOptions = commandOptions & "us"
                        End If
                        If (formUDF_FID.Value = Checked) Then
                            commandOptions = commandOptions & "uf"
                        End If
                        If (formUDF_Input.Value = Checked) Then
                            commandOptions = commandOptions & "ue"
                        End If
                    End If
                    
                    
                    If (formUDF_Allocation.Value = Checked) Then
                        commandOptions = commandOptions & "yl"
                    End If
                    If (formUDF_VideoZone.Value = Checked) Then
                        commandOptions = commandOptions & "uv"
                    End If
                End If
                
                If (formGen_limit.Value = Checked) Then
                    commandOptions = commandOptions & "m"
                End If
                If (formGen_GMT.Value = Checked) Then
                    commandOptions = commandOptions & "g"
                End If
                If (formGen_Hide.Value = Checked) Then
                    commandOptions = commandOptions & "h"
                End If
                If (formGen_ANSI.Value = Checked) Then
                    commandOptions = commandOptions & "c"
                End If
                If (formGen_Force.Value = Checked) Then
                    commandOptions = commandOptions & "k"
                End If
                
                If (formBoot_Floppy.Value = Checked) Then
                    commandOptions = commandOptions & "e"
                End If
                If (formBoot_ID.Value = Checked) Then
                    commandOptions = commandOptions & "p" & formBoot_InputID.Text
                End If
                
                If (commandOptions = "-") Then
                    commandOptions = ""
                End If
                If (formLabel.Text <> "") Then
                    commandOptions = commandOptions & " -l" & """" & formLabel.Text & """"
                End If
                
                If (formGen_DateCheck.Value = Checked) Then
                    Dim month, day, year, hour, min, sec
                    Dim m1, d1, y1, h1, n1, s1
                    
                    m1 = ""
                    d1 = ""
                    y1 = ""
                    h1 = ""
                    n1 = ""
                    s1 = ""
                    
                    day = formGen_Date.day
                    month = formGen_Date.month
                    year = formGen_Date.year
                    hour = formGen_Hour.hour
                    min = formGen_Hour.Minute
                    sec = formGen_Hour.Second
                    
                    If (day < 10) Then
                        d1 = "0" & CStr(day)
                    Else
                        d1 = CStr(day)
                    End If
                    If (month < 10) Then
                        m1 = "0" & CStr(month)
                    Else
                        m1 = CStr(month)
                    End If
                    If (year < 10) Then
                        y1 = "0" & CStr(year)
                    Else
                        y1 = CStr(year)
                    End If
                    If (hour < 10) Then
                        h1 = "0" & CStr(hour)
                    Else
                        h1 = CStr(hour)
                    End If
                    If (min < 10) Then
                        n1 = "0" & CStr(min)
                    Else
                        n1 = CStr(min)
                    End If
                    If (sec < 10) Then
                        s1 = "0" & CStr(sec)
                    Else
                        s1 = CStr(sec)
                    End If
                    
                    commandOptions = commandOptions & " -t" & m1 & "/" & d1 & "/" & y1 & "," & h1 & ":" & n1 & ":" & s1
                End If
                
                If (formBoot_Secteur.Value = Checked) Then
                    If (formBoot_FileBoot <> "") Then
                        commandOptions = commandOptions & " -b" & """" & formBoot_FileBoot.Text & """"
                    End If
                End If
                
                commandPath = App.Path & "\cdimage.exe"
                Dim ess
                ess = FileExists(commandPath)
                
                If (ess = True) Then
                    Dim Options
                    Dim hwnd
                    Dim pathA
                    Dim folderFile
                    Dim fileSave
                    Dim i
                    folderFile = ""
                    fileSave = ""
                    hwnd = Me.hwnd
                    
                    pathA = Split(formSaveAs.Text, "\")
                    
                    For i = LBound(pathA) To (UBound(pathA) - 1) Step 1
                        folderFile = folderFile & pathA(i) & "\"
                    Next
                    
                    fileSave = pathA(UBound(pathA))
                    
                    'MsgBox (fileSave)
                    
                    Options = commandOptions & " " & formListOpticalDisks.List(formListOpticalDisks.ListIndex) & " " & fileSave
                    
                    'MsgBox (LBound(pathA))
                    'MsgBox (UBound(pathA))
                    'MsgBox (folderFile)
                    'MsgBox (Options)
                    
                    Call ShellExecute(hwnd, "open", commandPath, Options, folderFile, 1)
                Else
                    rMes = MsgBox("L'application CDIMAGE.EXE n'existe pas dans " & App.Path, vbExclamation + vbOKOnly, "Erreur")
                End If
                
            Else
                rMes = MsgBox("Aucun label indiqué", vbExclamation + vbOKOnly, "Erreur")
            End If
        Else
            rMes = MsgBox("Aucun emplacement sélectionné", vbExclamation + vbOKOnly, "Erreur")
        End If
    Else
        rMes = MsgBox("Aucun lecteur sélectionné", vbExclamation + vbOKOnly, "Erreur")
    End If
    
End Sub

Private Sub formFileISO9660_Click()
    If (formFileISO9660.Value = True) Then
        groupISO.Enabled = True
        groupJoliet.Enabled = False
        groupUDF.Enabled = False
    End If
End Sub

Private Sub formFileJoliet_Click()
    If (formFileJoliet.Value = True) Then
        If (formJoliet_ISO.Value = True) Then
            groupISO.Enabled = True
        Else
            groupISO.Enabled = False
        End If
        groupJoliet.Enabled = True
        groupUDF.Enabled = False
    End If
End Sub

Private Sub formFileUDF_Click()
    If (formFileUDF.Value = True) Then
        If (formUDF_ISO.Value = True) Then
            groupISO.Enabled = True
        Else
            groupISO.Enabled = False
        End If
        groupJoliet.Enabled = False
        groupUDF.Enabled = True
    End If
End Sub

Private Sub formGen_DateCheck_Click()
    If (formGen_DateCheck.Value = Checked) Then
        formGen_Date.Enabled = True
        formGen_Hour.Enabled = True
    Else
        formGen_Date.Enabled = False
        formGen_Hour.Enabled = False
    End If
End Sub

Private Sub formISO_DOS_Click()
    If (formISO_LongNT.Value = True) Then
        formISO_Maj.Enabled = False
    Else
        formISO_Maj.Enabled = True
    End If
End Sub

Private Sub formISO_Long_Click()
    If (formISO_LongNT.Value = True) Then
        formISO_Maj.Enabled = False
    Else
        formISO_Maj.Enabled = True
    End If
End Sub

Private Sub formISO_LongNT_Click()
    If (formISO_LongNT.Value = True) Then
        formISO_Maj.Enabled = False
    Else
        formISO_Maj.Enabled = True
    End If
End Sub

Private Sub formJoliet_ISO_Click()
    If (formJoliet_ISO.Value = True) Then
        groupISO.Enabled = True
    Else
        groupISO.Enabled = False
    End If
End Sub

Private Sub formJoliet_Unicode_Click()
    If (formJoliet_Unicode.Value = True) Then
        groupISO.Enabled = False
    Else
        groupISO.Enabled = True
    End If
End Sub

Private Sub formLabel_LostFocus()
    Dim t
    t = UCase(formLabel.Text)
    t = Replace(t, " ", "_")
    formLabel.Text = t
End Sub

Private Sub formListOpticalDisks_Click()

    Dim letter
    letter = formListOpticalDisks.Text

    Dim volname As String   ' receives volume name
    Dim sn As Long          ' receives serial number
    Dim snstr As String     ' display form of serial number
    Dim maxcomplen As Long  ' receives maximum component length
    Dim sysflags As Long    ' receives file system flags
    Dim sysname As String   ' receives the file system name
    Dim retval As Long      ' return value

    ' Initialize string buffers.
    volname = Space(256)
    sysname = Space(256)
    
    ' Get information about drive volume.
    retval = GetVolumeInformation(letter, volname, Len(volname), sn, maxcomplen, _
        sysflags, sysname, Len(sysname))
        
    If (retval = 1) Then
        ' Remove the trailing nulls from the two strings.
        volname = Left(volname, InStr(volname, vbNullChar) - 1)
        sysname = Left(sysname, InStr(sysname, vbNullChar) - 1)
        
        ' Format the serial number properly.
        snstr = Trim(Hex(sn))
        snstr = String(8 - Len(snstr), "0") & snstr
        snstr = Left(snstr, 4) & "-" & Right(snstr, 4)
        
        ' Display the volume name, serial number, and file system name.
        formInfos.Text = "Lecteur : " & letter & vbNewLine
        formInfos.Text = formInfos.Text & "Volume : " & volname & vbNewLine
        formInfos.Text = formInfos.Text & "Numéro de série : " & snstr & vbNewLine
        formInfos.Text = formInfos.Text & "Système de fichiers : " & sysname & vbNewLine
        
        formLabel.Text = UCase(volname)
        
    End If
End Sub

Private Sub formRefresh_Click()
    Dim i As Integer
    
    Dim letters As Variant
    letters = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    
    formListOpticalDisks.Clear
    
    For i = 0 To 25 Step 1
        Dim existDisk
        
        Dim letter As String
        letter = letters(i) & ":\"
        
        existDisk = isExist(letter)
        
        
        If (existDisk = True) Then
        
            Dim fso, d
            
            Set fso = CreateObject("Scripting.FileSystemObject")
            Set d = fso.GetDrive(fso.GetDriveName(fso.GetAbsolutePathName(letter)))
                
            If (d.DriveType = 4) Then
                formListOpticalDisks.AddItem (letter)
            End If
            
        End If
    Next
    
    If (formListOpticalDisks.ListCount > 0) Then
        formListOpticalDisks.ListIndex = 0
    End If
End Sub

Private Sub formSaveButton_Click()
    CommonDialog1.Filter = "Image ISO (*.iso)|*.iso|Fichier image (*.img)|*.img"
    CommonDialog1.DefaultExt = "iso"
    CommonDialog1.DialogTitle = "Enregistrer sous..."
    CommonDialog1.ShowSave
    
    Dim f
    f = CommonDialog1.FileName
    
    formSaveAs.Text = f
    
End Sub

Private Sub formUDF_ISO_Click()
    If (formUDF_ISO.Value = True) Then
        formUDF_Input.Enabled = False
        formUDF_FID.Enabled = False
        formUDF_Clairseme.Enabled = False
        groupISO.Enabled = True
    Else
        formUDF_Input.Enabled = True
        formUDF_FID.Enabled = True
        formUDF_Clairseme.Enabled = True
        groupISO.Enabled = False
    End If
End Sub

Private Sub formUDF_UDF_Click()
    If (formUDF_UDF.Value = True) Then
        formUDF_Input.Enabled = True
        formUDF_FID.Enabled = True
        formUDF_Clairseme.Enabled = True
        groupISO.Enabled = False
    Else
        formUDF_Input.Enabled = False
        formUDF_FID.Enabled = False
        formUDF_Clairseme.Enabled = False
        groupISO.Enabled = True
    End If
End Sub
