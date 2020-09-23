VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frm_main 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Simple Way to read VBScript - Jonas Persson"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSScriptControlCtl.ScriptControl script_1 
      Left            =   4080
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.CommandButton cmd_Run 
      Caption         =   "Run VBScript"
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox txt_script 
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frm_main.frx":0000
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//How to execute a VBScript in one line of code
'//By Jonas Persson aka Mrsomeone aka Snutte =)
'//Using Microsoft Windows ScriptControl [Msscript.ocx]
'//Can Be downloaded at: http://msdn.microsoft.com/downloads/default.asp?url=/downloads/sample.asp?url=/MSDN-FILES/027/001/732/msdncompositedoc.xml&frame=true

Private Sub cmd_Run_Click()
    '//so simple as it can be
    script_1.ExecuteStatement (txt_script.Text)
End Sub
