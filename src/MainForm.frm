VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "AddConfigurations 23.1"
   ClientHeight    =   12375
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9510.001
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub btnCancel_Click()
    ExitApp
End Sub

Private Sub btnOk_Click()
    Run MainForm.txtConf.text
    ExitApp
End Sub

Private Sub UserForm_Initialize()
    Me.Hint.Caption = "Введите имена конфигураций, каждую на отдельной строке." + vbNewLine + _
        "Допускаются комментарии в кавычках." + vbNewLine + _
        "Пример:" + vbTab + "01" + vbNewLine + _
        vbTab + vbTab + "02 ""для ширины 600"""
End Sub
