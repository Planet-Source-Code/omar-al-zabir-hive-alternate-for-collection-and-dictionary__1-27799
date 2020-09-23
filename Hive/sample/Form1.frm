VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Dim objHive As New CHive
    Dim objCollection As New Collection
    With objHive
        .Add "Item Value", "Key"
        .Add "Item Before Key", "KeyBefore", "Key"
        .Add "Itme After Key", "KeyAfter", , "Key"
        
        .Add objCollection, "123"
        
        MsgBox .Exist("Key")
        
        .CompareMode = vbTextCompare
        MsgBox .Exist("keybefore")
        
        .AllowDuplicate = True
         .Add objCollection, "Key"
                
        MsgBox .Item("123").Count
        
        .Remove "123"
        MsgBox .Item("key")
        
        MsgBox IsObject(.FindFirst("key"))
        MsgBox IsObject(.FindNext("key"))
        .Clear
    
        
        
        
    End With
End Sub
