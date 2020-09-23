Attribute VB_Name = "mBrowseForFolder"
'**************************************
' Name: BrowseFolder Module
' Description:For those that like to create objects on the fly and avoid extra controls, this shell interface module provides easy access to the BrowseForFolder dialog with drop down list support for constants and all system folders. It even creates them with their corresponding desktop.ini file and proper icon if not present already. Windows handles the NewFolder errors and error handling is included to rule out invalid paths.
' By: D.W.
'
' Assumes:Copy and paste code into notepad and save as Browse.bas then run. Sample code is in the Sub Main() procedure.
'
'This code is copyrighted and has' limited warranties.Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=60460&lngWId=1'for details.'**************************************

Option Explicit

Private Const BIF_EDITBOX = &H10
Private Const BIF_NEWDIALOGSTYLE = &H40

Public Enum ShellSpecialFolderConstants
    ssfALTSTARTUP = 29
    ssfAPPDATA = 26
    ssfBITBUCKET = 10
    ssfCOMMONALTSTARTUP = 30
    ssfCOMMONAPPDATA = 35
    ssfCOMMONDESKTOPDIR = 25
    ssfCOMMONFAVORITES = 31
    ssfCOMMONPROGRAMS = 23
    ssfCOMMONSTARTMENU = 22
    ssfCOMMONSTARTUP = 24
    ssfCONTROLS = 3
    ssfCOOKIES = 33
    ssfDESKTOP = 0
    ssfDESKTOPDIRECTORY = 16
    ssfDRIVES = 17
    ssfFAVORITES = 6
    ssfFONTS = 20
    ssfHISTORY = 34
    ssfINTERNETCACHE = 32
    ssfLOCALAPPDATA = 28
    ssfMYPICTURES = 39
    ssfNETHOOD = 19
    ssfNETWORK = 18
    ssfPERSONAL = 5
    ssfPRINTERS = 4
    ssfPRINTHOOD = 27
    ssfPROFILE = 40
    ssfPROGRAMFILES = 38
    ssfPROGRAMFILESx86 = 48
    ssfPROGRAMS = 2
    ssfRECENT = 8
    ssfSENDTO = 9
    ssfSTARTMENU = 11
    ssfSTARTUP = 7
    ssfSYSTEM = 37
    ssfSYSTEMx86 = 41
    ssfTEMPLATES = 21
    ssfWINDOWS = 36
End Enum

Public Function BrowseFolder(Optional OpenAt As ShellSpecialFolderConstants) As String
On Error Resume Next
    Dim ShellApplication As Object
    Dim Folder As Object
    Set ShellApplication = CreateObject("Shell.Application")
    Set Folder = ShellApplication.BrowseForFolder(0, "Select Folder...", BIF_EDITBOX Or BIF_NEWDIALOGSTYLE, CInt(OpenAt))
    BrowseFolder = Folder.Items.Item.Path
    On Error GoTo 0
    If Left(BrowseFolder, 2) = "::" Or InStr(1, BrowseFolder, "\") = 0 Then
    BrowseFolder = vbNullString
    End If
End Function

