' freza.net 2005
' DChat Project
' IRC colors to RTF Converter
' RTF colors to IRC Converter
'------------------------------------------------------------------------------------
' Programmed by Garikk
' UNI_plugin API
' DC2Converter v.1.5.2.4
'------------------------------------------------------------------------------------
' ЗАРЕЗЕРВИРОВАНИЫЕ ASCII КОДЫ!!
' 2,3,31
'------------------------------------------------------------------------------------
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Drawing.Design
Imports UNI_Plugin
Imports UNI_Plugin.UNI_DChat
Imports UNI_Plugin.UNI_DChat.DCB_V1

#Region " Plugin Initialisation UNI_Plugin API (FMS)"
Public Class UNI_PLUGININFO
    Implements UNI_Universal.UNI_iPluginGetInfo

    Public ReadOnly Property DCB_GetInfo() As UNI_Universal.UNI_PluginInfo Implements UNI_Universal.UNI_iPluginGetInfo.DCB_GetInfo
        Get
            Dim Ret As UNI_Universal.UNI_PluginInfo
            Ret.INF_NAME = "DC2Conv"
            Ret.INF_PID = "/dc2conv"
            Ret.INF_VER = ModVER
            Ret.INF_DESCRIPTION = RTF2IRC_DESC
            Ret.INF_OPTIONS = Nothing
            Return Ret
        End Get
    End Property
End Class

Public Class UNI_Plugin_Connect
    Implements UNI_V1.UNI_iPluginConnector

    Public Function ExecCMD(ByVal cmd As String, Optional ByVal Param As Object = Nothing) As Object Implements UNI_V1.UNI_iPluginConnector.ExecCMD
        Return DC2CONV_Plugin_Coordinator.UNI_SE_Exec(cmd)
    End Function
    Public Function PluginINIT(ByVal RegPID As String, ByVal BaseSE As UNI_V1.UNI_BaseSEConnector) As Object Implements UNI_V1.UNI_iPluginConnector.PluginInit
        InitPlugin(RegPID, BaseSE)
    End Function
End Class
#End Region

#Region " Plugin Initialisation UNI_Plugin.DCB_Plugin API (DChat)"
Public Class DCB_PLUGININFO
    Implements DCB_Universal.DCB_UNI_iPluginGetInfo

    Public ReadOnly Property DCB_GetInfo() As UNI_Plugin.UNI_DChat.DCB_Universal.DCB_UNI_PluginInfo Implements UNI_Plugin.UNI_DChat.DCB_Universal.DCB_UNI_iPluginGetInfo.DCB_GetInfo
        Get
            Dim Ret As DCB_Universal.DCB_UNI_PluginInfo
            Ret.INF_NAME = "DC2Conv"
            Ret.INF_CMDEXEC = "/dc2conv"
            Ret.INF_VER = New Version(ModVER)
            Ret.INF_DESCRIPTION = RTF2IRC_DESC
            Ret.INF_OPTIONS = Nothing
            Ret.INF_PLUGINTYPE = DCB_Universal.DCB_UNI_PluginTypes.DCB_V2_m
            Ret.INF_PLUGINTYPESTR = "DCB_V2m"
            Return Ret
        End Get
    End Property
End Class

Public Class DCB_Plugin
    ' Связь с базой DChat (UNI V2_m)
    Implements DCB_Universal.DCB_iPlugin

    Public Function DCB_PLUGIN_Cmd(ByVal PID As String, ByVal Cmd As String) As Object Implements DCB_Universal.DCB_iPlugin.DCB_DCSE_Exec
        ' Выполнение комманд пришедших с базы
        Return DC2CONV_Plugin_Coordinator.UNI_SE_Exec(Cmd)
    End Function

    Public ReadOnly Property DCB_GetInfo() As DCB_Universal.DCB_UNI_PluginInfo Implements DCB_Universal.DCB_iPlugin.UNI_GetInfo
        Get
            Dim Ret As DCB_Universal.DCB_UNI_PluginInfo
            Ret.INF_NAME = "DC2Conv"
            Ret.INF_CMDEXEC = "/dc2conv"
            Ret.INF_VER = New Version(ModVER)
            Ret.INF_DESCRIPTION = RTF2IRC_DESC
            Ret.INF_OPTIONS = Nothing
            Ret.INF_PLUGINTYPE = DCB_Universal.DCB_UNI_PluginTypes.DCB_V2_m
            Ret.INF_PLUGINTYPESTR = "DCB_V2m"
            Return Ret
        End Get
    End Property
    Public Sub DCB_PLUGIN_Init(ByRef dcb_CHLCTL As DCB_V1.DCB_Channels2.DCB_iChannelsControl, ByRef dcb_USRCTL As DCB_Channels2.DCB_iUsersControl, ByRef DCGUI As DCB_Base.DCB_GUI_Control, ByRef NETWORK As DCB_Base.DCB_NetSocket, ByRef dcb_DCSE As DCB_Base.DCB_DCSE) Implements DCB_Universal.DCB_iPlugin.DCB_DCSE_V1_c_Plugininit
        ' Не используется
    End Sub
End Class
#End Region

#Region "UNI_SE"
' Universal Scripting Engine Ext
Module DC2CONV_Plugin_Coordinator
    Public Const RTF2IRC_DESC = "RTF to IRC / IRC to RTF text converter, ver. 1.5.1.4, DZFS.net 2005, Created by Garikk "
    Public Const ModVER = "1.5.2.4"
    Dim RegistredPID As String
    Dim BSE As UNI_V1.UNI_BaseSEConnector
    Friend Function InitPlugin(ByVal RegPID As String, ByVal BaseSE As UNI_V1.UNI_BaseSEConnector) As Object
        BSE = BaseSE
        RegistredPID = RegPID
        Return "DC2Conv Loaded"
    End Function
    Friend Function UNI_SE_Exec(ByVal cmd As String) As Object
        'If Not cmd.ToLower.Trim.StartsWith(RegistredPID) Then Return "1"
        If cmd.StartsWith("/dc2conv" & Chr(1)) Then cmd = cmd.Substring(9)
        If RegistredPID <> "" Then cmd = cmd.ToLower.Trim.Substring(RegistredPID.Length + 1)
        Select Case cmd.Split(Chr(1))(0)
            Case "rtftodc2"
                Dim TL As String = cmd.Split(Chr(1))(1)
                Return DC2Converter.RTFToDC2(TL)

            Case "dc2tortf"
                Dim FNT As String = cmd.Split(Chr(1))(1)
                Dim SIZ As Integer = cmd.Split(Chr(1))(2)
                Dim TL As String = cmd.Split(Chr(1))(3)
                Dim RET As String
                RET = DC2Converter.DC2ToRTF(tl, fnt, siz)
           
                Return RET
            Case "txttortf"

                Dim FNT As String = cmd.Split(Chr(1))(1)
                Dim SIZ As Integer = cmd.Split(Chr(1))(2)
                Dim TL As String = cmd.Split(Chr(1))(3)
                Return DC2Converter.TxtToRTF(TL, fnt, siz)

            Case "rtftotxt"
                Dim TL As String = cmd.Split(Chr(1))(1)
                Return DC2Converter.RtfToTxt(TL)
            Case "getargbfromirccode"
                Return DC2Converter.GetARGBfromIRCCode(cmd.Split(Chr(1))(1))
            Case "getirccodefroemargb"
                Return DC2Converter.GetIRCColorCode(cmd.Split(Chr(1))(1))
            Case "descriptionline"
                Return "DC2Conv, IRC to RTF/RTF to IRC Colored text converter, Version 1.5, Created by Garikk"
            Case "removeirccodes"
                Return DC2Converter.RemoveIRCCodes(cmd.Split(Chr(1))(1))
            Case "uni_pmgr_init"
                Return UNI_SE_Exec("descriptionline")
            Case Else
                Return "1"
        End Select
    End Function
End Module

#End Region







Public Class DC2Converter
    Dim RTFBox As New RichTextBox
    ' RTF-mIRC-mIRC-RTF Colored text converter v.1.5
    Dim R As Short
    Private Structure DC2Indexer
        Dim StartPosition As Integer
        Dim LengthString As Integer
        Dim IndexerType As DC2IndexerTypes
        Dim SetColor As Integer
    End Structure
    Private Structure IRC2RTFColorOffsetSet
        Dim ClrString As String
        Dim ClrOffset As Integer
    End Structure

    Private Enum DC2IndexerTypes
        DC2_RENDER_BOLD = 0
        DC2_RENDER_UNDERL = 1
        DC2_RENDER_COLOR = 2
        DC2_RENDER_REV = 3
        DC2_RENDER_ITALIC = 4
    End Enum

    Public Shared Function RTFToDC2(ByVal RTF As String) As String
        Static RTFbox As New RichTextBox
        Dim DC2Text As String
        Dim SelBold As Boolean = False
        Dim SelItal As Boolean = False
        Dim SelUnderl As Boolean = False
        Dim SelColor As Integer '= RTFBox.SelectionColor.ToArgb
        Dim Tmp As Integer
        Dim STEP1 As String()

        RTFbox.Rtf = RTF
        RTFbox.SelectionStart = 0
        RTFbox.SelectionLength = 1
        SelColor = RTFbox.SelectionColor.ToArgb

        '----------
        Dim REP1 As Integer = GetIRCColorCode(RTFbox.SelectionColor.ToArgb)
        If REP1 <> 99 Then
            DC2Text += Chr(3) & CStr(REP1)
            SelColor = RTFbox.SelectionColor.ToArgb
        End If
        'DC2Text += Chr(3) & GetIRCColorCode(RTFBox.SelectionColor.ToArgb) : SelColor = RTFBox.SelectionColor.ToArgb
        '---------
        For Tmp = 0 To RTFbox.Text.TrimEnd.Length
            RTFbox.Select(Tmp, 1)
            If RTFbox.SelectionFont.Bold = True And SelBold = False Then DC2Text += Chr(2) : SelBold = True
            If RTFbox.SelectionFont.Underline = True And SelUnderl = False Then DC2Text += Chr(31) : SelUnderl = True
            If RTFbox.SelectionFont.Bold = False And SelBold = True Then DC2Text += Chr(2) : SelBold = False
            If RTFbox.SelectionFont.Underline = False And SelUnderl = True Then DC2Text += Chr(31) : SelUnderl = False

            If RTFbox.SelectionColor.ToArgb <> SelColor Then
                Dim REP As Integer = GetIRCColorCode(RTFbox.SelectionColor.ToArgb)
                If REP <> 99 Then
                    DC2Text += Chr(3) & CStr(REP)
                End If
                SelColor = RTFbox.SelectionColor.ToArgb
            End If
            DC2Text += RTFbox.SelectedText
        Next
        Return DC2Text
    End Function
    Public Shared Function RemoveIRCCodes(ByVal DC2 As String) As String
        Return DC2Converter.RtfToTxt(DC2Converter.DC2ToRTF(DC2, "", 1))
    End Function
    Public Shared Function DC2ToRTF(ByVal RTF As String, ByVal FontName As String, ByVal FontSize As Integer) As String
        ' Dim RTFBox As New RichTextBox
        Dim RTFtext As String
        Dim SetBold(1) As String
        Dim SetUnderl(1) As String
        Dim CurStep As Integer
        Dim LastStep As Integer
        Dim BoldStat As Byte = 0
        Dim ColorString As String
        Dim ClrParams As IRC2RTFColorOffsetSet
        Dim strTmp() As String
        Dim strTmp1 As String
        SetBold(0) = "\b "
        SetBold(1) = "\b0 "
        SetUnderl(0) = "\ul "
        SetUnderl(1) = "\ulnone "
        Const RTF_BOLD = "\'02"
        Const RTF_UNDER = "\'1f"
        Const RTF_HIGHLIGHT = "\highlight"
        Const RTF_COLOR = "\'03"
        Const IRC_Color = "\cf"
        Dim TTR As Boolean = False

        If RTF.IndexOf(Chr(3)) < 0 And RTF.IndexOf(RTF_BOLD) < 0 And RTF.IndexOf(RTF_UNDER) < 0 Then
            Return IRCColToRTFCol(RTF, FontName, FontSize)
        End If

        '--------------
        'Colors std
        '--------------
        CurStep = RTF.IndexOf(Chr(3))
        RTFtext = IRCColToRTFCol(RTF, FontName, FontSize)
        '-------------
        'Bold
        '-------------
        CurStep = RTFtext.IndexOf(RTF_BOLD)
        If CurStep > -1 Then
            Do
                If CurStep = -1 Then Exit Do
                RTFtext = Mid(RTFtext, 1, CurStep) & SetBold(BoldStat) & Mid(RTFtext, CurStep + 5)
                CurStep = RTFtext.IndexOf(RTF_BOLD)
                If BoldStat = 0 Then BoldStat = 1 Else BoldStat = 0
            Loop Until CurStep = -1
        End If
        '-------------
        'Underline
        '-------------
        BoldStat = 0
        CurStep = RTFtext.IndexOf(RTF_UNDER)
        If CurStep > -1 Then
            Do
                If CurStep = -1 Then Exit Do
                RTFtext = Mid(RTFtext, 1, CurStep) & SetUnderl(BoldStat) & Mid(RTFtext, CurStep + 5)
                CurStep = RTFtext.IndexOf(RTF_UNDER)
                If BoldStat = 0 Then BoldStat = 1 Else BoldStat = 0
            Loop Until CurStep = -1
        End If
        ' RTFtext = RTFtext.Replace(Windows.Forms.RichTextBox.DefaultFont.Name, FontName)
        Return RTFtext
    End Function

    Public Shared Function GetRTFCLBL(ByVal RTF As String) As String
        Dim StartCTBL As Integer
        Dim StopCTBL As Integer
        Dim tmp As Integer
        StartCTBL = RTF.IndexOf("{\colortbl")
        If StartCTBL < 0 Then Return RTF
        StopCTBL = RTF.IndexOf("}", StartCTBL)
        Return RTF.Substring(StartCTBL, StopCTBL - StartCTBL + 1)
    End Function

    Public Shared Function TxtToRTF(ByVal Txt As String, ByVal FontName As String, ByVal FontSize As Integer) As String
        Return IRCColToRTFCol(Txt, FontName, FontSize)
    End Function
    Public Shared Function RtfToTxt(ByVal Rtf As String) As String
        Static A As New RichTextBox
        A.Rtf = Rtf
        A.SelectAll()
        Return A.SelectedText
    End Function

    Private Shared Function IRCColToRTFCol(ByVal IRCClr As String, ByVal FontName As String, ByVal FontSize As Integer) As String
        Static RTFBox As New RichTextBox
        'Dim Fix As Byte = 1
        RTFBox.Rtf = "{\rtf1\ansi\ansicpg1251\deff0\deflang1049{\fonttbl{\f0\fnil\fcharset204 Microsoft Sans Serif;}}" & vbCrLf & "{\colortbl ;\red255\green255\blue255;\red0\green0\blue0;\red0\green0\blue128;\red0\green128\blue0;\red255\green0\blue0;\red128\green0\blue0;\red128\green0\blue128;\red128\green128\blue0;\red255\green255\blue0;\red0\green255\blue0;\red0\green128\blue128;\red0\green255\blue255;\red0\green0\blue255;\red255\green0\blue255;\red128\green128\blue128;\red192\green192\blue192;}" & vbCrLf & "\viewkind4\uc1\pard\cf1\f0\fs" & Int(FontSize) * 2 & " 0\cf2 0\cf3 0\cf4 0\cf5 0\cf6 0\cf7 0\cf8 0\cf9 0\cf10 0\cf11 0\cf12 0\cf13 0\cf14 0\cf15 0\cf16 0\cf0\par" & vbCrLf & "}"
        Dim tmp As Integer
        'For tmp = 1 To 16
        '    If IRCClr.StartsWith(tmp.ToString) Then Fix = 0 : Exit For
        'Next

        IRCClr = Chr(3) & "1 " & IRCClr
        IRCClr = IRCCNN_MRTF_PREPROC(IRCClr)
        RTFBox.AppendText(IRCClr)
        'Dim RBTmp As String = RTFBox.Rtf
        RTFBox.Rtf = IRCColToRTFCol_D(RTFBox.Rtf)
        RTFBox.Rtf = IRCCNN_MRTF_POSTPROC(RTFBox.Rtf)
        RTFBox.SelectionStart = 16 + 1
        RTFBox.SelectionLength = RTFBox.TextLength
        RTFBox.SelectionFont = New Font(FontName, FontSize)
        Return RTFBox.SelectedRtf
    End Function
    Private Shared Function IRCColToRTFCol_D(ByVal IRCClr As String) As String
        Dim Tmp As Integer
        Dim Tmp2 As Integer
        Dim t As String
        Dim R() As String
        Dim result As String
        Dim TTC As String
        Dim TTC2 As String

        For Tmp = 15 To 0 Step -1
            For Tmp2 = 15 To 0 Step -1
                TTC = Tmp
                TTC2 = Tmp2
                If Tmp < 10 Then TTC = "0" & Tmp
                If Tmp2 < 10 Then TTC2 = "0" & Tmp2
                ' Замена \'030x,0x
                IRCClr = IRCClr.Replace("\'03" & TTC & "," & TTC2, "\highlight1 \cf1 \cf" & Tmp + 1 & " \highlight" & Tmp2 + 1)
                ' Замена \'03xx,xx
                IRCClr = IRCClr.Replace("\'03" & Tmp & "," & Tmp2, "\highlight1 \cf1 \cf" & Tmp + 1 & " \highlight" & Tmp2 + 1)
            Next

            'Замена \'030x
            IRCClr = IRCClr.Replace("\'03" & TTC, "\cf" & Tmp + 1)
            ' Замена \'03xx
            IRCClr = IRCClr.Replace("\'03" & Tmp, "\cf" & Tmp + 1)

        Next
        IRCClr = IRCClr.Replace("\'03", "\cf2")
        IRCClr = IRCClr.Replace("\'0f", "\cf2 \highlight0")
        Return IRCClr
    End Function
    Public Shared Function IRCCNN_MRTF_PREPROC(ByVal Inp As String) As String
        Static Ret As String
        Ret = Inp
        Ret = Ret.Replace(" ", "|spC")
        Return Ret
    End Function
    Public Shared Function IRCCNN_MRTF_POSTPROC(ByVal Inp As String) As String
        Static Ret As String
        Ret = Inp
        Ret = Ret.Replace("|spC", " ")
        Return Ret
    End Function
    Public Shared Function GetIRCColorCode(ByVal ARGB As Integer) As Integer
        Select Case ARGB
            Case Color.White.ToArgb
                Return 0
            Case Color.FromArgb(0, 1, 0).ToArgb
                Return 1
            Case Color.Navy.ToArgb
                Return 2
            Case Color.Green.ToArgb
                Return 3
            Case Color.Red.ToArgb
                Return 4
            Case Color.Maroon.ToArgb
                Return 5
            Case Color.Purple.ToArgb
                Return 6
            Case Color.Orange.ToArgb
                Return 7
            Case Color.Yellow.ToArgb
                Return 8
            Case Color.Lime.ToArgb
                Return 9
            Case Color.Teal.ToArgb
                Return 10
            Case Color.Aqua.ToArgb
                Return 11
            Case Color.Blue.ToArgb
                Return 12
            Case Color.Fuchsia.ToArgb
                Return 13
            Case Color.Gray.ToArgb
                Return 14
            Case Color.Silver.ToArgb
                Return 15
                '----------------------------------
            Case Color.Black.ToArgb
                Return 99
            Case Else
                Return Color.Black.ToArgb
        End Select
    End Function

    Public Shared Function GetARGBfromIRCCode(ByVal IRCColorCode As Integer) As Integer
        Select Case IRCColorCode
            Case 0
                Return Color.White.ToArgb
            Case 1
                Return Color.FromArgb(0, 1, 0).ToArgb
            Case 2
                Return Color.Navy.ToArgb
            Case 3
                Return Color.Green.ToArgb
            Case 4
                Return Color.Red.ToArgb
            Case 5
                Return Color.Maroon.ToArgb
            Case 6
                Return Color.Purple.ToArgb
            Case 7
                Return Color.Orange.ToArgb
            Case 8
                Return Color.Yellow.ToArgb
            Case 9
                Return Color.Lime.ToArgb
            Case 10
                Return Color.Teal.ToArgb
            Case 11
                Return Color.Aqua.ToArgb
            Case 12
                Return Color.Blue.ToArgb
            Case 13
                Return Color.Fuchsia.ToArgb
            Case 14
                Return Color.Gray.ToArgb
            Case 15
                Return Color.Silver.ToArgb
            Case Else
                Return Color.Black.ToArgb
        End Select
    End Function
End Class
