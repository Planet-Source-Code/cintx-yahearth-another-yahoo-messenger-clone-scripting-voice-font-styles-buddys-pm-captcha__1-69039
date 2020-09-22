Attribute VB_Name = "mod_Options"
Type Opt
    blScripting As Boolean
    blSpamFilter As Boolean
    blSpamAndUser As Boolean
    intScriptTimeOut As Integer
    blBlockDuplicates As Boolean
    blDisableFontStyle As Boolean
    blReconnect As Boolean
    blUseBG As Boolean
End Type
Public Options As Opt

Public Function SaveOption(strSection As String, strValue As String)
    SaveSetting "YahEarth", "Settings", strSection, strValue
End Function

Public Function GetOption(strSection As String, Optional strDefault As String = "") As String
    GetOption = GetSetting("YahEarth", "Settings", strSection, strDefault)
End Function
