<div align="center">

## Active Directory Search Functions


</div>

### Description

There are 4 different function that allow you to search your companies active directory in different ways. These function will allow you to search active directory by user or by group to determine permissions. I am currently using these in my enterprise applications so that I can set up security at a very granular level. Down to a specific control if i want to.
 
### More Info
 
You will need to have active directory set up at your company. You will see a few place that have the following words "{YOUR DC HERE}" that you will need to replace with your domain controller name.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Paul Gorman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/paul-gorman.md)
**Level**          |Advanced
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/paul-gorman-active-directory-search-functions__1-36322/archive/master.zip)





### Source Code

```
Option Explicit
Public Enum Enum_adscAccessType
  adscDenyedAccess = 0
  adscDataReader = 1
  adscDataWriter = 2
End Enum
Public Function AllowAccess(LoginID As String, Group As String) As Boolean
Dim oCN As ADODB.Connection, oCM As ADODB.Command, oRS As ADODB.Recordset, oField As ADODB.Field
Dim oUser As IADs, oParent As IADs, oGroup As IADs
Dim oPropList As IADsPropertyList, oPropEntry As IADsPropertyEntry, oPropVal As IADsPropertyValue
Dim sPath As String, v As Variant, i As Variant
'This function checks a specific users rights via their login and what ever group you pass in.
'You will need to replace the {YOUR DC HERE} with your own domain controller to active directory.
  Set oCN = New ADODB.Connection
  Set oCM = New ADODB.Command
  Set oRS = New ADODB.Recordset
  oCN.Provider = "ADsDSOObject"
  oCN.Open
  Set oCM.ActiveConnection = oCN
  oCM.CommandText = "SELECT AdsPath FROM 'LDAP://OU=Branches,OU=Corp,DC={YOUR DC HERE},DC=com' " & _
           "WHERE objectCategory='person' AND cn='" & LoginID & "'"
  oCM.Properties("searchscope") = 2
  Set oRS = oCM.Execute
  If Not oRS.EOF Then
    Set oUser = GetObject(oRS("AdsPath").Value)
    oUser.GetInfo
    Set oParent = GetObject(oUser.Parent)
    Set oParent = GetObject(oParent.Parent)
    For i = 0 To oUser.PropertyCount - 1
      Set oPropEntry = oUser.Item(i)
      If oPropEntry.Name = "memberOf" Then
        For Each v In oPropEntry.Values
          Set oPropVal = v
          sPath = oPropVal.DNString
          Set oGroup = GetObject("LDAP://" & sPath)
          If oGroup.Name = "CN=" & Group Then
            AllowAccess = True
            GoTo ShutDown
          End If
          Set oGroup = Nothing
        Next
      End If
      oUser.Next
    Next
  End If
  AllowAccess = False
ShutDown:
Set oCN = Nothing
Set oRS = Nothing
Set oCM = Nothing
Set oField = Nothing
Set oUser = Nothing
Set oParent = Nothing
Set oGroup = Nothing
Set oPropList = Nothing
Set oPropEntry = Nothing
Set oPropVal = Nothing
Set v = Nothing
End Function
Public Function ADSCAllowAccessByGroup(Group As String, UserName As String) As Boolean
On Error Resume Next
Dim oGroup As ActiveDs.IADsGroup
Dim oUser As ActiveDs.IADsUser
'This function checks whether or not a user is in a specific group. It will return a true or false
'You will need to replace the {YOUR DC HERE} with your own domain controller to active directory.
  Set oGroup = GetObject("WinNT://{YOUR DC HERE}.com/" & Group)
  If oGroup Is Nothing Then
    ADSCAllowAccessByGroup = False
    Exit Function
  End If
  For Each oUser In oGroup.Members
    Debug.Print oUser.Name
    If UCase(oUser.Name) = UCase(UserName) Then
      ADSCAllowAccessByGroup = True
      Exit Function
    End If
  Next
  ADSCAllowAccessByGroup = False
End Function
Public Function ADSCAllowAccessByUser(UserName As String, Group As String) As Boolean
On Error Resume Next
Dim oGroup As ActiveDs.IADsGroup
Dim oUser As ActiveDs.IADsUser
  Set oUser = GetObject("WinNT://{YOUR DC HERE}.com/" & UCase(UserName) & ",user")
  If oUser Is Nothing Then
    ADSCAllowAccessByUser = False
    Exit Function
  End If
  For Each oGroup In oUser.Groups
    If UCase(oGroup.Name) = UCase(Group) Then
      ADSCAllowAccessByUser = True
      Exit Function
    End If
  Next
End Function
Public Function ADSCAccessType(Location As String, UserName As String, Module As String, AppName As String) As Enum_adscAccessType
On Error Resume Next
Dim oGroup As ActiveDs.IADsGroup
Dim oUser As ActiveDs.IADsUser
'This function assumes that you already have 2 types of groups set up. One that has DataReader at the end and another
'that has datawriter at the end. It also assumes that you have set up your group name in the following
'order: Location_AppName & Module & DataReader/DataWriter.
'You can change this to fit your needs. The main part is the first line of code that sets the oUser
'You will need to replace the {YOUR DC HERE} with your own domain controller to active directory.
  Set oUser = GetObject("WinNT://{YOUR DC HERE}.com/" & UCase(UserName) & ",user")
  If oUser Is Nothing Then
    ADSCAccessType = adscDenyedAccess
    Exit Function
  End If
  For Each oGroup In oUser.Groups
    Select Case oGroup.Name
      Case Location & "_" & AppName & Module & "DataReader"
        ADSCAccessType = adscDataReader
        Exit Function
      Case Location & "_" & AppName & Module & "DataWriter"
        ADSCAccessType = adscDataWriter
        Exit Function
    End Select
  Next
  ADSCAccessType = adscDenyedAccess
End Function
```

