Attribute VB_Name = "Base64Helpers"
'@Folder "MVVM.SortOrder.Helpers"
Option Explicit

Public Function StringtoBase64(ByVal StringValue As String) As String
    Dim ByteArray() As Byte
    ByteArray = StrConv(StringValue, vbFromUnicode)
    
    Dim XMLObject As Object
    Set XMLObject = CreateObject("MSXML2.DOMDocument")
    
    Dim XMLElement As Object
    Set XMLElement = XMLObject.CreateElement("Base64")
    With XMLElement
        .DataType = "bin.base64"
        .NodeTypedValue = ByteArray
    End With
    
    StringtoBase64 = Replace(XMLElement.text, vbLf, vbNullString)
    
    Set XMLElement = Nothing
    Set XMLObject = Nothing
End Function

Public Function Base64toString(ByVal Base64Value As String) As String
    Dim XMLObject As Object
    Set XMLObject = CreateObject("MSXML2.DOMDocument")
    
    Dim XMLElement As Object
    Set XMLElement = XMLObject.CreateElement("Base64")
    With XMLElement
        .DataType = "bin.base64"
        .text = Base64Value
    End With
    
    Base64toString = StrConv(XMLElement.NodeTypedValue, vbUnicode)

    Set XMLElement = Nothing
    Set XMLObject = Nothing
End Function

