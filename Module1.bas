Attribute VB_Name = "Module1"
Declare Function GetShortPathName Lib "kernel32" Alias _
  "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal _
  lpszShortPath As String, ByVal cchBuffer As Long) As Long


