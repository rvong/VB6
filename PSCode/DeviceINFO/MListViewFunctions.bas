Attribute VB_Name = "MListViewFunctions"
Option Explicit

Const LVM_FIRST As Long = &H1000
Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Const LVSCW_AUTOSIZE As Long = -1
Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Sub AutosizeColumns(ListViewControl As ListView)
' This sub resizes the columns in a ListView: based on a sample code by Barcode (Andy D.)
Dim vColumn As Variant
Dim iColumn As Byte

    ' Resize each column in the ListView
    For Each vColumn In ListViewControl.ColumnHeaders
        ' Lock the ListView area
        LockWindowUpdate ListViewControl.hWnd
        ' Autosize the column
        SendMessage ListViewControl.hWnd, LVM_SETCOLUMNWIDTH, iColumn, LVSCW_AUTOSIZE_USEHEADER
        ' Release the ListView area
        LockWindowUpdate 0
        ' Update the column number
        iColumn = iColumn + 1
    Next
End Sub
