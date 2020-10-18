Attribute VB_Name = "m_paintOnOffTest"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private wb As Workbook
Private ws As Worksheet

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    
    Application.ScreenUpdating = False
    
    Workbooks.Add
    Set wb = ActiveWorkbook
    Set ws = ActiveSheet
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
    wb.Close SaveChanges:=False
    
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'this method runs before every test in the module.
    ws.Cells.Interior.Color = xlNone
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub
'@TestMethod("Uncategorized")
Private Sub TestMethod1()                        'TODO Rename test
    On Error GoTo TestFail
    
    'Arrange:

    'Act:
    Dim result As Variant
    result = fn_paintOnOff(ws.Range("B2:C3"))

    'Assert:
    Call Assert.AreEqual(True, result(0)) ' ê≥èÌèIóπ
    Call Assert.AreEqual(xlNone, ws.Range("A1").Interior.ColorIndex)
    Call Assert.AreEqual(xlNone, ws.Range("A2").Interior.ColorIndex)
    Call Assert.AreEqual(xlNone, ws.Range("A3").Interior.ColorIndex)
    Call Assert.AreEqual(xlNone, ws.Range("A4").Interior.ColorIndex)
    Call Assert.AreEqual(xlNone, ws.Range("B1").Interior.ColorIndex)
    Dim wkRGB As Variant: wkRGB = RGB(255, 255, 0)
    Dim wkINDEX As Variant: wkINDEX = ws.Range("B2").Interior.Color

'    Call Assert.AreEqual(RGB(255, 255, 0), ws.Range("B2").Interior.Color)
'    Call Assert.AreEqual(RGB(255, 255, 0), ws.Range("B3").Interior.Color)
    Call Assert.AreEqual(xlNone, ws.Range("B4").Interior.ColorIndex)
    Call Assert.AreEqual(xlNone, ws.Range("C1").Interior.ColorIndex)
'    Call Assert.AreEqual(RGB(255, 255, 0), ws.Range("C2").Interior.Color)
'    Call Assert.AreEqual(RGB(255, 255, 0), ws.Range("C3").Interior.Color)
    Call Assert.AreEqual(xlNone, ws.Range("C4").Interior.ColorIndex)
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

