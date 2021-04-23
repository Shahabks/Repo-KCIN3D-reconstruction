using System;
Private Sub SurroundingSub()
    Dim[Class] As[Public]
    Dim Attribute As VisualBasicClass
    ' TODO: Skipped BadDirectiveTrivia    VB_Name = "Module1" "D"c "F"c
    Dim [Declare] As[Public]
    Dim[Sub] As PtrSafe
    Dim [Lib] As Sleep
    Call "kernel32"([ByVal], dwMilliseconds, [As], [Long])
    ' TODO: Skipped BadDirectiveTrivia"F"c
    Dim[Declare] As[Public]
    Dim Sleep As [Sub]
[Lib]
' TODO: Skipped BadDirectiveTrivia
    Call "kernel32.dll"([ByVal], dwMilliseconds, [As], [Long])
        ''' Cannot convert LocalFunctionStatementSyntax, CONVERSION ERROR: Conversion for LocalFunctionStatement not implemented, please report this issue in 'Sub GetmyDatatoAutoCADC()\r\n ' at character 400
''' 
''' 
''' Input:
''' #End If
''' 
''' Sub GetmyDatatoAutoCADC()
''' 
''' "D"c
    Dim acadApp As [Dim]
Dim[Object] As[As]
    Dim acadDoc As [Dim]
Dim[Object] As[As]
    Dim acadCmd As [Dim]
Dim[String] As[As]
    Dim sht As [Dim]
Dim Worksheet As [As]
Dim LastRow As [Dim]
Dim[Long] As[As]
    Dim LastColumn As [Dim]
Dim[Integer] As[As]
    Dim i As [Dim]
Dim[Long] As[As]
    Dim j As [Dim]
Dim[Integer] As[As]
    "S"c
    Dim sht As [Set] = ThisWorkbook.Sheets("Sheet")
    "A"c
    Dim sht As [With]
' TODO: Error SkippedTokensTrivia '.'Dim LastRow As Activate = _.Cells(Rows.Count, "A").End(xlUp).Row
    Dim[With] As[End]
    "C"c
            ''' Cannot convert LocalFunctionStatementSyntax, CONVERSION ERROR: Conversion for LocalFunctionStatement not implemented, please report this issue in 'If LastRow< ' at character 1020
''' 
''' 
''' Input:
'''     If LastRow < 
''' 
    3
    Dim MsgBox As [Then]
    "There are no commands to send!"
    vbCritical
    "No Commands Error"
    sht.Range("C13").Select
    Dim[Sub] As [Exit]
Dim[If] As[End]
    "C"c
    Dim[Error] As [On]
Dim[Next] As[Resume]
    Dim acadApp As [Set] = GetObject(_, "AutoCAD.Application")
    Dim acadApp As [If]
Dim[Nothing] As[Is]
    Dim[Set] As[Then]
    acadApp = CreateObject("AutoCAD.Application")
    acadApp.Visible = [True]
    Dim[If] As[End]
    "C"c
    Dim acadApp As [If]
Dim[Nothing] As[Is]
    Dim MsgBox As [Then]
"Sorry, it was impossible to start AutoCAD!"
    vbCritical
    "AutoCAD Error"
    Dim[Sub] As[Exit]
    Dim[If] As[End]
    "M"c
    acadApp.WindowState = 3
    "3"c
    Dim [Error] As [On]
    [GoTo]
    0
    "I"c
    Dim [Error] As [On]
    Dim [Next] As [Resume]
    Dim acadDoc As [Set] = acadApp.ActiveDocument
    Dim acadDoc As[If]
    Dim [Nothing] As [Is]
    Dim [Set] As [Then]
    acadDoc = acadApp.Documents.Add
    Dim[If] As [End]
    Dim [Error] As [On]
    [GoTo]
    0
    "C"c
    Dim acadDoc As [If]
    ActiveSpace = 0
    [Then]
    "0"c
    acadDoc.ActiveSpace = 1
    "1"c
    Dim [If] As [End]
    Dim sht As [With]
    "L"c
    Dim i As [For] = 3
    Dim LastRow As[To]
    "F"c
    LastColumn = _.Cells(i, _.Columns.Count).End(xlToLeft).Column
    "C"c
    Dim LastColumn As [If]
            ''' Cannot convert BinaryExpressionSyntax, System.NullReferenceException: Object reference not set to an instance of an object.
'''    at ICSharpCode.CodeConverter.VB.NodesVisitor.VisitBinaryExpression(BinaryExpressionSyntax node) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\NodesVisitor.cs:line 1272
'''    at Microsoft.CodeAnalysis.CSharp.Syntax.BinaryExpressionSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
'''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
'''    at ICSharpCode.CodeConverter.VB.CommentConvertingVisitorWrapper`1.Accept(SyntaxNode csNode, Boolean addSourceMapping) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\CommentConvertingVisitorWrapper.cs:line 26
''' 
''' Input:
''' > 2 
''' 
    [Then]
    "C"c
    acadCmd = ""
    Dim j As[For] = 1
    Dim LastColumn As[To]
    Dim [Not] As [If]
    IsEmpty(Cells(i, j).Value)
    Dim acadCmd As [Then] = acadCmd And _.Cells(i, j).Value And vbCr
    Dim [If] As[End]
    Dim j As [Next]
"C"c
                ''' Cannot convert LocalFunctionStatementSyntax, CONVERSION ERROR: Conversion for LocalFunctionStatement not implemented, please report this issue in 'If Val(acadApp.Version) ' at character 3039
''' 
''' 
''' Input:
'''                 If Val(acadApp.Version) 
''' 
            ''' Cannot convert BinaryExpressionSyntax, System.NullReferenceException: Object reference not set to an instance of an object.
'''    at ICSharpCode.CodeConverter.VB.NodesVisitor.VisitBinaryExpression(BinaryExpressionSyntax node) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\NodesVisitor.cs:line 1272
'''    at Microsoft.CodeAnalysis.CSharp.Syntax.BinaryExpressionSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
'''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
'''    at ICSharpCode.CodeConverter.VB.CommentConvertingVisitorWrapper`1.Accept(SyntaxNode csNode, Boolean addSourceMapping) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\CommentConvertingVisitorWrapper.cs:line 26
''' 
''' Input:
''' < 20 
''' 
    [Then]
    "P"c "c"c
    vbCr
    " "c "I"c
    Chr
    27
    " "c
            ''' Cannot convert LocalFunctionStatementSyntax, CONVERSION ERROR: Conversion for LocalFunctionStatement not implemented, please report this issue in 'If InStr(' at character 3445
''' 
''' 
''' Input:
'''                     If InStr(
''' 
    1
    acadCmd
    "SELECT"
    vbTextCompare
                ''' Cannot convert BinaryExpressionSyntax, System.NullReferenceException: Object reference not set to an instance of an object.
'''    at ICSharpCode.CodeConverter.VB.NodesVisitor.VisitBinaryExpression(BinaryExpressionSyntax node) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\NodesVisitor.cs:line 1272
'''    at Microsoft.CodeAnalysis.CSharp.Syntax.BinaryExpressionSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
'''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
'''    at ICSharpCode.CodeConverter.VB.CommentConvertingVisitorWrapper`1.Accept(SyntaxNode csNode, Boolean addSourceMapping) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\CommentConvertingVisitorWrapper.cs:line 26
''' 
''' Input:
''' > 0 
''' 
            ''' Cannot convert LocalFunctionStatementSyntax, CONVERSION ERROR: Conversion for LocalFunctionStatement not implemented, please report this issue in 'Or InStr(' at character 3495
''' 
''' 
''' Input:
''' Or InStr(
''' 
    1
    acadCmd
    "AI_SELALL"
    vbTextCompare
    Dim _ As [Then]
Dim acadCmd As acadDoc.SendCommand
            ''' Cannot convert PrefixUnaryExpressionSyntax, System.NotSupportedException: AddressOfExpression is not supported!
'''    at ICSharpCode.CodeConverter.VB.SyntaxKindExtensions.ConvertToken(SyntaxKind t, TokenContext context)
'''    at ICSharpCode.CodeConverter.VB.NodesVisitor.VisitPrefixUnaryExpression(PrefixUnaryExpressionSyntax node) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\NodesVisitor.cs:line 954
'''    at Microsoft.CodeAnalysis.CSharp.Syntax.PrefixUnaryExpressionSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
'''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
'''    at ICSharpCode.CodeConverter.VB.CommentConvertingVisitorWrapper`1.Accept(SyntaxNode csNode, Boolean addSourceMapping) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\CommentConvertingVisitorWrapper.cs:line 26
''' 
''' Input:
''' & vbCr
''' 
''' 
    Dim _ As [Else]
Dim acadCmd As acadDoc.SendCommand
            ''' Cannot convert PrefixUnaryExpressionSyntax, System.NotSupportedException: AddressOfExpression is not supported!
'''    at ICSharpCode.CodeConverter.VB.SyntaxKindExtensions.ConvertToken(SyntaxKind t, TokenContext context)
'''    at ICSharpCode.CodeConverter.VB.NodesVisitor.VisitPrefixUnaryExpression(PrefixUnaryExpressionSyntax node) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\NodesVisitor.cs:line 954
'''    at Microsoft.CodeAnalysis.CSharp.Syntax.PrefixUnaryExpressionSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
'''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
'''    at ICSharpCode.CodeConverter.VB.CommentConvertingVisitorWrapper`1.Accept(SyntaxNode csNode, Boolean addSourceMapping) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\CommentConvertingVisitorWrapper.cs:line 26
''' 
''' Input:
''' & Chr
''' 
    27
    Dim [If] As[End]
    [Else]
    "I"c "c"c
    vbCr
    " "c
    Dim acadCmd As acadDoc.SendCommand
            ''' Cannot convert PrefixUnaryExpressionSyntax, System.NotSupportedException: AddressOfExpression is not supported!
'''    at ICSharpCode.CodeConverter.VB.SyntaxKindExtensions.ConvertToken(SyntaxKind t, TokenContext context)
'''    at ICSharpCode.CodeConverter.VB.NodesVisitor.VisitPrefixUnaryExpression(PrefixUnaryExpressionSyntax node) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\NodesVisitor.cs:line 954
'''    at Microsoft.CodeAnalysis.CSharp.Syntax.PrefixUnaryExpressionSyntax.Accept[TResult](CSharpSyntaxVisitor`1 visitor)
'''    at Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
'''    at ICSharpCode.CodeConverter.VB.CommentConvertingVisitorWrapper`1.Accept(SyntaxNode csNode, Boolean addSourceMapping) in D:\GitWorkspace\CodeConverter\CodeConverter\VB\CommentConvertingVisitorWrapper.cs:line 26
''' 
''' Input:
''' & vbCr
''' 
''' 
    Dim [If] As[End]
    Dim[If] As[End]
    "P"c "H"c
    Sleep
    20
    Dim i As [Next]
Dim[With] As[End]
    "I"c
    MsgBox
    "The user commands were successfully sent to AutoCAD!"
    vbInformation
    "Done"
    Dim [End] As[End]
    [Class]
End Sub