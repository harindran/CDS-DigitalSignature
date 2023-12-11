Imports iTextSharp.text.pdf.parser
Imports System.Globalization

Public Class ClsPDFText

    Public Class RectAndText
        Public Rect As iTextSharp.text.Rectangle
        Public Text As String

        Public Sub New(ByVal rect As iTextSharp.text.Rectangle, ByVal text As String)
            Me.Rect = rect
            Me.Text = text
        End Sub
    End Class

    Public Class MyLocationTextExtractionStrategy
        Inherits LocationTextExtractionStrategy

        Public myPoints As List(Of RectAndText) = New List(Of RectAndText)()
        Public Property TextToSearchFor As String
        Public Property CompareOptions As System.Globalization.CompareOptions

        Public Sub New(ByVal textToSearchFor As String, Optional ByVal compareOptions As System.Globalization.CompareOptions = System.Globalization.CompareOptions.None)
            Me.TextToSearchFor = textToSearchFor
            Me.CompareOptions = compareOptions

        End Sub

        Public Overrides Sub RenderText(ByVal renderInfo As TextRenderInfo)
            Try
                MyBase.RenderText(renderInfo)
                Dim startPosition = System.Globalization.CultureInfo.CurrentCulture.CompareInfo.IndexOf(renderInfo.GetText(), Me.TextToSearchFor, Me.CompareOptions)
                If startPosition < 0 Then
                    Return
                End If
                Dim chars = renderInfo.GetCharacterRenderInfos().Skip(startPosition).Take(Me.TextToSearchFor.Length).ToList()
                Dim firstChar = chars.First()
                Dim lastChar = chars.Last()
                Dim bottomLeft = firstChar.GetDescentLine().GetStartPoint() 'renderInfo.GetDescentLine().GetStartPoint() 
                Dim topRight = lastChar.GetAscentLine().GetEndPoint() 'renderInfo.GetAscentLine.GetEndPoint() 
                Dim rect = New iTextSharp.text.Rectangle(bottomLeft(Vector.I1), bottomLeft(Vector.I2), topRight(Vector.I1), topRight(Vector.I2))
                Me.myPoints.Add(New RectAndText(rect, Me.TextToSearchFor))
            Catch ex As Exception

            End Try

        End Sub


    End Class
End Class
