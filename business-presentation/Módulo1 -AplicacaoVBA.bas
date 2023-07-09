Attribute VB_Name = "M�dulo1"
Sub CriarApresentacao()
    
    ' Definir o caminho e nome do arquivo da apresenta��o do PowerPoint
    Dim caminho As String
    caminho = "C:\Users\usuario\Downloads\Apresenta��o-Empresarial\apresenta��o_2023.pptx"
    
    ' Criar uma nova apresenta��o do PowerPoint
    Dim pptApp As Object
    Dim pptPresentation As Object
    Set pptApp = CreateObject("PowerPoint.Application")
    Set pptPresentation = pptApp.Presentations.Add
    
    ' Adicionar slides � apresenta��o
    pptPresentation.Slides.Add 1, 12 ' Slide 1 (capa)
    pptPresentation.Slides.Add 2, 12 ' Slide 2 (t�picos)
    pptPresentation.Slides.Add 3, 12
    pptPresentation.Slides.Add 4, 12
    pptPresentation.Slides.Add 5, 12
    pptPresentation.Slides.Add 6, 12
    pptPresentation.Slides.Add 7, 12
    
    'O valor 12 usado no par�metro "layout" corresponde ao "ppLayoutBlank" - o que criar� slides em branco
    'documenta��o informativa:https://learn.microsoft.com/pt-br/office/vba/api/powerpoint.ppslidelayout
    
    ' Copiar e colar os gr�ficos em cada slide correspondente
    Dim ws As Worksheet
    Dim chart As ChartObject
    Dim slideIndex As Integer
    slideIndex = 3 ' Come�ar no Slide 3
    
     ' Aba A
    Set ws = ThisWorkbook.Sheets("A")
    Set chart1 = ws.ChartObjects("Gr�fico 1")
    Set chart2 = ws.ChartObjects("Gr�fico 2")
    chart1.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    
    ' Colar o primeiro gr�fico no slide
    pptPresentation.Slides(slideIndex).Shapes.PasteSpecial
    With pptPresentation.Slides(slideIndex).Shapes(pptPresentation.Slides(slideIndex).Shapes.Count)
        .Left = 8.86 * 28.35 ' 8,86 cm
        .Top = 0.51 * 28.35 ' 0,51 cm
        .Height = 8.84 * 28.35 ' 8,84 cm
        .Width = 14.92 * 28.35 ' 14,92 cm
    End With
    
    chart2.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    
    ' Colar o segundo gr�fico no slide
    pptPresentation.Slides(slideIndex).Shapes.PasteSpecial
    With pptPresentation.Slides(slideIndex).Shapes(pptPresentation.Slides(slideIndex).Shapes.Count)
        .Left = 8.86 * 28.35 '8,86 cm Ajuste o valor para posicionar o segundo gr�fico
        .Top = 10.06 * 28.35 ' 10,06 cm Ajuste o valor para posicionar o segundo gr�fico
        .Height = 8.84 * 28.35 ' 8,84 cm convertendo pixels para cent�metros
        .Width = 14.92 * 28.35 ' 14,92 cm
    End With
    
    'explicando o *28.35
    'A unidade de medida no PowerPoint � o ponto (point), onde 1 ponto � igual a 1/72 polegada.
    'Para converter cent�metros em pontos, � comum usar uma aproxima��o,
    'considerando que 1 polegada � igual a 2,54 cent�metros. Portanto,
    '1 cent�metro � aproximadamente igual a 28.35 pontos.
    
    slideIndex = slideIndex + 1
    
    ' Aba B
    Set ws = ThisWorkbook.Sheets("B")
    Set chart3 = ws.ChartObjects("Gr�fico 3")
    Set chart4 = ws.ChartObjects("Gr�fico 4")
    chart3.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    
    pptPresentation.Slides(slideIndex).Shapes.PasteSpecial
    With pptPresentation.Slides(slideIndex).Shapes(pptPresentation.Slides(slideIndex).Shapes.Count)
        .Left = 8.86 * 28.35 ' 8,86 cm
        .Top = 0.51 * 28.35 ' 0,51 cm
        .Height = 8.84 * 28.35 ' 8,84 cm
        .Width = 14.92 * 28.35 ' 14,92 cm
    End With
    
    'colar o segundo gr�fico
    chart4.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    
    pptPresentation.Slides(slideIndex).Shapes.PasteSpecial
    With pptPresentation.Slides(slideIndex).Shapes(pptPresentation.Slides(slideIndex).Shapes.Count)
        .Left = 8.86 * 28.35
        .Top = 10.06 * 28.35
        .Height = 8.84 * 28.35
        .Width = 14.92 * 28.35
    End With
    slideIndex = slideIndex + 1
    
    ' Aba C
    Set ws = ThisWorkbook.Sheets("C")
    Set chart5 = ws.ChartObjects("Gr�fico 5")
    Set chart6 = ws.ChartObjects("Gr�fico 6")
    chart5.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    
    pptPresentation.Slides(slideIndex).Shapes.PasteSpecial
    With pptPresentation.Slides(slideIndex).Shapes(pptPresentation.Slides(slideIndex).Shapes.Count)
        .Left = 8.86 * 28.35 ' 8,86 cm
        .Top = 0.51 * 28.35 ' 0,51 cm
        .Height = 8.84 * 28.35 ' 8,84 cm
        .Width = 14.92 * 28.35 ' 14,92 cm
    End With
    
    'colar o segundo gr�fico no slide
    chart6.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    pptPresentation.Slides(slideIndex).Shapes.PasteSpecial
    With pptPresentation.Slides(slideIndex).Shapes(pptPresentation.Slides(slideIndex).Shapes.Count)
        .Left = 8.86 * 28.35
        .Top = 10.06 * 28.35
        .Height = 8.84 * 28.35
        .Width = 14.92 * 28.35
    End With
    slideIndex = slideIndex + 1
    
    ' Aba D
    Set ws = ThisWorkbook.Sheets("D")
    Set chart7 = ws.ChartObjects("Gr�fico 7")
    Set chart8 = ws.ChartObjects("Gr�fico 8")
    chart7.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    
    pptPresentation.Slides(slideIndex).Shapes.PasteSpecial
    With pptPresentation.Slides(slideIndex).Shapes(pptPresentation.Slides(slideIndex).Shapes.Count)
        .Left = 8.86 * 28.35 ' 8,86 cm
        .Top = 0.51 * 28.35 ' 0,51 cm
        .Height = 8.84 * 28.35 ' 8,84 cm
        .Width = 14.92 * 28.35 ' 14,92 cm
    End With
    
    'colar o segundo gr�fico
    chart8.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    pptPresentation.Slides(slideIndex).Shapes.PasteSpecial
    With pptPresentation.Slides(slideIndex).Shapes(pptPresentation.Slides(slideIndex).Shapes.Count)
        .Left = 8.86 * 28.35
        .Top = 10.06 * 28.35
        .Height = 8.84 * 28.35
        .Width = 14.92 * 28.35
    End With
    slideIndex = slideIndex + 1
    
    ' Aba E
    Set ws = ThisWorkbook.Sheets("E")
    Set chart9 = ws.ChartObjects("Gr�fico 9")
    Set chart10 = ws.ChartObjects("Gr�fico 10")
    chart9.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    
    'colar o primeiro gr�fico
    pptPresentation.Slides(slideIndex).Shapes.PasteSpecial
    With pptPresentation.Slides(slideIndex).Shapes(pptPresentation.Slides(slideIndex).Shapes.Count)
        .Left = 8.86 * 28.35 ' 8,86 cm
        .Top = 0.51 * 28.35 ' 0,51 cm
        .Height = 8.84 * 28.35 ' 8,84 cm
        .Width = 14.92 * 28.35 ' 14,92 cm
    End With
    
    'colar o segundo gr�fico no slide
    chart10.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    pptPresentation.Slides(slideIndex).Shapes.PasteSpecial
    With pptPresentation.Slides(slideIndex).Shapes(pptPresentation.Slides(slideIndex).Shapes.Count)
        .Left = 8.86 * 28.35
        .Top = 10.06 * 28.35
        .Height = 8.84 * 28.35
        .Width = 14.92 * 28.35
    End With
    slideIndex = slideIndex + 1
End Sub

