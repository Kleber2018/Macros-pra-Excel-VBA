Sub IMPRESSÃO()

' MACRO para imprimir 2 formulários A5 frente e verso em uma folha A4
' Está configurado para imprimir 60 formulário frente e verso em 30 folhas de A4
  'Ao imprimir utilizei o PDF Creator para unir todas as paginas e depois imprimir o arquivo,
  'pois a impressoa não trata muito bem arquivos separados

Dim StartNumber, Meio, TempNumber, TempAnswer As Integer
StartNumber = Range("AR3").Value + 1 'para continuar a numeração de onde parou

Do
  
  TempAnswer = MsgBox("Vão ser impressas " & (StartNumber + 60) - StartNumber & _
  " folhas! OK?", vbYesNoCancel, "Confirmar números...")

    Select Case TempAnswer
        Case vbCancel
            Exit Sub
        Case vbYes
            Exit Do
    End Select
Loop
 
  For TempNumber = StartNumber To (StartNumber + 29)
      ActiveWorkbook.Sheets("FRENTE").Range("T3").Value = TempNumber
      ActiveWorkbook.Sheets("FRENTE").Range("AR3").Value = TempNumber + 30
      ActiveWorkbook.Sheets("FRENTE").PrintOut
      ActiveWorkbook.Sheets("VERSO").PrintOut
  Next TempNumber

End Sub
