' ----------------------------------------------------------------------------------------------------------------------
' Script: Teste de Ping por arquivo
' Autor: Thiago Estole
' Data da Criação: 11/08/2020
' ----------------------------------------------------------------------------------------------------------------------
'
on error resume next

Public Sub Grava(comp)
 Const ForAppending = 8

arq_ext = "StatusPing.txt"
 Set fso = CreateObject("Scripting.FileSystemObject")
 Set arq_int = fso.OpenTextFile(arq_ext , ForAppending, true)
 arq_int.write (comp & vbcrlf)
 arq_int.close
End Sub

Public Sub Ping()
data = now()
aMachines = ("8.8.8.8")

Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}")._
        ExecQuery("select * from Win32_PingStatus where address = '"& amachines & "'")
    For Each objStatus in objPing
      If IsNull(objStatus.StatusCode) or objStatus.StatusCode<>0 Then
   result = ("Não foi possível efetuar ping - " & data)
   grava(result)
 else
   result = ("Ping OK - " &  data)
   grava(result)
 end if
 next
End Sub

Do
 While Counter < 2
  Ping()
  Counter = Counter + 1
  wscript.sleep (5000)
 Wend
 Counter = 0
Loop Until Counter = 2
'
' Observações:
' Data        N.      Autor               Informação
' ----------  ------  -------------       ------------------------------------------------------------------------------
' XX/XX/XXXX  XX      Thiago Estole       Informacões.
' ----------  ------  -------------       ------------------------------------------------------------------------------
'
' Informações:
' Linhas  Função
' ------  --------------------------------------------------------------------------------------------------------------
' 12-17   Colocar o local onde irá salvar o log
' 23      Colocar o IP que deseja realizar o PING
' 52      Definir de quanto em quanto tempo será executado, para cinco minutos alterar o valor para 300000
' ------  --------------------------------------------------------------------------------------------------------------
'
' Revisões:
' Data        Versão  Autor               Detalhes
' ----------  ------  -------------       ------------------------------------------------------------------------------
' 11/08/2020  1.0     Thiago Estole       Criação do Documento.
' ----------  ------  -------------       ------------------------------------------------------------------------------


