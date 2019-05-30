Function procC(ByVal oq As String, ByVal onde As Range, Optional absl As Integer, Optional absc As Integer, Optional desl As Integer, Optional desc As Integer)
        
        ' Procura um determinado valor, em determinado lugar, retornando o valor de uma celula relativa, definindo linha x coluna com valores absolutos ou relativos
        ' Exemplo: =procC(D1;B:B;;4;1;)  > Procura o valor contido na célula D1, na coluna B e, retorna o valor de uma linha abaixo, sempre na coluna D
        Set q = onde.Find(oq)
        If absl = 0 Then l = q.Row + desl
        If absc = 0 Then c = q.Column + desc
        
        
        procC = Cells(absl + l, absc + c).Value
        

End Function
