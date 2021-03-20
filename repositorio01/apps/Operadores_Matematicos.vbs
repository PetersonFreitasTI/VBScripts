Dim n(3), op, resposta, resultado, pontos, acertou 'variavel para calculo matem�tico
Dim resp 'Variavel para escolhas msgbox
Dim exec, path 'variavel para usar fun��es do sistema windows #shell

call inicio

sub inicio
	
	'inicia a pontua��o com zero
	pontos = 0
	
	'Cria o objeto de executar o shell e volta ao menu
	set exec = createobject("wscript.shell")
	path = exec.CurrentDirectory
	
	call randomizacao
	
end sub

'Rotina de escolha de numeros aleat�rios
sub randomizacao()

	randomize(second(time)) 'Default - sortear de maneira ramdomica conforme relogio do SO
	n(0) = int(rnd * 10) + 1 'Escolhendo o primeiro n�mero 
	n(1) = int(rnd * 3) + 1
	n(2) = int(rnd * 10) + 1
		
	select case	n(1)
		case 1
			op = n(0) & " + " & n(2)
			resultado = n(0) + n(2)
		case 2
			op = n(0) & " - " & n(2)
			resultado = n(0) - n(2)
		case 3
			op = n(0) & " * " & n(2)
			resultado = n(0) * n(2)
		'case 4
		'	op = n(0) & " / " & n(2)
		'	resultado = n(0) / n(2)
		case else
			msgbox "Erro na randomiza��o."
			
		end select
	
	'Chama rotina do jogo
	call jogo
	
end sub


'Rotina jogo
sub jogo()
	
	'Inicio: escolhe um n�mero
	'Tratamento de erro caso digite caracter
	on error resume next
	resposta = cint(inputbox(vbNewLine &_
		"=============================" & vbNewLine &_
		"Acerte o resultado Matem�tico" & vbNewLine &_
		"=============================" & vbNewLine &_
		"        " & op & " = ??" & vbNewLine &_
		"=============================", "Advinha��o"))
	
	'retorno nulo em caso de erro continuando para pr�xima instru��o
	on error goto 0
	'Fim: escolha n�mero
	
	'Inicio: Verifica se acertou a opera��o matem�tica
	if resposta = resultado then
		pontos = pontos + 1
		msgbox ("Parab�ns!" & vbNewLine &_
			"Voc� tem "  & pontos & " ponto(s)."), vbInformation + vbOKOnly, "Pontua��o"
		acertou = true
	else
		msgbox ("Que pena, voc� errou." & vbNewLine &_
			"Seu placar � de " & pontos & " ponto(s)."), vbError + vbOKOnly, "Pontua��o"
		acertou = false
	end if
	'Fim da verifica��o do acerto
	
	'Verifica se acertou e volta o jogo
	if acertou then
		call randomizacao
	else
		call continuar
	end if
	
end sub

'Rotina Continuar
sub continuar()
	
	resp=msgbox("Deseja continuar?",vbQuestion + vbYesNo,"ATEN��O")
	if resp = vbYes then
		call inicio
	else
		'Encerra Aplica��o'
		'wscript.quit
		
		'Chama menu
		exec.run (path & "..\menu.vbs")
	end if
	
end sub