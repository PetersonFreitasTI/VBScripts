Dim jogador, lista_palavras(30), palavra, resultado, i, score
Dim audio, status_voz, status_pular
Dim exec, path 'variavel para usar funções do sistema windows #shell

call inicializador1

sub inicializador1()
		
	'Inicializa o score
	score = 0
		
	'Chances
	status_voz = 1
	status_pular = 1
	
	'Cria o objeto de executar o shell e volta ao menu
	set exec = createobject("wscript.shell")
	path = exec.CurrentDirectory
	
	'Criando o recurso de voz
	set audio = createobject("SAPI.SPVOICE") 'O Objeto de fala foi criado
	audio.rate = 1
	audio.volume = 100
	
	'Nome jogador
	jogador = inputbox("Digite seu nome: ","Soletrando")
	
	if jogador = "" then
		'Encerra Aplicação'
		call continuar
	else
		'Chama rotina e carrega uma palavra da lista
		call escolhe_palavra
	end if
	
end sub


'Carrega a lista de palavras
sub escolhe_palavra()

	'Lista Palavras
	lista_palavras(0) = "machado"
	lista_palavras(1) = "abacaxi"
	lista_palavras(2) = "zezinho"
	lista_palavras(3) = "melancia"
	lista_palavras(4) = "ameixa"
	lista_palavras(5) = "morango"
	lista_palavras(6) = "esquisito"
	lista_palavras(7) = "barbearia"
	lista_palavras(8) = "corrimão"
	lista_palavras(9) = "estilhaço"
	lista_palavras(10) = "camaleão"
	lista_palavras(11) = "acolchoado"
	lista_palavras(12) = "jabuticaba"
	lista_palavras(13) = "balanço"
	lista_palavras(14) = "sossegado"
	lista_palavras(15) = "abençoado"
	lista_palavras(16) = "marmita"
	lista_palavras(17) = "vizinhança"
	lista_palavras(18) = "televisão"
	lista_palavras(19) = "caminhonete"
	lista_palavras(20) = "parafuso"
	lista_palavras(21) = "guilhotina"
	lista_palavras(22) = "paralisado"
	lista_palavras(23) = "almoço"
	lista_palavras(24) = "oxitona"
	lista_palavras(25) = "jazigo"
	lista_palavras(26) = "entulho"
	lista_palavras(27) = "cambalacho"
	lista_palavras(28) = "xícara"
	lista_palavras(29) = "acentuação"
	
	call inicializador2

end sub

sub inicializador2()
		
	'Escolhendo um nome da lista
	randomize(second(time))
	i = int(rnd * 30)
	palavra = lista_palavras(i)
	
	'Soletra a palavra
	audio.speak (palavra)
	
	'Chamando a função soletrando
	call soletrar
	
end sub

sub soletrar
				
	'Inicio: escolhe um número
	resultado = inputbox("Digite a palavra ouvida: " & vbNewLine &_
		"=============================" & vbNewLine &_
		"Opções:" & vbNewLine &_
		"[O]uvir novamente a palavra (" & status_voz & " chance)" & vbNewLine &_
		"[P]ular a palavra (" & status_pular & " chance)" & vbNewLine &_
		"=============================", "Jogador: " & jogador)
	'Fim: escolha número
	
	select case resultado
		case ""
			'Encerra Aplicação'
			call continuar
		
		case "O","o"
			if status_voz = 1 then
				status_voz = 0
				audio.speak (palavra)
			else
				audio.speak (jogador & ", você não tem mais chances para ouvir a palavra novamente.")
			end if
			
			'retorna ao jogo
			call soletrar
			
		case "P","p"
			if status_pular = 1 then
				status_pular = 0
				call inicializador2
			else
				audio.speak (jogador & ", acabou suas chances de pular uma palavra.")
			end if
			
			'retorna ao jogo
			call soletrar
		
		case else
			if resultado = palavra then
				score = score + 1
				msgbox (jogador & ", você acertou e seu score é " & score & " de 12")				
				call pontuacao
			else
				audio.speak (jogador & ", que pena. Você perdeu.")
				msgbox ("A palavra correta era " & palavra & " e seu score foi " & score & " de 12")
				call continuar
			end if
		end select	
end sub

sub pontuacao()

	'Caso score seja 12 venceu
	if score = 12 then
		audio.speak ("Parabéns " & jogador & ". Você venceu o jogo Soletrando.")
		call continuar
	else
		call inicializador2
	end if

end sub

'Rotina para decidir se retorna ao inicio ou fecha o sistema
sub continuar()

	resp=msgbox("Deseja continuar?",vbQuestion + vbYesNo,"ATENÇÃO")
	if resp = vbYes then
		call inicializador1
	else
		'Encerra Aplicação'
		'wscript.quit
		
		'Chama menu
		exec.run (path & "..\menu.vbs")
	end if

end sub