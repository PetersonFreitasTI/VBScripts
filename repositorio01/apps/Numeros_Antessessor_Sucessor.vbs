Dim valor, resp
Dim audio, status
Dim exec, path 'variavel para usar funções do sistema windows #shell

call recurso_voz

'Inicianco a rotina de Recurso de Voz
sub recurso_voz()
	
	'Cria o objeto de executar o shell e volta ao menu
	set exec = createobject("wscript.shell")
	path = exec.CurrentDirectory
	
	resp = msgbox("Deseja ativar o recurso de voz?", vbQuestion+vbYesNo,"ATENÇÃO")
	if resp = vbYes then
		'Criando o recurso de voz
		set audio = createobject("SAPI.SPVOICE") 'O Objeto de fala foi criado
		audio.rate = 1
		audio.volume = 100
		status = 1

		msgbox "Recurso de voz ativado",vbInformation,"Recurso de Voz"
	else
		status = 2
		msgbox "Recurso de voz desativado",vbInformation,"Recurso de Voz"
	end if
		
	'Chamando a função número
	call numero

end sub

'Inicianco a rotina função número
sub numero()
	
	
	'Inpur do número inteiro
	'Tratamento de erro caso digite caracter
	on error resume next
	valor = cint(inputbox("Digite um número: ", "Numeros Inteiros"))
	
	if valor = "" then valor = 0
	
	'retorno nulo em caso de erro continuando para próxima instrução
	on error goto 0
	'fim da rotina input numero inteiro
	
	'Condicional para analisar se o retorno será voz ou msgbox
	if status <> 1 then
		msgbox("O valor escolhido é: " & valor & "" + vbNewLine &_
			"Seu antecessor é: " & valor - 1 & "" + vbNewLine &_
			"Seu sucessor é: " & valor + 1 & ""), + vbInformation, "Números Inteiros"
	else
		audio.speak("O valor escolhido é: " & valor & "" + vbNewLine &_
			"Seu antecessor é: " & valor - 1 & "" + vbNewLine &_
			"Seu sucessor é: " & valor + 1 & "")
	end if
		
	'Rotina continuar
	call continuar

end sub

'Rotina para decidir se retorna ao inicio ou fecha o sistema
sub continuar()

	resp=msgbox("Deseja continuar?",vbQuestion + vbYesNo,"ATENÇÃO")
	if resp = vbYes then
		call recurso_voz
	else
		'Encerra Aplicação'
		'wscript.quit
		
		'Chama menu
		exec.run (path & "..\menu.vbs")
	end if

end sub