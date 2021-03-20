Dim valor, resp
Dim audio, status
Dim exec, path 'variavel para usar fun��es do sistema windows #shell

call recurso_voz

'Inicianco a rotina de Recurso de Voz
sub recurso_voz()
	
	'Cria o objeto de executar o shell e volta ao menu
	set exec = createobject("wscript.shell")
	path = exec.CurrentDirectory
	
	resp = msgbox("Deseja ativar o recurso de voz?", vbQuestion+vbYesNo,"ATEN��O")
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
		
	'Chamando a fun��o n�mero
	call numero

end sub

'Inicianco a rotina fun��o n�mero
sub numero()
	
	
	'Inpur do n�mero inteiro
	'Tratamento de erro caso digite caracter
	on error resume next
	valor = cint(inputbox("Digite um n�mero: ", "Numeros Inteiros"))
	
	if valor = "" then valor = 0
	
	'retorno nulo em caso de erro continuando para pr�xima instru��o
	on error goto 0
	'fim da rotina input numero inteiro
	
	'Condicional para analisar se o retorno ser� voz ou msgbox
	if status <> 1 then
		msgbox("O valor escolhido �: " & valor & "" + vbNewLine &_
			"Seu antecessor �: " & valor - 1 & "" + vbNewLine &_
			"Seu sucessor �: " & valor + 1 & ""), + vbInformation, "N�meros Inteiros"
	else
		audio.speak("O valor escolhido �: " & valor & "" + vbNewLine &_
			"Seu antecessor �: " & valor - 1 & "" + vbNewLine &_
			"Seu sucessor �: " & valor + 1 & "")
	end if
		
	'Rotina continuar
	call continuar

end sub

'Rotina para decidir se retorna ao inicio ou fecha o sistema
sub continuar()

	resp=msgbox("Deseja continuar?",vbQuestion + vbYesNo,"ATEN��O")
	if resp = vbYes then
		call recurso_voz
	else
		'Encerra Aplica��o'
		'wscript.quit
		
		'Chama menu
		exec.run (path & "..\menu.vbs")
	end if

end sub