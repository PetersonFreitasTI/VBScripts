Dim n1,n2,n3,maior
Dim audio, status
Dim exec, path 'variavel para usar funções do sistema windows #shell

call recurso_voz

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
	
	call numero_maior

end sub

'Rotina Entrada
sub numero_maior()
	
	'Guarda numeros para comparar maior
	'Tratamento de erro caso digite caracter
	on error resume next
	n1 = cdbl(inputbox("Digite o primeiro número: ","Valor 1"))
	n2 = cdbl(inputbox("Digite o primeiro número: ","Valor 2"))
	n3 = cdbl(inputbox("Digite o primeiro número: ","Valor 3"))
	'retorno nulo em caso de erro continuando para próxima instrução
	
	if n1 = "" then n1 = 0
	if n2 = "" then n2 = 0
	if n3 = "" then n3 = 0
	
	'retorno nulo em caso de erro continuando para próxima instrução
	on error goto 0
	'Fim da rotina guarda numeros para comparar maior

	'Chama rotina para comparar os numeros
	call comparar_numeros

end sub

'Rotina Processamento e Saída
sub comparar_numeros()
	
	if n1 > n2 and n1 > n3 then
		maior = n1	
	elseif n2 > n3 then
		maior = n2
	else
		maior = n3
	end if

	if status <> 1 then
		msgbox("Entre os números " & n1 & ", " & n2 & " e " & n3 &_
			" o maior número é " & maior &""), vbInformation+vbOKOnly, "Resposta"
	else
		audio.speak("Entre os números " & n1 & ", " & n2 & " e " & n3 &_
			" o maior número é " & maior &"")
	end if
	
	call continuar
	
end sub

'Rotina Continuar
sub continuar()
	
	resp = msgbox("Deseja continuar?",vbQuestion + vbYesNo,"ATENÇÃO")
	if resp = vbYes then
		call recurso_voz
	else
		'Encerra Aplicação'
		'wscript.quit
		
		'Chama menu
		exec.run (path & "..\menu.vbs")
	end if
	
end sub