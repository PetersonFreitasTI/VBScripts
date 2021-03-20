Dim lado, area, perimetro 'variavel para calculo matemático
Dim resp 'Variavel para escolhas msgbox
Dim audio, status 'variavel para comandos de voz
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
	
	call quadrado

end sub


'Rotina Quadrado
sub quadrado()
	
	'Inicia calculo de area e perimetro
	'Tratamento de erro caso digite caracter
	on error resume next
	lado = cdbl(inputbox("Entre com o tamanho do lado do quadrado: ","Lado do Quadrado"))
	area = (lado * lado)
	perimetro = (lado * 4)
	
	if status <> 1 then
		msgbox("A área do quadrado é " & area & "" + vbNewLine &_
			"O Perímetro do quadrado é " & perimetro & ""), vbInformation+vbOKOnly, "Perímetro do Quadrado"
	else
		audio.speak("A área do quadrado é " & area & "" + vbNewLine &_
			"O Perímetro do quadrado é " & perimetro & "")
	end if
	
	'retorno nulo em caso de erro continuando para próxima instrução
	on error goto 0
	'fim do calculo de area e perimetro
	
	'Chama rotina para continuar algoritmo
	call continuar
	
end sub

'Rotina Continuar
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