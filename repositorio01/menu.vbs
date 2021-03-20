Dim op 'variavel para definir a operação
Dim resp 'variavel para definir a resposta do msgbox
Dim exec, path 'variavel para usar funções do sistema windows #shell

call carregar_menu

sub carregar_menu()
	
	set exec = createobject("wscript.shell") 'Cria o objeto shell para chamar arquivos externos
	path = exec.CurrentDirectory
	
	op = inputbox("==========================" + vbNewLine &_
		"[N] Numero Antecessor e Sucessor" + vbNewLine &_
		"[A] Area do Quadrado e Perímetro" + vbNewLine &_
		"[M] Numero Maior" + vbNewLine &_
		"==========================" + vbNewLine &_
		"[O] Jogo Operadores Matemáticos" + vbNewLine &_
		"[S] Jogo Soletrando" + vbNewLine &_
		"==========================" + vbNewLine &_
		"[F] Finalizar o Script" + vbNewLine &_
		"==========================","Escolha uma Opçãoo:")

	select case	op
		case "N","n":
			exec.run (path & "\apps\Numeros_Antessessor_Sucessor.vbs")
		case "A","a":
			exec.run (path & "\apps\Area_Perimetro_Quadrado.vbs")
		case "M","m":
			exec.run (path & "\apps\Numero_Maior.vbs")
		case "O","o":
			exec.run (path & "\apps\Operadores_Matematicos.vbs")
		case "S","s":
			exec.run (path & "\apps\Soletrando.vbs")
		case "F","f":
			resp=msgbox("Deseja realmente sair?",vbQuestion + vbYesNo,"ATENÇÃO")
				if resp = vbYes then
					wscript.quit 'Encerra Aplicação'
				else
					call carregar_menu
				end if
		case else
			msgbox "Opção inválida!",vbExclamation + vbOKOnly,"ATENÇÃO"
			call carregar_menu
	end select
	
	wscript.quit 'Fecha o menu e abre o próximo
end sub
