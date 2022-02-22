dim audio, pular, nome, basico(10), intermediario(10), avancado(10), palavras(15), indice, leitor, validacao, numero, nivel, resposta
dim classificacao, novamente, resp, novo,nivel_final

call carregar_audio
sub carregar_audio()
set audio=createobject("SAPI.SPVOICE")
audio.volume=100
audio.rate=1
call name
end sub

sub name()
nivel= 1
pular= 0

nome=inputbox("Digite o seu nome: ", vbinformation + vbokonly)
call rotina
end sub

sub rotina()
basico(1)="MAIO"
basico(2)="GARFO" 
basico(3)="POMADA"
basico(4)="GATO"
basico(5)="TEIA"
basico(6)="JUÍZ"
basico(7)="ALGODÃO"
basico(8)="ZEBRA"
basico(9)="TECIDO"
basico(10)="PORCO"
intermediario(1)="MADRID"
intermediario(2)="TRICICLO" 
intermediario(3)="MAJESTADE"
intermediario(4)="VARSÓVIA"
intermediario(5)="RELAXANTE"
intermediario(6)="MOUSSELINE"
intermediario(7)="DESODORANTE"
intermediario(8)="CONSTANTINOPLA"
intermediario(9)="HEADSET"
intermediario(10)="LANCINANTE"
avancado(1)="TORDESILHAS"
avancado(2)="LAMBORGHINI" 
avancado(3)="CHEVETTE"
avancado(4)="SCENIC"
avancado(5)="RESSUREIÇÃO"
avancado(6)="PNEUMOULTRAMICROSCOPICOSSILICOVULCANOCONIÓTICO"
avancado(7)="VICISSITUDES"
avancado(8)="ELUCUBRAÇÕES"
avancado(9)="PROLEGÔMENOS"
avancado(10)="OTORRINOLARINGOLOGISTA"

leitor= 1
do while leitor <=5
    validacao= 0
    randomize(second(time))
    numero=int(rnd * 10) + 1
    for indice= 1 to 5 step 1
        if palavras(indice) = basico(numero) then 
            validacao= 1
        end if
    next
    if validacao= 0 then   
        palavras(leitor) = basico(numero)
        leitor=leitor + 1
    end if
loop
leitor= 6
do while leitor <=10
    validacao= 0
    randomize(second(time))
    numero=int(rnd * 10) + 1
    for indice= 6 to 10 step 1
        if palavras(indice) = intermediario(numero) then 
            validacao= 1
        end if
    next
    if validacao= 0 then   
        palavras(leitor) = intermediario(numero)
        leitor=leitor + 1
    end if
loop
leitor= 11
do while leitor <=15
    validacao= 0
    randomize(second(time))
    numero=int(rnd * 10) + 1
    for indice= 11 to 15 step 1
        if palavras(indice) = avancado(numero) then 
            validacao= 1
        end if
    next
    if validacao= 0 then   
        palavras(leitor) = avancado(numero)
        leitor=leitor + 1
    end if
loop
call falar
end sub

sub falar()
    novamente= 0
    if nivel <= 5 Then
        classificacao = "Basico"
    elseif nivel <= 10 then 
        classificacao = "intermediario"
    else    
        classificacao = "Avançado"
    end if
    audio.speak ("A palavra é " + palavras(nivel))
    do while true
        resposta=UCase(inputbox("Digite a palavra ouvida: " + vbnewline & _
                        "=================================" + vbnewline & _
                        "NOME: "& nome &"" + vbnewline & _
                        "[O]uvir novamente a palavra" + vbnewline & _
                        "[P]ular a palavra uma Unica vez" + vbnewline & _
                        "=================================", vbinformation + vbokonly))

        if palavras(nivel) = resposta then
            msgbox("Parabéns, você acertou! A palavra era: " & palavras(nivel) & vbnewline & _
                "Pontuação: "& nivel & vbnewline & _ 
                "Nível: " & classificacao), vbinformation + vbokonly,"AVISO"
            nivel = nivel + 1
            if nivel > 15 then
                call ganhou
            end if
        call falar
        elseif resposta = "O" then
            if novamente= 0 then
                novamente=novamente + 1
                audio.speak ("A palavra é " + palavras(nivel))
            else
                msgbox("Você já utilizou essa opção, somente disponível na próxima palavra"),vbinformation + vbokonly,"AVISO"    
            end if
        elseif resposta = "P" then
            if pular= 0 then 
                pular=pular + 1
                 msgbox("Você pulou, a palavra era: " & palavras(nivel))
                do while nivel <=5
                    validacao= 0
                    randomize(second(time))
                    novo=int(rnd * 10) + 1
                    for indice= 1 to 5 step 1
                        if palavras(indice) = basico(novo) then 
                            validacao= 1
                        end if
                    next
                    if validacao= 0 then   
                        palavras(nivel) = basico(novo)
                    end if

                    call falar
                loop
                do while nivel <=10
                    validacao= 0
                    randomize(second(time))
                    novo=int(rnd * 10) + 1
                    for indice= 6 to 10 step 1
                        if palavras(indice) = intermediario(novo) then 
                            validacao= 1
                        end if
                    next
                    if validacao= 0 then   
                        palavras(nivel) = intermediario(novo)
                    end if

                    call falar
                loop
                do while nivel <=15
                    validacao= 0
                    randomize(second(time))
                    novo=int(rnd * 10) + 1
                    for indice= 11 to 15 step 1
                        if palavras(indice) = avancado(novo) then 
                            validacao= 1
                        end if
                    next
                    if validacao= 0 then   
                        palavras(nivel) = avancado(novo)
                    end if

                    call falar
                loop	
            else
             msgbox("Você já utilizou essa opção!!!"),vbinformation + vbokonly,"AVISO"  
            end if
        else
		nivel_final=nivel-1
		 if nivel_final <= 5 then
                classificacao= "Basico"
            elseif nivel_final <= 10 then
                classificacao= "Intermediario"
            else
                classificacao= "Avançado"
            end if
            audio.speak("Que pena "& nome &", você perdeu!") 
            resp=msgbox("Que pena, você perdeu!" +vbnewline & _
                        "Pontuação: "& nivel_final &""   +vbnewline & _ 
                        "Nível: "& classificacao &"" +vbnewline & _
                        "Deseja jogar novamente?",vbQuestion + vbyesno,"AVISO")
            if resp= vbyes then
                call name
            else   
            wscript.quit
            end if
        end if
    loop
end sub 

sub ganhou()
        audio.speak("Parabéns "& nome &", você ganhou o soletrando!")
        resp=msgbox("Parabéns, você zerou o game! Deseja jogar novamente?",vbQuestion + vbyesno,"AVISO")
        if resp= vbyes then
            call name
        else
        wscript.quit
        end if
end sub    