import random
from colorama import init, Fore, Back
import sys
import time

init()


class Personagem:
    def __init__(self, nome, vida_base, mana_base, vigor_base, força_base, destreza_base, inteligencia_base, arma_inicial):
        self.nome = nome
        self.vida_base = vida_base
        self.mana_base = mana_base
        self.vigor_base = vigor_base
        self.força_base = força_base
        self.destreza_base = destreza_base
        self.inteligencia_base = inteligencia_base
        self.vida = vida_base
        self.mana = mana_base
        self.vigor = vigor_base
        self.força = força_base
        self.destreza = destreza_base
        self.inteligencia = inteligencia_base
        self.nivel = 1
        self.experiencia = 0
        self.experiencia_necessaria = 10
        self.pontos_atributo = 0
        self.arma_inicial = arma_inicial
        self.inventario_armas = [arma_inicial]
        self.ataques = self.determinar_ataques(arma_inicial)
        self.habilidades = self.determinar_habilidades(arma_inicial)

    def determinar_ataques(self, arma):
        ataques = {"ataque basico": 0}
        if arma == "espadão":
            ataques.update({"ataque forte": 3, "ataque giratório": 5,
                            "pulo com golpe": 3, "golpe em area": 3})
        elif arma == "duas adagas":
            ataques.update(
                {"back stab": 3, "corte rapido": 3, "dança das adagas": 6})
        elif arma == "cajado":
            ataques.update({"ataque basico": 0})
            self.elementos = {"fogo": {"bola de fogo": 3, "cuspe do dragão": 5},
                              "gelo": {"floco de gelo": 3, "explosão de gelo": 5},
                              "raio": {"ataque de raio": 3, "choque do trovão": 5}}
        return ataques

    def determinar_habilidades(self, arma):
        habilidades = []
        if arma == "espadão":
            habilidades = ["ataque giratório",
                           "pulo com golpe", "golpe em area"]
        elif arma == "duas adagas":
            habilidades = ["back stab", "dança das adagas", "corte rapido"]
        elif arma == "cajado":
            habilidades = {"fogo": ["bola de fogo", "cuspe do dragão"],
                           "gelo": ["floco de gelo", "explosão de gelo"],
                           "raio": ["ataque de raio", "choque do trovão"]}
        return habilidades

    def ganhar_experiencia(self, exp):
        self.experiencia += exp
        while self.experiencia >= self.experiencia_necessaria:
            self.subir_nivel()
            self.experiencia -= self.experiencia_necessaria
            self.experiencia_necessaria += 15

    def subir_nivel(self):
        self.nivel += 1
        print(Fore.GREEN + "\nVocê subiu de nível, parabéns!")
        time.sleep(1)
        self.pontos_atributo += 5
        self.vida = self.vida_base
        self.mana = self.mana_base
        print(Fore.GREEN + "Sua vida e mana foram totalmente restauradas.")
        time.sleep(1)
        self.distribuir_pontos()

    def distribuir_pontos(self):
        print(Fore.GREEN + "\nVocê ganhou 5 pontos de habilidades, onde deseja inserir?")
        time.sleep(1)
        while self.pontos_atributo > 0:
            escolha = input("""
1 - Vida
2 - Mana
3 - Vigor
4 - Força
5 - Destreza
6 - Inteligência
Digite o número do atributo: """)
            try:
                escolha = int(escolha)
                if 1 <= escolha <= 6:
                    pontos = input(
                        Fore.GREEN + f"Quantos pontos deseja adicionar em {['Vida', 'Mana', 'Vigor', 'Força', 'Destreza', 'Inteligência'][escolha - 1]}? (Você tem {self.pontos_atributo} pontos): ")
                    pontos = int(pontos)
                    if 0 < pontos <= self.pontos_atributo:
                        if escolha == 1:
                            self.vida_base += pontos
                            self.vida += pontos
                        elif escolha == 2:
                            self.mana_base += pontos
                            self.mana += pontos
                        elif escolha == 3:
                            self.vigor_base += pontos
                            self.vigor += pontos
                        elif escolha == 4:
                            self.força_base += pontos
                            self.força += pontos
                        elif escolha == 5:
                            self.destreza_base += pontos
                            self.destreza += pontos
                        elif escolha == 6:
                            self.inteligencia_base += pontos
                            self.inteligencia += pontos
                        self.pontos_atributo -= pontos
                        print(
                            f"{pontos} pontos adicionados em {['Vida', 'Mana', 'Vigor', 'Força', 'Destreza', 'Inteligência'][escolha - 1]}.")
                        time.sleep(1)
                    else:
                        print("Quantidade de pontos inválida.")
                        time.sleep(1)
                else:
                    print("Escolha inválida.")
                    time.sleep(1)
            except ValueError:
                print("Entrada inválida.")
                time.sleep(1)

    def usar_mana(self, custo):
        if self.mana >= custo:
            self.mana -= custo
            return True
        else:
            print(Fore.RED + "Mana insuficiente!")
            time.sleep(1)
            return False

    def atacar(self, alvo, ataque="ataque basico", habilidade=None):
        dano = 0
        nome_golpe = ataque
        if ataque == "ataque basico":
            if "espadão" in self.inventario_armas:
                dano = (self.força * 0.45) + (self.destreza * 0.05)
            elif "duas adagas" in self.inventario_armas:
                dano = (self.destreza * 0.45) + (self.força * 0.05)
                if random.random() < 0.05:
                    dano += self.destreza * 0.35
                    nome_golpe = "ataque basico (crítico)"
            elif "cajado" in self.inventario_armas:
                dano = self.inteligencia * 0.2
        elif ataque in self.ataques and self.ataques[ataque] > 0:
            custo = self.ataques[ataque]
            if self.usar_mana(custo):
                if ataque == "ataque forte":
                    dano = (self.força * 0.7)
                elif ataque == "ataque giratório":
                    dano = (self.força * 0.5) + (self.vigor * 0.1)
                elif ataque == "pulo com golpe":
                    dano = (self.força * 0.5) + (self.vigor * 0.1)
                elif ataque == "golpe em area":
                    dano = (self.força * 0.5) + (self.vigor * 0.1)
                elif ataque == "back stab":
                    dano = (self.destreza * 0.6)
                elif ataque == "corte rapido":
                    dano = (self.destreza * 0.5) + (self.força * 0.1)
                elif ataque == "dança das adagas":
                    dano = (self.destreza * 0.5)
                elif ataque == "bola de fogo":
                    dano = (self.inteligencia * 0.45)
                elif ataque == "cuspe do dragão":
                    dano = (self.inteligencia * 0.45)
                elif ataque == "floco de gelo":
                    dano = (self.inteligencia * 0.45)
                elif ataque == "explosão de gelo":
                    dano = (self.inteligencia * 0.45)
                elif ataque == "ataque de raio":
                    dano = (self.inteligencia * 0.45)
                elif ataque == "choque do trovão":
                    dano = (self.inteligencia * 0.45)
                else:
                    return 0, None
            else:
                return 0, None
        elif self.arma_inicial == "cajado" and habilidade:
            for elemento, lista_habilidades in self.habilidades.items():
                if habilidade in lista_habilidades:
                    if habilidade in self.elementos[elemento]:
                        custo = self.elementos[elemento][habilidade]
                        if self.usar_mana(custo):
                            dano = self.inteligencia * 0.65
                            nome_golpe = habilidade
                        else:
                            print(
                                Fore.RED + "Mana insuficiente para esta habilidade!")
                            time.sleep(1)
                            return 0, None
                        break
                    else:
                        print(
                            Fore.RED + f"Erro interno: Habilidade '{habilidade}' não encontrada no elemento '{elemento}'.")
                        time.sleep(1)
                        return 0, None
            if dano == 0 and nome_golpe == habilidade:
                return 0, None
            elif dano == 0:
                print("Habilidade inválida!")
                time.sleep(1)
                return 0, None
        else:
            print("Ataque inválido!")
            time.sleep(1)
            return 0, None

        if dano > 0:
            alvo.vida -= dano
            alvo.vida = max(0, alvo.vida)
        return dano, nome_golpe


class Monstros:
    def __init__(self, nome, vida, mana, vigor, força, destreza, inteligencia):
        self.nome = nome
        self.vida = vida
        self.mana = mana
        self.vigor = vigor
        self.força = força
        self.destreza = destreza
        self.inteligencia = inteligencia
        self.dado = ["1", "2", "3", "4", "5", "6"]

    def rolar_dado(self):
        return random.choice(self.dado)

    def atacar(self, personagem):
        rolagem = self.rolar_dado()
        print(f"\n{self.nome} rola o dado: {rolagem}")
        time.sleep(1)
        if int(rolagem) >= 4:
            print(f"{self.nome} teve sucesso na rolagem!")
            time.sleep(1)
            return self.realizar_ataque(personagem)
        else:
            print(f"{self.nome} falhou na rolagem!")
            time.sleep(1)
            return 0

    def realizar_ataque(self, personagem):
        raise NotImplementedError(
            "Subclasses devem implementar o método realizar_ataque")


class Esqueleto(Monstros):
    def __init__(self):
        super().__init__("Esqueleto", 20, 5, 10, 14, 4, 4)

    def realizar_ataque(self, personagem):
        dano = (self.força * 0.3) + (self.destreza * 0.05)
        personagem.vida -= dano
        personagem.vida = max(0, personagem.vida)
        print(
            f"O Esqueleto ataca com um golpe básico, causando {dano:.2f} de dano.")
        time.sleep(1)
        return dano


class Encapuzado(Monstros):
    def __init__(self):
        super().__init__("Encapuzado", 14, 12, 6, 5, 4, 16)

    def realizar_ataque(self, personagem):
        dano = self.inteligencia * 0.4
        personagem.vida -= dano
        personagem.vida = max(0, personagem.vida)
        print(f"O Encapuzado ataca com magia, causando {dano:.2f} de dano.")
        time.sleep(1)
        return dano


class Goblin(Monstros):
    def __init__(self):
        super().__init__("Goblin", 18, 8, 8, 8, 12, 6)

    def realizar_ataque(self, personagem):
        dano = (self.destreza * 0.3) + (self.força * 0.05)
        if random.random() < 0.05:
            dano += self.destreza * 0.2
            print("Ataque crítico do Goblin!")
            time.sleep(1)
        personagem.vida -= dano
        personagem.vida = max(0, personagem.vida)
        print(f"O Goblin ataca com suas adagas, causando {dano:.2f} de dano.")
        time.sleep(1)
        return dano


class Aranhao(Monstros):
    def __init__(self):
        super().__init__("Aranhao", 16, 8, 12, 7, 11, 5)

    def realizar_ataque(self, personagem):
        dano = (self.destreza * 0.4) + (self.vigor * 0.2)
        if random.random() < 0.1:
            print("O Aranhao tece uma teia pegajosa!")
            time.sleep(1)
        personagem.vida -= dano
        personagem.vida = max(0, personagem.vida)
        print(
            f"O Aranhao ataca com suas patas afiadas, causando {dano:.2f} de dano.")
        time.sleep(1)
        return dano


class Ademiro(Monstros):
    def __init__(self):
        super().__init__("Ademiro", 50, 8, 7, 6, 10, 8)

    def realizar_ataque(self, personagem):
        dano = (self.destreza * 0.3) + (self.força * 0.05) + \
               (self.inteligencia * 0.6)
        if random.random() < 0.05:
            dano += self.destreza * 0.2
            print("Ataque crítico do Ademiro!")
            time.sleep(1)
        personagem.vida -= dano
        personagem.vida = max(0, personagem.vida)
        print(
            f"Ademiro ataca com suas garras feitas de ossos, causando {dano:.2f} de dano.")
        time.sleep(1)
        return dano


chefao = Ademiro()


def combate(jogador, monstro, resistencias, dado_jogador):
    turno = 0
    while jogador.vida > 0 and monstro.vida > 0:
        print(Fore.YELLOW + "\n--- Novo Turno ---")
        time.sleep(1)
        print(
            Fore.WHITE + f"{jogador.nome} (Nível {jogador.nivel}): Vida {jogador.vida:.2f}, Mana {jogador.mana:.2f}")
        time.sleep(1)
        print(f"{monstro.nome}: Vida {monstro.vida:.2f}")
        time.sleep(1)
        if turno % 2 == 0:
            print(Fore.GREEN + "Seu turno!")
            time.sleep(1)
            ataque_escolhido = "ataque basico"
            habilidade_escolhida = None
            if jogador.mana == 0:
                print(Fore.RED + "Sem mana! Apenas o ataque básico está disponível.")
                time.sleep(1)
            else:
                print(Fore.GREEN + "Escolha seu ataque:")
                time.sleep(1)
                opcoes_ataque = list(jogador.ataques.keys())
                if jogador.arma_inicial == "cajado":
                    print("0. ataque basico (Sem custo)")

                    opcoes_habilidade = []
                    contador = 1
                    for elemento, habilidades in jogador.habilidades.items():
                        print(f"--- {elemento.capitalize()} ---")

                        for h in habilidades:
                            opcoes_habilidade.append(h)
                            print(
                                f"{contador}. {h} (Mana: {jogador.elementos[elemento][h]})")

                            contador += 1
                    try:
                        escolha = int(
                            input(Fore.GREEN + "Digite o número do ataque/habilidade: "))
                        if escolha == 0:
                            ataque_escolhido = "ataque basico"
                        else:
                            habilidade_escolhida = opcoes_habilidade[escolha - 1]
                            ataque_escolhido = None
                    except:
                        pass
                else:
                    for i, atk in enumerate(opcoes_ataque):
                        custo = f" (Mana: {jogador.ataques[atk]})" if jogador.ataques[atk] > 0 else ""
                        print(f"{i + 1}. {atk}{custo}")

                    try:
                        escolha = int(
                            input(Fore.GREEN + "Digite o número do ataque: "))
                        ataque_escolhido = opcoes_ataque[escolha - 1]
                    except:
                        pass
                if input(Fore.BLUE + "Deseja rolar o dado para atacar? (sim/não): ").lower() == "sim":
                    if int(random.choice(dado_jogador)) >= 3:
                        print(Fore.GREEN + "Sucesso na rolagem do dado!")
                        time.sleep(1)
                        dano, golpe = jogador.atacar(
                            monstro, ataque_escolhido, habilidade_escolhida)
                        if golpe:
                            tipo = jogador.arma_inicial
                            inimigo = monstro.nome.lower()
                            if tipo in resistencias and inimigo in resistencias[tipo]:
                                dano *= resistencias[tipo][inimigo]
                                print(
                                    f"(Dano modificado pela resistência: {dano:.2f})")
                                time.sleep(1)
                            monstro.vida = max(0, monstro.vida)
                            print(
                                Fore.GREEN + f"Seu golpe '{golpe}' causou {dano:.2f} de dano.")
                            time.sleep(1)
                            print(
                                Fore.RED + f"{monstro.nome} vida restante: {monstro.vida:.2f}")
                            time.sleep(1)
                        else:
                            print(Fore.RED + "O ataque falhou!")
                            time.sleep(1)
                    else:
                        print(Fore.RED + "Falha na rolagem do dado!")
                        time.sleep(1)
                else:
                    print(Fore.RED + "Você perdeu o turno!")
                    time.sleep(1)
        else:
            print(Fore.RED + f"\nTurno do {monstro.nome}!")
            time.sleep(1)
            dano = monstro.atacar(jogador)
            print(Fore.RED + f"Sua vida restante: {jogador.vida:.2f}")
            time.sleep(1)
            if jogador.vida <= 0:
                print(Fore.RED + "Você foi derrotado!\n ----- FIM DE JOGO -----")
                time.sleep(10)
                sys.exit()
        turno += 1
    if monstro.vida <= 0:
        print(Fore.GREEN + f"Você derrotou o {monstro.nome}!")
        time.sleep(1)
        jogador.ganhar_experiencia(10)
        print(Fore.MAGENTA + "Você derrotou o inimigo!")
        time.sleep(1)


print(Fore.YELLOW + " APRESENTANDO O BETA ABERTO DE: A GUERRA DOS NOBRES TRABALHADORES. ")
time.sleep(1)
print(Fore.YELLOW + " --------------- CARREGANDO ---------------")
time.sleep(1)
print(Fore.YELLOW + " Seja bem vindo ao mundo fantasioso de: Ass Alariados, o reino dos CLTs. ")
time.sleep(1)

nome_jogador = input(Fore.GREEN + "Escolha o nome do jogador: ")

while True:
    print(Fore.GREEN + "Com qual arma deseja iniciar?")
    time.sleep(1)
    escolha_arma = input(Fore.GREEN + "Espadão\nDuas Adagas\nCajado\n").lower()
    if escolha_arma == "espadão":
        jogador = Personagem(nome_jogador, 18, 11, 11, 18, 6, 4, "espadão")
        break
    elif escolha_arma == "duas adagas":
        jogador = Personagem(nome_jogador, 15, 12, 10, 8, 18, 6, "duas adagas")
        break
    elif escolha_arma == "cajado":
        jogador = Personagem(nome_jogador, 13, 15, 8, 5, 4, 18, "cajado")
        break
    else:
        print(Fore.RED + "Arma inválida. Digite novamente.")
        time.sleep(1)

print(Fore.MAGENTA +
      f" A história de {nome_jogador} se inicia nos reinos de Ass Alariados, O Reino dos CLT'S. Sendo apresentado como o novo integrante da DIAFI. ")
time.sleep(1)

resistencias = {
    "espadão": {"esqueleto": 0.9, "encapuzado": 1.1, "goblin": 0.9, "aranha gigante": 1.0},
    "cajado": {"esqueleto": 1.1, "encapuzado": 0.9, "goblin": 1.1, "aranha gigante": 0.8},
    "duas adagas": {"esqueleto": 1.0, "encapuzado": 1.0, "goblin": 1.0, "aranha gigante": 1.2}
}

inimigos = [Esqueleto, Encapuzado, Goblin, Aranhao]
inimigo1_aleatorio = random.choice(inimigos)
inimigo1_instanciado = inimigo1_aleatorio()
dado_jogador = ["1", "2", "3", "4", "5", "6"]

print(Fore.MAGENTA +
      f" Após ser escolhido para a sua primeira jornada, {nome_jogador} decide partir imediatamente.")
time.sleep(1)
print(Fore.MAGENTA +
      "Você encontra uma caverna e sente calafrios só de olhar para dentro dela. ")
time.sleep(1)
print(Fore.MAGENTA + "Mas você decide entrar...")
time.sleep(1)
print(Fore.MAGENTA +
      f" Caminhando mais a frente, {nome_jogador} se depara com uma porta enorme, escrito NAJ ")
time.sleep(1)
print(Fore.MAGENTA + "Deseja entrar?")
time.sleep(1)

decisao_jogador = input(Fore.CYAN + "sim ou não? ").lower()

if decisao_jogador == "sim":
    print(Fore.RED +
          f"Ao ultrapassar a porta, {nome_jogador} se depara com um {inimigo1_instanciado.nome}!")
    time.sleep(1)
    print(Fore.YELLOW + "Iniciando combate!")
    time.sleep(1)
    monstro_ativo = inimigo1_instanciado
    turno = 0

    combate(jogador, monstro_ativo, resistencias, dado_jogador)

else:
    print(Fore.BLUE + "Então você continua a explorar.")
    time.sleep(1)
    print(Fore.RED + "Você passa por corredores e mais corredores, a cada segundo que passa você fica mais perdido. ")
    time.sleep(1)
    print(Fore.RED + "Finalmente, no fim de algum corredor explorado por acaso, você encontra uma saída. ")
    time.sleep(1)
    print(Fore.RED + "Uma porta bem a sua frente, angustiado por vagar solitariamente por horas, você abre a porta e corre para o fim dessa dungeon. ")
    time.sleep(1)
    print(Fore.RED + "Controlado pelo desespero, você vai direto para uma armadilha. ")
    time.sleep(1)
    print(Fore.RED +
          f"{nome_jogador} pisa em um fundo falso e cai em direção a um abismo.")
    time.sleep(1)
    print(Fore.RED +
          f"Com pouco tempo antes de sua inevitável morte, {nome_jogador} jurou nunca mais fugir de uma batalha. ")
    time.sleep(1)
    print(Fore.RED + "Mas em sua próxima vida...")
    time.sleep(1)
    print(Fore.RED + "--------- VOCÊ MORREU ---------")
    time.sleep(1)
    print(Fore.RED + "--------- FIM DE JOGO ---------")
    time.sleep(10)
    sys.exit()


print(Fore.LIGHTBLUE_EX +
      "Após derrotar o inimigo, você se deparada com uma porta a direita e um caminho a sua frente.")
time.sleep(1)
print(Fore.LIGHTBLUE_EX + "Por onde deseja prosseguir?")
time.sleep(1)
print(Fore.LIGHTBLUE_EX + "1- Entrar na porta a direita. ")
time.sleep(1)
print(Fore.LIGHTBLUE_EX + "2- Seguir em frente. ")
time.sleep(1)
escolha_caminho = input(Fore.MAGENTA + "Por onde deseja seguir? ")

if escolha_caminho == "1":
    inimigos = [Esqueleto, Encapuzado, Goblin, Aranhao]
    inimigo2_aleatorio = random.choice(inimigos)
    inimigo2_instanciado = inimigo2_aleatorio()
    dado_jogador = ["1", "2", "3", "4", "5", "6"]
    print(Fore.BLUE + "Você decide entrar na porta a direita... ")
    time.sleep(1)
    print(Fore.BLUE + "Ao olhar no fundo da sala, você encontra um baú")
    time.sleep(1)
    decisao_bau = input(
        Fore.MAGENTA + "Deseja abrir o baú? (sim) (não)\n ").lower()
    if decisao_bau == "sim":
        print(Fore.BLUE + "Abrindo o baú...")
        time.sleep(1)
        print(Fore.BLUE + "Ao chegar perto do baú, você percebe que tem algo de errado...")
        time.sleep(1)
        print(Fore.BLUE + "Você sente que... está sendo observado...")
        time.sleep(1)
        print(
            Fore.RED + f" Era uma armadilha, você encontrou um {inimigo2_instanciado.nome}.")
        time.sleep(1)
        print(Fore.YELLOW + "Iniciando combate!")
        time.sleep(1)

        monstro_ativo = inimigo2_instanciado
        turno = 0

        combate(jogador, monstro_ativo, resistencias, dado_jogador)

        print(Fore.BLUE + "Após uma luta intensa, algo na parede se mexe, chama sua atenção, pois não parecia comum.")
        time.sleep(1)
        print(Fore.BLUE + "Ao chegar perto, você percebe uma anomalia.")
        time.sleep(1)
        print(Fore.BLUE + "Investigando parede...")
        time.sleep(1)
        print(Fore.BLUE + "Uma passagem se abre...")
        time.sleep(1)
        print(Fore.BLUE + "Você encontra uma jovem garota, aparentemente perdida. ")
        time.sleep(1)
        print(Fore.BLUE + "Garota Desconhecida: Pode... mee... ajudar?")
        time.sleep(1)
        decisao_jovem = input(
            Fore.MAGENTA + "Deseja ajuda-lá? (sim) (não)\n").lower()
        if decisao_jovem == "sim":
            inimigos = [Esqueleto, Encapuzado, Goblin, Aranhao]
            inimigo4_aleatorio = random.choice(inimigos)
            inimigo4_instanciado = inimigo4_aleatorio()
            dado_jogador = ["1", "2", "3", "4", "5", "6"]
            print(Fore.BLUE +
                  f"{nome_jogador}: Uma jovem garota aqui? sozinha?")
            time.sleep(1)
            print(Fore.BLUE + "Se aproximando da garota...")
            time.sleep(1)
            print(Fore.BLUE +
                  f"{nome_jogador}? O que faz aqui, sozinha? você está b...")
            time.sleep(1)
            print(Fore.BLUE + "Ela se revela ser um monstro. ")
            time.sleep(1)
            print(
                Fore.RED + f"{inimigo4_instanciado}: Não acredito que caiu nessa HAHAHAHAHAHA")
            time.sleep(1)
            print(Fore.YELLOW + "Iniciando combate!")
            time.sleep(1)
            monstro_ativo = inimigo4_instanciado
            turno = 0

            combate(jogador, monstro_ativo, resistencias, dado_jogador)

            print(Fore.CYAN + "Por que uma garota estaria sozinha e perdida dentro dessa caverna? você se pergunta. ")
            time.sleep(1)
            print(Fore.CYAN + "Não tem problema, vi alguém em perigo e decidi ajudar, não posso me culpar pela armadilha. ")
            time.sleep(1)
            print(Fore.MAGENTA +
                  f"Então {nome_jogador} decide seguir adiante...")
            time.sleep(1)
            print(
                Fore.CYAN + "Você sente calafrios, a cada passo que da, sente um terrível poder a sua frente. ")
            time.sleep(1)
            print(Fore.CYAN + "Uma porta enorme está no fim do corredor, e você percebe que é de lá que vem seus calafrios. ")
            time.sleep(1)
            print(
                Fore.CYAN + "Ao abrir a porta, você se depara com Ademiro Santiro, Chefe dos Estagiários. ")
            time.sleep(1)
            print(Fore.LIGHTMAGENTA_EX +
                  f"Ademiro: Achei que nunca iria chegar, {nome_jogador} ")
            time.sleep(1)
            print(Fore.LIGHTMAGENTA_EX +
                  "Ademiro: Como ousa invadir meu santuário? ")
            time.sleep(1)
            print(Fore.LIGHTMAGENTA_EX +
                  "Ademiro: Não permitirei que tal ato de rebeldia saia impune. ")
            time.sleep(1)
            print(Fore.LIGHTMAGENTA_EX +
                  f"Ademiro: EU TE FAREI MEU ESTAGIÁRIO {nome_jogador} MUAHAHAHAHAHAHAHAHA")
            time.sleep(1)
            monstro_ativo = chefao
            turno = 0

            combate(jogador, monstro_ativo, resistencias, dado_jogador)

        else:
            print(Fore.CYAN + "Por que uma garota estaria sozinha e perdida dentro dessa caverna? você se pergunta. ")
            time.sleep(1)
            print(
                Fore.CYAN + "Com certeza era uma armadilha, não vou me arriscar desse jeito. ")
            time.sleep(1)
            print(Fore.MAGENTA +
                  f"Então {nome_jogador} decide seguir a diante...")
            time.sleep(1)
            print(
                Fore.CYAN + "Você sente calafrios, a cada passo que da, sente um terrível poder a sua frente. ")
            time.sleep(1)
            print(Fore.CYAN + "Uma porta enorme está no fim do corredor, e você percebe que é de lá que vem seus calafrios. ")
            time.sleep(1)
            print(
                Fore.CYAN + "Ao abrir a porta, você se depara com Ademiro Santiro, Chefe dos Estagiários. ")
            time.sleep(1)
            print(Fore.LIGHTMAGENTA_EX +
                  f"Ademiro: Achei que nunca iria chegar, {nome_jogador} ")
            time.sleep(1)
            print(Fore.LIGHTMAGENTA_EX +
                  "Ademiro: Como ousa invadir meu santuário? ")
            time.sleep(1)
            print(Fore.LIGHTMAGENTA_EX +
                  "Ademiro: Não permitirei que tal ato de rebeldia saia impune. ")
            time.sleep(1)
            print(Fore.LIGHTMAGENTA_EX +
                  f"Ademiro: EU TE FAREI MEU ESTAGIÁRIO {nome_jogador} MUAHAHAHAHAHAHAHAHA")
            time.sleep(1)
            monstro_ativo = chefao
            turno = 0

            combate(jogador, monstro_ativo, resistencias, dado_jogador)

    else:
        print(Fore.BLUE + "Passando pelo baú...")
        time.sleep(1)
        print(
            Fore.BLUE + f"{nome_jogador}: Não sei o por que, mas aquele baú não me parecia uma boa ideia.")
        time.sleep(1)
        print(Fore.BLUE + "Ao chegar no fim da sala, você percebe uma anomalia na parede.")
        time.sleep(1)
        print(Fore.BLUE + f"{nome_jogador}: Talvez uma passagem secreta? ")
        time.sleep(1)
        print(Fore.BLUE + "Investigando parede...")
        time.sleep(1)
        print(Fore.BLUE + "Você encontrou uma jovem garota, aparentemente perdida. ")
        time.sleep(1)
        print(Fore.BLUE + "Garota Desconhecida: Pode... mee... ajudar?")
        time.sleep(1)
        decisao_jovem = input(
            Fore.MAGENTA + "Deseja ajuda-lá? (sim) (não)\n").lower()
        if decisao_jovem == "sim":
            inimigos = [Esqueleto, Encapuzado, Goblin, Aranhao]
            inimigo3_aleatorio = random.choice(inimigos)
            inimigo3_instanciado = inimigo3_aleatorio()
            dado_jogador = ["1", "2", "3", "4", "5", "6"]
            print(Fore.BLUE +
                  f"{nome_jogador}: Uma jovem garota aqui? sozinha?")
            time.sleep(1)
            print(Fore.BLUE + "Se aproximando da garota...")
            time.sleep(1)
            print(Fore.BLUE +
                  f"{nome_jogador}? O que faz aqui, sozinha? você está b...")
            time.sleep(1)
            print(Fore.BLUE + "Ela se revela ser um monstro. ")
            time.sleep(1)
            print(
                Fore.RED + f"{inimigo3_instanciado}: Não acredito que caiu nessa HAHAHAHAHAHA")

            time.sleep(1)
            print(Fore.YELLOW + "Iniciando combate!")
            time.sleep(1)
            monstro_ativo = inimigo3_instanciado
            turno = 0

            combate(jogador, monstro_ativo, resistencias, dado_jogador)

            print(Fore.CYAN + "Ao chegar perto do monstro, você escuta um sussurro. ")
            time.sleep(1)
            print(Fore.BLACK + "Me desculpe... Ademiro")
            time.sleep(1)
            print(
                Fore.CYAN + f"{nome_jogador}: Como alguém tem coragem de usar uma linda jovem para algo tão cruel? ")
            time.sleep(1)
            print(Fore.CYAN + "Não deixarei que Ademiro fique vivo, EU IREI MATÁ-LO! ")
            time.sleep(1)
            print(
                Fore.CYAN + "No fim da sala, você encontra uma pequena porta, talvez uma entrada alternativa? ")
            time.sleep(1)
            print(Fore.CYAN + "Ao abrir a porta, você se depara com um ser enorme, de cabelos grisalhos e garras de ossos. ")
            time.sleep(1)
            print(Fore.RED + f"{nome_jogador}: ADEMIROOOOOOOOOOOO!!!")
            time.sleep(1)
            print(Fore.RED + "EU O DERROTAREI, SINTA MINHA FÚRIA!!!!!!!!")
            time.sleep(1)
            print(Fore.RED + "Ademiro: HAHAHAHAHAHAHA ")
            time.sleep(1)
            print(Fore.RED + f"Venha {nome_jogador}, te farei meu estiário. ")
            time.sleep(1)
            print(Fore.RED + " VENHA JOVENZINHO. MUAHAHAHHAHA")
            time.sleep(1)
            monstro_ativo = chefao
            turno = 0

            combate(jogador, monstro_ativo, resistencias, dado_jogador)

        else:
            print(Fore.BLUE + "Você decide então não ajudar a garota, sua astúcia lhe alertou sobre a possível armadilha. ")
            time.sleep(1)
            print(
                Fore.BLUE + "Seguindo em frente, com total certeza de que era uma armadilha.")
            time.sleep(1)
            print(Fore.BLUE + "Você encontra uma porta enorme, de vidro.")
            time.sleep(1)
            print(Fore.BLUE + "Ao olhar mais de perto, você se depara com um salão enorme e no meio dele, um ser misterioso, de cabelos grisalhos. ")
            time.sleep(1)
            decisao_porta_vidro = input(
                "Deseja entrar no salão? (sim) (não)\n").lower()
            if decisao_porta_vidro == "sim":
                print(
                    Fore.RED + "Você encontrou Ademiro Santistiro, o REI dos Estagiários.")
                time.sleep(1)
                print(Fore.RED +
                      f"Ademiro: Estava te esperando {nome_jogador}.")
                time.sleep(1)
                print(Fore.RED + "Ninguém entra no NAJ sem ser notado por mim.")
                time.sleep(1)
                print(
                    Fore.RED + f"Batalhe comigo {nome_jogador}, se vencer, te deixarei ir, mas se perder...")
                time.sleep(1)
                print(Fore.RED +
                      "VOCÊ ME SERVIRÁ PARA SEMPRE JOVENZINHO!!! HAHAHAHAHAHA")
                time.sleep(1)
                print(Fore.YELLOW + "Iniciando combate!")
                time.sleep(1)
                monstro_ativo = chefao

                combate(jogador, monstro_ativo, resistencias, dado_jogador)

            else:
                print(Fore.YELLOW + " Você decide não entrar. ")
                time.sleep(1)
                print(
                    Fore.YELLOW + " Enquanto estava observando o salão, do lado de fora, você escuta alguém te chamando. ")
                time.sleep(1)
                print(Fore.YELLOW + "Investigando a voz misteriosa... ")
                time.sleep(1)
                print(Fore.YELLOW +
                      "Você encontrou Julianiri Chefirini, a lider da DIAFI. ")
                time.sleep(1)
                print(
                    Fore.YELLOW + f"Julianiri: Se afaste {nome_jogador}, eu espero por essa batalha a anos, deixe Ademiro Santistiro comigo. ")
                time.sleep(1)
                print(Fore.YELLOW + "Enquanto uma longa batalha é travada entre Ademiro e Julianiri, você volta em segurança para DIAFI, onde seus colegas o esperam. ")
                time.sleep(1)
                print(
                    Fore.BLUE + f"{nome_jogador}: Eu ficarei mais forte!!! nunca mais precisarei ser socorrido de uma caverna.")
                time.sleep(1)
                print(Fore.RED + " ---------------- FIM DO BETA ----------------")
                time.sleep(1)
                sys.exit()


else:
    print(Fore.RED + "Você encontrou a sala do boss!!!")
    time.sleep(1)
    print(Fore.RED + "Ademiro: Eu sou Ademiro Santiro, REI dos Estagiários. Vejo que derrotou meu guarda-costas. ")
    time.sleep(1)
    print(Fore.RED + "Ademiro: Afinal, o que deseja?")
    time.sleep(1)
    print(Fore.RED + "Ademiro: Está em busca de poder? fama? não importa... ")
    time.sleep(1)
    print(Fore.RED + "Ademiro: Não deixarei que um simples CLT saia vivo do NAJ...")
    time.sleep(1)
    print(Fore.RED + "Ademiro: VOCÊ MORRERÁ, JOVENZINHO!!!!!!!")
    time.sleep(1)
    print(Fore.YELLOW + "------------- Iniciando combate! -------------")
    time.sleep(1)
    monstro_ativo = chefao

    combate(jogador, monstro_ativo, resistencias, dado_jogador)


print(Fore.LIGHTWHITE_EX +
      f"Parabéns {nome_jogador}, você derrotou Ademiro Santiro")
time.sleep(1)
print(Fore.LIGHTWHITE_EX +
      "Sua primeira jornada foi travada com muita luta e muita habilidade")
time.sleep(1)
print(Fore.LIGHTWHITE_EX +
      "Você retornará para a DIAFI como um herói, pois finalmente nasceu alguém que pudesse combater este mal")
time.sleep(1)
print(Fore.LIGHTWHITE_EX + "Voltando para DIAFI...")
time.sleep(1)
print(Fore.RED + "          ------------- FIM DE JOGO -------------          ")
time.sleep(1)
print(Fore.RED + "------------- OBRIGADO POR JOGADOR A BETA -------------")

time.sleep(10)
sys.exit()
