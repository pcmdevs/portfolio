import pygame
import sys
import random
import os
import time
from pygame.locals import *
import math


pygame.init()
pygame.font.init()
pygame.mixer.init()

LARGURA, ALTURA = 800, 600
tela = pygame.display.set_mode((LARGURA, ALTURA))
pygame.display.set_caption("A LENDA DO CLT")


PRETO = (0, 0, 0)
BRANCO = (255, 255, 255)
VERMELHO = (255, 0, 0)
VERDE = (0, 255, 0)
AZUL = (0, 0, 255)
AMARELO = (255, 255, 0)
CIANO = (0, 255, 255)
MAGENTA = (255, 0, 255)
CINZA_ESCURO = (50, 50, 50)
CINZA_CLARO = (180, 180, 180)


FONTE_PADRAO = "Arial"
try:
    fonte_pequena = pygame.font.SysFont(FONTE_PADRAO, 16)
    fonte_media = pygame.font.SysFont(FONTE_PADRAO, 22)
    fonte_grande = pygame.font.SysFont(FONTE_PADRAO, 30)
    fonte_titulo = pygame.font.SysFont(FONTE_PADRAO, 46)
except pygame.error:
    print(
        f"Aviso: Fonte {FONTE_PADRAO} não encontrada, usando fonte padrão do Pygame.")
    fonte_pequena = pygame.font.Font(None, 18)
    fonte_media = pygame.font.Font(None, 24)
    fonte_grande = pygame.font.Font(None, 32)
    fonte_titulo = pygame.font.Font(None, 48)


musica_jogo = os.path.join("musicas", "cinematic_tension_002.mp3")
musica_chefao = os.path.join("musicas", "gladiator.mp3")
musica_atual = None


def tocar_musica(caminho_musica, volume=0.3, loop=-1, forcar_troca=False):
    global musica_atual
    if not pygame.mixer.get_init():
        print("Mixer de áudio não inicializado.")
        return
    try:

        if forcar_troca or caminho_musica != musica_atual:
            pygame.mixer.music.load(caminho_musica)
            pygame.mixer.music.set_volume(volume)
            pygame.mixer.music.play(loop)
            musica_atual = caminho_musica

        elif not pygame.mixer.music.get_busy():
            pygame.mixer.music.play(loop)

    except pygame.error as e:
        print(f"Erro ao carregar ou tocar música {caminho_musica}: {e}")
        musica_atual = None


def carregar_imagem(nome_arquivo):
    try:
        caminho = os.path.join("imagens", nome_arquivo)
        return pygame.image.load(caminho).convert_alpha()
    except pygame.error as e:
        print(f"Erro ao carregar imagem {nome_arquivo}: {e}")

        surface_erro = pygame.Surface((50, 50), pygame.SRCALPHA)
        surface_erro.fill(MAGENTA + (128,))
        pygame.draw.line(surface_erro, PRETO, (0, 0), (49, 49), 2)
        pygame.draw.line(surface_erro, PRETO, (0, 49), (49, 0), 2)
        return surface_erro


# --- Carregamento de Imagens ---
img_jogador_cajado = carregar_imagem("jogador_cajado.png")
img_jogador_espadao = carregar_imagem("jogador_espadão.png")
img_jogador_adagas = carregar_imagem("jogador_adagas.png")

img_esqueleto = carregar_imagem("esqueleto.png")
img_encapuzado = carregar_imagem("encapuzado.png")
img_goblin = carregar_imagem("goblin.png")
img_aranhao = carregar_imagem("aranhao.png")
img_chefao = carregar_imagem("chefao.png")
img_esqueleto_rh = carregar_imagem("esqueleto.png")  # Placeholder
img_aranhao_ti = carregar_imagem("aranhao.png")  # Placeholder
img_goao_jilherme = carregar_imagem(
    "goao_jilherme.png")  # Imagem do novo chefe
img_orc = carregar_imagem("orc.png")

img_dado = carregar_imagem("dado.png")
img_fundo_batalha = carregar_imagem("fundo_batalha_dungeon.png")


def redimensionar_imagem(imagem, largura, altura):
    if imagem:
        return pygame.transform.smoothscale(imagem, (largura, altura))
    return None


# --- Redimensionamento de Imagens ---
img_jogador_cajado = redimensionar_imagem(img_jogador_cajado, 150, 150)
img_jogador_espadao = redimensionar_imagem(img_jogador_espadao, 150, 150)
img_jogador_adagas = redimensionar_imagem(img_jogador_adagas, 150, 150)
img_esqueleto = redimensionar_imagem(img_esqueleto, 150, 150)
img_encapuzado = redimensionar_imagem(img_encapuzado, 150, 150)
img_goblin = redimensionar_imagem(img_goblin, 150, 150)
img_aranhao = redimensionar_imagem(img_aranhao, 150, 150)
img_chefao = redimensionar_imagem(img_chefao, 150, 150)
img_esqueleto_rh = redimensionar_imagem(img_esqueleto_rh, 150, 150)
img_aranhao_ti = redimensionar_imagem(img_aranhao_ti, 150, 150)
img_goao_jilherme = redimensionar_imagem(
    img_goao_jilherme, 180, 180)  # Redimensiona a imagem do chefe

img_dado = redimensionar_imagem(img_dado, 100, 100)
img_fundo_batalha = redimensionar_imagem(img_fundo_batalha, LARGURA, ALTURA)


class TextoGradual:
    def __init__(self, texto, fonte, cor, velocidade_palavras=10):
        self.texto_completo = texto
        self.palavras = texto.split(" ")
        self.texto_atual = ""
        self.indice_palavra = 0
        self.fonte = fonte
        self.cor = cor
        self.velocidade = velocidade_palavras
        self.concluido = False
        self.ultima_atualizacao = pygame.time.get_ticks()
        self.delay_palavra = 1000 / self.velocidade if self.velocidade > 0 else 0

    def atualizar(self):
        if self.concluido or self.delay_palavra <= 0:
            return
        agora = pygame.time.get_ticks()
        if agora - self.ultima_atualizacao > self.delay_palavra:
            if self.indice_palavra < len(self.palavras):
                self.texto_atual += self.palavras[self.indice_palavra] + " "
                self.indice_palavra += 1
                self.ultima_atualizacao = agora
            else:
                self.concluido = True

    def desenhar(self, tela, x, y, largura_maxima):
        palavras = self.texto_atual.split(" ")
        linhas = []
        linha_atual = ""

        for palavra in palavras:

            if "\n" in palavra:
                partes = palavra.split("\n")
                for i, parte in enumerate(partes):
                    if i > 0:
                        linhas.append(linha_atual.strip())
                        linha_atual = ""

                    texto_teste = linha_atual + parte + " "
                    try:
                        largura_teste, _ = self.fonte.size(texto_teste)
                    except pygame.error:
                        largura_teste = largura_maxima + 1

                    if largura_teste <= largura_maxima:
                        linha_atual = texto_teste
                    else:
                        linhas.append(linha_atual.strip())
                        linha_atual = parte + " "
            else:
                texto_teste = linha_atual + palavra + " "
                try:
                    largura_teste, _ = self.fonte.size(texto_teste)
                except pygame.error:
                    largura_teste = largura_maxima + 1

                if largura_teste <= largura_maxima:
                    linha_atual = texto_teste
                else:
                    linhas.append(linha_atual.strip())
                    linha_atual = palavra + " "

        linhas.append(linha_atual.strip())

        altura_linha = self.fonte.get_height()
        for i, linha in enumerate(linhas):
            try:
                texto_surface = self.fonte.render(linha, True, self.cor)
                tela.blit(texto_surface, (x, y + i * altura_linha))
            except pygame.error as e:
                print(f"Erro ao renderizar texto: {e}")

    def esta_concluido(self):
        return self.concluido

    def concluir(self):
        self.texto_atual = self.texto_completo
        self.indice_palavra = len(self.palavras)
        self.concluido = True


class Botao:
    def __init__(self, x, y, largura, altura, texto, cor_normal, cor_hover, cor_texto, acao=None, fonte=fonte_media):
        self.rect = pygame.Rect(x, y, largura, altura)
        self.texto_original = texto
        self.cor_normal = cor_normal
        self.cor_hover = cor_hover
        self.cor_texto = cor_texto
        self.cor_atual = cor_normal
        self.acao = acao
        self.fonte_original = fonte
        self.fonte_render = fonte
        self.texto_render = texto
        self.ajustar_texto()

    def ajustar_texto(self):
        fonte_atual = self.fonte_original
        texto_atual = self.texto_original
        tamanho_fonte = self.fonte_original.get_height()
        padding = 10

        while tamanho_fonte > 8:
            try:
                largura_texto, altura_texto = fonte_atual.size(texto_atual)
                if largura_texto <= self.rect.width - padding and altura_texto <= self.rect.height - padding:
                    self.fonte_render = fonte_atual
                    self.texto_render = texto_atual
                    return
            except pygame.error:
                pass

            tamanho_fonte -= 1
            try:
                fonte_atual = pygame.font.SysFont(FONTE_PADRAO, tamanho_fonte)
            except pygame.error:
                try:
                    fonte_atual = pygame.font.Font(None, tamanho_fonte)
                except pygame.error:

                    self.fonte_render = fonte_pequena

                    while True:
                        try:
                            largura_texto, _ = self.fonte_render.size(
                                texto_atual + "...")
                            if largura_texto <= self.rect.width - padding or len(texto_atual) <= 1:
                                self.texto_render = texto_atual + \
                                    "..." if len(texto_atual) < len(
                                        self.texto_original) else texto_atual
                                return
                            texto_atual = texto_atual[:-1]
                        except pygame.error:
                            self.texto_render = "?"
                            return

        self.fonte_render = fonte_pequena
        texto_atual = self.texto_original
        while True:
            try:
                largura_texto, _ = self.fonte_render.size(texto_atual + "...")
                if largura_texto <= self.rect.width - padding or len(texto_atual) <= 1:
                    self.texto_render = texto_atual + \
                        "..." if len(texto_atual) < len(
                            self.texto_original) else texto_atual
                    return
                texto_atual = texto_atual[:-1]
            except pygame.error:
                self.texto_render = "?"
                return

    def desenhar(self, tela):
        pygame.draw.rect(tela, self.cor_atual, self.rect, border_radius=10)
        pygame.draw.rect(tela, BRANCO, self.rect, 2, border_radius=10)

        try:
            texto_surface = self.fonte_render.render(
                self.texto_render, True, self.cor_texto)
            texto_rect = texto_surface.get_rect(center=self.rect.center)
            tela.blit(texto_surface, texto_rect)
        except pygame.error as e:
            print(f"Erro ao renderizar texto do botão: {e}")

    def verificar_hover(self, pos_mouse):
        if self.rect.collidepoint(pos_mouse):
            self.cor_atual = self.cor_hover
            return True
        else:
            self.cor_atual = self.cor_normal
            return False

    def verificar_clique(self, pos_mouse):

        if self.rect.collidepoint(pos_mouse):
            return self.acao
        return None


class Barra:
    def __init__(self, x, y, largura, altura, valor_max, cor):
        self.rect = pygame.Rect(x, y, largura, altura)
        self.valor_max = valor_max
        self.valor_atual = valor_max
        self.cor = cor

    def atualizar(self, valor_atual):
        self.valor_atual = max(0, min(valor_atual, self.valor_max))

    def desenhar(self, tela):

        pygame.draw.rect(tela, BRANCO, self.rect, 2)

        largura_preenchimento = int(
            self.rect.width * (self.valor_atual / self.valor_max)) if self.valor_max > 0 else 0
        rect_preenchimento = pygame.Rect(
            self.rect.x, self.rect.y, largura_preenchimento, self.rect.height)
        pygame.draw.rect(tela, self.cor, rect_preenchimento)

        texto = f"{int(self.valor_atual)}/{int(self.valor_max)}"
        try:
            texto_surface = fonte_pequena.render(texto, True, BRANCO)
            texto_rect = texto_surface.get_rect(center=self.rect.center)

            sombra_surface = fonte_pequena.render(texto, True, PRETO)
            tela.blit(sombra_surface, (texto_rect.x + 1, texto_rect.y + 1))
            tela.blit(texto_surface, texto_rect)
        except pygame.error as e:
            print(f"Erro ao renderizar texto da barra: {e}")


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
        self.ataques = {}
        self.habilidades = {}
        self.elementos = {}
        self.definir_ataques_habilidades(arma_inicial)

        if arma_inicial == "espadão":
            self.imagem = img_jogador_espadao
        elif arma_inicial == "duas adagas":
            self.imagem = img_jogador_adagas
        elif arma_inicial == "cajado":
            self.imagem = img_jogador_cajado
        else:
            self.imagem = img_jogador_espadao

    def definir_ataques_habilidades(self, arma):
        self.ataques = {"ataque basico": 0}
        self.habilidades = []
        self.elementos = {}
        if arma == "espadão":
            self.ataques.update({"ataque forte": 3, "ataque giratório": 5,
                                 "pulo com golpe": 3, "golpe em area": 3})
            self.habilidades = [
                "ataque forte", "ataque giratório", "pulo com golpe", "golpe em area"]
        elif arma == "duas adagas":
            self.ataques.update(
                {"back stab": 3, "corte rapido": 3, "dança das adagas": 6})
            self.habilidades = ["back stab",
                                "corte rapido", "dança das adagas"]
        elif arma == "cajado":

            self.elementos = {"fogo": {"bola de fogo": 3, "cuspe do dragão": 5},
                              "gelo": {"floco de gelo": 3, "explosão de gelo": 5},
                              "raio": {"ataque de raio": 3, "choque do trovão": 5}}

            self.habilidades = {}
            for elemento, magias in self.elementos.items():
                self.habilidades[elemento] = list(magias.keys())
                self.ataques.update(magias)

    def ganhar_experiencia(self, exp):
        self.experiencia += exp
        while self.experiencia >= self.experiencia_necessaria:
            pontos_ganhos = 5  # Guarda os pontos antes de zerar
            self.experiencia -= self.experiencia_necessaria
            self.experiencia_necessaria = int(
                self.experiencia_necessaria * 1.5)
            self.subir_nivel(pontos_ganhos)

    def subir_nivel(self, pontos_ganhos):
        self.nivel += 1
        self.pontos_atributo += pontos_ganhos  # Adiciona os pontos ganhos
        self.vida = self.vida_base  # Restaura vida base
        self.mana = self.mana_base  # Restaura mana base

        mensagem = f"\nVocê subiu para o nível {self.nivel}, parabéns!"
        mensagem += "\nSua vida e mana foram totalmente restauradas."
        mensagem += f"\nVocê ganhou {pontos_ganhos} pontos de atributo!"
        exibir_texto(mensagem)

        # Chama a tela de distribuição de pontos
        tela_distribuir_pontos(self)

    # Modificado para aplicar um ponto por vez, chamado pela interface gráfica
    def distribuir_ponto_atributo(self, atributo_escolhido):
        if self.pontos_atributo > 0:
            if atributo_escolhido == "vida":
                self.vida_base += 1
                self.vida += 1
            elif atributo_escolhido == "mana":
                self.mana_base += 1
                self.mana += 1
            elif atributo_escolhido == "vigor":
                self.vigor_base += 1
                self.vigor += 1
            elif atributo_escolhido == "força":
                self.força_base += 1
                self.força += 1
            elif atributo_escolhido == "destreza":
                self.destreza_base += 1
                self.destreza += 1
            elif atributo_escolhido == "inteligencia":
                self.inteligencia_base += 1
                self.inteligencia += 1
            else:
                return False  # Atributo inválido

            self.pontos_atributo -= 1
            return True  # Ponto distribuído com sucesso
        else:
            return False  # Sem pontos para distribuir

    def usar_mana(self, custo):
        if self.mana >= custo:
            self.mana -= custo
            return True
        else:
            return False

    def atacar(self, alvo, ataque="ataque basico"):
        dano = 0
        nome_golpe = ataque
        custo_mana = 0
        mensagem_extra = ""

        if ataque == "ataque basico":
            custo_mana = self.ataques.get(ataque, 0)
            if self.arma_inicial == "espadão":
                dano = (self.força * 0.45) + (self.destreza * 0.05)
            elif self.arma_inicial == "duas adagas":
                dano = (self.destreza * 0.45) + (self.força * 0.05)
                if random.random() < 0.15:
                    dano *= 1.5
                    nome_golpe = "ataque basico (CRÍTICO!)"
            elif self.arma_inicial == "cajado":
                dano = self.inteligencia * 0.3
        elif ataque in self.ataques:
            custo_mana = self.ataques[ataque]
            if self.usar_mana(custo_mana):
                if self.arma_inicial == "espadão":
                    if ataque == "ataque forte":
                        dano = (self.força * 0.7)
                    elif ataque == "ataque giratório":
                        dano = (self.força * 0.5) + (self.vigor * 0.1)
                    elif ataque == "pulo com golpe":
                        dano = (self.força * 0.5) + (self.vigor * 0.1)
                    elif ataque == "golpe em area":
                        dano = (self.força * 0.5) + (self.vigor * 0.1)
                elif self.arma_inicial == "duas adagas":
                    if ataque == "back stab":
                        dano = (self.destreza * 0.6)
                    elif ataque == "corte rapido":
                        dano = (self.destreza * 0.5) + (self.força * 0.1)
                    elif ataque == "dança das adagas":
                        dano = (self.destreza * 0.5)
                elif self.arma_inicial == "cajado":

                    if ataque == "bola de fogo":
                        dano = (self.inteligencia * 0.5)
                    elif ataque == "cuspe do dragão":
                        dano = (self.inteligencia * 0.6)
                    elif ataque == "floco de gelo":
                        dano = (self.inteligencia * 0.5)
                    elif ataque == "explosão de gelo":
                        dano = (self.inteligencia * 0.6)
                    elif ataque == "ataque de raio":
                        dano = (self.inteligencia * 0.5)
                    elif ataque == "choque do trovão":
                        dano = (self.inteligencia * 0.6)
                nome_golpe = ataque
            else:
                return 0, "Mana insuficiente!", 0
        else:
            return 0, "Ataque inválido!", 0

        dano_final = max(1, int(dano)) if dano > 0 else 0
        alvo.vida -= dano_final
        alvo.vida = max(0, alvo.vida)
        return dano_final, nome_golpe, custo_mana


class Monstros:
    def __init__(self, nome, vida, mana, vigor, força, destreza, inteligencia, imagem, exp_concedida=10):
        self.nome = nome
        self.vida_max = vida
        self.vida = vida
        self.mana = mana
        self.vigor = vigor
        self.força = força
        self.destreza = destreza
        self.inteligencia = inteligencia
        self.imagem = imagem
        self.exp_concedida = exp_concedida  # XP que o monstro dá ao ser derrotado

    def atacar(self, personagem):
        dano, mensagem_extra = self.realizar_ataque(personagem)
        # Garante dano mínimo 1 para o monstro também
        dano_final = max(1, int(dano)) if dano > 0 else 0
        personagem.vida -= dano_final
        personagem.vida = max(0, personagem.vida)
        return dano_final, f"{self.nome} {mensagem_extra}, causando {dano_final:.0f} de dano."

    def realizar_ataque(self, personagem):
        # Método base, retorna dano 0 e mensagem padrão
        # As subclasses devem sobrescrever este método
        return 0, "ataca"


class Esqueleto(Monstros):
    def __init__(self):
        super().__init__("Esqueleto", 10, 5, 10, 10, 8, 3, img_esqueleto, exp_concedida=10)

    def realizar_ataque(self, personagem):
        dano = (self.força * 0.5) + (self.vigor * 0.1)
        mensagem = "ataca com sua espada enferrujada"
        if random.random() < 0.1:
            dano *= 1.5
            mensagem = "ataca com um golpe CRÍTICO"
        # Dano é aplicado na função atacar da classe base
        return dano, mensagem


class Encapuzado(Monstros):
    def __init__(self):
        super().__init__("Encapuzado", 12, 10, 8, 6,
                         10, 12, img_encapuzado, exp_concedida=12)

    def realizar_ataque(self, personagem):
        dano = (self.inteligencia * 0.4) + (self.destreza * 0.2)
        mensagem = "lança um feitiço sombrio"
        if random.random() < 0.15:
            mensagem = "lança uma magia venenosa"
        # Dano é aplicado na função atacar da classe base
        return dano, mensagem


class Goblin(Monstros):
    def __init__(self):
        super().__init__("Goblin", 8, 6, 7, 8, 12, 4, img_goblin, exp_concedida=8)

    def realizar_ataque(self, personagem):
        dano = (self.destreza * 0.5) + (self.força * 0.1)
        mensagem = "ataca com suas adagas"
        if random.random() < 0.2:
            dano *= 1.5
            mensagem = "ataca com adagas CRÍTICAS"
        # Dano é aplicado na função atacar da classe base
        return dano, mensagem


class Aranhao(Monstros):
    def __init__(self):
        super().__init__("Aranhao", 16, 8, 12, 7, 11, 5, img_aranhao, exp_concedida=15)

    def realizar_ataque(self, personagem):
        dano = (self.destreza * 0.4) + (self.vigor * 0.2)
        mensagem = "ataca com suas patas afiadas"
        if random.random() < 0.1:
            mensagem = "tece uma teia pegajosa"
        # Dano é aplicado na função atacar da classe base
        return dano, mensagem

# --- Novos Monstros Capítulo 2 ---


class EsqueletoRH(Monstros):
    def __init__(self):
        super().__init__("Esqueleto do RH", 25, 10, 12, 15,
                         10, 8, img_esqueleto_rh, exp_concedida=20)

    def realizar_ataque(self, personagem):
        ataques = [
            ("ataca com um grampeador enferrujado", (self.força * 0.6)),
            ("exige sua assinatura em um documento ilegível",
             (self.inteligencia * 0.3) + (self.destreza * 0.1)),
            ("ameaça cortar seu vale-refeição",
             (self.vigor * 0.2) + (self.força * 0.2))
        ]
        mensagem, dano = random.choice(ataques)
        if random.random() < 0.1:  # Chance de crítico
            dano *= 1.5
            mensagem += " (CRÍTICO!)"
        # Dano é aplicado na função atacar da classe base
        return dano, mensagem


class AranhaoTI(Monstros):
    def __init__(self):
        super().__init__("Aranhao do TI", 30, 15, 10, 10,
                         18, 15, img_aranhao_ti, exp_concedida=25)

    def realizar_ataque(self, personagem):
        ataques = [
            ("lança um cabo de rede emaranhado",
             (self.destreza * 0.5) + (self.inteligencia * 0.1)),
            ("tece uma teia de firewalls impenetráveis",
             (self.inteligencia * 0.4) + (self.vigor * 0.1)),
            ("reseta suas configurações de ataque", (self.destreza * 0.3) +
             (self.inteligencia * 0.3))  # Poderia ter um efeito especial
        ]
        mensagem, dano = random.choice(ataques)
        if random.random() < 0.15:  # Chance de crítico
            dano *= 1.5
            mensagem += " (CRÍTICO!)"
        # Dano é aplicado na função atacar da classe base
        return dano, mensagem


class OrcTrabalhista(Monstros):
    def __init__(self):
        super().__init__("Orc Trabalhista", 40, 15, 10,
                         15, 20, 10, img_orc, exp_concedida=25)

    def realizar_ataque(self, personagem):
        ataques = [
            ("Lança uma maleta cheia de processos",
             (self.destreza * 0.5) + (self.inteligencia * 0.1)),
            ("Arremessa uma pilha de processos para ser digitalizados",
             (self.inteligencia * 0.4) + (self.vigor * 0.1)),
            ("Te obriga a ir fazer o café", (self.destreza * 0.3) +
             (self.inteligencia * 0.3))  # Poderia ter um efeito especial
        ]
        mensagem, dano = random.choice(ataques)
        if random.random() < 0.15:  # Chance de crítico
            dano *= 1.5
            mensagem += " (CRÍTICO!)"
        # Dano é aplicado na função atacar da classe base
        return dano, mensagem

# --- Fim Novos Monstros ---


class Ademiro(Monstros):
    def __init__(self):
        super().__init__("Ademiro", 50, 20, 7, 6, 10, 15, img_chefao, exp_concedida=50)

    def realizar_ataque(self, personagem):
        ataque_escolhido = random.choice(["garra", "magia"])
        dano = 0
        mensagem = ""
        if ataque_escolhido == "garra":
            dano = (self.destreza * 0.3) + (self.força * 0.1)
            mensagem = "ataca com suas garras de ossos"
            if random.random() < 0.15:
                dano *= 1.5
                mensagem = "ataca com garras CRÍTICAS"
        else:
            dano = self.inteligencia * 0.5
            mensagem = "conjura uma magia profana"
        # Dano é aplicado na função atacar da classe base
        return dano, mensagem

# --- Novo Chefe Capítulo 2: Goão Jilherme ---


class GoaoJilherme(Monstros):
    def __init__(self):
        super().__init__("Goão Jilherme", 80, 30, 15, 18,
                         12, 20, img_goao_jilherme, exp_concedida=100)
        self.dialogos_entrada = [
            f"Oxente, {jogador.nome}! Tu pensa que vai aonde, abestalhado?",
            f"Eita! Chegou mais um pra levar uma pisa! Pensou que aqui era a baixa da égua?",
            f"Ó paí! Acunhou demais, agora vai levar o farelo!"
        ]
        self.dialogos_ataque = [
            ("te dá um 'chega pra lá' com a pasta de processos", (self.força * 0.7)),
            ("te enrola com uma conversa mole, um 'desdrobo' danado",
             (self.inteligencia * 0.5) + (self.destreza * 0.2)),
            ("ameaça 'rebolar no mato' suas férias",
             (self.vigor * 0.3) + (self.inteligencia * 0.3)),
            ("grita 'Aí dento!' e te joga um grampeador", (self.destreza * 0.6)),
            ("diz que a situação tá 'barril dobrado' e te dá uma cabeçada",
             (self.força * 0.5) + (self.vigor * 0.2))
        ]
        self.dialogos_derrota_jogador = [
            "Eita ferro! Num aguentou o rojão, foi? Vaza daqui, canelau!",
            "Ficou aí, môca? Te avisei que ia ser barril!",
            "Correu frouxo, é? Vai simbora, abestalhado!"
        ]
        self.dialogos_vitoria_jogador = [
            "Oxente! Mas que surra eu levei...",
            "Aí dento... perdi... Mas volto já, visse?",
            "Eita! Me pegou desprevenido... Pode ir, cabra da peste..."
        ]

    def dialogo_entrada_aleatorio(self):
        return random.choice(self.dialogos_entrada)

    def dialogo_derrota_aleatorio(self):
        return random.choice(self.dialogos_derrota_jogador)

    def dialogo_vitoria_aleatorio(self):
        return random.choice(self.dialogos_vitoria_jogador)

    def realizar_ataque(self, personagem):
        mensagem, dano = random.choice(self.dialogos_ataque)
        if random.random() < 0.2:  # Chance de crítico
            dano *= 1.5
            mensagem += " (Eita, foi CRÍTICO!)"
        # Dano é aplicado na função atacar da classe base
        return dano, mensagem

# --- Fim Novo Chefe ---


def combate(jogador, monstro, resistencias, dado_jogador, tela):
    turno = 0
    mensagens = []
    resultado_combate = None
    valor_dado = None
    acao_jogador_realizada = False
    ataque_escolhido = "ataque basico"
    mostrar_resultado_dado = False
    ultimo_ataque_jogador = ""

    # Cor da barra de vida do jogador
    barra_vida_jogador = Barra(50, 450, 200, 20, jogador.vida_base, VERDE)
    barra_mana_jogador = Barra(50, 480, 200, 20, jogador.mana_base, AZUL)

    # Posição e cor ajustadas
    barra_vida_monstro = Barra(575, 205, 200, 20, monstro.vida_max, VERMELHO)
    barra_vida_monstro.atualizar(monstro.vida)

    botoes_ataque = []
    y_botao = 510
    x_botao_inicial = 50
    largura_botao = 140
    altura_botao = 30
    espaco_botao = 10
    botoes_por_linha = 5

    botoes_ataque.append(Botao(x_botao_inicial, y_botao, largura_botao, altura_botao,
                         f"Ataque Básico (0)", AZUL, (100, 100, 255), BRANCO, acao="ataque basico", fonte=fonte_pequena))

    habilidades_disponiveis = []
    if jogador.arma_inicial == "cajado":
        for elemento, magias in jogador.habilidades.items():
            habilidades_disponiveis.extend(magias)
    else:
        habilidades_disponiveis = jogador.habilidades

    for i, hab in enumerate(habilidades_disponiveis):
        custo = jogador.ataques.get(hab, 0)
        linha = (i + 1) // botoes_por_linha
        coluna = (i + 1) % botoes_por_linha
        botoes_ataque.append(Botao(
            x_botao_inicial + coluna * (largura_botao + espaco_botao),
            y_botao + linha * (altura_botao + espaco_botao),
            largura_botao, altura_botao,
            f"{hab} ({custo})", AZUL, (100, 100, 255), BRANCO, acao=hab, fonte=fonte_pequena
        ))

    botao_dado = Botao(LARGURA // 2 - 75, 350, 150, 40, "Rolar Dado",
                       VERMELHO, (255, 100, 100), BRANCO, acao="rolar_dado")

    botao_passar = Botao(LARGURA // 2 - 75, 400, 150, 40, "Passar Turno",
                         AMARELO, (255, 255, 100), PRETO, acao="passar_turno")

    rodando = True
    clock = pygame.time.Clock()
    while rodando:

        # Verifica condição de fim de combate ANTES de processar eventos ou desenhar
        if jogador.vida <= 0:
            if isinstance(monstro, GoaoJilherme):
                mensagens.append(monstro.dialogo_derrota_aleatorio())
            else:
                mensagens.append(f"Você foi derrotado pelo {monstro.nome}!")
            resultado_combate = "derrota"
            rodando = False
            break  # Sai do loop imediatamente
        elif monstro.vida <= 0:
            if isinstance(monstro, GoaoJilherme):
                mensagens.append(monstro.dialogo_vitoria_aleatorio())
            else:
                mensagens.append(f"Você derrotou o {monstro.nome}!")
            jogador.ganhar_experiencia(
                monstro.exp_concedida)  # Usa XP do monstro
            resultado_combate = "vitoria"
            rodando = False
            break  # Sai do loop imediatamente

        barra_vida_jogador.atualizar(jogador.vida)
        barra_mana_jogador.atualizar(jogador.mana)
        barra_vida_monstro.atualizar(monstro.vida)

        mouse_pos = pygame.mouse.get_pos()
        for evento in pygame.event.get():
            if evento.type == pygame.QUIT:
                pygame.quit()
                sys.exit()

            if turno % 2 == 0 and rodando and evento.type == pygame.MOUSEBUTTONDOWN and evento.button == 1:
                acao_clicada = None

                for botao in botoes_ataque:
                    acao_botao = botao.verificar_clique(mouse_pos)
                    if acao_botao:
                        acao_clicada = acao_botao
                        ataque_escolhido = acao_botao
                        mensagens.append(
                            f"Ataque/Habilidade: {ataque_escolhido}")
                        ultimo_ataque_jogador = ataque_escolhido
                        break

                if not acao_clicada:
                    acao_dado = botao_dado.verificar_clique(mouse_pos)
                    if acao_dado and ultimo_ataque_jogador:
                        acao_clicada = acao_dado
                        valor_dado = random.choice(dado_jogador)
                        mensagens.append(f"Você rolou o dado: {valor_dado}")
                        mostrar_resultado_dado = True
                        acao_jogador_realizada = False

                        if int(valor_dado) >= 3:
                            mensagens.append("Sucesso na rolagem do dado!")
                            dano, golpe, custo = jogador.atacar(
                                monstro, ultimo_ataque_jogador)

                            if golpe == "Mana insuficiente!":
                                mensagens.append(golpe)
                            elif golpe == "Ataque inválido!":
                                mensagens.append(golpe)
                            elif golpe:
                                tipo = jogador.arma_inicial
                                inimigo = monstro.nome.lower()
                                dano_modificado = dano  # Dano já é o final calculado

                                # Aplica resistência se houver
                                if tipo in resistencias and inimigo in resistencias[tipo]:
                                    resistencia = resistencias[tipo][inimigo]
                                    # Ajuste na lógica de resistência: dano_modificado *= (1 - resistencia) se for % ou dano_modificado /= resistencia se for multiplicador
                                    # Assumindo que o valor em 'resistencias' é um multiplicador (0.9 = 10% resistente, 1.1 = 10% vulnerável)
                                    dano_modificado /= resistencia
                                    # Garante dano mínimo 1 pós-resistência
                                    dano_modificado = max(
                                        1, int(dano_modificado)) if dano_modificado > 0 else 0
                                    mensagens.append(
                                        f"({monstro.nome} tem resistência/vulnerabilidade: {resistencia:.1f}x -> Dano: {dano_modificado:.0f})")
                                else:
                                    # Garante dano mínimo 1 se não houver resistência
                                    dano_modificado = max(
                                        1, int(dano)) if dano > 0 else 0

                                # Aplica o dano final ao monstro (já feito na função atacar)
                                # monstro.vida -= dano_modificado # Linha removida, dano já aplicado em atacar()
                                # monstro.vida = max(0, monstro.vida) # Linha removida, já feito em atacar()

                                mensagens.append(f"Seu golpe ")
                                mensagens.append(
                                    f"\'{golpe}\' causou {dano_modificado:.0f} de dano.")
                                if custo > 0:
                                    mensagens.append(f"(Custo: {custo} Mana)")
                                mensagens.append(
                                    f"{monstro.nome} vida restante: {monstro.vida:.0f}")
                                acao_jogador_realizada = True
                            else:
                                mensagens.append(
                                    "O ataque falhou misteriosamente!")
                                acao_jogador_realizada = True
                        else:
                            mensagens.append(
                                "Falha na rolagem do dado! Seu ataque não teve efeito.")
                            acao_jogador_realizada = True

                        # Verifica fim de combate APÓS o ataque do jogador
                        if monstro.vida <= 0:
                            if isinstance(monstro, GoaoJilherme):
                                mensagens.append(
                                    monstro.dialogo_vitoria_aleatorio())
                            else:
                                mensagens.append(
                                    f"Você derrotou o {monstro.nome}!")
                            jogador.ganhar_experiencia(monstro.exp_concedida)
                            resultado_combate = "vitoria"
                            rodando = False
                            # break # Não precisa de break aqui, o loop principal vai parar
                        elif acao_jogador_realizada:
                            turno += 1
                            ultimo_ataque_jogador = ""

                    elif acao_dado and not ultimo_ataque_jogador:
                        mensagens.append(
                            "Escolha um ataque/habilidade primeiro!")

                if not acao_clicada:
                    acao_passar = botao_passar.verificar_clique(mouse_pos)
                    if acao_passar:
                        acao_clicada = acao_passar
                        mensagens.append("Você passou seu turno.")
                        turno += 1
                        mostrar_resultado_dado = False
                        ultimo_ataque_jogador = ""

            if evento.type == pygame.MOUSEMOTION:
                for botao in botoes_ataque:
                    botao.verificar_hover(mouse_pos)
                botao_dado.verificar_hover(mouse_pos)
                botao_passar.verificar_hover(mouse_pos)

        # Turno do Monstro
        if turno % 2 == 1 and rodando:
            pygame.time.delay(1000)
            dano, msg_monstro = monstro.atacar(jogador)
            mensagens.append(msg_monstro)
            mensagens.append(f"Sua vida restante: {jogador.vida:.0f}")

            # Verifica fim de combate APÓS o ataque do monstro
            if jogador.vida <= 0:
                if isinstance(monstro, GoaoJilherme):
                    mensagens.append(monstro.dialogo_derrota_aleatorio())
                else:
                    mensagens.append(
                        f"Você foi derrotado pelo {monstro.nome}!")
                resultado_combate = "derrota"
                rodando = False
                # break # Não precisa de break aqui, o loop principal vai parar
            else:
                turno += 1
                valor_dado = None
                mostrar_resultado_dado = False

        if len(mensagens) > 6:
            mensagens = mensagens[-6:]

        # --- Desenho --- (Só desenha se o jogo ainda estiver rodando)
        if rodando:
            # Desenha o fundo da batalha primeiro
            if img_fundo_batalha:
                tela.blit(img_fundo_batalha, (0, 0))
            else:
                tela.fill(PRETO)  # Fallback para cor sólida se a imagem falhar

            # Desenha os personagens e outros elementos sobre o fundo
            if jogador.imagem:
                tela.blit(jogador.imagem, (50, 300))
            if monstro.imagem:
                # Ajusta a posição Y do chefe para melhor encaixe
                pos_y_monstro = 30 if isinstance(monstro, GoaoJilherme) else 50
                tela.blit(monstro.imagem, (600, pos_y_monstro))

            if img_dado:
                tela.blit(img_dado, (LARGURA // 2 - 50, 250))

            if valor_dado and mostrar_resultado_dado:
                try:
                    texto_dado = fonte_grande.render(
                        str(valor_dado), True, BRANCO)
                    dado_rect = texto_dado.get_rect(center=(LARGURA // 2, 230))

                    sombra_dado = fonte_grande.render(
                        str(valor_dado), True, CINZA_ESCURO)
                    tela.blit(sombra_dado, (dado_rect.x + 2, dado_rect.y + 2))
                    tela.blit(texto_dado, dado_rect)
                except pygame.error as e:
                    print(f"Erro ao renderizar valor do dado: {e}")

            barra_vida_jogador.desenhar(tela)
            barra_mana_jogador.desenhar(tela)
            barra_vida_monstro.desenhar(tela)

            try:
                texto_jogador = fonte_media.render(
                    # Cor alterada para verde
                    f"{jogador.nome} (Nível {jogador.nivel})", True, VERDE)
                texto_monstro = fonte_media.render(
                    monstro.nome, True, VERMELHO)  # Cor alterada para vermelho
                tela.blit(texto_jogador, (50, 420))
                # Ajusta posição Y do nome do monstro
                pos_y_nome_monstro = 185
                if isinstance(monstro, GoaoJilherme):
                    pos_y_nome_monstro = 215  # Desce um pouco para não sobrepor a imagem maior
                tela.blit(texto_monstro, (575, pos_y_nome_monstro))
            except pygame.error as e:
                print(f"Erro ao renderizar nomes: {e}")

            # Caixa de mensagens semi-transparente
            caixa_msg_rect = pygame.Rect(40, 40, LARGURA - 400, 150)
            s = pygame.Surface(
                (caixa_msg_rect.width, caixa_msg_rect.height), pygame.SRCALPHA)
            s.fill((50, 50, 50, 180))  # Cinza escuro com alpha
            tela.blit(s, (caixa_msg_rect.x, caixa_msg_rect.y))
            pygame.draw.rect(tela, CINZA_CLARO, caixa_msg_rect, 1)  # Borda

            for i, msg in enumerate(mensagens):
                try:
                    texto_msg = fonte_pequena.render(msg, True, BRANCO)
                    tela.blit(texto_msg, (caixa_msg_rect.x + 10,
                              caixa_msg_rect.y + 10 + i * 20))
                except pygame.error as e:
                    print(f"Erro ao renderizar mensagem: {e}")

            if turno % 2 == 0 and rodando:
                for botao in botoes_ataque:
                    botao.desenhar(tela)
                botao_dado.desenhar(tela)
                botao_passar.desenhar(tela)

            pygame.display.flip()
            clock.tick(30)
        # --- Fim Desenho ---

    # --- Pós-Combate ---
    # Exibe a última tela com a mensagem final (vitória/derrota)
    if img_fundo_batalha:
        tela.blit(img_fundo_batalha, (0, 0))
    else:
        tela.fill(PRETO)
    if jogador.imagem:
        tela.blit(jogador.imagem, (50, 300))
    if monstro.imagem:
        pos_y_monstro = 30 if isinstance(monstro, GoaoJilherme) else 50
        tela.blit(monstro.imagem, (600, pos_y_monstro))

    barra_vida_jogador.desenhar(tela)
    barra_mana_jogador.desenhar(tela)
    barra_vida_monstro.desenhar(tela)
    try:
        texto_jogador = fonte_media.render(
            f"{jogador.nome} (Nível {jogador.nivel})", True, VERDE)
        texto_monstro = fonte_media.render(monstro.nome, True, VERMELHO)
        tela.blit(texto_jogador, (50, 420))
        pos_y_nome_monstro = 185
        if isinstance(monstro, GoaoJilherme):
            pos_y_nome_monstro = 215
        tela.blit(texto_monstro, (575, pos_y_nome_monstro))
    except pygame.error as e:
        print(f"Erro ao renderizar nomes pós-combate: {e}")

    caixa_msg_rect = pygame.Rect(40, 40, LARGURA - 400, 150)
    s = pygame.Surface(
        (caixa_msg_rect.width, caixa_msg_rect.height), pygame.SRCALPHA)
    s.fill((50, 50, 50, 180))
    tela.blit(s, (caixa_msg_rect.x, caixa_msg_rect.y))
    pygame.draw.rect(tela, CINZA_CLARO, caixa_msg_rect, 1)
    for i, msg in enumerate(mensagens[-6:]):  # Mostra as últimas mensagens
        try:
            texto_msg = fonte_pequena.render(msg, True, BRANCO)
            tela.blit(texto_msg, (caixa_msg_rect.x + 10,
                      caixa_msg_rect.y + 10 + i * 20))
        except pygame.error as e:
            print(f"Erro ao renderizar mensagem pós-combate: {e}")

    pygame.display.flip()
    # Aumenta o delay para ler a mensagem final, especialmente do chefe
    pygame.time.delay(4000)
    return resultado_combate


def exibir_texto(texto, cor=BRANCO, velocidade_palavras=10):
    texto_obj = TextoGradual(texto, fonte_media, cor, velocidade_palavras)
    continuar_botao = Botao(LARGURA // 2 - 75, ALTURA - 60, 150, 40,
                            "Continuar", AZUL, (100, 100, 255), BRANCO, acao="continuar")

    rodando = True
    clock = pygame.time.Clock()
    while rodando:
        mouse_pos = pygame.mouse.get_pos()
        for evento in pygame.event.get():
            if evento.type == pygame.QUIT:
                pygame.quit()
                sys.exit()

            if evento.type == pygame.MOUSEBUTTONDOWN and evento.button == 1:

                if not texto_obj.esta_concluido():
                    texto_obj.concluir()

                elif continuar_botao.verificar_clique(mouse_pos) == "continuar":
                    rodando = False

            if evento.type == pygame.MOUSEMOTION:
                continuar_botao.verificar_hover(mouse_pos)

            if evento.type == pygame.KEYDOWN:
                if evento.key == pygame.K_SPACE or evento.key == pygame.K_RETURN:
                    if not texto_obj.esta_concluido():
                        texto_obj.concluir()
                    else:
                        rodando = False

        texto_obj.atualizar()

        tela.fill(PRETO)
        texto_obj.desenhar(tela, 50, 50, LARGURA - 100)

        if texto_obj.esta_concluido():
            continuar_botao.desenhar(tela)

        pygame.display.flip()
        clock.tick(30)

# --- Nova Função: Tela de Distribuição de Pontos ---


def tela_distribuir_pontos(jogador):
    atributos = ["vida", "mana", "vigor", "força", "destreza", "inteligencia"]
    nomes_atributos = {
        "vida": "Vida", "mana": "Mana", "vigor": "Vigor",
        "força": "Força", "destreza": "Destreza", "inteligencia": "Inteligência"
    }
    botoes_atributo = []
    y_inicial = 150
    x_botao_mais = 450
    largura_botao_mais = 40
    altura_botao = 30

    for i, attr in enumerate(atributos):
        botoes_atributo.append(Botao(x_botao_mais, y_inicial + i * (altura_botao + 15),
                                     largura_botao_mais, altura_botao, "+",
                                     VERDE, (100, 255, 100), BRANCO, acao=f"add_{attr}", fonte=fonte_grande))

    botao_confirmar = Botao(LARGURA // 2 - 75, ALTURA - 80, 150, 40,
                            "Confirmar", AZUL, (100, 100, 255), BRANCO, acao="confirmar")

    rodando = True
    clock = pygame.time.Clock()
    while rodando:
        mouse_pos = pygame.mouse.get_pos()
        for evento in pygame.event.get():
            if evento.type == pygame.QUIT:
                pygame.quit()
                sys.exit()

            if evento.type == pygame.MOUSEBUTTONDOWN and evento.button == 1:
                acao_clicada = None
                for botao in botoes_atributo:
                    acao = botao.verificar_clique(mouse_pos)
                    if acao and acao.startswith("add_"):
                        atributo_escolhido = acao.split("_")[1]
                        if jogador.distribuir_ponto_atributo(atributo_escolhido):
                            # Ponto adicionado, a interface será atualizada no loop de desenho
                            pass
                        else:
                            # Feedback visual/sonoro de que não há mais pontos?
                            pass
                        acao_clicada = acao
                        break

                if not acao_clicada:
                    acao_confirmar = botao_confirmar.verificar_clique(
                        mouse_pos)
                    if acao_confirmar == "confirmar":
                        # Permite confirmar mesmo com pontos restantes
                        rodando = False

            if evento.type == pygame.MOUSEMOTION:
                for botao in botoes_atributo:
                    botao.verificar_hover(mouse_pos)
                botao_confirmar.verificar_hover(mouse_pos)

        tela.fill(PRETO)

        # Título
        texto_titulo_surf = fonte_grande.render(
            "Distribua seus Pontos!", True, AMARELO)
        titulo_rect = texto_titulo_surf.get_rect(center=(LARGURA // 2, 50))
        tela.blit(texto_titulo_surf, titulo_rect)

        # Pontos restantes
        texto_pontos_surf = fonte_media.render(
            f"Pontos restantes: {jogador.pontos_atributo}", True, BRANCO)
        pontos_rect = texto_pontos_surf.get_rect(center=(LARGURA // 2, 100))
        tela.blit(texto_pontos_surf, pontos_rect)

        # Atributos e botões
        for i, attr in enumerate(atributos):
            nome_attr = nomes_atributos[attr]
            valor_base = getattr(jogador, f"{attr}_base")
            texto_attr_surf = fonte_media.render(
                f"{nome_attr}: {valor_base}", True, BRANCO)
            tela.blit(texto_attr_surf, (LARGURA // 2 - 150,
                      y_inicial + i * (altura_botao + 15) + 5))
            # Desabilita botão '+' se não houver pontos
            if jogador.pontos_atributo > 0:
                botoes_atributo[i].cor_normal = VERDE
                botoes_atributo[i].cor_hover = (100, 255, 100)
            else:
                botoes_atributo[i].cor_normal = CINZA_ESCURO
                botoes_atributo[i].cor_hover = CINZA_ESCURO
            botoes_atributo[i].desenhar(tela)

        # Botão Confirmar
        botao_confirmar.desenhar(tela)

        pygame.display.flip()
        clock.tick(30)


def obter_escolha(pergunta, opcoes):
    texto_obj = TextoGradual(pergunta, fonte_media, BRANCO, 10)
    botoes = []

    altura_total_botoes = len(opcoes) * 50 - 10
    y_inicial_botoes = (ALTURA - altura_total_botoes) // 2 + 50

    for i, opcao_texto in enumerate(opcoes):
        acao_botao = opcao_texto.lower().replace(" ", "_")  # Ação baseada no texto
        botoes.append(Botao(LARGURA // 2 - 100, y_inicial_botoes + i * 50, 200,
                      40, opcao_texto, AZUL, (100, 100, 255), BRANCO, acao=acao_botao))

    escolha_feita = None
    rodando = True
    clock = pygame.time.Clock()
    while rodando:
        mouse_pos = pygame.mouse.get_pos()
        for evento in pygame.event.get():
            if evento.type == pygame.QUIT:
                pygame.quit()
                sys.exit()

            if evento.type == pygame.MOUSEBUTTONDOWN and evento.button == 1:

                if not texto_obj.esta_concluido():
                    texto_obj.concluir()

                else:
                    for botao in botoes:
                        acao_clicada = botao.verificar_clique(mouse_pos)
                        if acao_clicada:
                            escolha_feita = acao_clicada
                            rodando = False
                            break

            if evento.type == pygame.MOUSEMOTION:
                for botao in botoes:
                    botao.verificar_hover(mouse_pos)

            if evento.type == pygame.KEYDOWN:
                if (evento.key == pygame.K_SPACE or evento.key == pygame.K_RETURN) and not texto_obj.esta_concluido():
                    texto_obj.concluir()

        texto_obj.atualizar()

        tela.fill(PRETO)
        texto_obj.desenhar(tela, 50, 50, LARGURA - 100)

        if texto_obj.esta_concluido():
            for botao in botoes:
                botao.desenhar(tela)

        pygame.display.flip()
        clock.tick(30)

    return escolha_feita


def obter_nome():
    nome = ""
    texto_obj = TextoGradual(
        "Escolha o nome do jogador:", fonte_media, BRANCO, 10)
    botao_confirmar = Botao(LARGURA // 2 - 75, 300, 150, 40,
                            "Confirmar", VERDE, (100, 255, 100), BRANCO, acao="confirmar")
    input_rect = pygame.Rect(LARGURA // 2 - 150, 200, 300, 40)
    ativo = True

    rodando = True
    clock = pygame.time.Clock()
    while rodando:
        mouse_pos = pygame.mouse.get_pos()
        for evento in pygame.event.get():
            if evento.type == pygame.QUIT:
                pygame.quit()
                sys.exit()

            if ativo and evento.type == pygame.KEYDOWN:
                if evento.key == pygame.K_BACKSPACE:
                    nome = nome[:-1]
                elif evento.key == pygame.K_RETURN:
                    if nome.strip():
                        rodando = False
                elif evento.key == pygame.K_SPACE:
                    if nome and nome[-1] != " ":  # Evita espaços duplos
                        nome += " "
                else:
                    # Permite letras, números e espaço (mas não no início)
                    if len(nome) < 20 and (evento.unicode.isalnum() or (evento.unicode == " " and nome)):
                        nome += evento.unicode

            if evento.type == pygame.MOUSEBUTTONDOWN and evento.button == 1:

                if not texto_obj.esta_concluido():
                    texto_obj.concluir()

                elif botao_confirmar.verificar_clique(mouse_pos) == "confirmar" and nome.strip():
                    rodando = False

                elif input_rect.collidepoint(mouse_pos):
                    ativo = True
                else:
                    ativo = False

            if evento.type == pygame.MOUSEMOTION:
                botao_confirmar.verificar_hover(mouse_pos)

        texto_obj.atualizar()

        tela.fill(PRETO)
        texto_obj.desenhar(tela, 50, 50, LARGURA - 100)

        cor_borda_input = BRANCO if ativo else (100, 100, 100)
        pygame.draw.rect(tela, cor_borda_input, input_rect, 2)
        try:
            texto_nome_surface = fonte_media.render(nome, True, BRANCO)
            tela.blit(texto_nome_surface, (input_rect.x + 5, input_rect.y + 5))
        except pygame.error as e:
            print(f"Erro ao renderizar nome: {e}")

        # Cursor piscando
        if ativo and pygame.time.get_ticks() % 1000 < 500:
            try:
                cursor_pos = input_rect.x + 5 + texto_nome_surface.get_width()
                pygame.draw.line(
                    tela, BRANCO, (cursor_pos, input_rect.y + 5), (cursor_pos, input_rect.y + 35), 2)
            except NameError:  # Se texto_nome_surface ainda não foi criado
                pygame.draw.line(
                    tela, BRANCO, (input_rect.x + 5, input_rect.y + 5), (input_rect.x + 5, input_rect.y + 35), 2)
            except pygame.error:
                pass

        if texto_obj.esta_concluido():
            botao_confirmar.desenhar(tela)

        pygame.display.flip()
        clock.tick(30)

    return nome.strip()


def menu_principal():
    botao_novo_jogo = Botao(LARGURA // 2 - 100, 250, 200, 50,
                            "Novo Jogo", VERDE, (100, 255, 100), BRANCO, acao="novo_jogo")
    botao_sair = Botao(LARGURA // 2 - 100, 320, 200, 50,
                       "Sair", VERMELHO, (255, 100, 100), BRANCO, acao="sair")
    botoes = [botao_novo_jogo, botao_sair]

    tocar_musica(musica_jogo, volume=0.3)

    rodando = True
    clock = pygame.time.Clock()
    while rodando:
        mouse_pos = pygame.mouse.get_pos()
        for evento in pygame.event.get():
            if evento.type == pygame.QUIT:
                pygame.quit()
                sys.exit()

            if evento.type == pygame.MOUSEBUTTONDOWN and evento.button == 1:
                for botao in botoes:
                    acao = botao.verificar_clique(mouse_pos)
                    if acao == "novo_jogo":
                        return True
                    elif acao == "sair":
                        pygame.quit()
                        sys.exit()

            if evento.type == pygame.MOUSEMOTION:
                for botao in botoes:
                    botao.verificar_hover(mouse_pos)

        tela.fill(PRETO)

        try:
            texto_titulo = fonte_titulo.render("A Lenda do CLT", True, BRANCO)
            titulo_rect = texto_titulo.get_rect(center=(LARGURA // 2, 150))

            sombra_titulo = fonte_titulo.render(
                "A Lenda do CLT", True, CINZA_ESCURO)
            tela.blit(sombra_titulo, (titulo_rect.x + 3, titulo_rect.y + 3))
            tela.blit(texto_titulo, titulo_rect)
        except pygame.error as e:
            print(f"Erro ao renderizar título: {e}")

        for botao in botoes:
            botao.desenhar(tela)

        pygame.display.flip()
        clock.tick(30)


# Variável global para o jogador, para ser acessível na classe GoaoJilherme
jogador = None


def jogo_principal():
    global jogador  # Declara que vamos usar a variável global

    tocar_musica(musica_jogo, volume=0.3)

    nome_jogador = obter_nome()
    if not nome_jogador:
        return

    arma_escolhida = obter_escolha("Escolha sua arma inicial:", [
                                   "Espadão", "Duas Adagas", "Cajado"])
    if not arma_escolhida:
        return

    # Instancia o jogador na variável global
    if arma_escolhida == "espadão":
        jogador = Personagem(nome_jogador, 25, 11, 11, 20, 6, 4, "espadão")
    elif arma_escolhida == "duas_adagas":  # Corrigido para usar o valor retornado por obter_escolha
        jogador = Personagem(nome_jogador, 22, 12, 10, 8, 20, 6, "duas adagas")
    elif arma_escolhida == "cajado":
        jogador = Personagem(nome_jogador, 18, 15, 8, 5, 4, 20, "cajado")
    else:
        print(f"Escolha de arma inválida: {arma_escolhida}")
        return

    # --- Capítulo 1 --- (Adicionado)
    exibir_texto("Capítulo 1: O Chamado do Estagiário", cor=AMARELO)

    exibir_texto(
        f" A história de {jogador.nome} se inicia nos reinos de Ass Alariados, O Reino dos CLT'S. Sendo apresentado como o novo integrante da DIAFI. ")

    resistencias = {
        "espadão": {"esqueleto": 0.9, "encapuzado": 1.1, "goblin": 0.9, "aranhao": 1.0, "ademiro": 1.0, "esqueleto do rh": 0.9, "aranhao do ti": 1.1, "goão jilherme": 1.0},
        "cajado": {"esqueleto": 1.1, "encapuzado": 0.9, "goblin": 1.1, "aranhao": 0.8, "ademiro": 1.0, "esqueleto do rh": 1.2, "aranhao do ti": 0.8, "goão jilherme": 1.1},
        "duas adagas": {"esqueleto": 1.0, "encapuzado": 1.0, "goblin": 1.0, "aranhao": 1.2, "ademiro": 1.0, "esqueleto do rh": 1.0, "aranhao do ti": 1.0, "goão jilherme": 0.9}
    }

    inimigos_cap1 = [Esqueleto, Encapuzado, Goblin, Aranhao]
    inimigo1_aleatorio = random.choice(inimigos_cap1)
    inimigo1_instanciado = inimigo1_aleatorio()
    dado_jogador = ["1", "2", "3", "4", "5", "6"]

    exibir_texto(
        f" Após ser escolhido para a sua primeira jornada, {jogador.nome} decide partir imediatamente.\nVocê encontra uma caverna e sente calafrios só de olhar para dentro dela. \nMas você decide entrar...\n Caminhando mais a frente, {jogador.nome} se depara com uma porta enorme, escrito NAJ ")

    decisao_jogador = obter_escolha("Deseja entrar?", ["Sim", "Não"])
    if not decisao_jogador:
        return

    if decisao_jogador == "sim":
        exibir_texto(
            f"Ao ultrapassar a porta, {jogador.nome} se depara com um {inimigo1_instanciado.nome}!")
        exibir_texto("Iniciando combate!")

        resultado = combate(jogador, inimigo1_instanciado,
                            resistencias, dado_jogador, tela)

        if resultado == "derrota":
            exibir_texto("Você foi derrotado! Fim de jogo.")
            return
        elif resultado is None:  # Caso o jogador feche a janela durante o combate
            return

        exibir_texto(
            "Após derrotar o inimigo, você se deparada com uma porta a direita e um caminho a sua frente.")
        escolha_caminho = obter_escolha("Por onde deseja prosseguir?", [
                                        "Entrar na porta a direita", "Seguir em frente"])
        if not escolha_caminho:
            return

        if escolha_caminho == "entrar_na_porta_a_direita":  # Corrigido para usar o valor retornado

            inimigo2_aleatorio = random.choice(inimigos_cap1)
            inimigo2_instanciado = inimigo2_aleatorio()

            exibir_texto(
                "Você decide entrar na porta a direita... \nAo olhar no fundo da sala, você encontra um baú")
            decisao_bau = obter_escolha("Deseja abrir o baú?", ["Sim", "Não"])
            if not decisao_bau:
                return

            if decisao_bau == "sim":
                exibir_texto("Abrindo o baú...")
                exibir_texto(
                    f"Era uma armadilha, você encontrou um {inimigo2_instanciado.nome}.")
                exibir_texto("Iniciando combate!")

                resultado = combate(
                    jogador, inimigo2_instanciado, resistencias, dado_jogador, tela)

                if resultado == "derrota":
                    exibir_texto("Você foi derrotado! Fim de jogo.")
                    return
                elif resultado is None:
                    return

                exibir_texto("Após uma luta intensa, algo na parede se mexe, chama sua atenção, pois não parecia comum.\nAo chegar perto, você percebe uma anomalia.\nInvestigando parede...\nUma passagem se abre...\nVocê encontra uma jovem garota, aparentemente perdida. \nGarota Desconhecida: Pode... mee... ajudar?")
                decisao_jovem = obter_escolha(
                    "Deseja ajudá-la?", ["Sim", "Não"])
                if not decisao_jovem:
                    return

                if decisao_jovem == "sim":
                    inimigo4_aleatorio = random.choice(inimigos_cap1)
                    inimigo4_instanciado = inimigo4_aleatorio()

                    exibir_texto(
                        "Chegando perto da garota, ela se revela ser um monstro. ")
                    exibir_texto("Iniciando combate!")

                    resultado = combate(
                        jogador, inimigo4_instanciado, resistencias, dado_jogador, tela)

                    if resultado == "derrota":
                        exibir_texto("Você foi derrotado! Fim de jogo.")
                        return
                    elif resultado is None:
                        return

                    exibir_texto(
                        "Por que uma garota estaria sozinha e perdida dentro dessa caverna? você se pergunta. \nNão tem problema, vi alguém em perigo e decidi ajudar, não posso me culpar pela armadilha. ")
                    exibir_texto(
                        f"Então {jogador.nome} decide seguir a diante...")

                else:  # Não ajudou a garota
                    exibir_texto(
                        "Você decide não ajudar a garota e segue em frente.")

                # Caminho comum após a sala do baú/garota ou se não ajudou a garota
                exibir_texto(
                    "Você sente calafrios, a cada passo que da, sente um terrível poder a sua frente. \nUma porta enorme está no fim do corredor, e você percebe que é de lá que vem seus calafrios. ")
                exibir_texto(
                    "Ao abrir a porta, você se depara com Ademiro Santiro, Chefe dos Estagiários. ")
                exibir_texto(
                    f"Ademiro: Achei que nunca iria chegar, {jogador.nome} \nAdemiro: Como ousa invadir meu santuário? \nAdemiro: Não permitirei que tal ato de rebeldia saia impune. \nAdemiro: EU TE FAREI MEU ESTAGIÁRIO {jogador.nome} MUAHAHAHAHAHAHAHAHA")

                tocar_musica(musica_chefao, forcar_troca=True)
                chefao_ademiro = Ademiro()  # Instancia Ademiro aqui
                resultado_ademiro = combate(
                    jogador, chefao_ademiro, resistencias, dado_jogador, tela)
                tocar_musica(musica_jogo, forcar_troca=True)

                if resultado_ademiro == "derrota":
                    exibir_texto(
                        "Você foi derrotado pelo chefão! Fim de jogo.")
                    return
                elif resultado_ademiro == "vitoria":
                    exibir_texto(f"Parabéns, {jogador.nome}! Você derrotou Ademiro Santiro!\nSua primeira jornada foi travada com muita luta e muita habilidade.\nVocê retornará para a DIAFI como um herói, pois finalmente nasceu alguém que pudesse combater este mal.\nVoltando para DIAFI...")
                    # --- Continuação da História (Capítulo 2) --- (Adicionado)
                    iniciar_capitulo_2(
                        jogador, resistencias, dado_jogador, tela)
                    return  # Fim do jogo após capítulo 2
                else:  # Se o jogador fechou a janela durante a luta contra Ademiro
                    return

            else:  # Não abriu o baú
                exibir_texto("Você decide não abrir o baú e segue em frente.")
                # Caminho para o chefão
                exibir_texto(
                    "Você sente calafrios, a cada passo que da, sente um terrível poder a sua frente. \nUma porta enorme está no fim do corredor, e você percebe que é de lá que vem seus calafrios. ")
                exibir_texto(
                    "Ao abrir a porta, você se depara com Ademiro Santiro, Chefe dos Estagiários. ")
                exibir_texto(
                    f"Ademiro: Achei que nunca iria chegar, {jogador.nome} \nAdemiro: Como ousa invadir meu santuário? \nAdemiro: Não permitirei que tal ato de rebeldia saia impune. \nAdemiro: EU TE FAREI MEU ESTAGIÁRIO {jogador.nome} MUAHAHAHAHAHAHAHAHA")

                tocar_musica(musica_chefao, forcar_troca=True)
                chefao_ademiro = Ademiro()  # Instancia Ademiro aqui
                resultado_ademiro = combate(
                    jogador, chefao_ademiro, resistencias, dado_jogador, tela)
                tocar_musica(musica_jogo, forcar_troca=True)

                if resultado_ademiro == "derrota":
                    exibir_texto(
                        "Você foi derrotado pelo chefão! Fim de jogo.")
                    return
                elif resultado_ademiro == "vitoria":
                    exibir_texto(f"Parabéns, {jogador.nome}! Você derrotou Ademiro Santiro!\nSua primeira jornada foi travada com muita luta e muita habilidade.\nVocê retornará para a DIAFI como um herói, pois finalmente nasceu alguém que pudesse combater este mal.\nVoltando para DIAFI...")
                    # --- Continuação da História (Capítulo 2) --- (Adicionado)
                    iniciar_capitulo_2(
                        jogador, resistencias, dado_jogador, tela)
                    return  # Fim do jogo após capítulo 2
                else:
                    return

        elif escolha_caminho == "seguir_em_frente":  # Corrigido para usar o valor retornado
            inimigo3_aleatorio = random.choice(inimigos_cap1)
            inimigo3_instanciado = inimigo3_aleatorio()

            exibir_texto(
                f"Você decide seguir em frente... \nAo virar o corredor, você se depara com um {inimigo3_instanciado.nome}.")
            exibir_texto("Iniciando combate!")

            resultado = combate(jogador, inimigo3_instanciado,
                                resistencias, dado_jogador, tela)

            if resultado == "derrota":
                exibir_texto("Você foi derrotado! Fim de jogo.")
                return
            elif resultado is None:
                return

            # Caminho para o chefão
            exibir_texto(
                "Você sente calafrios, a cada passo que da, sente um terrível poder a sua frente. \nUma porta enorme está no fim do corredor, e você percebe que é de lá que vem seus calafrios. ")
            exibir_texto(
                "Ao abrir a porta, você se depara com Ademiro Santiro, Chefe dos Estagiários. ")
            exibir_texto(
                f"Ademiro: Achei que nunca iria chegar, {jogador.nome} \nAdemiro: Como ousa invadir meu santuário? \nAdemiro: Não permitirei que tal ato de rebeldia saia impune. \nAdemiro: EU TE FAREI MEU ESTAGIÁRIO {jogador.nome} MUAHAHAHAHAHAHAHAHA")

            tocar_musica(musica_chefao, forcar_troca=True)
            chefao_ademiro = Ademiro()  # Instancia Ademiro aqui
            resultado_ademiro = combate(
                jogador, chefao_ademiro, resistencias, dado_jogador, tela)
            tocar_musica(musica_jogo, forcar_troca=True)

            if resultado_ademiro == "derrota":
                exibir_texto("Você foi derrotado pelo chefão! Fim de jogo.")
                return
            elif resultado_ademiro == "vitoria":
                exibir_texto(f"Parabéns, {jogador.nome}! Você derrotou Ademiro Santiro!\nSua primeira jornada foi travada com muita luta e muita habilidade.\nVocê retornará para a DIAFI como um herói, pois finally nasceu alguém que pudesse combater este mal.\nVoltando para DIAFI...")
                # --- Continuação da História (Capítulo 2) --- (Adicionado)
                iniciar_capitulo_2(jogador, resistencias, dado_jogador, tela)
                return  # Fim do jogo após capítulo 2
            else:
                return

    else:  # Não entrou na porta NAJ
        exibir_texto(
            "Você decide não entrar e volta para a entrada da caverna, mas algo bloqueia seu caminho...")
        inimigo_covarde = random.choice(inimigos_cap1)()
        exibir_texto(f"Um {inimigo_covarde.nome} aparece!")
        resultado = combate(jogador, inimigo_covarde,
                            resistencias, dado_jogador, tela)
        if resultado == "derrota":
            exibir_texto("Você foi derrotado por fugir! Fim de jogo.")
        else:
            exibir_texto(
                "Você derrotou o monstro, mas não ache que sua jornada terminou.")
        return

# --- Nova Função: Capítulo 2 ---


def iniciar_capitulo_2(jogador, resistencias, dado_jogador, tela):
    exibir_texto("Capítulo 2: A Vingança do Departamento", cor=AMARELO)
    exibir_texto(
        f"Após {jogador.nome} voltar como um herói para a Diafi, Julianiri Chefiri já o manda para um novo combate. O herói ainda tem mais batalhas a travar...")
    exibir_texto(
        f"Julianiri Chefiri: Excelente trabalho, {jogador.nome}! Mas não há tempo para descanso! Novos problemas surgiram nos departamentos vizinhos.")
    exibir_texto("Julianiri Chefiri: O Departamento de RH está assombrado por esqueletos burocráticos e o TI está infestado por aranhas de rede! Contamos com você!")

    inimigos_cap2 = [EsqueletoRH, AranhaoTI]

    # Encontro 1
    exibir_texto("Você se dirige ao sombrio Departamento de RH...")
    inimigo_rh = EsqueletoRH()
    exibir_texto(f"Um {inimigo_rh.nome} surge das pilhas de papelada!")
    resultado_rh = combate(jogador, inimigo_rh,
                           resistencias, dado_jogador, tela)
    if resultado_rh == "derrota":
        exibir_texto("A burocracia venceu... Fim de jogo.")
        return
    elif resultado_rh is None:
        return
    exibir_texto("Você superou a papelada infernal!")

    # Encontro 2
    exibir_texto(
        "Agora, você adentra os corredores cabeados do Departamento de TI...")
    inimigo_ti = AranhaoTI()
    exibir_texto(f"Cuidado! Um {inimigo_ti.nome} desce do teto!")
    resultado_ti = combate(jogador, inimigo_ti,
                           resistencias, dado_jogador, tela)
    if resultado_ti == "derrota":
        exibir_texto("Erro 404: Herói não encontrado... Fim de jogo.")
        return
    elif resultado_ti is None:
        return
    exibir_texto("Bug corrigido! Você limpou a área.")

    # Encontro 3
    exibir_texto("Após passar pelo Departamento de Ti, você avista uma sala com baixa iluminação. \nAo chegar perto da sala, você percebe um ser estranho, enfurecido, cheio de processos na mesa e esperando algum estagiário para obrigá-lo a fazeru seu trabalho")
    inimigo_vice_diretor = OrcTrabalhista()
    resultado_trabalhista = combate(
        jogador, inimigo_vice_diretor, resistencias, dado_jogador, tela)
    if resultado_trabalhista == "derrota":
        exibir_texto("Faça meu trabalho, IMEDIATAMENTE!!!... Fim de jogo.")
        return
    elif resultado_trabalhista is None:
        return
    exibir_texto(
        f"{jogador.nome} Jamais farei o trabalho por você, recebendo menos do que deveria.")
    # Preparação para o Chefão do Capítulo 2
    exibir_texto(f"Julianiri Chefiri: Incrível, {jogador.nome}! Mas o verdadeiro problema ainda reside nas sombras... \nJulianiri Chefiri: O líder dessa bagunça toda, Goão Jilherme, está camuflado na sala da diretoria! \nJulianiri Chefiri: Dizem que ele é meio... preguiçoso, mas não se deixe enganar. Vá e derrote o chefe.")

    # Encontro com o Chefão Goão Jilherme
    exibir_texto(
        "Você chega à imponente porta da diretoria. Um cheiro de café forte e processos acumulados paira no ar. \nFAÇA, APENAS FAÇA. FAÇA MINHA EMPRESA LUCRAR BRILHÕES E VOCÊ NÃO RECEBERÁ NADA EM TROCA HAHAHHAHAHA.")
    chefao_jilherme = GoaoJilherme()
    # Diálogo de entrada
    exibir_texto(chefao_jilherme.dialogo_entrada_aleatorio())
    exibir_texto(f"Prepare-se para enfrentar {chefao_jilherme.nome}!")

    tocar_musica(musica_chefao, forcar_troca=True,
                 volume=0.4)  # Toca música de chefão
    resultado_jilherme = combate(
        jogador, chefao_jilherme, resistencias, dado_jogador, tela)
    tocar_musica(musica_jogo, forcar_troca=True)  # Volta para música normal

    if resultado_jilherme == "derrota":
        # A mensagem de derrota já é exibida dentro da função combate com o diálogo do chefe
        exibir_texto(
            "Goão Jilherme provou ser barril dobrado... Fim de jogo.")
        return
    elif resultado_jilherme == "vitoria":
        # A mensagem de vitória já é exibida dentro da função combate com o diálogo do chefe
        exibir_texto(
            f"Parabéns, {jogador.nome}! Você derrotou Goão Jilherme e trouxe ordem aos departamentos!\nSua lenda como o Herói CLT cresce a cada dia! \nVocê começa a inspirar gerações, a servir de exemplo para meros trabalhores comuns.")
        exibir_texto(
            "         ------------- FIM DO CAPÍTULO 2 -------------        \n------------- OBRIGADO POR JOGAR! -------------")
    elif resultado_jilherme is None:
        return  # Jogador fechou a janela


if __name__ == "__main__":
    if menu_principal():
        jogo_principal()
    pygame.quit()
    sys.exit()
