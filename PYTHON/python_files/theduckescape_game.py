import pygame
import random

# Inicialização do Pygame
pygame.init()

# Configurações da tela
WIDTH = 800
HEIGHT = 400
screen = pygame.display.set_mode((WIDTH, HEIGHT))
pygame.display.set_caption("Jogo do Pato Saltitante")

# Cores
BLUE = (135, 206, 235)
DARK_BLUE = (70, 130, 180)
BLACK = (0, 0, 0)
GREEN = (0, 255, 0)

# Pato
duck = pygame.Rect(50, HEIGHT - 60, 40, 40)
duck_y_velocity = 0
JUMP_STRENGTH = 15
GRAVITY = 0.8

# Jacarés
alligators = []
ALLIGATOR_WIDTH = 60
ALLIGATOR_HEIGHT = 30
alligator_speed = 5

# Pontuação
score = 0
high_score = 0
font = pygame.font.Font(None, 36)

# Clock
clock = pygame.time.Clock()

# Loop principal do jogo
running = True
while running:
    for event in pygame.event.get():
        if event.type == pygame.QUIT:
            running = False
        if event.type == pygame.KEYDOWN:
            if event.key == pygame.K_SPACE and duck.bottom >= HEIGHT - 40:
                duck_y_velocity = -JUMP_STRENGTH

    # Atualiza a posição do pato
    duck_y_velocity += GRAVITY
    duck.y += duck_y_