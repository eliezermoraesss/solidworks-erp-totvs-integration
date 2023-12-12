-- Inserir registros na tabela perfil_acesso
INSERT INTO perfil_acesso (ds_autoridade) VALUES 
('Admin'), ('Moderador'), ('Usuário Padrão'), ('Visitante'), 
('Editor'), ('Supervisor'), ('Convidado'), ('Analista');

-- Inserir registros na tabela tb_usuarios
INSERT INTO tb_usuarios (ds_nome, ds_genero, dt_nascimento, st_doador, ds_email, ds_senha) VALUES 
('João Silva', 'Masculino', '1990-01-15', 'N', 'joao.silva@email.com', 'senha123'),
('Maria Oliveira', 'Feminino', '1985-05-22', 'S', 'maria.oliveira@email.com', 'senha456'),
('Carlos Santos', 'Masculino', '1995-08-10', 'N', 'carlos.santos@email.com', 'senha789'),
('Ana Pereira', 'Feminino', '1992-03-30', 'S', 'ana.pereira@email.com', 'senhaabc'),
('Lucas Rodrigues', 'Masculino', '1988-11-18', 'N', 'lucas.rodrigues@email.com', 'senhaxyz'),
('Camila Souza', 'Feminino', '1987-07-25', 'S', 'camila.souza@email.com', 'senhajkl'),
('Fernando Lima', 'Masculino', '1998-04-05', 'N', 'fernando.lima@email.com', 'senha321'),
('Mariana Costa', 'Feminino', '1993-09-12', 'S', 'mariana.costa@email.com', 'senhawxy');

-- Inserir registros na tabela tb_enderecos
INSERT INTO tb_enderecos (nr_cep, ds_logradouro, nr_numero, ds_complemento, ds_cidade, sg_estado, ds_pais, id_usuario) VALUES 
('12345-678', 'Rua A', 123, 'Apto 101', 'Cidade A', 'CA', 'Brasil', 1),
('54321-876', 'Avenida B', 456, 'Casa 202', 'Cidade B', 'CB', 'Brasil', 2),
('98765-432', 'Travessa C', 789, 'Sala 303', 'Cidade C', 'CC', 'Brasil', 3),
('45678-901', 'Rua D', 234, 'Casa 404', 'Cidade D', 'CD', 'Brasil', 4),
('21098-765', 'Avenida E', 567, 'Apartamento 505', 'Cidade E', 'CE', 'Brasil', 5),
('87654-321', 'Rua F', 890, 'Casa 606', 'Cidade F', 'CF', 'Brasil', 6),
('34567-890', 'Travessa G', 123, 'Casa 707', 'Cidade G', 'CG', 'Brasil', 7),
('78901-234', 'Avenida H', 456, 'Apartamento 808', 'Cidade H', 'CH', 'Brasil', 8);

-- Inserir registros na tabela tb_doacoes
INSERT INTO tb_doacoes (dt_doacao, id_usuario, ds_observacao, tb_usuarios_tb_usuarios_id) VALUES 
('2023-01-05', 1, 'Doação para animais abandonados', 1),
('2023-02-10', 2, 'Doação mensal para ONG', 2),
('2023-03-15', 3, 'Doação de ração para cães', 3),
('2023-04-20', 4, 'Doação para campanha de adoção', 4),
('2023-05-25', 5, 'Doação para resgate de animais', 5),
('2023-06-30', 6, 'Doação para tratamento veterinário', 6),
('2023-07-05', 7, 'Doação para abrigo de gatos', 7),
('2023-08-10', 8, 'Doação para castração de animais', 8);

-- Inserir registros na tabela tb_adocoes
INSERT INTO tb_adocoes (dt_adocao, id_usuario, ds_observacao, tb_usuarios_tb_usuarios_id) VALUES 
('2023-01-01', 1, 'Adoção de um gato preto', 1),
('2023-02-02', 2, 'Adoção de um cachorro vira-lata', 2),
('2023-03-03', 3, 'Adoção de um coelho', 3),
('2023-04-04', 4, 'Adoção de um pássaro', 4),
('2023-05-05', 5, 'Adoção de um peixe', 5),
('2023-06-06', 6, 'Adoção de um hamster', 6),
('2023-07-07', 7, 'Adoção de uma iguana', 7),
('2023-08-08', 8, 'Adoção de um porquinho-da-índia', 8);

-- Inserir registros na tabela tb_animais
INSERT INTO tb_animais (nm_nome, nr_idade, ds_especie, st_adotado, tb_adocoes_id_adotante, ds_observacao, tb_doacoes_id_doador) VALUES 
('Frajola', 2, 'Gato', 'S', 1, 'Frajola é um gato muito brincalhão e carinhoso.', 1),
('Totó', 3, 'Cachorro', 'S', 2, 'Totó adora correr e brincar no parque.', 2),
('Pernalonga', 1, 'Coelho', 'S', 3, 'Pernalonga é um coelho dócil e amigável.', 3),
('Piu-Piu', 1, 'Pássaro', 'S', 4, 'Piu-Piu canta lindas melodias pela manhã.', 4),
('Nemo', 2, 'Peixe', 'S', 5, 'Nemo é um peixinho colorido e cheio de energia.', 5),
('Bolt', 4, 'Hamster', 'S', 6, 'Bolt adora correr na sua rodinha durante a noite.', 6),
('Luna', 3, 'Iguana', 'S', 7, 'Luna é uma iguana tranquila e curiosa.', 7),
('Bolinha', 2, 'Porquinho-da-índia', 'S', 8, 'Bolinha é um porquinho-da-índia adorável.', 8);

-- Inserir registros na tabela tb_posts
INSERT INTO tb_posts (ds_mensagem, id_usuario, nm_curtidas, tb_usuarios_tb_usuarios_id) VALUES 
('Olá, pessoal! Estou muito feliz com a adoção do Frajola!', 1, 10, 1),
('Conheçam o novo membro da família: Totó! 🐾', 2, 15, 2),
('Meu coelhinho Pernalonga é a coisa mais fofa do mundo!', 3, 8, 3),
('Piu-Piu, o pássaro cantor! 🎶', 4, 12, 4),
('Nemo, o peixinho colorido, está fazendo a festa no aquário!', 5, 18, 5),
('Bolt, o hamster veloz, conquistando corações.', 6, 7, 6),
('Luna, a iguana curiosa, explorando o território.', 7, 22, 7),
('Bolinha, o porquinho-da-índia, fazendo graça!', 8, 5, 8);

-- Inserir registros na tabela tb_respostas
INSERT INTO tb_respostas (ds_msg_resposta, tb_posts_id_post, id_usuario, nm_curtidas, tb_usuarios_tb_usuarios_id) VALUES 
('Que legal, João! Frajola é um gato encantador!', 1, 2, 5, 2),
('Totó é lindo demais! Parabéns pela adoção, Maria!', 2, 1, 8, 1),
('Pernalonga é uma fofura! 😍', 3, 4, 3, 4),
('Piu-Piu tem um canto incrível mesmo!', 4, 3, 10, 3),
('Nemo está maravilhoso! Adoro peixes coloridos.', 5, 6, 15, 6),
('Bolt é fofo demais! Tenho um hamster também.', 6, 5, 2, 5),
('Luna parece muito curiosa mesmo. Que legal!', 7, 8, 18, 8),
('Bolinha é adorável! Tenho um porquinho-da-índia também.', 8, 7, 7, 7);

-- Inserir registros na tabela tb_telefones
INSERT INTO tb_telefones (nr_telefone, id_usuario, tb_tipo_telefone_id_tipo_telefone, tb_usuarios_tb_usuarios_id) VALUES 
(987654321, 1, 1, 1),
(123456789, 2, 2, 2),
(555555555, 3, 1, 3),
(999999999, 4, 2, 4),
(111111111, 5, 1, 5),
(777777777, 6, 2, 6),
(333333333, 7, 1, 7),
(888888888, 8, 2, 8);
