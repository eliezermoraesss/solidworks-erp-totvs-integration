-- predefined type, no DDL - MDSYS.SDO_GEOMETRY
-- predefined type, no DDL - XMLTYPE

CREATE TABLE perfil_acesso (
    id_perfil     SERIAL PRIMARY KEY,
    ds_autoridade VARCHAR(20) NOT NULL
);

CREATE TABLE tb_adocoes (
    id_adotante                SERIAL PRIMARY KEY,
    dt_adocao                  DATE NOT NULL,
    id_usuario                 INTEGER NOT NULL,
    ds_observacao              VARCHAR(2000),
    FOREIGN KEY (id_usuario) REFERENCES tb_usuarios(id_usuario)
);

CREATE TABLE tb_animais (
    id_animal              SERIAL PRIMARY KEY,
    nm_nome                VARCHAR(50) NOT NULL,
    nr_idade               INTEGER NOT NULL,
    ds_especie             VARCHAR(10) NOT NULL,
    st_adotado             VARCHAR(1) NOT NULL,
    tb_adocoes_id_adotante INTEGER NOT NULL,
    ds_observacao          VARCHAR(2000),
    tb_doacoes_id_doador   INTEGER NOT NULL,
    FOREIGN KEY (tb_adocoes_id_adotante) REFERENCES tb_adocoes(id_adotante),
    FOREIGN KEY (tb_doacoes_id_doador) REFERENCES tb_doacoes(id_doador)
);

CREATE TABLE tb_doacoes (
    id_doador                  SERIAL PRIMARY KEY,
    dt_doacao                  DATE NOT NULL,
    id_usuario                 INTEGER NOT NULL,
    ds_observacao              VARCHAR(2000),
    FOREIGN KEY (id_usuario) REFERENCES tb_usuarios(id_usuario)
);

CREATE TABLE tb_enderecos (
    id_endereco                 SERIAL PRIMARY KEY,
    nr_cep                      VARCHAR(10) NOT NULL,
    ds_logradouro               VARCHAR(50) NOT NULL,
    nr_numero                   INTEGER NOT NULL,
    ds_complemento              VARCHAR(20),
    ds_cidade                   VARCHAR(20) NOT NULL,
    sg_estado                   VARCHAR(2) NOT NULL,
    ds_pais                     VARCHAR(20) NOT NULL,
	id_usuario					INTEGER NOT NULL,
	FOREIGN KEY (id_usuario) REFERENCES tb_usuarios(id_usuario)
);

CREATE TABLE tb_posts (
    id_post                    SERIAL PRIMARY KEY,
    ds_mensagem                VARCHAR(1000) NOT NULL,
    id_usuario                 INTEGER NOT NULL,
    nm_curtidas                INTEGER,
    FOREIGN KEY (id_usuario) REFERENCES tb_usuarios(id_usuario)
);

CREATE TABLE tb_respostas (
    id_resposta                SERIAL PRIMARY KEY,
    ds_msg_resposta            VARCHAR(1000) NOT NULL,
    tb_posts_id_post           INTEGER NOT NULL,
    id_usuario                 INTEGER NOT NULL,
    nm_curtidas                INTEGER,
    FOREIGN KEY (tb_posts_id_post) REFERENCES tb_posts(id_post),
    FOREIGN KEY (id_usuario) REFERENCES tb_usuarios(id_usuario)
);

CREATE TABLE tb_telefones (
    id_telefone                       SERIAL PRIMARY KEY,
    nr_telefone                       BIGINT NOT NULL,
    id_usuario                        INTEGER NOT NULL,
    tb_tipo_telefone_id_tipo_telefone INTEGER NOT NULL,
    FOREIGN KEY (tb_tipo_telefone_id_tipo_telefone) REFERENCES tb_tipo_telefone(id_tipo_telefone),
    FOREIGN KEY (id_usuario) REFERENCES tb_usuarios(id_usuario)
);

CREATE TABLE tb_tipo_telefone (
    id_tipo_telefone SERIAL PRIMARY KEY,
    ds_tipo          VARCHAR(10) NOT NULL
);

CREATE TABLE tb_usuarios (
    id_usuario               SERIAL PRIMARY KEY,
    ds_nome                  VARCHAR(50) NOT NULL,
    ds_genero                VARCHAR(10) NOT NULL,
    dt_nascimento            DATE NOT NULL,
    st_doador                VARCHAR(1),
    ds_email                 VARCHAR(50) NOT NULL,
    ds_senha                 VARCHAR(30) NOT NULL
);

CREATE TABLE usuario_possuir_perfil_acesso (
    tb_usuarios_id_usuario     INTEGER NOT NULL,
    perfil_acesso_id_perfil    INTEGER NOT NULL,
    tb_usuarios_tb_usuarios_id INTEGER NOT NULL,
    PRIMARY KEY (tb_usuarios_id_usuario, perfil_acesso_id_perfil, tb_usuarios_tb_usuarios_id),
    FOREIGN KEY (perfil_acesso_id_perfil) REFERENCES perfil_acesso(id_perfil),
    FOREIGN KEY (tb_usuarios_tb_usuarios_id) REFERENCES tb_usuarios(tb_usuarios_id)
);

CREATE SEQUENCE tb_usuarios START 0;

CREATE OR REPLACE FUNCTION trigger_tb_usuarios()
RETURNS TRIGGER AS $$
BEGIN
    IF NEW.id_usuario IS NULL THEN
        NEW.id_usuario := NEXTVAL('tb_usuarios');
    END IF;
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER tb_usuarios
BEFORE INSERT ON tb_usuarios
FOR EACH ROW
EXECUTE FUNCTION trigger_tb_usuarios();
