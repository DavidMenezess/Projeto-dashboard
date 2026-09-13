"""
criar_usuario.py

Script de linha de comando para criar (ou atualizar) um usuário que pode
logar no dashboard. O login é feito com Nome + Sobrenome (não e-mail).

Por que existe um script em vez de só a rota da API:
A rota POST /usuarios só pode ser usada por um administrador já logado —
ou seja, alguém precisa ser o primeiro administrador. Este script cria
esse primeiro usuário (ou qualquer outro, se for mais prático rodar
direto no servidor). Depois que existir ao menos um administrador, o
resto da equipe pode ser cadastrado tanto por aqui quanto pela rota
POST /usuarios (usada pela própria tela de "Usuários" no dashboard).

Uso:
  python -m scripts.criar_usuario
  (vai perguntar nome, sobrenome, cargo, tipo de acesso e senha, interativamente)
"""

import getpass
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app.database import SessionLocal, Usuario, criar_tabelas
from app.auth import gerar_hash_senha, gerar_chave_login


def main():
    criar_tabelas()
    sessao = SessionLocal()

    primeiro_nome = input("Nome: ").strip()
    sobrenome = input("Sobrenome: ").strip()
    cargo = input("Cargo no escritório (ex: Advogado, Estagiário, Sócio): ").strip()

    tipo = ""
    while tipo not in ("admin", "normal"):
        tipo = input("Tipo de acesso ('admin' ou 'normal'): ").strip().lower()

    senha = getpass.getpass("Senha (não aparece na tela): ")
    confirmacao = getpass.getpass("Confirme a senha: ")

    if senha != confirmacao:
        print("As senhas não coincidem. Nada foi salvo.")
        return

    if len(senha) < 8:
        print("Use uma senha com pelo menos 8 caracteres. Nada foi salvo.")
        return

    chave_login = gerar_chave_login(primeiro_nome, sobrenome)
    usuario_existente = sessao.query(Usuario).filter(Usuario.chave_login == chave_login).first()

    if usuario_existente:
        usuario_existente.senha_hash = gerar_hash_senha(senha)
        usuario_existente.cargo = cargo
        usuario_existente.tipo = tipo
        sessao.commit()
        print(f"Usuário '{chave_login}' atualizado com sucesso (tipo: {tipo}).")
        print(f"Login: Nome = '{primeiro_nome}', Sobrenome = '{sobrenome}'")
    else:
        novo_usuario = Usuario(
            primeiro_nome=primeiro_nome, sobrenome=sobrenome, chave_login=chave_login,
            senha_hash=gerar_hash_senha(senha), cargo=cargo, tipo=tipo, ativo=1,
        )
        sessao.add(novo_usuario)
        sessao.commit()
        print(f"Usuário criado com sucesso (tipo: {tipo}).")
        print(f"Login: Nome = '{primeiro_nome}', Sobrenome = '{sobrenome}'")

    sessao.close()


if __name__ == "__main__":
    main()
