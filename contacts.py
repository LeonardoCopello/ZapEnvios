# cria classe para armazenar contatos para os quais serão enviadas as mensagens
class Contact:
    def __init__(self, nome, mensagem, arquivo, imagem, telefone, birthday):
        """

        :rtype: object
        """
        self.__nome = nome
        self.__mensagem = mensagem
        self.__arquivo = arquivo
        self.__imagem = imagem
        self.__telefone = telefone
        self.__birthday = birthday

    @property
    def nome(self):
        return self.__nome
    @nome.setter
    def nome(self, nome):
        self.__nome = nome
    @property
    def mensagem(self):
        return self.__mensagem
    @mensagem.setter
    def mensagem(self, mensagem):
        self.__mensagem = mensagem
    @property
    def arquivo(self):
        return self.__arquivo
    @arquivo.setter
    def arquivo(self, arquivo):
        self.__arquivo = arquivo
    @property
    def imagem(self):
        return self.__imagem
    @imagem.setter
    def imagem(self, imagem):
        self.__imagem = imagem
    @property
    def telefone(self):
        return self.__telefone
    @telefone.setter
    def telefone(self, telefone):
        self.__telefone = telefone
    @property
    def birthday(self):
        return self.__birthday
    @telefone.setter
    def telefone(self, birthday):
        self.__birthday = birthday