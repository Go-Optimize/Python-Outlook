
![Logo](https://dev-to-uploads.s3.amazonaws.com/uploads/articles/th5xamgrr6se0x5ro4g6.png)


# Integração Python + Outlook para envio de e-mails

Este repositório oferece uma solução simples para integração entre Python e o Microsoft Outlook, permitindo o envio automatizado de e-mails diretamente de scripts Python. Com essa integração, os usuários podem incorporar facilmente funcionalidades de envio de e-mails em seus aplicativos Python, automatizando processos de comunicação por e-mail.


## Introdução
Primeiro precisamos entender a estrutura de um e-mail e os parâmetros que podem ser adicionados:
![App Screenshot](https://via.placeholder.com/468x300?text=App+Screenshot+Here)

Seguindo a estrutura, os atributos que você deve especificar são:
- Remetente (por padrão o e-mail principal é utilizado, mas em caso de múltiplos e-mails logados no Outlook é possível especificar qual você gostaria de utilizar)
- Destinatários (obrigatório)
- Destinatários em CC
- Título 
- Corpo do E-mail
- Imagem incorporada no Corpo do E-mail (prints, assinaturas, etc.)
- Anexos


## Integração

O módulo que usaremos para a integração entre as duas ferramentas é o [win32com](https://pypi.org/project/pywin32/), ele fornece uma forma de interagir com uma tecnologia da Microsoft (COM - Component Object Model) que permite a comunicação entre diferentes programas e componentes do ambiente Windows. 

Com o módulo win32com é possível acessar recursos como o Microsoft Office, Outlook, IExplorer e diversos aplicativos Windows, Sendo frequentemente usado para automatização de tarefas repetitivas.

## Instalação

Para utlizar o módulo win32com é necessário instalar a biblioteca que ele está inserido, neste caso a biblioteca pywin32.

Para instalar uma biblioteca basta acessar o "Prompt de Comando" e colar o código abaixo:

```
pip install pywin32
```


## Documentação
O primeiro passo é importar a o módulo para que possamos utilizar suas funcionalidades:

```python
import win32com.client
```

Agora para acessar o Outlook usaremos o método Dispatch que é usado para instanciar um programa do sistema Windows:

```python
outlook = win32com.client.Dispatch('Outlook.Application')
```

Feito isso nós vamos criar um novo item usando o método CreateItem, existem alguns tipos de itens que podemos criar, como:

- CreateItem(0): Cria um novo e-mail. Este é o valor padrão quando nenhum tipo é especificado.
- CreateItem(1): Cria uma nova reunião.
- CreateItem(2): Cria uma nova tarefa.
- CreateItem(3): Cria um novo contato.
- CreateItem(4): Cria um novo lembrete.
- CreateItem(5): Cria um novo registro de diário.

Como criaremos um novo e-mail vamos usar: 
```python
email = outlook.CreateItem(0)
```
