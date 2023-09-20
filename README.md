# Monitor de E-mails no Outlook com Pushbullet

Este é um programa Python que monitora a caixa de entrada do Outlook em busca de novos e-mails não lidos e os envia como notificações para o seu celular usando o Pushbullet.

## Pré-requisitos

Antes de usar este programa, certifique-se de ter o seguinte instalado:

- Python 3.x (https://www.python.org/downloads/)
- Bibliotecas Python: pywin32, pushbullet.py

Você pode instalar as bibliotecas Python necessárias usando o pip:

```bash
pip install pywin32 pushbullet.py
```

## Configuração

1. Certifique-se de ter uma conta no Pushbullet (https://www.pushbullet.com/).
2. Obtenha seu token de API do Pushbullet seguindo as instruções fornecidas pelo Pushbullet.
3. Substitua `'seu_api_key'` pelo seu token de API do Pushbullet no código Python.

## Como usar

Execute o programa Python usando o seguinte comando:

```bash
python notificacao.py
```

O programa irá monitorar sua caixa de entrada do Outlook em intervalos regulares. Quando um novo e-mail não lido é encontrado, ele enviará uma notificação para o seu celular com o conteúdo do e-mail e o remetente.

## Notas

- Certifique-se de que o Outlook esteja em execução enquanto o programa estiver sendo executado.
- O código foi desenvolvido para funcionar com o Outlook, mas pode ser ajustado para outros programas de e-mail, se necessário.

**Aviso:** Este código é fornecido apenas para fins educacionais e de demonstração. Certifique-se de usar este programa de acordo com as políticas de privacidade e segurança de sua organização e de acordo com as leis locais aplicáveis.
```
