# Automacao de Baixa de NF - Dealer.net

Ferramenta de automacao para dar baixa em Notas Fiscais no Dealer.net (Grupo Indiana). Le um arquivo Excel com as NFs e valores, abre o navegador, e faz a baixa automaticamente de cada NF.

## Funcionalidades

- Interface web local para controle da automacao
- Suporte para 3 lojas: Mandarim Iguatemi, Itabuna e Lauro de Freitas
- Upload de arquivo Excel direto pela interface
- Monitoramento em tempo real das baixas
- Relatorio com graficos de pizza
- Aba de analise (apenas consulta, sem fazer baixa)
- Exportacao de relatorios em Excel
- Pausa/Retomada da automacao

## Requisitos

- Python 3.12 ou superior
- Windows 10/11
- Acesso ao Dealer.net do Grupo Indiana

## Instalacao

### 1. Instalar Python

Baixe e instale o Python 3.12 do [python.org](https://www.python.org/downloads/).

Ou via winget:
```
winget install Python.Python.3.12
```

**Importante:** Durante a instalacao, marque "Add Python to PATH".

### 2. Clonar o repositorio

```
git clone https://github.com/Krozs1607/baixa-nf.git
cd baixa-nf
```

### 3. Instalar dependencias

```
pip install -r requirements.txt
```

### 4. Instalar o Chromium do Playwright

```
python -m playwright install chromium
```

## Como usar

### Metodo 1: Atalho (Recomendado)

1. Abra o arquivo `Iniciar Automacao.bat` clicando 2 vezes
2. A janela do terminal vai abrir e o navegador vai abrir automaticamente em `http://localhost:5000`

### Metodo 2: Via terminal

```
python servidor.py
```

Depois abra `http://localhost:5000` no seu navegador.

## Fluxo de uso

1. **Selecionar loja** no dropdown (Iguatemi, Itabuna ou Lauro de Freitas)
2. **Fazer upload do Excel** com as NFs a dar baixa
3. **Clicar em "Configurar Dealer"**
4. **Clicar em "Comecar Baixas"** - vai abrir o Chromium
5. **Fazer login no Dealer.net** manualmente
6. **Selecionar a loja** no Dealer
7. **Ir em Financeiro > Contas a Receber > Titulo a Receber**
8. **Expandir o Filtro Avancado**
9. **Clicar em "INICIAR BAIXAS"** no painel web
10. Acompanhar o progresso em tempo real

## Estrutura do Excel de entrada

O Excel deve ter as seguintes colunas:

| CNPJ | Referencia (NF) | Texto | Total Geral |
|------|----------------|-------|-------------|
| 42.616.980/0001-00 | 2461 | VLR.REF.NF 2461... | -92038,48 |

- Coluna A: CNPJ
- Coluna B: Numero da NF (Referencia)
- Coluna C: Texto descritivo
- Coluna D: Valor (pode ser negativo)

## Observacoes importantes

- **Mandarim Itabuna e Lauro de Freitas**: as NFs sao formatadas automaticamente para 15 digitos com prefixo "2026" (ex: NF 132 vira 202600000000132)
- **Mandarim Iguatemi**: usa o numero da NF original
- O valor da baixa e sempre positivo (mesmo que no Excel esteja negativo)
- O Historico e preenchido automaticamente como "Baixa Garantia Nota: XXX"
- A ferramenta pula automaticamente NFs que ja estao pagas
- Os arquivos Excel enviados ficam em `uploads/` (nao sao versionados)

## Campos preenchidos no Dealer

- **Tipo de Credito/Debito**: RECEBIMENTO DE TITULO
- **Tipo de Documento**: AVISO DE LANCAMENTO
- **Agente Cobrador**: CONTA MOVIMENTO FABRICA (3.06.60)
- **Historico**: Baixa Garantia Nota: [numero da NF]
- **Valor**: valor absoluto do Excel

## Suporte

Em caso de problemas:
1. Verifique se o servidor esta rodando (janela preta aberta)
2. Verifique os logs na propria interface web
3. Reinicie o servidor fechando e reabrindo o .bat
