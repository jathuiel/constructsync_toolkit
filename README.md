# Construct Sync Toolkit

Plugin para **Autodesk Navisworks Simulate 2026** que permite gerenciar atributos personalizados em elementos BIM diretamente pelo ambiente Navisworks.

---

## Funcionalidades

### Gravar Atributos
Cria, edita ou exclui categorias de atributos personalizados nos elementos selecionados do modelo.

- **Gravar** — Cria novas categorias de atributos com pares nome/valor
- **Editar** — Mescla novos valores com atributos já existentes
- **Excluir** — Remove categorias inteiras de atributos (com confirmação)
- Sidebar com árvore hierárquica das categorias e propriedades existentes
- Operações em lote sobre múltiplos elementos selecionados simultaneamente

### Selection Inspector
Inspeciona e exporta propriedades dos elementos selecionados em formato tabular.

- Filtragem por categorias de propriedades via checkboxes
- Grid dinâmica com colunas por propriedade e linhas por elemento
- Exportação para **CSV** (com BOM UTF-8 para compatibilidade com Excel)
- Cópia para área de transferência em formato TSV (para colar em planilhas)
- Processamento em background para não travar a interface

---

## Tecnologias

| Camada | Tecnologia |
|---|---|
| Linguagem | C# (.NET Framework 4.8) |
| Interface | WPF + XAML |
| API BIM | Autodesk Navisworks API 2026 |
| Build | MSBuild + InnoSetup |
| Plataforma | Windows x64 |

---

## Estrutura do Projeto

```
Plugin Set/
├── Plugin/                          # Entry points do plugin (AddInPlugin)
├── Views/                           # Janelas WPF (XAML + code-behind)
├── Models/                          # Modelos de dados (CheckableItem)
├── Themes/                          # Design system (tokens e estilos globais)
├── Resources/                       # Ícones do ribbon
├── Build/                           # Script InnoSetup e logs de build
├── SetAtributesToolkit.addin        # Manifesto do plugin para o Navisworks
└── SetAtributesToolkit.xaml         # Definição do ribbon
```

---

## Instalação

### Pré-requisitos
- Autodesk Navisworks Simulate 2026
- .NET Framework 4.8

### Manual
1. Compile o projeto no Visual Studio 2022 (`Release | x64`)
2. Copie os arquivos da pasta `Output/` para:
   ```
   %AppData%\Autodesk\Navisworks Simulate 2026\Plugins\SetAtributesToolkit\
   ```
3. Reinicie o Navisworks — os botões aparecerão automaticamente no ribbon

### Instalador
Execute o script `Build/Script para gerar instalador.iss` com o **InnoSetup** para gerar um instalador `.exe`.

---

## Como usar

1. Abra um modelo no Navisworks Simulate 2026
2. Selecione os elementos desejados na cena ou na árvore de seleção
3. Acesse os botões **Gravar Atributos** ou **Selection Inspector** no ribbon
4. Utilize os modos Gravar / Editar / Excluir conforme necessário

---

## Desenvolvimento

```bash
# Clonar o repositório
git clone https://github.com/jathuiel/constructsync_toolkit.git

# Abrir no Visual Studio
Plugin Simulate.sln
```

> As referências ao Navisworks API são resolvidas via GAC. Certifique-se de ter o Navisworks 2026 instalado antes de compilar.

---

## Licença

Projeto proprietário — todos os direitos reservados.
