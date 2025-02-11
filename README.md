# cfop-to-json

Este projeto converte dados de um arquivo Excel (CFOP's) para o formato JSON.

## Descrição

O programa lê um arquivo Excel chamado 160314_Tabela_CFOP.xls, extrai os dados da planilha "CFOP" e os converte em um arquivo JSON chamado cfops.json.

## Dependências

- [calamine](https://crates.io/crates/calamine) - Biblioteca para leitura de arquivos Excel.
- [serde](https://crates.io/crates/serde) - Biblioteca para serialização de dados.
- [serde_json](https://crates.io/crates/serde_json) - Biblioteca para manipulação de dados JSON.

## Estrutura do Código

```rust
use calamine::{open_workbook, Reader, Xls}; // Importa as bibliotecas necessárias para ler arquivos Excel
use serde::Serialize; // Importa a biblioteca para serialização de dados
use std::fs::File; // Importa a biblioteca para manipulação de arquivos
use std::io::Write; // Importa a biblioteca para escrita em arquivos

#[derive(Serialize)]
struct CFOP {
    codigo: String,
    descricao: String,
    ind_nfe: String,
    ind_comunica: String,
    ind_transp: String,
    ind_devol: String,
}

fn main() {
    // Abre o arquivo Excel
    let mut excel: Xls<_> = open_workbook("160314_Tabela_CFOP.xls").unwrap();
    // Obtém a planilha específica dentro do arquivo Excel
    let range = excel.worksheet_range("CFOP").unwrap();

    // Índices das colunas A, B, C, D, E e F
    let coluna_a_index = 0;
    let coluna_b_index = 1;
    let coluna_c_index = 2;
    let coluna_d_index = 3;
    let coluna_e_index = 4;
    let coluna_f_index = 5;

    // Cria um vetor para os CFOP
    let mut cfops: Vec<CFOP> = Vec::new();

    // Itera sobre as linhas da planilha, pulando as primeiras 1 linhas
    for linha in range.rows().skip(1) {
        // Obtém os valores das colunas A, B, C, D, E e F
        let codigo = linha[coluna_a_index].to_string();
        let descricao = linha[coluna_b_index].to_string();
        let ind_nfe = linha[coluna_c_index].to_string();
        let ind_comunica = linha[coluna_d_index].to_string();
        let ind_transp = linha[coluna_e_index].to_string();
        let ind_devol = linha[coluna_f_index].to_string();

        // Cria um novo CFOP
        let cfop = CFOP {
            codigo,
            descricao,
            ind_nfe,
            ind_comunica,
            ind_transp,
            ind_devol,
        };

        // Adiciona o CFOP ao vetor
        cfops.push(cfop);
    }
    // Serializa o vetor de CFOPs para JSON
    let json = serde_json::to_string_pretty(&cfops).unwrap();

    // Cria um arquivo JSON
    let mut file = File::create("cfops.json").unwrap();

    // Escreve o JSON no arquivo
    file.write_all(json.as_bytes()).unwrap();

    println!("Arquivo cfops.json criado com sucesso!");
}
```

## Como Executar

1. Certifique-se de ter o Rust instalado em seu sistema. Se não tiver, siga as instruções em [rust-lang.org](https://www.rust-lang.org/).
2. Clone este repositório.
3. Navegue até o diretório do projeto.
4. Execute o comando `cargo run` para compilar e executar o programa.

```sh
cargo run
```

Após a execução, o arquivo cfops.json será gerado no diretório do projeto.

GustavoOta: quero que gere a documentação para o readme.md do git

Workspace: Collecting workspace information# cfop-to-json

Este projeto converte dados de um arquivo Excel (CFOP's) para o formato JSON.

## Descrição

O programa lê um arquivo Excel chamado 160314_Tabela_CFOP.xls, extrai os dados da planilha "CFOP" e os converte em um arquivo JSON chamado cfops.json.

## Dependências

- [calamine](https://crates.io/crates/calamine) - Biblioteca para leitura de arquivos Excel.
- [serde](https://crates.io/crates/serde) - Biblioteca para serialização de dados.
- [serde_json](https://crates.io/crates/serde_json) - Biblioteca para manipulação de dados JSON.

## Estrutura do Código

```rust
use calamine::{open_workbook, Reader, Xls}; // Importa as bibliotecas necessárias para ler arquivos Excel
use serde::Serialize; // Importa a biblioteca para serialização de dados
use std::fs::File; // Importa a biblioteca para manipulação de arquivos
use std::io::Write; // Importa a biblioteca para escrita em arquivos

#[derive(Serialize)]
struct CFOP {
    codigo: String,
    descricao: String,
    ind_nfe: String,
    ind_comunica: String,
    ind_transp: String,
    ind_devol: String,
}

fn main() {
    // Abre o arquivo Excel
    let mut excel: Xls<_> = open_workbook("160314_Tabela_CFOP.xls").unwrap();
    // Obtém a planilha específica dentro do arquivo Excel
    let range = excel.worksheet_range("CFOP").unwrap();

    // Índices das colunas A, B, C, D, E e F
    let coluna_a_index = 0;
    let coluna_b_index = 1;
    let coluna_c_index = 2;
    let coluna_d_index = 3;
    let coluna_e_index = 4;
    let coluna_f_index = 5;

    // Cria um vetor para os CFOP
    let mut cfops: Vec<CFOP> = Vec::new();

    // Itera sobre as linhas da planilha, pulando as primeiras 1 linhas
    for linha in range.rows().skip(1) {
        // Obtém os valores das colunas A, B, C, D, E e F
        let codigo = linha[coluna_a_index].to_string();
        let descricao = linha[coluna_b_index].to_string();
        let ind_nfe = linha[coluna_c_index].to_string();
        let ind_comunica = linha[coluna_d_index].to_string();
        let ind_transp = linha[coluna_e_index].to_string();
        let ind_devol = linha[coluna_f_index].to_string();

        // Cria um novo CFOP
        let cfop = CFOP {
            codigo,
            descricao,
            ind_nfe,
            ind_comunica,
            ind_transp,
            ind_devol,
        };

        // Adiciona o CFOP ao vetor
        cfops.push(cfop);
    }
    // Serializa o vetor de CFOPs para JSON
    let json = serde_json::to_string_pretty(&cfops).unwrap();

    // Cria um arquivo JSON
    let mut file = File::create("cfops.json").unwrap();

    // Escreve o JSON no arquivo
    file.write_all(json.as_bytes()).unwrap();

    println!("Arquivo cfops.json criado com sucesso!");
}
```

## Como Executar

1. Certifique-se de ter o Rust instalado em seu sistema. Se não tiver, siga as instruções em [rust-lang.org](https://www.rust-lang.org/).
2. Clone este repositório.
3. Navegue até o diretório do projeto.
4. Execute o comando `cargo run` para compilar e executar o programa.

```sh
cargo run
```

Após a execução, o arquivo cfops.json será gerado no diretório do projeto.