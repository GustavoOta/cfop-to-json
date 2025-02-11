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
