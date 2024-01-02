
async function main(workbook: ExcelScript.Workbook) {
    // Declaração de Variáveis
    let w = workbook.getActiveWorksheet();
    let cabecalho:string[] = ['Logradouro','Bairro','Localidade','UF'];
    // código para capturar a última linha
    let ultlin:number = w.getUsedRange(true).getLastRow().getRowIndex();
    // criando uma constante para a api
    const url: string = 'https://viacep.com.br/ws/'

    // inserindo os cabeçalhos
    w.getCell(0,1).getResizedRange(0,3).setValues([cabecalho])

    // iterando para cada CEP
    for(let i:number = 1; i<= ultlin; i++){
        let cep:string = w.getCell(i,0).getValue().toString();
        let link:string = url + cep + '/json/'
       
        let fetchResult = await fetch(link);
        let resultado:string = await fetchResult.json();

        // armazenando as respostas em variáveis
        let respostas: string[] = [resultado['logradouro'], resultado['bairro'], resultado['localidade'], resultado['uf']]
        // inserindos as respostas obtidas na api
        w.getCell(i,1).getResizedRange(0,3).setValues([respostas])
        // ajustando o tamanho da coluna a cada iteração
        w.getCell(i,0).getResizedRange(0,4).getFormat().autofitColumns();
      sleep(0.75);
};
// código para capturar a última coluna
let ultcol: number = w.getUsedRange(true).getLastColumn().getColumnIndex();

// selecionando todas as células
let celulas_preenchidas = w.getCell(0,0).getResizedRange(ultlin,ultcol);
// criando a tabela
let novaTabela = workbook.addTable(w.getRange(celulas_preenchidas.getAddress()), true);
// Retirando as Linhas de Grade do Fundo
w.setShowGridlines(false);
celulas_preenchidas.getFormat().autofitColumns();

}

// função de espera
// https://www.reddit.com/r/excel/comments/tbmzav/office_scripts_add_waiting_time_before_running/
function sleep(seconds:number) {
  var waitUntil = new Date().getTime() + seconds * 1000;
  while (new Date().getTime() < waitUntil) { };
}
