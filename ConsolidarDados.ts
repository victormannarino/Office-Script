function main(workbook: ExcelScript.Workbook) {
  // Listar todas as abas que serão combinadas
  const sheetNames: string[] = ['ABA1', 'ABA2', 'ABA3'  ];

  // Obter ou criar a aba 'MACRO'
  let sheetMacro: ExcelScript.Worksheet = workbook.getWorksheet('MACRO');
  if (!sheetMacro) {
    sheetMacro = workbook.addWorksheet('MACRO');
  }

  // Limpar a aba 'MACRO' 
  sheetMacro.getRange().clear(ExcelScript.ClearApplyTo.contents);

  // Definir os cabeçalhos das colunas selecionadas
  const cabecalhos: string[][] = [
    ['DATA', 'CENTRO', 'NF', 'FORNECEDOR', 'PEDIDO', 'OBS']
  ];
  const headerRange: ExcelScript.Range = sheetMacro.getRange('A1:F1');
  headerRange.setValues(cabecalhos);

  // Definir os cabeçalhos para a verificação
  const verificacaoCabecalhos: string[][] = [
    ['ABA', 'Status', 'Dados']
  ];
  const verificacaoHeaderRange: ExcelScript.Range = sheetMacro.getRange('Q1:S1'); // Colunas 17, 18 e 19
  verificacaoHeaderRange.setValues(verificacaoCabecalhos);

  let startRow: number = 2; // Começar na linha 2, pois a 1 é o cabeçalho

  // Índices das colunas que queremos copiar
  const columnIndexes = {
    DATA: 0,
    NF: 3,
    FORNECEDOR: 4,
    PEDIDO: 6,
    OBS: 7,
    CENTRO: 5 // Índice para a nova coluna 'CENTRO'
  };

  // Percorrer cada aba e obter os valores das linhas
  sheetNames.forEach((sheetName: string, index: number) => {
    const worksheet: ExcelScript.Worksheet = workbook.getWorksheet(sheetName);
    let linhasEncontradas: number = 0;
    let dadosExistem: boolean = false;

    if (worksheet) { // Verificar se a aba existe
      const usedRange: ExcelScript.Range = worksheet.getUsedRange();
      if (usedRange) {
        const values: string[][] = usedRange.getValues() as string[][];

        // Verificar se existem dados na aba (exceto cabeçalhos)
        if (values && values.length > 1) {
          dadosExistem = true;
        }

        // Filtrar e adicionar as linhas onde a coluna PEDIDO (índice 6) está vazia e NF (índice 3) não está vazio
        if (values) {
          for (let i: number = 1; i < values.length; i++) { // Começar em 1 para pular o cabeçalho
            if (values[i] &&
              values[i][columnIndexes.PEDIDO] === '' &&
              values[i][columnIndexes.NF] !== '') {
              const selectedValues: (string | number)[] = [
                values[i][columnIndexes.DATA],
                sheetName, // Adicionar o nome da aba na coluna 'CENTRO'
                values[i][columnIndexes.NF],
                values[i][columnIndexes.FORNECEDOR],
                values[i][columnIndexes.PEDIDO],
                values[i][columnIndexes.OBS]
              ];
              const targetRange: ExcelScript.Range = sheetMacro.getRange(`A${startRow}:F${startRow}`);
              targetRange.setValues([selectedValues]);
              startRow++;
              linhasEncontradas++;
            }
          }
        }
      }
    }

    // Adicionar as informações de verificação na aba 'MACRO'
    const verificacaoValues: (string | number)[][] = [
      [sheetName, dadosExistem ? 'OK' : 'problema', linhasEncontradas]
    ];
    const verificacaoTargetRange: ExcelScript.Range = sheetMacro.getRange(`Q${index + 2}:S${index + 2}`);
    verificacaoTargetRange.setValues(verificacaoValues);
  });

  // Adicionar a data e o horário da execução no índice 8
  const dataHoraExecucao: Date = new Date();
  const dataFormatada: string = `${String(dataHoraExecucao.getDate()).padStart(2, '0')}/${String(dataHoraExecucao.getMonth() + 1).padStart(2, '0')}`;
  const horaFormatada: string = dataHoraExecucao.toLocaleTimeString();
  const dataHoraValores: string[][] = [
    ['Data de Execução', dataFormatada],
    ['Hora de Execução', horaFormatada]
  ];
  const dataHoraRange: ExcelScript.Range = sheetMacro.getRange('H1:I2');
  dataHoraRange.setValues(dataHoraValores);

  // Informar ao usuário para ajustar as colunas manualmente
  console.log("O ajuste automático das colunas não é suportado no Excel Online. Ajuste manualmente as colunas se necessário.");
}
