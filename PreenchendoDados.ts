function main(workbook: ExcelScript.Workbook) {
  // Definir a aba 'MACRO' onde estão os dados
  const sheetMacroName: string = 'MACRO';
  let sheetMacro: ExcelScript.Worksheet | undefined = workbook.getWorksheet(sheetMacroName);

  if (!sheetMacro) {
    return;
  }

  // Obter os dados da aba 'MACRO'
  const rangeMacro: ExcelScript.Range = sheetMacro.getUsedRange();
  const valuesMacro: (string | number | boolean)[][] = rangeMacro.getValues();

  // Índices das colunas na aba 'MACRO'
  const columnIndexes: { [key: string]: number } = {
    DATA: 0,
    CENTRO: 1,
    NF: 2,
    FORNECEDOR: 3,
    PEDIDO: 4,
    OBS: 5
  };

  // Cache das abas para melhorar a performance
  const sheetNames: string[] = [
    'ABA1', 'ABA2', 'ABA3'
  ];

  const allowedSheets: { [key: string]: ExcelScript.Worksheet } = {};
  sheetNames.forEach(name => {
    allowedSheets[name] = workbook.getWorksheet(name) as ExcelScript.Worksheet;
  });

  // Obter a data atual no formato DD/MM/AAAA
  const today: Date = new Date();
  const day: string = today.getDate().toString().padStart(2, '0');
  const month: string = (today.getMonth() + 1).toString().padStart(2, '0');
  const year: string = today.getFullYear().toString();

  // Array para guardar linhas a serem removidas
  const rowsToRemove: number[] = [];

  // Percorrer cada linha da aba 'MACRO' (começando da segunda linha, a primeira é o cabeçalho)
  for (let i: number = 1; i < valuesMacro.length; i++) {
    const centro: string = String(valuesMacro[i][columnIndexes.CENTRO]);
    const nfMacro: string | number | boolean = valuesMacro[i][columnIndexes.NF];
    const fornecedorMacro: string | number | boolean = valuesMacro[i][columnIndexes.FORNECEDOR];
    const pedido: string | number | boolean = valuesMacro[i][columnIndexes.PEDIDO];

    // Verificar se o valor de 'PEDIDO' na 'MACRO' está vazio
    if (!pedido) {
      continue; // Pular esta linha se PEDIDO estiver vazio
    }

    // Verificar se o valor de 'CENTRO' não está vazio e se a aba correspondente existe
    if (centro && allowedSheets[centro]) {
      const sheetDest: ExcelScript.Worksheet = allowedSheets[centro];

      // Obter os dados da aba de destino
      const rangeDest: ExcelScript.Range = sheetDest.getUsedRange();
      const valuesDest: (string | number | boolean)[][] = rangeDest.getValues();

      // Procurar a linha correspondente na aba de destino onde NF e FORNECEDOR são iguais aos valores na 'MACRO'
      for (let j: number = 1; j < valuesDest.length; j++) {
        const nfDest: string | number | boolean = valuesDest[j][3]; // Coluna NF na aba de destino
        const fornecedorDest: string | number | boolean = valuesDest[j][4]; // Coluna FORNECEDOR na aba de destino

        if (nfMacro === nfDest && fornecedorMacro === fornecedorDest) {
          // Definir o valor de 'PEDIDO' na aba de destino (coluna G é a sétima coluna)
          sheetDest.getRange(`G${j + 1}`).setValue(pedido);
          // Definir o valor de 'STATUS' na aba de destino (coluna I é a nona coluna)
          sheetDest.getRange(`I${j + 1}`).setValue("Integrada e corrigida");
          // Definir a data como texto diretamente com o formato correto
          sheetDest.getRange(`J${j + 1}`).setNumberFormatLocal("@");
          sheetDest.getRange(`J${j + 1}`).setValue(`${day}/${month}/${year}`);
          // Definir o valor de 'INTEGRADO POR' na aba de destino (coluna K é a décima primeira coluna)
          sheetDest.getRange(`K${j + 1}`).setValue("Laura");
          rowsToRemove.push(i + 1); // Adicionar linha para remover na 'MACRO'
          break;
        }
      }
    }
  }

  // Remover as linhas da 'MACRO' que foram processadas
  rowsToRemove.reverse().forEach(row => {
    const rangeToDelete = sheetMacro.getRange(`A${row}:F${row}`);
    rangeToDelete.delete(ExcelScript.DeleteShiftDirection.up); // Deleta apenas as 6 primeiras colunas
  });
}
