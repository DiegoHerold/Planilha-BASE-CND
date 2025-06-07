import XLSX from 'xlsx'

export function atualizarTodosDocumentos(workbookBase, linhas, cabecalhoCompleto) {
  const novaAba = [
    cabecalhoCompleto,
    ...linhas.map(linha => {
      const nova = new Array(cabecalhoCompleto.length).fill('')
      for (let i = 0; i < 10 && i < cabecalhoCompleto.length; i++) nova[i] = linha[i]
      return nova
    })
  ]

  workbookBase.Sheets['Todos os Documentos'] = XLSX.utils.aoa_to_sheet(novaAba)
}
