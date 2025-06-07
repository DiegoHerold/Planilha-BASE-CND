import XLSX from 'xlsx'

export function preencherResponsaveis(workbookBase, matrizPath) {
  const workbookMatriz = XLSX.readFile(matrizPath)
  const sheet = workbookMatriz.Sheets[workbookMatriz.SheetNames[0]]
  const dados = XLSX.utils.sheet_to_json(sheet, { header: 1, range: 2 })

  const mapa = new Map()
  dados.forEach(linha => {
    const doc = String(linha[9] || '').replace(/\D/g, '')
    if (doc) mapa.set(doc, linha[1] || '')
  })

  const aba = workbookBase.Sheets['Todos os Documentos']
  const data = XLSX.utils.sheet_to_json(aba, { header: 1 })
  const cabecalho = data[0]
  const corpo = data.slice(1)

  const atualizado = corpo.map(linha => {
    const doc = String(linha[2] || '').replace(/\D/g, '')
    if (mapa.has(doc)) linha[3] = mapa.get(doc)
    return linha
  })

  workbookBase.Sheets['Todos os Documentos'] = XLSX.utils.aoa_to_sheet([cabecalho, ...atualizado])
}
