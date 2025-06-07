import XLSX from 'xlsx'
import path from 'path'

export function preencherEmails(workbookBase) {
  const emailPath = path.resolve('./src/data/email responsaveis.xlsx')
  const workbook = XLSX.readFile(emailPath)
  const sheet = workbook.Sheets[workbook.SheetNames[0]]
  const dados = XLSX.utils.sheet_to_json(sheet, { header: 1, range: 1 })

  const mapa = new Map()
  dados.forEach(linha => {
    const nome = String(linha[0] || '').trim().toUpperCase()
    const email = String(linha[1] || '').trim()
    if (nome && email) mapa.set(nome, email)
  })

  const aba = workbookBase.Sheets['Todos os Documentos']
  const data = XLSX.utils.sheet_to_json(aba, { header: 1 })
  const cabecalho = data[0]
  const corpo = data.slice(1)

  const atualizado = corpo.map(linha => {
    const responsavel = String(linha[3] || '').trim().toUpperCase()
    if (mapa.has(responsavel)) linha[4] = mapa.get(responsavel)
    return linha
  })

  workbookBase.Sheets['Todos os Documentos'] = XLSX.utils.aoa_to_sheet([cabecalho, ...atualizado])
}
