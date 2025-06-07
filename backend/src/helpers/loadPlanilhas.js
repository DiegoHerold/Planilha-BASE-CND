import XLSX from 'xlsx'

export function carregarPlanilhaConferencia(path) {
  const workbook = XLSX.readFile(path)
  const sheet = workbook.Sheets['Todos Documentos']
  if (!sheet) throw new Error('Aba "Todos Documentos" não encontrada na conferência.')
  return XLSX.utils.sheet_to_json(sheet, { header: 1, range: 12 })
}

export function carregarPlanilhaBase(path) {
  return XLSX.readFile(path, { bookVBA: true })
}
