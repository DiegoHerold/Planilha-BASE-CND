import XLSX from 'xlsx'

// 🔧 Aqui você define quais grupos serão removidos da aba "Todos os Documentos"
export const gruposRemoverDoTodos = [
  'JAISE',
  'SERVIÇO EXTRA',
  'ASSESSORIA'
]

export function aplicarPoliticasFinais(workbookBase) {
  const aba = workbookBase.Sheets['Todos os Documentos']
  const dados = XLSX.utils.sheet_to_json(aba, { header: 1 })
  const cabecalho = dados[0] || []
  const corpo = dados.slice(1)

  const removerHubcount = []
  const manterEmTodos = []

  for (const linha of corpo) {
    const grupo = (linha[0] || '').toUpperCase().trim()
    const responsavel = (linha[3] || '').trim()

    if (!responsavel) {
      removerHubcount.push(linha)
      continue
    }

    if (gruposRemoverDoTodos.includes(grupo)) {
      // ← Remoção condicional com base na lista acima
      continue
    }

    manterEmTodos.push(linha)
  }

  // Atualiza aba "Todos os Documentos" com os que permanecem
  workbookBase.Sheets['Todos os Documentos'] = XLSX.utils.aoa_to_sheet([
    cabecalho,
    ...manterEmTodos
  ])

  // Atualiza ou cria aba "REMOVER HUBCOUNT"
  const abaRemover = workbookBase.Sheets['REMOVER HUBCOUNT']
  const antiga = abaRemover
    ? XLSX.utils.sheet_to_json(abaRemover, { header: 1 }).slice(1)
    : []

  const novaAbaRemover = [cabecalho, ...antiga, ...removerHubcount]
  workbookBase.Sheets['REMOVER HUBCOUNT'] = XLSX.utils.aoa_to_sheet(novaAbaRemover)
}
