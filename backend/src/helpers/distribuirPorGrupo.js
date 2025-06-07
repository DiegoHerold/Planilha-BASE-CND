import XLSX from 'xlsx'

const grupoParaAba = {
  'ASSESSORIA': 'ASSESSORIA',
  'JAISE': 'JAISE',
  'SERVIÇO EXTRA': 'Serviço Extra',
  'CONTABILIDADE INTERNA - CI': 'Contabilidade Interna',
  'CONTABILIDADE EXTERNA - CE': 'Contabilidade Externa',
  'RH - RECURSOS HUMANOS': 'RH'
}

export function distribuirPorGrupo(workbookBase, linhasTransformadas) {
  for (const [grupo, nomeAba] of Object.entries(grupoParaAba)) {
    const linhasGrupo = linhasTransformadas.filter(l => (l[0] || '').toUpperCase().trim() === grupo)
    const aba = workbookBase.Sheets[nomeAba]
    if (!aba) continue

    const existente = XLSX.utils.sheet_to_json(aba, { header: 1 })
    const cabecalho = existente[0] || []

    const novas = linhasGrupo.map(linha => {
      const nova = new Array(cabecalho.length).fill('')
      for (let i = 0; i < 10 && i < cabecalho.length; i++) nova[i] = linha[i]
      return nova
    })

    workbookBase.Sheets[nomeAba] = XLSX.utils.aoa_to_sheet([cabecalho, ...novas])
  }
}
