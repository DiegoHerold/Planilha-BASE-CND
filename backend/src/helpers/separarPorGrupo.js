import XLSX from 'xlsx'

export function separarLinhasPorGrupo(workbookBase, linhas, grupoParaAba, gruposParaRecorte) {
  const linhasRestantes = []

  for (const linha of linhas) {
    const grupo = (linha[0] || '').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim()
    const abaDestino = grupoParaAba[grupo]

    if (!abaDestino || !workbookBase.Sheets[abaDestino]) {
      linhasRestantes.push(linha)
      continue
    }

    const dadosAba = XLSX.utils.sheet_to_json(workbookBase.Sheets[abaDestino], { header: 1 })
    const cab = dadosAba[0] || []
    const corpo = dadosAba.slice(1)

    const novaLinha = new Array(cab.length).fill('')
    for (let i = 0; i < linha.length && i < cab.length; i++) {
      novaLinha[i] = linha[i]
    }

    const novaAba = [cab, ...corpo, novaLinha]
    workbookBase.Sheets[abaDestino] = XLSX.utils.aoa_to_sheet(novaAba)

    if (!gruposParaRecorte.includes(grupo)) {
      linhasRestantes.push(linha)
    }
  }

  return linhasRestantes
}
