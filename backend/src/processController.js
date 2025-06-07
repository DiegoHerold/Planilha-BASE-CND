import XLSX from 'xlsx'
import path from 'path'

export async function processFiles(conferenciaPath, matrizPath, outputPath = './output/resultado.xlsm') {
  try {
    // [1] Conferência
    const workbookConf = XLSX.readFile(conferenciaPath)
    const sheetConf = workbookConf.Sheets['Todos Documentos']
    if (!sheetConf) throw new Error('Aba "Todos Documentos" não encontrada na conferência.')

    const dataConf = XLSX.utils.sheet_to_json(sheetConf, { header: 1, range: 12 })
    const linhasTransformadas = dataConf.map(linha => [
      linha[0],  // A
      linha[1],  // B
      linha[2],  // C
      '',        // D (Responsável)
      '',        // E (Email)
      linha[10], // K → F
      linha[11], // L → G
      linha[12], // M → H
      linha[13], // N → I
      linha[14], // O → J
    ])

    // [2] Carrega Planilha BASE original com macros
    const basePath = path.resolve('./src/data/Planilha BASE.xlsm')
    const workbookBase = XLSX.readFile(basePath, { bookVBA: true })

    // [3] Atualiza "Todos os Documentos"
    const sheetBase = workbookBase.Sheets['Todos os Documentos']
    const dadosBase = XLSX.utils.sheet_to_json(sheetBase, { header: 1 })
    const cabecalhoCompleto = dadosBase[0] || []

    const todosAtualizados = [cabecalhoCompleto, ...linhasTransformadas.map(linha => {
      const nova = new Array(cabecalhoCompleto.length).fill('')
      for (let i = 0; i < 10 && i < cabecalhoCompleto.length; i++) nova[i] = linha[i]
      return nova
    })]
    workbookBase.Sheets['Todos os Documentos'] = XLSX.utils.aoa_to_sheet(todosAtualizados)

    // [4] Distribuição por grupo (aba por valor na Coluna A)
    const grupoParaAba = {
      'ASSESSORIA': 'ASSESSORIA',
      'JAISE': 'JAISE',
      'SERVIÇO EXTRA': 'Serviço Extra',
      'CONTABILIDADE INTERNA - CI': 'Contabilidade Interna',
      'CONTABILIDADE EXTERNA - CE': 'Contabilidade Externa',
      'RH - RECURSOS HUMANOS': 'RH'
    }

    for (const [valorColunaA, nomeAba] of Object.entries(grupoParaAba)) {
      const linhasGrupo = linhasTransformadas.filter(l => (l[0] || '').toUpperCase().trim() === valorColunaA)
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

    // [5] Responsáveis da Matriz de Serviços
    const workbookMatriz = XLSX.readFile(matrizPath)
    const sheetMatriz = workbookMatriz.Sheets[workbookMatriz.SheetNames[0]]
    const dadosMatriz = XLSX.utils.sheet_to_json(sheetMatriz, { header: 1, range: 2 })

    const mapaResponsaveis = new Map()
    dadosMatriz.forEach(linha => {
      const doc = String(linha[9] || '').replace(/\D/g, '')
      const responsavel = linha[1] || ''
      if (doc) mapaResponsaveis.set(doc, responsavel)
    })

    const sheetTodos = workbookBase.Sheets['Todos os Documentos']
    const dadosTodos = XLSX.utils.sheet_to_json(sheetTodos, { header: 1 })
    const cabecalho = dadosTodos[0] || []
    const corpo = dadosTodos.slice(1)

    const corpoComResponsaveis = corpo.map(linha => {
      const doc = String(linha[2] || '').replace(/\D/g, '')
      if (mapaResponsaveis.has(doc)) {
        linha[3] = mapaResponsaveis.get(doc) // coluna D
      }
      return linha
    })

    // [6] Emails dos responsáveis (email responsaveis.xlsx)
    const emailPath = path.resolve('./src/data/email responsaveis.xlsx')
    const workbookEmails = XLSX.readFile(emailPath)
    const sheetEmail = workbookEmails.Sheets[workbookEmails.SheetNames[0]]
    const dadosEmail = XLSX.utils.sheet_to_json(sheetEmail, { header: 1, range: 1 })

    const mapaEmails = new Map()
    dadosEmail.forEach(linha => {
      const nome = String(linha[0] || '').trim()
      const email = String(linha[1] || '').trim()
      if (nome && email) mapaEmails.set(nome.toUpperCase(), email)
    })

    const corpoFinal = corpoComResponsaveis.map(linha => {
      const responsavel = String(linha[3] || '').trim().toUpperCase()
      if (mapaEmails.has(responsavel)) {
        linha[4] = mapaEmails.get(responsavel) // coluna E
      }
      return linha
    })

    const finalTodos = [cabecalho, ...corpoFinal]
    workbookBase.Sheets['Todos os Documentos'] = XLSX.utils.aoa_to_sheet(finalTodos)

    // [7] Salvar como .xlsm com macros
    let finalPath = path.resolve(outputPath)
    if (!finalPath.toLowerCase().endsWith('.xlsm')) {
      finalPath = finalPath.replace(/\.\w+$/, '.xlsm')
    }

    XLSX.writeFile(workbookBase, finalPath, {
      bookType: 'xlsm',
      bookVBA: true
    })

    console.log(`[✅] Planilha processada e salva com sucesso em: ${finalPath}`)
    return finalPath
  } catch (error) {
    console.error('[❌] Erro no processamento:', error)
    throw error
  }
}
