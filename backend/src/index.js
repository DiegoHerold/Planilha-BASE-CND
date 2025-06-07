import path from 'path'
import XLSX from 'xlsx'
import { loadConferencia } from './helpers/loadConferencia.js'
import { atualizarTodosDocumentos } from './helpers/atualizarTodosDocumentos.js'
import { distribuirPorGrupo } from './helpers/distribuirPorGrupo.js'
import { preencherResponsaveis } from './helpers/preencherResponsaveis.js'
import { preencherEmails } from './helpers/preencherEmails.js'
import { aplicarPoliticasFinais } from './helpers/aplicarPoliticasFinais.js'

export async function processFiles(conferenciaPath, matrizPath, outputPath = './output/resultado.xlsm') {
  try {
    const linhas = loadConferencia(conferenciaPath)

    const basePath = path.resolve('./src/data/Planilha BASE.xlsm')
    const workbookBase = XLSX.readFile(basePath, { bookVBA: true })

    const sheet = workbookBase.Sheets['Todos os Documentos']
    const cabecalho = XLSX.utils.sheet_to_json(sheet, { header: 1 })[0] || []

    atualizarTodosDocumentos(workbookBase, linhas, cabecalho)
    preencherResponsaveis(workbookBase, matrizPath)
    preencherEmails(workbookBase)

    const sheetTodosAtualizado = workbookBase.Sheets['Todos os Documentos']
    const dadosAtualizados = XLSX.utils.sheet_to_json(sheetTodosAtualizado, { header: 1 }).slice(1)

    distribuirPorGrupo(workbookBase, dadosAtualizados)
    aplicarPoliticasFinais(workbookBase)

    let finalPath = path.resolve(outputPath)
    if (!finalPath.toLowerCase().endsWith('.xlsm')) {
      finalPath = finalPath.replace(/\.\w+$/, '.xlsm')
    }

    XLSX.writeFile(workbookBase, finalPath, { bookType: 'xlsm', bookVBA: true })

    return finalPath
  } catch (err) {
    throw err
  }
}
