import XLSX from 'xlsx'
import path from 'path'

export function salvarComoXLSM(workbook, outputPath) {
  const finalPath = path.resolve(outputPath).replace(/\.\w+$/, '.xlsm')
  XLSX.writeFile(workbook, finalPath, { bookType: 'xlsm', bookVBA: true })
  return finalPath
}
