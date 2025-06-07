import express from 'express'
import multer from 'multer'
import cors from 'cors'
import { processFiles } from './processController.js'

const app = express()
const upload = multer({ dest: 'uploads/' })

app.use(cors())

app.post(
  '/upload',
  upload.fields([
    { name: 'conferencia', maxCount: 1 },
    { name: 'matriz', maxCount: 1 },
  ]),
  async (req, res) => {
    try {
      const conferenciaFile = req.files['conferencia']?.[0]
      const matrizFile = req.files['matriz']?.[0]

      if (!conferenciaFile || !matrizFile) {
        return res.status(400).json({ error: 'Arquivos obrigatórios não enviados.' })
      }

      const outputPath = await processFiles(conferenciaFile.path, matrizFile.path)
      res.download(outputPath)
    } catch (error) {
      console.error('Erro no processamento:', error)
      res.status(500).json({ error: 'Erro ao processar as planilhas.' })
    }
  }
)

app.listen(3001, () => {
  console.log('Servidor rodando em http://localhost:3001')
})
