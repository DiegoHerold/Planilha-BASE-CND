export function preencherResponsaveisEmail(linhas, mapaResponsaveis, mapaEmails) {
  return linhas.map(linha => {
    const doc = String(linha[2] || '').replace(/\D/g, '')
    const nome = mapaResponsaveis.get(doc)
    if (nome) {
      linha[3] = nome
      linha[4] = mapaEmails.get(nome) || ''
    }
    return linha
  })
}
