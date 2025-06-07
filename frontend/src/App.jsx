import { useState } from 'react'

function App() {
  const [conferencia, setConferencia] = useState(null)
  const [matriz, setMatriz] = useState(null)
  const [loading, setLoading] = useState(false)

  const handleSubmit = async (e) => {
    e.preventDefault()
    setLoading(true)

    try {
      const formData = new FormData()
      formData.append('conferencia', conferencia)
      formData.append('matriz', matriz)

      const res = await fetch('http://localhost:3001/upload', {
        method: 'POST',
        body: formData,
      })

      const blob = await res.blob()
      const url = window.URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = 'resultado.xlsm'
      a.click()
    } catch (err) {
      alert('Erro ao processar o arquivo.')
      console.error(err)
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-100 to-slate-200 flex items-center justify-center px-4">
      <div className="bg-white shadow-xl rounded-3xl p-10 w-full max-w-lg">
        <h1 className="text-3xl font-bold text-center text-gray-800 mb-8">
          Enviar <span className="text-blue-600">Planilhas</span>
        </h1>

        <form onSubmit={handleSubmit} className="flex flex-col gap-6">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">
              Arquivo de Conferência (.xlsx)
            </label>
            <input
              type="file"
              accept=".xlsx"
              onChange={(e) => setConferencia(e.target.files[0])}
              required
              className="w-full border border-gray-300 rounded-lg px-4 py-2 shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
          </div>

          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">
              Matriz de Serviços (.xlsx)
            </label>
            <input
              type="file"
              accept=".xlsx"
              onChange={(e) => setMatriz(e.target.files[0])}
              required
              className="w-full border border-gray-300 rounded-lg px-4 py-2 shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
          </div>

          <button
            type="submit"
            className="bg-gradient-to-r from-blue-600 to-blue-500 hover:from-blue-700 hover:to-blue-600 text-white font-semibold py-3 rounded-lg shadow-lg transition duration-200 flex justify-center items-center"
            disabled={loading}
          >
            {loading ? (
              <div className="flex gap-2 items-center">
                <svg
                  className="animate-spin h-5 w-5 text-white"
                  xmlns="http://www.w3.org/2000/svg"
                  fill="none"
                  viewBox="0 0 24 24"
                >
                  <circle
                    className="opacity-25"
                    cx="12"
                    cy="12"
                    r="10"
                    stroke="currentColor"
                    strokeWidth="4"
                  />
                  <path
                    className="opacity-75"
                    fill="currentColor"
                    d="M4 12a8 8 0 018-8v8H4z"
                  />
                </svg>
                Processando...
              </div>
            ) : (
              'Enviar e baixar resultado'
            )}
          </button>
        </form>
      </div>
    </div>
  )
}

export default App
