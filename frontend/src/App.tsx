import { useEffect, useState } from 'react'
import type { ChangeEvent } from 'react'

import DocumentSections from './DocumentSections'
import DocumentPreview from './DocumentPreview'

import type {
  DocumentSection,
  LoadedDocument,
} from './types'

import './App.css'

const API_URL = 'http://127.0.0.1:8000'
const MAX_FILE_SIZE = 10 * 1024 * 1024

type PendingReplacement = {
  id: string
  source: DocumentSection
  target: DocumentSection
}

type TransferCounts = {
  copied_paragraphs: number
  copied_tables: number
  copied_images: number
  created_lists: number
}

type ReplacementSummary = TransferCounts & {
  source_title: string
  target_title: string
}

type BatchResult = LoadedDocument &
  TransferCounts & {
    replacement_count: number
    replacements: ReplacementSummary[]
    download_url: string
  }

async function readApiResponse(response: Response) {
  const data = await response.json().catch(() => null)

  if (!response.ok) {
    throw new Error(
      typeof data?.detail === 'string'
        ? data.detail
        : `La solicitud falló: HTTP ${response.status}.`,
    )
  }

  if (!data) {
    throw new Error('El backend no devolvió un JSON válido.')
  }

  return data
}

function sectionsOverlap(
  first: DocumentSection,
  second: DocumentSection,
) {
  return (
    first.start_index < second.end_index &&
    second.start_index < first.end_index
  )
}

function App() {
  const [connection, setConnection] = useState(
    'Comprobando conexión...',
  )

  const [fileA, setFileA] = useState<File | null>(null)
  const [fileB, setFileB] = useState<File | null>(null)

  const [documentA, setDocumentA] =
    useState<LoadedDocument | null>(null)

  const [documentB, setDocumentB] =
    useState<LoadedDocument | null>(null)

  const [sectionA, setSectionA] =
    useState<DocumentSection | null>(null)

  const [sectionB, setSectionB] =
    useState<DocumentSection | null>(null)

  const [pending, setPending] =
    useState<PendingReplacement[]>([])

  const [errorA, setErrorA] = useState('')
  const [errorB, setErrorB] = useState('')
  const [transferError, setTransferError] = useState('')
  const [transferMessage, setTransferMessage] = useState('')

  const [operation, setOperation] = useState<
    'uploadA' | 'uploadB' | 'transfer' | null
  >(null)

  const [outputFile, setOutputFile] = useState<File | null>(null)
  const [downloadUrl, setDownloadUrl] = useState('')

  // Si el lote se genera pero falla la obtención del archivo,
  // conservamos un enlace al resultado en el backend.
  const [serverDownloadPath, setServerDownloadPath] = useState('')

  const busy = operation !== null

  const [lastSummary, setLastSummary] =
  useState<BatchResult | null>(null)

  const canAdd =
    !busy &&
    documentA !== null &&
    documentB !== null &&
    sectionA?.transfer_enabled === true &&
    sectionB?.transfer_enabled === true

  const canApply =
    !busy &&
    documentA !== null &&
    documentB !== null &&
    pending.length > 0

  useEffect(() => {
    const controller = new AbortController()

    async function checkBackend() {
      try {
        const response = await fetch(`${API_URL}/health`, {
          signal: controller.signal,
        })

        const data = await readApiResponse(response)

        if (data.status !== 'ok') {
          throw new Error('El backend no está listo.')
        }

        setConnection('Backend conectado')
      } catch {
        if (!controller.signal.aborted) {
          setConnection('No se pudo conectar con el backend')
        }
      }
    }

    void checkBackend()

    return () => controller.abort()
  }, [])

  useEffect(() => {
    if (!outputFile) {
      setDownloadUrl('')
      return
    }

    const url = URL.createObjectURL(outputFile)
    setDownloadUrl(url)

    return () => URL.revokeObjectURL(url)
  }, [outputFile])

  function clearPending() {
    setPending([])
    setSectionA(null)
    setSectionB(null)
    setTransferError('')
    setTransferMessage('')
  }

  function selectFile(
    event: ChangeEvent<HTMLInputElement>,
    label: 'A' | 'B',
  ) {
    const file = event.target.files?.[0] ?? null

    const setFile = label === 'A' ? setFileA : setFileB
    const setError = label === 'A' ? setErrorA : setErrorB

    // Las parejas pertenecen a los documentos cargados actuales.
    // Si cambia un archivo, dejan de ser válidas.
    clearPending()
    setError('')

    if (label === 'A') {
      setDocumentA(null)
    } else {
      setDocumentB(null)
      setOutputFile(null)
      setServerDownloadPath('')
      setLastSummary(null)
    }

    if (file && !file.name.toLowerCase().endsWith('.docx')) {
      setFile(null)
      setError('Selecciona un archivo con extensión .docx.')
      event.target.value = ''
      return
    }

    if (file && file.size > MAX_FILE_SIZE) {
      setFile(null)
      setError('El archivo supera el límite de 10 MB.')
      event.target.value = ''
      return
    }

    setFile(file)
  }

  async function uploadDocument(label: 'A' | 'B') {
    const file = label === 'A' ? fileA : fileB

    if (!file || busy) {
      return
    }

    const setError = label === 'A' ? setErrorA : setErrorB

    clearPending()
    setError('')

    if (label === 'A') {
      setDocumentA(null)
    } else {
      setDocumentB(null)
    }

    setOperation(label === 'A' ? 'uploadA' : 'uploadB')

    try {
      const formData = new FormData()
      formData.append('file', file)

      const response = await fetch(`${API_URL}/documents`, {
        method: 'POST',
        body: formData,
      })

      const data = await readApiResponse(response)

      if (
        typeof data.document_id !== 'string' ||
        !Array.isArray(data.sections)
      ) {
        throw new Error('La respuesta de carga es inesperada.')
      }

      const loaded = data as LoadedDocument

      if (label === 'A') {
        setDocumentA(loaded)
      } else {
        setDocumentB(loaded)
      }

      setConnection('Backend conectado')
    } catch (error) {
      setError(
        error instanceof Error
          ? error.message
          : 'No se pudo cargar el documento.',
      )
    } finally {
      setOperation(null)
    }
  }

  function addReplacement() {
    if (!canAdd || !sectionA || !sectionB) {
      return
    }

    setTransferError('')
    setTransferMessage('')

    if (pending.length >= 50) {
      setTransferError('El límite es de 50 parejas por lote.')
      return
    }

    const repeated = pending.find(
      (item) => item.target.id === sectionB.id,
    )

    if (repeated) {
      setTransferError(
        `El destino «${sectionB.title}» ya está en la lista.`,
      )
      return
    }

    const overlapping = pending.find(
      (item) => sectionsOverlap(item.target, sectionB),
    )

    if (overlapping) {
      setTransferError(
        `El destino «${sectionB.title}» se superpone con ` +
          `«${overlapping.target.title}». ` +
          'No puedes reemplazar una sección y una subsección ' +
          'dentro de ella en el mismo lote.',
      )
      return
    }

    setPending((current) => [
      ...current,
      {
        id: crypto.randomUUID(),
        source: sectionA,
        target: sectionB,
      },
    ])

    // Conservamos el origen por si se reutiliza en otro destino.
    setSectionB(null)

    setTransferMessage('Pareja agregada a la lista pendiente.')
  }

  function removeReplacement(id: string) {
    setPending((current) =>
      current.filter((item) => item.id !== id),
    )

    setTransferError('')
    setTransferMessage('')
  }

  async function applyBatch() {
    if (!canApply || !documentA || !documentB) {
      return
    }

    const summary = pending
      .map(
        (item, index) =>
          `${index + 1}. ${item.source.title} → ${item.target.title}`,
      )
      .join('\n')

    const confirmed = window.confirm(
      `Se aplicarán ${pending.length} reemplazos:\n\n` +
        summary +
        '\n\nSe conservarán los títulos de destino y se ' +
        'reemplazará su contenido, incluidas sus subsecciones.\n\n' +
        'Se generará una nueva copia de B. ¿Continuar?',
    )

    if (!confirmed) {
      return
    }

    setOperation('transfer')
    setTransferError('')
    setTransferMessage('')

    try {
      const response = await fetch(`${API_URL}/transfers/batch`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          source_document_id: documentA.document_id,
          target_document_id: documentB.document_id,
          replacements: pending.map((item) => ({
            source_section_id: item.source.id,
            target_section_id: item.target.id,
          })),
        }),
      })

      const data = await readApiResponse(response)

      if (
        typeof data.document_id !== 'string' ||
        typeof data.download_url !== 'string' ||
        !Array.isArray(data.sections)
      ) {
        throw new Error('La respuesta del lote es inesperada.')
      }

      const result = data as BatchResult

      // El lote ya se completó en Python. Conservamos su ruta
      // antes de intentar obtener el archivo.
      setServerDownloadPath(result.download_url)

      const fileResponse = await fetch(
        `${API_URL}${result.download_url}`,
      )

      if (!fileResponse.ok) {
        throw new Error(
          'El lote se generó, pero no se pudo obtener el DOCX. ' +
            'Usa el enlace de recuperación para descargarlo. ' +
            'El panel B todavía muestra la versión anterior.',
        )
      }

      const blob = await fileResponse.blob()

      const updatedFile = new File(
        [blob],
        result.filename || 'B_actualizado.docx',
        {
          type:
            'application/vnd.openxmlformats-officedocument.' +
            'wordprocessingml.document',
        },
      )

      // Solo cambia B; A permanece intacto.
      setFileB(updatedFile)
      setDocumentB(result)
      setOutputFile(updatedFile)
      setLastSummary(result)

      // Los identificadores de sección de B se recalcularon.
      setSectionA(null)
      setSectionB(null)
      setPending([])
      setErrorB('')
      setServerDownloadPath('')

      setTransferMessage(
        `Lote completado: ${result.replacement_count} secciones ` +
          'reemplazadas. El panel B muestra el resultado.',
      )
    } catch (error) {
      // Si hay un error, conservamos las parejas pendientes.
      setTransferError(
        error instanceof Error
          ? error.message
          : 'No se pudo completar el lote.',
      )
    } finally {
      setOperation(null)
    }
  }

  function renderDocumentPanel(label: 'A' | 'B') {
    const isSource = label === 'A'

    const file = isSource ? fileA : fileB
    const loadedDocument = isSource ? documentA : documentB
    const selectedSection = isSource ? sectionA : sectionB
    const error = isSource ? errorA : errorB
    const uploading =
      operation === (isSource ? 'uploadA' : 'uploadB')

    return (
      <section className="document-panel">
        <h2>
          Documento {label} — {isSource ? 'Origen' : 'Destino'}
        </h2>

        <label htmlFor={`document-${label}`}>
          Seleccionar archivo {label}
        </label>

        <input
          id={`document-${label}`}
          type="file"
          accept=".docx"
          disabled={busy}
          onChange={(event) => selectFile(event, label)}
        />

        {file && (
          <>
            <p className="file-info">
              Archivo actual: <strong>{file.name}</strong>
            </p>

            <button
              className="action-button"
              type="button"
              disabled={busy}
              onClick={() => uploadDocument(label)}
            >
              {uploading
                ? 'Cargando...'
                : loadedDocument
                  ? `Volver a cargar ${label}`
                  : `Cargar documento ${label}`}
            </button>
          </>
        )}

        {error && (
          <p className="error" role="alert">
            {error}
          </p>
        )}

        <DocumentSections
          document={loadedDocument}
          selectedId={selectedSection?.id ?? null}
          documentLabel={label}
          disabled={busy}
          onSelect={(section) => {
            if (isSource) {
              setSectionA(section)
            } else {
              setSectionB(section)
            }

            setTransferError('')
            setTransferMessage('')
          }}
        />

        {file ? (
          <DocumentPreview
            file={file}
            sections={loadedDocument?.sections}
            selectedSectionId={selectedSection?.id ?? null}
          />
        ) : (
          <div className="preview-placeholder">
            Selecciona {label} para ver su contenido.
          </div>
        )}
      </section>
    )
  }

  return (
    <main className="app">
      <header className="app-header">
        <h1>Transferencia de secciones Word</h1>

        <p>
          Prepara varias parejas de secciones y aplica todos los
          reemplazos en una única copia de B.
        </p>

        <p className="connection" role="status">
          {connection}
        </p>
      </header>

      <div className="document-grid">
        {renderDocumentPanel('A')}
        {renderDocumentPanel('B')}
      </div>

      <section className="transfer-panel">
        <h2>Preparar reemplazo</h2>

        <p>
          <strong>Origen A: </strong>
          {sectionA?.title ?? 'Sin seleccionar'}
        </p>

        <p>
          <strong>Destino B: </strong>
          {sectionB?.title ?? 'Sin seleccionar'}
        </p>

        <button
          type="button"
          disabled={!canAdd}
          onClick={addReplacement}
        >
          Agregar a la lista
        </button>

        <h2 className="pending-heading">
          Transferencias pendientes ({pending.length})
        </h2>

        {pending.length === 0 ? (
          <p>
            Selecciona una sección de A y su destino en B,
            y agrégalos a la lista.
          </p>
        ) : (
          <>
            <div className="pending-table-wrapper">
              <table className="pending-table">
                <caption>
                  Parejas que se aplicarán al documento B
                </caption>

                <thead>
                  <tr>
                    <th scope="col">#</th>
                    <th scope="col">Origen A</th>
                    <th scope="col">Destino B</th>
                    <th scope="col">Acción</th>
                  </tr>
                </thead>

                <tbody>
                  {pending.map((item, index) => (
                    <tr key={item.id}>
                      <td>{index + 1}</td>
                      <td>{item.source.title}</td>
                      <td>{item.target.title}</td>
                      <td>
                        <button
                          className="remove-button"
                          type="button"
                          disabled={busy}
                          onClick={() =>
                            removeReplacement(item.id)
                          }
                        >
                          Quitar
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="batch-actions">
              <button
                type="button"
                disabled={!canApply}
                onClick={applyBatch}
              >
                {operation === 'transfer'
                  ? 'Aplicando cambios...'
                  : `Aplicar todos los cambios (${pending.length})`}
              </button>

              <button
                className="secondary-button"
                type="button"
                disabled={busy}
                onClick={clearPending}
              >
                Vaciar lista
              </button>
            </div>
          </>
        )}

        {transferError && (
          <p className="error" role="alert">
            {transferError}
          </p>
        )}

        {transferMessage && (
          <p role="status">{transferMessage}</p>
        )}

        {lastSummary && (
          <section
            className="result-summary"
            aria-labelledby="result-summary-title"
          >
            <h2 id="result-summary-title">
              Resumen del último lote completado
            </h2>

            <dl className="summary-grid">
              <div>
                <dt>Secciones reemplazadas</dt>
                <dd>{lastSummary.replacement_count}</dd>
              </div>

              <div>
                <dt>Párrafos copiados</dt>
                <dd>{lastSummary.copied_paragraphs}</dd>
              </div>

              <div>
                <dt>Tablas copiadas</dt>
                <dd>{lastSummary.copied_tables}</dd>
              </div>

              <div>
                <dt>Imágenes copiadas</dt>
                <dd>{lastSummary.copied_images}</dd>
              </div>

              <div>
                <dt>Listas creadas</dt>
                <dd>{lastSummary.created_lists}</dd>
              </div>
            </dl>

            <div className="pending-table-wrapper">
              <table className="pending-table">
                <caption>
                  Detalle de los reemplazos completados
                </caption>

                <thead>
                  <tr>
                    <th scope="col">Origen A</th>
                    <th scope="col">Destino B</th>
                    <th scope="col">Párrafos</th>
                    <th scope="col">Tablas</th>
                    <th scope="col">Imágenes</th>
                    <th scope="col">Listas</th>
                  </tr>
                </thead>

                <tbody>
                  {lastSummary.replacements.map((item, index) => (
                    <tr key={index}>
                      <td>{item.source_title}</td>
                      <td>{item.target_title}</td>
                      <td>{item.copied_paragraphs}</td>
                      <td>{item.copied_tables}</td>
                      <td>{item.copied_images}</td>
                      <td>{item.created_lists}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <p className="summary-note">
              Los párrafos incluyen los de las celdas y los vacíos.
              Las cifras corresponden al contenido insertado por
              este lote, no al documento B completo.
            </p>
          </section>
        )}

        {serverDownloadPath && (
          <p>
            <a
              href={`${API_URL}${serverDownloadPath}`}
              target="_blank"
              rel="noopener noreferrer"
            >
              Recuperar el DOCX generado desde el backend
            </a>
          </p>
        )}

        {downloadUrl && outputFile && (
          <p>
            <a
              className="download-link"
              href={downloadUrl}
              download={outputFile.name}
            >
              Descargar B actualizado
            </a>
          </p>
        )}
      </section>

      <p className="stage-note">
        Versión de prueba: texto y subtítulos. No se admiten
        listas, imágenes ni tablas de origen. Cambiar o volver
        a cargar un documento vacía la lista pendiente.
        Los archivos originales no se sobrescriben.
      </p>
    </main>
  )
}

export default App