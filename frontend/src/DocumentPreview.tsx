import { useEffect, useRef, useState } from 'react'
import { renderAsync } from 'docx-preview'
import type { DocumentSection } from './types'

type Props = {
  file: File
  sections?: DocumentSection[]
  selectedSectionId?: string | null
}

const MAX_FILE_SIZE = 10 * 1024 * 1024
const EMPTY_SECTIONS: DocumentSection[] = []

function normalizeText(value: string) {
  return value
    .normalize('NFKC')
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
}

export default function DocumentPreview({
  file,
  sections = EMPTY_SECTIONS,
  selectedSectionId = null,
}: Props) {
  const containerRef = useRef<HTMLDivElement | null>(null)
  const viewportRef = useRef<HTMLDivElement | null>(null)

  const [loading, setLoading] = useState(true)
  const [error, setError] = useState('')
  const [navigationMessage, setNavigationMessage] = useState('')
  const [renderVersion, setRenderVersion] = useState(0)

  useEffect(() => {
    let active = true

    const visibleContainer = containerRef.current

    if (!visibleContainer) {
      return
    }

    visibleContainer.replaceChildren()
    setLoading(true)
    setError('')
    setNavigationMessage('')

    async function generatePreview() {
      try {
        if (file.size > MAX_FILE_SIZE) {
          throw new Error(
            'La previsualización admite archivos de hasta 10 MB.',
          )
        }

        const buffer = await file.arrayBuffer()

        if (!active) {
          return
        }

        const temporaryContainer = document.createElement('div')

        await renderAsync(
          buffer,
          temporaryContainer,
          temporaryContainer,
          {
            className: 'word-preview',
            inWrapper: true,
            ignoreWidth: false,
            ignoreHeight: false,
            breakPages: true,
          },
        )

        if (!active) {
          return
        }

        visibleContainer.replaceChildren(
          ...Array.from(temporaryContainer.childNodes),
        )

        setRenderVersion((version) => version + 1)
      } catch (error) {
        if (active) {
          console.error('Error al previsualizar el DOCX:', error)

          setError(
            error instanceof Error &&
              error.message ===
                'La previsualización admite archivos de hasta 10 MB.'
              ? error.message
              : 'No se pudo previsualizar este documento.',
          )
        }
      } finally {
        if (active) {
          setLoading(false)
        }
      }
    }

    void generatePreview()

    return () => {
      active = false
      visibleContainer.replaceChildren()
    }
  }, [file])

  useEffect(() => {
    const container = containerRef.current
    const viewport = viewportRef.current

    if (!container || !viewport || loading || error) {
      return
    }

    // Limpia los resaltados anteriores.
    container
      .querySelectorAll('.selected-word-heading')
      .forEach((element) => {
        element.classList.remove('selected-word-heading')
      })

    setNavigationMessage('')

    if (!selectedSectionId) {
      return
    }

    // Busca párrafos del cuerpo renderizado, excluyendo
    // encabezados y pies de página.
    const candidates = Array.from(
      container.querySelectorAll<HTMLElement>(
        'article p, article h1, article h2, article h3, ' +
          'article h4, article h5, article h6',
      ),
    )

    const targets = new Map<string, HTMLElement>()
    let nextCandidate = 0

    // Asocia los títulos siguiendo su orden en el documento.
    // Esto permite distinguir títulos repetidos.
    for (const section of sections) {
      const expectedText = normalizeText(section.title)

      for (
        let index = nextCandidate;
        index < candidates.length;
        index += 1
      ) {
        const candidate = candidates[index]
        const actualText = normalizeText(
          candidate.textContent ?? '',
        )

        if (actualText === expectedText) {
          targets.set(section.id, candidate)
          nextCandidate = index + 1
          break
        }
      }
    }

    const target = targets.get(selectedSectionId)

    if (!target) {
      setNavigationMessage(
        'La sección está seleccionada, pero no se pudo localizar ' +
          'su título en la previsualización.',
      )
      return
    }

    target.classList.add('selected-word-heading')

    const viewportRect = viewport.getBoundingClientRect()
    const targetRect = target.getBoundingClientRect()

    const top =
      viewport.scrollTop +
      targetRect.top -
      viewportRect.top -
      viewport.clientTop -
      24

    // Desplaza únicamente este panel, no toda la página.
    viewport.scrollTo({
      top: Math.max(0, top),
      behavior: 'smooth',
    })
  }, [
    sections,
    selectedSectionId,
    renderVersion,
    loading,
    error,
  ])

  return (
    <section className="document-preview">
      <h3>Previsualización</h3>

      {loading && (
        <p role="status">Generando previsualización...</p>
      )}

      {error && (
        <p className="error" role="alert">
          {error}
        </p>
      )}

      {navigationMessage && (
        <p className="navigation-note" role="status">
          {navigationMessage}
        </p>
      )}

      <div
        ref={viewportRef}
        className="preview-viewport"
        aria-busy={loading}
        tabIndex={0}
        aria-label={`Previsualización de ${file.name}`}
      >
        <div ref={containerRef} />
      </div>
    </section>
  )
}