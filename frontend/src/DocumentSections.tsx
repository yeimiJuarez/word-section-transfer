import type {
  DocumentSection,
  LoadedDocument,
} from './types'

type Props = {
  document: LoadedDocument | null
  selectedId: string | null
  documentLabel: 'A' | 'B'
  disabled: boolean
  onSelect: (section: DocumentSection) => void
}

export default function DocumentSections({
  document,
  selectedId,
  documentLabel,
  disabled,
  onSelect,
}: Props) {
  if (!document) {
    return null
  }

  const selectedSection = document.sections.find(
    (section) => section.id === selectedId,
  )

  return (
    <div className="section-browser">
      <p>
        {document.sections.length} secciones detectadas.
      </p>

      {document.sections.length === 0 && (
        <p>No se encontraron títulos en el cuerpo principal.</p>
      )}

      <ul className="section-list">
        {document.sections.map((section) => (
          <li
            key={section.id}
            style={{
              paddingLeft: Math.max(0, section.level - 1) * 12,
            }}
          >
            <button
              type="button"
              className="section-button"
              disabled={disabled || !section.transfer_enabled}
              aria-pressed={selectedId === section.id}
              onClick={() => onSelect(section)}
            >
              {section.title}

              {!section.transfer_enabled && (
                <span className="section-note">
                  {' — Transferencia deshabilitada'}
                </span>
              )}
            </button>
          </li>
        ))}
      </ul>

      {selectedSection && (
        <p>
          {documentLabel === 'A'
            ? 'Origen seleccionado: '
            : 'Destino seleccionado: '}
          <strong>{selectedSection.title}</strong>
        </p>
      )}
    </div>
  )
}