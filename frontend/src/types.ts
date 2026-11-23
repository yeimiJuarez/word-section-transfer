export type DocumentSection = {
  id: string
  title: string
  level: number
  parent_id: string | null
  transfer_enabled: boolean
  start_index: number
  content_start_index: number
  end_index: number
}

export type LoadedDocument = {
  document_id: string
  filename: string
  section_count: number
  sections: DocumentSection[]
}

export type DocumentSelection = {
  documentId: string
  filename: string
  section: DocumentSection
}