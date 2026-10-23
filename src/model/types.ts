/**
 * In-memory model of a Docentric AX template (.docx).
 *
 * A Docentric template is a plain Word package plus one custom XML part whose
 * root is `ProjectData` (namespace http://Docentric.org/ProjectData.xml). That
 * part holds everything the Docentric Designer add-in edits:
 *   - the data source package (report id, DSP class, languages, template labels)
 *   - the data sources (XSD + sample XML copied from the .ddsp)
 *   - the element registry (Field, If, List, Label, Var, ...) with their bindings
 * Each element is placed in the document as a content control (`w:sdt`) whose
 * tag is `<projectId>-<elementId>` and whose alias is the element type.
 */

export type ElementType =
  | 'Field'
  | 'If'
  | 'List'
  | 'Label'
  | 'Var'
  | 'SummaryField'
  | 'Image'
  | 'Group'
  | 'Chart'
  | 'Barcode'
  | 'Hyperlink'
  | string;

export interface Binding {
  /** Data source name, or '.' for the current context (enclosing List / SummaryField). */
  source: string | null;
  path: string;
  resultType: string | null;
}

export interface TemplateElement {
  id: string;
  type: ElementType;
  /** `Name` attribute — set on Var elements (the variable name) and some others. */
  name: string | null;
  binding: Binding | null;
  /** SummaryField only: the per-row expression, evaluated relative to `binding`. */
  aggregateBinding: Binding | null;
  /** Label elements: the label code they render, e.g. `@DOC_LABEL:Buyer`. */
  labelCode: string | null;
  childIds: string[];
  /** Remaining attributes (RenderMode, FormatString, AggregateFunction, ...). */
  attributes: Record<string, string>;
}

export interface TemplateLabel {
  name: string;
  code: string;
  section: string;
  /** language → text */
  texts: Record<string, string>;
}

export type DataSectionType = 'MainData' | 'GeneralData' | 'Parameters' | 'Unknown';

export interface SchemaNode {
  name: string;
  /** attribute name → xs type (without prefix) */
  attributes: Map<string, string>;
  children: Map<string, SchemaNode>;
  repeated: boolean;
}

export interface DataSource {
  name: string;
  section: DataSectionType;
  /** Root element of the XSD (e.g. ReportData). Paths with Source=<name> are relative to it. */
  root: SchemaNode | null;
  hasSampleXml: boolean;
}

export interface ReportInfo {
  dataSourceType: string | null;
  reportId: string | null;
  reportGroupId: string | null;
  dataSourceProviderClass: string | null;
  ddspUri: string | null;
}

export interface ContentControl {
  /** Package part, e.g. word/document.xml or word/header2.xml */
  part: string;
  tag: string;
  projectId: string | null;
  elementId: string | null;
  /** The element type Docentric writes as the content control title. */
  alias: string | null;
  /** Element id of the nearest enclosing Docentric content control, if any. */
  parentElementId: string | null;
}

export interface DocentricTemplate {
  productVersion: string | null;
  report: ReportInfo;
  defaultLanguage: string | null;
  languages: string[];
  labels: TemplateLabel[];
  dataSources: DataSource[];
  elements: Map<string, TemplateElement>;
  contentControls: ContentControl[];
  /** Part name of the ProjectData custom XML. */
  projectDataPart: string;
  /**
   * Data sources whose schema the docentric skill generated (no DDSP yet), from the
   * document property DocentricSkill.ProvisionalDataSources.
   */
  provisionalDataSources: string[];
}
