/**
 * A synthetic AxReport (SSRS) for a fictitious invoice: two datasets, a page
 * header, a header textbox reading First(), a Tablix with header / group /
 * detail / total rows, a database image and an expression the converter cannot
 * translate. Shaped like Visual Studio's AxReport metadata, written from scratch.
 */

const field = (name: string, type: string) =>
  `<AxReportDataSetField><Name>${name}</Name><Alias>Tmp.1.${name}</Alias><DataType>${type}</DataType></AxReportDataSetField>`;

const textbox = (name: string, value: string, top: string, left: string, width: string, style = '') =>
  `<Textbox Name="${name}"><Paragraphs><Paragraph><TextRuns><TextRun><Value>${value}</Value>` +
  `<Style>${style}</Style></TextRun></TextRuns></Paragraph></Paragraphs>` +
  `<Top>${top}</Top><Left>${left}</Left><Width>${width}</Width><Height>0.25in</Height></Textbox>`;

const cell = (name: string, value: string, style = '') =>
  `<TablixCell><CellContents><Textbox Name="${name}"><Paragraphs><Paragraph><TextRuns><TextRun>` +
  `<Value>${value}</Value><Style>${style}</Style></TextRun></TextRuns></Paragraph></Paragraphs></Textbox></CellContents></TablixCell>`;

const rdl = `<?xml version="1.0" encoding="utf-8"?>
<Report xmlns="http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition">
  <ReportParameters>
    <ReportParameter Name="AX_CompanyName"><DataType>String</DataType></ReportParameter>
    <ReportParameter Name="DocumentTitle"><DataType>String</DataType></ReportParameter>
  </ReportParameters>
  <ReportSections><ReportSection>
    <Body><ReportItems>
      ${textbox('InvoiceId', '=First(Fields!InvoiceId.Value, "ConInvoiceHeaderTmp")', '0in', '0in', '3in', '<FontWeight>Bold</FontWeight><FontSize>14pt</FontSize>')}
      ${textbox('InvoiceDate', '=Format(First(Fields!InvoiceDate.Value, "ConInvoiceHeaderTmp"), "d")', '0in', '4in', '2in', '<TextAlign>Right</TextAlign>')}
      ${textbox('Customer', '="Bill to: " &amp; First(Fields!CustName.Value, "ConInvoiceHeaderTmp")', '0.5in', '0in', '6in')}
      ${textbox('Weird', '=Code.FormatAddress(Fields!CustName.Value)', '0.9in', '0in', '6in')}
      <Image Name="Logo"><Source>Database</Source><Value>=First(Fields!CompanyLogo.Value, "ConInvoiceHeaderTmp")</Value>
        <MIMEType>image/png</MIMEType><Top>1.2in</Top><Left>0in</Left><Width>1.5in</Width><Height>0.6in</Height></Image>
      <Tablix Name="Lines">
        <TablixBody>
          <TablixColumns><TablixColumn><Width>1in</Width></TablixColumn><TablixColumn><Width>3in</Width></TablixColumn><TablixColumn><Width>1in</Width></TablixColumn></TablixColumns>
          <TablixRows>
            <TablixRow><Height>0.25in</Height><TablixCells>${cell('H1', 'Item', '<FontWeight>Bold</FontWeight>')}${cell('H2', '@SYS7576')}${cell('H3', 'Amount')}</TablixCells></TablixRow>
            <TablixRow><Height>0.25in</Height><TablixCells>${cell('G1', '=Fields!ItemGroup.Value')}${cell('G2', '')}${cell('G3', '=Sum(Fields!LineAmount.Value)')}</TablixCells></TablixRow>
            <TablixRow><Height>0.25in</Height><TablixCells>${cell('D1', '=Fields!ItemId.Value')}${cell('D2', '=Fields!Name.Value')}${cell('D3', '=Fields!LineAmount.Value', '<TextAlign>Right</TextAlign>')}</TablixCells></TablixRow>
            <TablixRow><Height>0.25in</Height><TablixCells>${cell('T1', 'Total')}${cell('T2', '')}${cell('T3', '=Sum(Fields!LineAmount.Value)')}</TablixCells></TablixRow>
          </TablixRows>
        </TablixBody>
        <TablixRowHierarchy><TablixMembers>
          <TablixMember><KeepWithGroup>After</KeepWithGroup></TablixMember>
          <TablixMember><Group Name="ItemGroup"><GroupExpressions><GroupExpression>=Fields!ItemGroup.Value</GroupExpression></GroupExpressions></Group>
            <TablixMembers><TablixMember/><TablixMember><Group Name="Details"/></TablixMember></TablixMembers></TablixMember>
          <TablixMember><KeepWithGroup>Before</KeepWithGroup></TablixMember>
        </TablixMembers></TablixRowHierarchy>
        <DataSetName>ConInvoiceLineTmp</DataSetName>
        <Top>2in</Top><Left>0in</Left><Width>5in</Width><Height>1in</Height>
      </Tablix>
    </ReportItems></Body>
    <Page><PageHeader><ReportItems>
      ${textbox('Company', '=Parameters!AX_CompanyName.Value', '0in', '0in', '3in')}
      ${textbox('Title', '=Parameters!DocumentTitle.Value', '0in', '3in', '3in', '<TextAlign>Right</TextAlign>')}
    </ReportItems></PageHeader></Page>
  </ReportSection></ReportSections>
</Report>`;

export const AX_REPORT = `<?xml version="1.0" encoding="utf-8"?>
<AxReport xmlns:i="http://www.w3.org/2001/XMLSchema-instance" xmlns="Microsoft.Dynamics.AX.Metadata.V2">
  <Name>ConInvoice</Name>
  <DataSets>
    <AxReportDataSet xmlns=""><Name>ConInvoiceHeaderTmp</Name><DataSourceType>ReportDataProvider</DataSourceType>
      <Fields>${field('InvoiceId', 'System.String')}${field('InvoiceDate', 'System.DateTime')}${field('CustName', 'System.String')}${field('CompanyLogo', 'System.Byte[]')}</Fields>
    </AxReportDataSet>
    <AxReportDataSet xmlns=""><Name>ConInvoiceLineTmp</Name><DataSourceType>ReportDataProvider</DataSourceType>
      <Fields>${field('ItemGroup', 'System.String')}${field('ItemId', 'System.String')}${field('Name', 'System.String')}${field('LineAmount', 'System.Double')}</Fields>
    </AxReportDataSet>
  </DataSets>
  <Designs>
    <AxReportDesign xmlns="" i:type="AxReportPrecisionDesign"><Name>Report</Name><Text><![CDATA[${rdl}]]></Text></AxReportDesign>
  </Designs>
</AxReport>`;
