/**
 * MS Office chart extension URIs and namespaces.
 *
 * Shared between the chart-space renderer (`chart-space-render.ts`) and the
 * chart-space parser (`chart-space-xform.ts`); kept in a neutral module so
 * neither side depends on the other for these constants.
 */

// MS Office 2010 pivot chart options extension — ECMA-376 MS-XLSX §2.3.11.
export const C14_PIVOT_OPTIONS_EXT_URI = "{781A3756-C4B2-4CAC-9D66-4F8BD8637D16}";
export const C14_CHART_NAMESPACE = "http://schemas.microsoft.com/office/drawing/2007/8/2/chart";

// Office 2014 pivot chart options16 extension — sibling to c14:pivotOptions.
export const C16_PIVOT_OPTIONS16_EXT_URI = "{E28EC0CA-F0BB-4C9C-879D-F8772B89E7AC}";
export const C16_CHART_NAMESPACE = "http://schemas.microsoft.com/office/drawing/2014/chart";

// Excel 2013+ "Value From Cells" / dataLabelsRange extension.
export const C15_DATA_LABELS_RANGE_EXT_URI = "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}";
export const C15_CHART_NAMESPACE = "http://schemas.microsoft.com/office/drawing/2012/chart";
