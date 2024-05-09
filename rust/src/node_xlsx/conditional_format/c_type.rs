use rust_xlsxwriter::{
    ConditionalFormat, ConditionalFormat2ColorScale, ConditionalFormat3ColorScale,
    ConditionalFormatAverage, ConditionalFormatBlank, ConditionalFormatCell,
    ConditionalFormatCustomIcon, ConditionalFormatDataBar, ConditionalFormatDate,
    ConditionalFormatDuplicate, ConditionalFormatError, ConditionalFormatFormula,
    ConditionalFormatIconSet, ConditionalFormatText, ConditionalFormatTop, ConditionalFormatValue,
    Worksheet, XlsxError,
};

pub enum NodeXlsxConditionalFormatType {
    TwoColorScale(ConditionalFormat2ColorScale),
    ThreeColorScale(ConditionalFormat3ColorScale),
    Average(ConditionalFormatAverage),
    Blank(ConditionalFormatBlank),
    Cell(ConditionalFormatCell),
    DataBar(ConditionalFormatDataBar),
    Date(ConditionalFormatDate),
    Duplicate(ConditionalFormatDuplicate),
    Error(ConditionalFormatError),
    Formula(ConditionalFormatFormula),
    IconSet(ConditionalFormatIconSet),
    Text(ConditionalFormatText),
    Top(ConditionalFormatTop),
}

impl NodeXlsxConditionalFormatType {
    pub fn set_conditional_format<'a>(
        &'a self,
        worksheet: &'a mut Worksheet,
        first_row: u32,
        last_row: u32,
        first_column: u16,
        last_column: u16,
    ) -> Result<&mut Worksheet, XlsxError> {
        match self {
            NodeXlsxConditionalFormatType::TwoColorScale(cf) => {
                worksheet.add_conditional_format(first_row, first_column, last_row, last_column, cf)
            }
            NodeXlsxConditionalFormatType::ThreeColorScale(cf) => {
                worksheet.add_conditional_format(first_row, first_column, last_row, last_column, cf)
            }
            NodeXlsxConditionalFormatType::Average(cf) => {
                worksheet.add_conditional_format(first_row, first_column, last_row, last_column, cf)
            }
            NodeXlsxConditionalFormatType::Blank(cf) => {
                worksheet.add_conditional_format(first_row, first_column, last_row, last_column, cf)
            }
            NodeXlsxConditionalFormatType::Cell(cf) => {
                worksheet.add_conditional_format(first_row, first_column, last_row, last_column, cf)
            }
            NodeXlsxConditionalFormatType::DataBar(cf) => {
                worksheet.add_conditional_format(first_row, first_column, last_row, last_column, cf)
            }
            NodeXlsxConditionalFormatType::Date(cf) => {
                worksheet.add_conditional_format(first_row, first_column, last_row, last_column, cf)
            }
            NodeXlsxConditionalFormatType::Duplicate(cf) => {
                worksheet.add_conditional_format(first_row, first_column, last_row, last_column, cf)
            }
            NodeXlsxConditionalFormatType::Error(cf) => {
                worksheet.add_conditional_format(first_row, first_column, last_row, last_column, cf)
            }
            NodeXlsxConditionalFormatType::Formula(cf) => {
                worksheet.add_conditional_format(first_row, first_column, last_row, last_column, cf)
            }
            NodeXlsxConditionalFormatType::IconSet(cf) => {
                worksheet.add_conditional_format(first_row, first_column, last_row, last_column, cf)
            }
            NodeXlsxConditionalFormatType::Text(cf) => {
                worksheet.add_conditional_format(first_row, first_column, last_row, last_column, cf)
            }
            NodeXlsxConditionalFormatType::Top(cf) => {
                worksheet.add_conditional_format(first_row, first_column, last_row, last_column, cf)
            }
        }
    }
}
