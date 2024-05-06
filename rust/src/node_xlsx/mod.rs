use neon::{
    context::FunctionContext,
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsArray, JsObject, JsValue},
};
use rust_xlsxwriter::{Url, Workbook, Worksheet};

use self::{error::NodeXlsxError, sheet::NodeXlsxSheet, types::NodeXlsxTypes};

mod cell;
mod error;
mod sheet;
mod types;

#[derive(Debug)]
pub struct NodeXlsxWorkbook {
    sheets: Vec<NodeXlsxSheet>,
}

impl NodeXlsxWorkbook {
    pub fn from_js_object<'a>(
        cx: &mut FunctionContext<'a>,
        obj: Handle<JsObject>,
    ) -> NeonResult<Self> {
        let mut inner_sheets = vec![];

        let sheets: Handle<JsArray> = obj.get(cx, "sheets")?;
        let sheets: Vec<Handle<JsValue>> = sheets.to_vec(cx)?;
        for sheet in sheets {
            let sheet = sheet.downcast_or_throw::<JsObject, FunctionContext>(cx)?;
            let sheet = NodeXlsxSheet::from_js_object(cx, sheet)?;
            inner_sheets.push(sheet);
        }
        return Ok(Self {
            sheets: inner_sheets,
        });
    }

    pub fn save_to_buffer(self) -> Result<Vec<u8>, NodeXlsxError> {
        let mut workbook = self.parse()?;
        let buffer = workbook.save_to_buffer()?;
        return Ok(buffer);
    }

    pub fn save_to_file(self, path: &str) -> Result<(), NodeXlsxError> {
        let mut workbook = self.parse()?;
        workbook.save(path)?;
        return Ok(());
    }

    fn parse(self) -> Result<Workbook, NodeXlsxError> {
        let mut workbook = rust_xlsxwriter::Workbook::new();
        for sheet in self.sheets {
            let mut worksheet = Worksheet::new();
            worksheet.set_name(&sheet.name)?;
            for cell in sheet.cells {
                match cell.cel_type {
                    NodeXlsxTypes::String => {
                        worksheet.write_string(cell.row, cell.col, &cell.value)?;
                    }
                    NodeXlsxTypes::Number => {
                        let value: f64 = cell.value.parse::<f64>()?;
                        worksheet.write_number(cell.row, cell.col, value)?;
                    }
                    NodeXlsxTypes::Link => {
                        let url = Url::new(&cell.value);
                        worksheet.write(cell.row, cell.col, url)?;
                    }
                }
            }
            workbook.push_worksheet(worksheet);
        }
        return Ok(workbook);
    }
}

#[cfg(test)]
mod test {
    use crate::node_xlsx::{
        cell::NodeXlsxCell, sheet::NodeXlsxSheet, types::NodeXlsxTypes, NodeXlsxWorkbook,
    };

    #[test]
    fn test_save_to_buffer() {
        let sheet = NodeXlsxSheet {
            name: "test".to_string(),
            cells: vec![
                NodeXlsxCell {
                    col: 0,
                    row: 0,
                    value: "header 1".to_string(),
                    cel_type: NodeXlsxTypes::String,
                },
                NodeXlsxCell {
                    col: 1,
                    row: 0,
                    value: "header 2".to_string(),
                    cel_type: NodeXlsxTypes::String,
                },
                NodeXlsxCell {
                    col: 0,
                    row: 1,
                    value: "1".to_string(),
                    cel_type: NodeXlsxTypes::Number,
                },
                NodeXlsxCell {
                    col: 1,
                    row: 1,
                    value: "2".to_string(),
                    cel_type: NodeXlsxTypes::Number,
                },
                NodeXlsxCell {
                    col: 0,
                    row: 2,
                    value: "https://example.com".to_string(),
                    cel_type: NodeXlsxTypes::Link,
                },
            ],
        };

        let workbook = NodeXlsxWorkbook {
            sheets: vec![sheet],
        };
        let buffer = workbook.save_to_buffer().unwrap();
        assert!(buffer.len() > 0);

        //write to file
        std::fs::write("temp/test_buff.xlsx", &buffer).unwrap();
    }

    #[test]
    fn test_save_to_file() {
        let sheet = NodeXlsxSheet {
            name: "test".to_string(),
            cells: vec![
                NodeXlsxCell {
                    col: 0,
                    row: 0,
                    value: "header 1".to_string(),
                    cel_type: NodeXlsxTypes::String,
                },
                NodeXlsxCell {
                    col: 1,
                    row: 0,
                    value: "header 2".to_string(),
                    cel_type: NodeXlsxTypes::String,
                },
                NodeXlsxCell {
                    col: 0,
                    row: 1,
                    value: "1".to_string(),
                    cel_type: NodeXlsxTypes::Number,
                },
                NodeXlsxCell {
                    col: 1,
                    row: 1,
                    value: "2".to_string(),
                    cel_type: NodeXlsxTypes::Number,
                },
                NodeXlsxCell {
                    col: 0,
                    row: 2,
                    value: "https://example.com".to_string(),
                    cel_type: NodeXlsxTypes::Link,
                },
            ],
        };

        let workbook = NodeXlsxWorkbook {
            sheets: vec![sheet],
        };
        workbook.save_to_file("temp/test.xlsx").unwrap();
    }
}
