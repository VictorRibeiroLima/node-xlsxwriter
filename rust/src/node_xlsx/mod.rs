use base64::{engine::general_purpose, Engine};
use neon::{
    context::FunctionContext,
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsArray, JsObject, JsValue},
};
use rust_xlsxwriter::{Workbook, Worksheet};

use self::{error::NodeXlsxError, sheet::NodeXlsxSheet, types::NodeXlsxTypes};

mod border;
mod cell;
mod color;
mod conditional_format;
mod error;
mod format;
mod sheet;
mod types;
mod util;

pub struct NodeXlsxWorkbook {
    sheets: Vec<NodeXlsxSheet>,
}

impl NodeXlsxWorkbook {
    pub fn from_js_object<'b>(
        cx: &mut FunctionContext<'b>,
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

    pub fn save_to_base64(self) -> Result<String, NodeXlsxError> {
        let buffer = self.save_to_buffer()?;
        let base64 = general_purpose::STANDARD.encode(buffer);
        return Ok(base64);
    }

    fn parse(self) -> Result<Workbook, NodeXlsxError> {
        let mut workbook = rust_xlsxwriter::Workbook::new();
        for sheet in self.sheets {
            let format_map = sheet.format_map;
            let conditional_format_map = sheet.conditional_format_map;
            let mut worksheet = Worksheet::new();
            worksheet.set_name(&sheet.name)?;
            for cf in sheet.conditional_formats {
                cf.set_conditional_format(&mut worksheet, &conditional_format_map)?;
            }
            for cell in sheet.cells {
                match cell.cell_type {
                    NodeXlsxTypes::String(value) | NodeXlsxTypes::Unknown(value) => {
                        if let Some(format) = cell.format {
                            let format = format_map.get(&format).unwrap();
                            worksheet
                                .write_string_with_format(cell.row, cell.col, &value, format)?;
                        } else {
                            worksheet.write_string(cell.row, cell.col, &value)?;
                        }
                    }
                    NodeXlsxTypes::Number(value) => {
                        if let Some(format) = cell.format {
                            let format = format_map.get(&format).unwrap();
                            worksheet
                                .write_number_with_format(cell.row, cell.col, value, format)?;
                        } else {
                            worksheet.write_number(cell.row, cell.col, value)?;
                        }
                    }
                    NodeXlsxTypes::Link(value) => {
                        if let Some(format) = cell.format {
                            let format = format_map.get(&format).unwrap();
                            worksheet.write_url_with_format(cell.row, cell.col, value, format)?;
                        } else {
                            worksheet.write_url(cell.row, cell.col, value)?;
                        }
                    }
                    NodeXlsxTypes::Date(value) => {
                        if let Some(format) = cell.format {
                            let format = format_map.get(&format).unwrap();
                            worksheet
                                .write_datetime_with_format(cell.row, cell.col, value, format)?;
                        } else {
                            worksheet.write_datetime(cell.row, cell.col, value)?;
                        }
                    }
                    NodeXlsxTypes::Formula(value) => {
                        if let Some(format) = cell.format {
                            let format = format_map.get(&format).unwrap();
                            worksheet
                                .write_formula_with_format(cell.row, cell.col, value, format)?;
                        } else {
                            worksheet.write_formula(cell.row, cell.col, value)?;
                        }
                    }
                }
            }
            workbook.push_worksheet(worksheet);
        }
        return Ok(workbook);
    }
}
