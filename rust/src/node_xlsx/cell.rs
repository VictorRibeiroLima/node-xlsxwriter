use std::collections::HashMap;

use super::{types::NodeXlsxTypes, util::create_format};

use neon::{
    context::{Context, FunctionContext},
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsNumber, JsObject, JsString, JsValue},
};

pub struct NodeXlsxCell {
    pub col: u16,
    pub row: u32,
    pub cell_type: NodeXlsxTypes,
    pub format: Option<u32>,
}

impl NodeXlsxCell {
    pub fn from_js_object(
        cx: &mut FunctionContext,
        obj: Handle<JsObject>,
        format_map: &mut HashMap<u32, rust_xlsxwriter::Format>,
    ) -> NeonResult<Self> {
        let result = Self::inner_from_js_object(cx, &obj, format_map);
        match result {
            Ok(cell) => Ok(cell),
            Err(error) => {
                let error = format!("Error parsing cell: {:?} with error:\n  {}", obj, error);
                let js_string = cx.string(error);
                cx.throw(js_string)
            }
        }
    }

    fn inner_from_js_object(
        cx: &mut FunctionContext,
        obj: &Handle<JsObject>,
        format_map: &mut HashMap<u32, rust_xlsxwriter::Format>,
    ) -> NeonResult<Self> {
        let col: Handle<JsNumber> = obj.get(cx, "col")?;
        let col = col.value(cx);
        if col < 0.0 || col >= 16_384.0 {
            let error = format!("Column with illegal number {}", col.to_string());
            let js_string = cx.string(error);
            return cx.throw(js_string);
        }
        let col = u16::try_from(col as u64);
        let col = match col {
            Ok(col) => col,
            Err(_) => {
                let js_string = cx.string("Column number is too large");
                return cx.throw(js_string);
            }
        };

        let row: Handle<JsNumber> = obj.get(cx, "row")?;
        let row = row.value(cx);
        if row < 0.0 || row >= 1_048_577.0 {
            let error = format!("Row with illegal number {}", row.to_string());
            let js_string = cx.string(error);
            return cx.throw(js_string);
        }
        let row = row as u32;

        let cel_type: Option<Handle<JsString>> = obj.get_opt(cx, "cellType")?;
        let value: Handle<JsValue> = obj.get(cx, "value")?;
        let cel_type = NodeXlsxTypes::from_js_string(cx, cel_type, value)?;

        let format: Option<Handle<JsObject>> = obj.get_opt(cx, "format")?;

        let format = match format {
            Some(format) => Some(create_format(cx, format, format_map)?),
            None => None,
        };

        Ok(Self {
            col,
            row,
            cell_type: cel_type,
            format,
        })
    }

    pub fn write_to_sheet(
        self,
        worksheet: &mut rust_xlsxwriter::Worksheet,
        format_map: &HashMap<u32, rust_xlsxwriter::Format>,
    ) -> Result<(), rust_xlsxwriter::XlsxError> {
        match self.cell_type {
            NodeXlsxTypes::String(value) | NodeXlsxTypes::Unknown(value) => {
                if let Some(format) = self.format {
                    let format = format_map.get(&format).unwrap();
                    worksheet.write_string_with_format(self.row, self.col, &value, format)?;
                } else {
                    worksheet.write_string(self.row, self.col, &value)?;
                }
            }
            NodeXlsxTypes::Number(value) => {
                if let Some(format) = self.format {
                    let format = format_map.get(&format).unwrap();
                    worksheet.write_number_with_format(self.row, self.col, value, format)?;
                } else {
                    worksheet.write_number(self.row, self.col, value)?;
                }
            }
            NodeXlsxTypes::Link(value) => {
                if let Some(format) = self.format {
                    let format = format_map.get(&format).unwrap();
                    worksheet.write_url_with_format(self.row, self.col, value, format)?;
                } else {
                    worksheet.write_url(self.row, self.col, value)?;
                }
            }
            NodeXlsxTypes::Date(value) => {
                if let Some(format) = self.format {
                    let format = format_map.get(&format).unwrap();
                    worksheet.write_datetime_with_format(self.row, self.col, value, format)?;
                } else {
                    worksheet.write_datetime(self.row, self.col, value)?;
                }
            }
            NodeXlsxTypes::Formula((value, dynamic)) => {
                let has_format = self.format.is_some();
                if dynamic && !has_format {
                    worksheet.write_dynamic_formula(self.row, self.col, value)?;
                } else if dynamic && has_format {
                    let format = self.format.unwrap();
                    let format = format_map.get(&format).unwrap();
                    worksheet
                        .write_dynamic_formula_with_format(self.row, self.col, value, format)?;
                } else if !dynamic && !has_format {
                    worksheet.write_formula(self.row, self.col, value)?;
                } else {
                    let format = self.format.unwrap();
                    let format = format_map.get(&format).unwrap();
                    worksheet.write_formula_with_format(self.row, self.col, value, format)?;
                }
            }
        }
        return Ok(());
    }
}
