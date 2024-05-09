use std::collections::HashMap;

use neon::{
    context::{Context, FunctionContext},
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsNumber, JsObject, JsString},
};
use rust_xlsxwriter::{Format, Worksheet};

use crate::node_xlsx::conditional_format::{
    average::Average,
    blank::Blank,
    c_type::NodeXlsxConditionalFormatType,
    scale::{ThreeColorScale, TwoColorScale},
};

pub struct ConditionalFormatSheetValue {
    pub first_row: u32,
    pub last_row: u32,
    pub first_column: u16,
    pub last_column: u16,
    pub format: u32,
}

impl ConditionalFormatSheetValue {
    pub fn from_js_object(
        cx: &mut FunctionContext,
        obj: Handle<JsObject>,
        format_map: &mut HashMap<u32, Format>,
        conditional_format_map: &mut HashMap<u32, NodeXlsxConditionalFormatType>,
    ) -> NeonResult<Self> {
        let first_row: Handle<JsNumber> = obj.get(cx, "firstRow")?;
        let first_row = first_row.value(cx);
        if first_row < 0.0 || first_row >= 1_048_577.0 {
            let error = format!("Row with illegal number {}", first_row.to_string());
            let js_string = cx.string(error);
            return cx.throw(js_string);
        }
        let first_row = first_row as u32;

        let last_row: Handle<JsNumber> = obj.get(cx, "lastRow")?;
        let last_row = last_row.value(cx);
        if last_row < 0.0 || last_row >= 1_048_577.0 {
            let error = format!("Row with illegal number {}", last_row.to_string());
            let js_string = cx.string(error);
            return cx.throw(js_string);
        }
        let last_row = last_row as u32;

        let first_column: Handle<JsNumber> = obj.get(cx, "firstColumn")?;
        let first_column = first_column.value(cx);
        if first_column < 0.0 || first_column >= 16_384.0 {
            let error = format!("Column with illegal number {}", first_column.to_string());
            let js_string = cx.string(error);
            return cx.throw(js_string);
        }
        let first_column = first_column as u16;

        let last_column: Handle<JsNumber> = obj.get(cx, "lastColumn")?;
        let last_column = last_column.value(cx);
        if last_column < 0.0 || last_column >= 16_384.0 {
            let error = format!("Column with illegal number {}", last_column.to_string());
            let js_string = cx.string(error);
            return cx.throw(js_string);
        }
        let last_column = last_column as u16;

        let format: Handle<JsObject> = obj.get(cx, "format")?;
        let f_type: Handle<JsString> = format.get(cx, "type")?;

        let f_type = f_type.value(cx);
        let format = match f_type.as_str() {
            "twoColorScale" => {
                let id = TwoColorScale::create_and_set_to_map(cx, format, conditional_format_map)?;
                id
            }
            "threeColorScale" => {
                let id =
                    ThreeColorScale::create_and_set_to_map(cx, format, conditional_format_map)?;
                id
            }
            "average" => {
                let id =
                    Average::create_and_set_to_map(cx, format, conditional_format_map, format_map)?;
                id
            }
            "blank" => {
                let id =
                    Blank::create_and_set_to_map(cx, format, conditional_format_map, format_map)?;
                id
            }
            _ => {
                let err = format!("Invalid ConditionalFormatType: {}", f_type);
                return cx.throw_error(err);
            }
        };

        return Ok(Self {
            first_row,
            last_row,
            first_column,
            last_column,
            format,
        });
    }

    pub fn set_conditional_format(
        &self,
        worksheet: &mut Worksheet,
        format_map: &HashMap<u32, NodeXlsxConditionalFormatType>,
    ) -> Result<(), rust_xlsxwriter::XlsxError> {
        let cf = format_map.get(&self.format).unwrap();
        cf.set_conditional_format(
            worksheet,
            self.first_row,
            self.last_row,
            self.first_column,
            self.last_column,
        )?;
        Ok(())
    }
}
