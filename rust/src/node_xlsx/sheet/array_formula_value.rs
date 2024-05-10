use std::collections::HashMap;

use neon::{
    context::{Context, FunctionContext},
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsNumber, JsObject},
};
use rust_xlsxwriter::{Format, Formula};

use crate::node_xlsx::util::{create_format, object_to_formula};

pub struct ArrayFormulaSheetValue {
    pub first_row: u32,
    pub last_row: u32,
    pub first_column: u16,
    pub last_column: u16,
    pub formula: Formula,
    pub dynamic: bool,
    pub format: Option<u32>,
}

impl ArrayFormulaSheetValue {
    pub fn from_js_object(
        cx: &mut FunctionContext,
        obj: Handle<JsObject>,
        format_map: &mut HashMap<u32, Format>,
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

        let formula: Handle<JsObject> = obj.get(cx, "formula")?;
        let (formula, dynamic) = object_to_formula(cx, formula)?;

        let format: Option<Handle<JsObject>> = obj.get_opt(cx, "format")?;
        let format = match format {
            Some(format) => Some(create_format(cx, format, format_map)?),
            None => None,
        };

        Ok(Self {
            first_row,
            last_row,
            first_column,
            last_column,
            formula,
            dynamic,
            format,
        })
    }
}
