use std::collections::HashMap;

use super::{format::NodeXlsxFormat, types::NodeXlsxTypes};
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
            Some(format) => Some(Self::create_format(cx, format, format_map)?),
            None => None,
        };

        Ok(Self {
            col,
            row,
            cell_type: cel_type,
            format,
        })
    }

    fn create_format(
        cx: &mut FunctionContext,
        obj: Handle<JsObject>,
        format_map: &mut HashMap<u32, rust_xlsxwriter::Format>,
    ) -> NeonResult<u32> {
        let id: Handle<JsNumber> = obj.get(cx, "id")?;
        let id = id.value(cx) as u32;
        if !format_map.contains_key(&id) {
            let format = NodeXlsxFormat::from_js_object(cx, obj)?;
            format_map.insert(id, format.into());
        }
        Ok(id)
    }
}
