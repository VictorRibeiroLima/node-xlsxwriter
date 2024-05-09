mod conditional_format_value;

use std::collections::HashMap;

use neon::{
    context::FunctionContext,
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsArray, JsObject, JsString, JsValue},
};
use rust_xlsxwriter::Format;

use self::conditional_format_value::ConditionalFormatSheetValue;

use super::{cell::NodeXlsxCell, conditional_format::c_type::NodeXlsxConditionalFormatType};

pub struct NodeXlsxSheet {
    pub name: String,
    pub cells: Vec<NodeXlsxCell>,
    pub conditional_formats: Vec<ConditionalFormatSheetValue>,
    pub format_map: HashMap<u32, Format>,
    pub conditional_format_map: HashMap<u32, NodeXlsxConditionalFormatType>,
}

impl NodeXlsxSheet {
    pub fn from_js_object(cx: &mut FunctionContext, obj: Handle<JsObject>) -> NeonResult<Self> {
        let name: Handle<JsString> = obj.get(cx, "name")?;
        let name = name.value(cx);

        let cells: Handle<JsArray> = obj.get(cx, "cells")?;
        let cells: Vec<Handle<JsValue>> = cells.to_vec(cx)?;

        let conditional_formats: Handle<JsArray> = obj.get(cx, "conditionalFormats")?;
        let conditional_formats: Vec<Handle<JsValue>> = conditional_formats.to_vec(cx)?;

        let mut inner_cells = vec![];
        let mut inner_conditional_formats = vec![];
        let mut format_map = HashMap::new();
        let mut conditional_format_map = HashMap::new();

        for cell in cells {
            let cell = cell.downcast_or_throw::<JsObject, FunctionContext>(cx)?;

            let cell = NodeXlsxCell::from_js_object(cx, cell, &mut format_map)?;
            inner_cells.push(cell);
        }

        for conditional_format in conditional_formats {
            let conditional_format =
                conditional_format.downcast_or_throw::<JsObject, FunctionContext>(cx)?;
            let conditional_format = ConditionalFormatSheetValue::from_js_object(
                cx,
                conditional_format,
                &mut format_map,
                &mut conditional_format_map,
            )?;
            inner_conditional_formats.push(conditional_format);
        }

        return Ok(Self {
            name,
            cells: inner_cells,
            conditional_formats: inner_conditional_formats,
            format_map,
            conditional_format_map,
        });
    }
}
