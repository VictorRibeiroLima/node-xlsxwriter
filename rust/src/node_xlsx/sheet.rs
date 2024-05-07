use std::collections::HashMap;

use neon::{
    context::FunctionContext,
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsArray, JsObject, JsString, JsValue},
};
use rust_xlsxwriter::Format;

use super::cell::NodeXlsxCell;

pub struct NodeXlsxSheet {
    pub name: String,
    pub cells: Vec<NodeXlsxCell>,
    pub format_map: HashMap<u32, Format>,
}

impl NodeXlsxSheet {
    pub fn from_js_object(cx: &mut FunctionContext, obj: Handle<JsObject>) -> NeonResult<Self> {
        let name: Handle<JsString> = obj.get(cx, "name")?;
        let name = name.value(cx);

        let cells: Handle<JsArray> = obj.get(cx, "cells")?;
        let cells: Vec<Handle<JsValue>> = cells.to_vec(cx)?;

        let mut inner_cells = vec![];
        let mut format_map = HashMap::new();

        for cell in cells {
            let cell = cell.downcast_or_throw::<JsObject, FunctionContext>(cx)?;
            let cell = NodeXlsxCell::from_js_object(cx, cell, &mut format_map)?;
            inner_cells.push(cell);
        }

        return Ok(Self {
            name,
            cells: inner_cells,
            format_map,
        });
    }
}
