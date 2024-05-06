use neon::{
    context::FunctionContext,
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsArray, JsObject, JsString, JsValue},
};

use super::cell::NodeXlsxCell;

#[derive(Debug)]
pub struct NodeXlsxSheet {
    pub name: String,
    pub cells: Vec<NodeXlsxCell>,
}

impl NodeXlsxSheet {
    pub fn from_js_object<'a>(
        cx: &mut FunctionContext<'a>,
        obj: Handle<JsObject>,
    ) -> NeonResult<Self> {
        let name: Handle<JsString> = obj.get(cx, "name")?;
        let name = name.value(cx);
        let mut inner_cells = vec![];

        let cells: Handle<JsArray> = obj.get(cx, "cells")?;
        let cells: Vec<Handle<JsValue>> = cells.to_vec(cx)?;
        for cell in cells {
            let cell = cell.downcast_or_throw::<JsObject, FunctionContext>(cx)?;
            let cell = NodeXlsxCell::from_js_object(cx, cell)?;
            inner_cells.push(cell);
        }
        return Ok(Self {
            name,
            cells: inner_cells,
        });
    }
}
