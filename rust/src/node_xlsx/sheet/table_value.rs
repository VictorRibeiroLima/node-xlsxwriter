use std::collections::HashMap;

use neon::{
    context::FunctionContext,
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsNumber, JsObject},
};
use rust_xlsxwriter::{Format, Table};

use crate::node_xlsx::table::NodeXlsxTable;

pub struct NodeXlsxTableValue {
    pub first_row: u32,
    pub last_row: u32,
    pub first_column: u16,
    pub last_column: u16,
    pub table: Table,
}

impl NodeXlsxTableValue {
    pub fn from_js_object(
        cx: &mut FunctionContext,
        obj: Handle<JsObject>,
        format_map: &mut HashMap<u32, Format>,
    ) -> NeonResult<Self> {
        let first_row: Handle<JsNumber> = obj.get(cx, "firstRow")?;
        let first_row = first_row.value(cx) as u32;

        let last_row: Handle<JsNumber> = obj.get(cx, "lastRow")?;
        let last_row = last_row.value(cx) as u32;

        let first_column: Handle<JsNumber> = obj.get(cx, "firstColumn")?;
        let first_column = first_column.value(cx) as u16;

        let last_column: Handle<JsNumber> = obj.get(cx, "lastColumn")?;
        let last_column = last_column.value(cx) as u16;

        let table: Handle<JsObject> = obj.get(cx, "table")?;
        let table = NodeXlsxTable::from_js_object(cx, table, format_map)?;

        Ok(Self {
            first_row,
            last_row,
            first_column,
            last_column,
            table: table.into(),
        })
    }
}
