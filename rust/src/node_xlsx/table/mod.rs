use std::collections::HashMap;

use neon::{
    context::FunctionContext,
    handle::Handle,
    object::Object,
    result::NeonResult,
    types::{JsArray, JsBoolean, JsObject, JsString},
};
use rust_xlsxwriter::{Format, Table, TableColumn, TableStyle};
use util::string_to_table_style;

mod column;
mod util;

pub struct NodeXlsxTable {
    name: Option<String>,
    auto_filter: bool,
    banded_columns: bool,
    banded_rows: bool,
    columns: Vec<TableColumn>,
    first_column_highlighted: bool,
    last_column_highlighted: bool,
    header_row: bool,
    style: Option<TableStyle>,
    total_row: bool,
}

impl NodeXlsxTable {
    pub fn from_js_object(
        cx: &mut FunctionContext,
        obj: Handle<JsObject>,
        format_map: &mut HashMap<u32, Format>,
    ) -> NeonResult<Self> {
        let name: Option<Handle<JsString>> = obj.get_opt(cx, "name")?;
        let name = name.map(|name| name.value(cx));

        let auto_filter: Handle<JsBoolean> = obj.get(cx, "autoFilter")?;
        let auto_filter = auto_filter.value(cx);

        let banded_columns: Handle<JsBoolean> = obj.get(cx, "bandedColumns")?;
        let banded_columns = banded_columns.value(cx);

        let banded_rows: Handle<JsBoolean> = obj.get(cx, "bandedRows")?;
        let banded_rows = banded_rows.value(cx);

        let mut inner_columns: Vec<TableColumn> = vec![];
        let columns: Handle<JsArray> = obj.get(cx, "columns")?;
        let columns = columns.to_vec(cx)?;
        for column in columns {
            let column = column.downcast_or_throw::<JsObject, _>(cx)?;
            let column = column::NodeXlsxTableColumn::create_and_into(cx, column, format_map)?;
            inner_columns.push(column);
        }

        let first_column_highlighted: Handle<JsBoolean> = obj.get(cx, "firstColumnHighlighted")?;
        let first_column_highlighted = first_column_highlighted.value(cx);

        let last_column_highlighted: Handle<JsBoolean> = obj.get(cx, "lastColumnHighlighted")?;
        let last_column_highlighted = last_column_highlighted.value(cx);

        let header_row: Handle<JsBoolean> = obj.get(cx, "headerRow")?;
        let header_row = header_row.value(cx);

        let style: Option<Handle<JsString>> = obj.get_opt(cx, "style")?;
        let style = match style {
            Some(style) => Some(string_to_table_style(cx, style)?),
            None => None,
        };

        let total_row: Handle<JsBoolean> = obj.get(cx, "totalRow")?;
        let total_row = total_row.value(cx);

        Ok(Self {
            name,
            auto_filter,
            banded_columns,
            banded_rows,
            columns: inner_columns,
            first_column_highlighted,
            last_column_highlighted,
            header_row,
            style,
            total_row,
        })
    }
}

impl Into<Table> for NodeXlsxTable {
    fn into(self) -> Table {
        let mut table = Table::new();
        if let Some(name) = self.name {
            table = table.set_name(name);
        }
        table = table.set_autofilter(self.auto_filter);
        table = table.set_banded_columns(self.banded_columns);
        table = table.set_banded_rows(self.banded_rows);
        table = table.set_columns(&self.columns);
        table = table.set_first_column(self.first_column_highlighted);
        table = table.set_last_column(self.last_column_highlighted);
        table = table.set_header_row(self.header_row);
        if let Some(style) = self.style {
            table = table.set_style(style);
        }
        table = table.set_total_row(self.total_row);
        table
    }
}
